#include "PolicyModule.h"

#include <atlbase.h>
#include <comdef.h>
#include <strsafe.h>

namespace {
constexpr wchar_t kPolicyDescription[] = L"ADCS policy module that enforces certificate expiry on Monday or Tuesday.";
constexpr wchar_t kNotAfterProperty[] = L"ExpirationDate";

void MoveBackDays(SYSTEMTIME& st, int days) {
    FILETIME ft{};
    SystemTimeToFileTime(&st, &ft);

    ULARGE_INTEGER value{};
    value.LowPart = ft.dwLowDateTime;
    value.HighPart = ft.dwHighDateTime;

    constexpr ULONGLONG kTicksPerDay = 24ULL * 60ULL * 60ULL * 10000000ULL;
    value.QuadPart -= static_cast<ULONGLONG>(days) * kTicksPerDay;

    ft.dwLowDateTime = value.LowPart;
    ft.dwHighDateTime = value.HighPart;

    FileTimeToSystemTime(&ft, &st);
}
}  // namespace

extern volatile LONG g_moduleRefs;

MidweekExpiryPolicy::MidweekExpiryPolicy()
    : m_refCount(1) {
    InterlockedIncrement(&g_moduleRefs);
}

MidweekExpiryPolicy::~MidweekExpiryPolicy() {
    InterlockedDecrement(&g_moduleRefs);
}

HRESULT MidweekExpiryPolicy::QueryInterface(REFIID riid, void** ppvObject) {
    if (ppvObject == nullptr) {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == IID_IUnknown || riid == IID_ICertPolicy) {
        *ppvObject = static_cast<ICertPolicy*>(this);
    } else if (riid == IID_ICertPolicy2) {
        *ppvObject = static_cast<ICertPolicy2*>(this);
    } else {
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG MidweekExpiryPolicy::AddRef() {
    return InterlockedIncrement(&m_refCount);
}

ULONG MidweekExpiryPolicy::Release() {
    const LONG count = InterlockedDecrement(&m_refCount);
    if (count == 0) {
        delete this;
    }

    return static_cast<ULONG>(count);
}

HRESULT MidweekExpiryPolicy::Initialize(const BSTR strConfig) {
    m_config = strConfig != nullptr ? strConfig : L"";
    return S_OK;
}

HRESULT MidweekExpiryPolicy::VerifyRequest(
    const BSTR strConfig,
    LONG Context,
    LONG /*bNewRequest*/,
    LONG /*Flags*/,
    LONG* pDisposition) {
    if (pDisposition == nullptr) {
        return E_POINTER;
    }

    *pDisposition = VR_INSTANT_OK;

    CComPtr<ICertServerPolicy> server;
    HRESULT hr = CoCreateInstance(
        __uuidof(CCertServerPolicy),
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&server));
    if (FAILED(hr)) {
        return hr;
    }

    hr = server->SetContext(strConfig, Context);
    if (FAILED(hr)) {
        return hr;
    }

    CComVariant varNotAfter;
    hr = server->GetCertificateProperty(const_cast<BSTR>(kNotAfterProperty), PROPTYPE_DATE, &varNotAfter);
    if (FAILED(hr) || varNotAfter.vt != VT_DATE) {
        return hr;
    }

    SYSTEMTIME originalSt{};
    if (!VariantTimeToSystemTime(varNotAfter.date, &originalSt)) {
        return E_FAIL;
    }

    if (IsMondayOrTuesday(originalSt)) {
        return S_OK;
    }

    FILETIME originalFt{};
    if (!SystemTimeToFileTime(&originalSt, &originalFt)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const FILETIME adjustedFt = AdjustNotAfterToMidweek(originalFt);
    hr = SetCertificateValidity(server, adjustedFt);
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

HRESULT MidweekExpiryPolicy::GetDescription(BSTR* pstrDescription) {
    if (pstrDescription == nullptr) {
        return E_POINTER;
    }

    *pstrDescription = SysAllocString(kPolicyDescription);
    return *pstrDescription == nullptr ? E_OUTOFMEMORY : S_OK;
}

HRESULT MidweekExpiryPolicy::ShutDown() {
    return S_OK;
}

HRESULT MidweekExpiryPolicy::GetManageModule(ICertManageModule** ppManageModule) {
    if (ppManageModule == nullptr) {
        return E_POINTER;
    }

    *ppManageModule = nullptr;
    return E_NOTIMPL;
}

bool MidweekExpiryPolicy::IsMondayOrTuesday(const SYSTEMTIME& st) {
    return st.wDayOfWeek == 1 || st.wDayOfWeek == 2;
}

FILETIME MidweekExpiryPolicy::AdjustNotAfterToMidweek(const FILETIME& notAfter) {
    SYSTEMTIME st{};
    FileTimeToSystemTime(&notAfter, &st);

    // wDayOfWeek: 0=Sunday, 1=Monday, 2=Tuesday ... 6=Saturday
    if (st.wDayOfWeek == 0) {
        MoveBackDays(st, 5);
    } else if (st.wDayOfWeek >= 3 && st.wDayOfWeek <= 6) {
        MoveBackDays(st, st.wDayOfWeek - 2);
    }

    FILETIME adjusted{};
    SystemTimeToFileTime(&st, &adjusted);
    return adjusted;
}

std::wstring MidweekExpiryPolicy::FileTimeToIso8601(const FILETIME& ft) {
    SYSTEMTIME st{};
    FileTimeToSystemTime(&ft, &st);

    wchar_t buffer[32] = {};
    StringCchPrintfW(
        buffer,
        _countof(buffer),
        L"%04u-%02u-%02uT%02u:%02u:%02uZ",
        st.wYear,
        st.wMonth,
        st.wDay,
        st.wHour,
        st.wMinute,
        st.wSecond);

    return buffer;
}

HRESULT MidweekExpiryPolicy::SetCertificateValidity(ICertServerPolicy* pServer, const FILETIME& adjustedNotAfter) const {
    SYSTEMTIME st{};
    if (!FileTimeToSystemTime(&adjustedNotAfter, &st)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DOUBLE variantDate = 0;
    if (!SystemTimeToVariantTime(&st, &variantDate)) {
        return E_FAIL;
    }

    CComVariant value(variantDate);
    return pServer->SetCertificateProperty(const_cast<BSTR>(kNotAfterProperty), PROPTYPE_DATE, &value);
}
