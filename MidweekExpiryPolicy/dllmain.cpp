#include <windows.h>
#include <combaseapi.h>
#include <new>
#include <string>

#include "ClassFactory.h"
#include "PolicyModule.h"

HINSTANCE g_hInstance = nullptr;
volatile LONG g_moduleRefs = 0;

namespace {
constexpr wchar_t kFriendlyName[] = L"MidweekExpiryPolicy";
constexpr wchar_t kProgId[] = L"MidweekExpiryPolicy.Module";

std::wstring GuidToString(REFGUID guid) {
    wchar_t buffer[64] = {};
    StringFromGUID2(guid, buffer, _countof(buffer));
    return buffer;
}

HRESULT SetRegistryString(HKEY root, const std::wstring& subKey, const wchar_t* valueName, const std::wstring& value) {
    HKEY key = nullptr;
    LONG result = RegCreateKeyExW(root, subKey.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &key, nullptr);
    if (result != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(result);
    }

    result = RegSetValueExW(
        key,
        valueName,
        0,
        REG_SZ,
        reinterpret_cast<const BYTE*>(value.c_str()),
        static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t)));

    RegCloseKey(key);
    return HRESULT_FROM_WIN32(result);
}

HRESULT RegisterServer() {
    wchar_t modulePath[MAX_PATH] = {};
    if (GetModuleFileNameW(g_hInstance, modulePath, _countof(modulePath)) == 0) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const std::wstring clsid = GuidToString(CLSID_MidweekExpiryPolicy);
    const std::wstring clsidKey = L"CLSID\\" + clsid;

    HRESULT hr = SetRegistryString(HKEY_CLASSES_ROOT, clsidKey, nullptr, kFriendlyName);
    if (FAILED(hr)) {
        return hr;
    }

    hr = SetRegistryString(HKEY_CLASSES_ROOT, clsidKey + L"\\InprocServer32", nullptr, modulePath);
    if (FAILED(hr)) {
        return hr;
    }

    hr = SetRegistryString(HKEY_CLASSES_ROOT, clsidKey + L"\\InprocServer32", L"ThreadingModel", L"Both");
    if (FAILED(hr)) {
        return hr;
    }

    hr = SetRegistryString(HKEY_CLASSES_ROOT, clsidKey + L"\\ProgID", nullptr, kProgId);
    if (FAILED(hr)) {
        return hr;
    }

    hr = SetRegistryString(HKEY_CLASSES_ROOT, kProgId, nullptr, kFriendlyName);
    if (FAILED(hr)) {
        return hr;
    }

    hr = SetRegistryString(HKEY_CLASSES_ROOT, kProgId + std::wstring(L"\\CLSID"), nullptr, clsid);
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

HRESULT UnregisterServer() {
    const std::wstring clsid = GuidToString(CLSID_MidweekExpiryPolicy);
    const std::wstring clsidKey = L"CLSID\\" + clsid;

    RegDeleteTreeW(HKEY_CLASSES_ROOT, clsidKey.c_str());
    RegDeleteTreeW(HKEY_CLASSES_ROOT, kProgId);
    return S_OK;
}
}  // namespace

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ulReasonForCall, LPVOID /*lpReserved*/) {
    if (ulReasonForCall == DLL_PROCESS_ATTACH) {
        g_hInstance = hModule;
        DisableThreadLibraryCalls(hModule);
    }

    return TRUE;
}

STDAPI DllCanUnloadNow() {
    return (g_moduleRefs == 0 && g_serverLocks == 0) ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    if (ppv == nullptr) {
        return E_POINTER;
    }

    *ppv = nullptr;

    if (rclsid != CLSID_MidweekExpiryPolicy) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    PolicyClassFactory* factory = new (std::nothrow) PolicyClassFactory();
    if (factory == nullptr) {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = factory->QueryInterface(riid, ppv);
    factory->Release();
    return hr;
}

STDAPI DllRegisterServer() {
    return RegisterServer();
}

STDAPI DllUnregisterServer() {
    return UnregisterServer();
}
