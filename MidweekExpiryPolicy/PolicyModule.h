#pragma once

#include <windows.h>
#include <certsrv.h>
#include <certpol.h>
#include <string>

// {B4AB7B5E-BD14-4F36-93B9-F4A5F1B6EA2B}
static const CLSID CLSID_MidweekExpiryPolicy = {
    0xb4ab7b5e,
    0xbd14,
    0x4f36,
    {0x93, 0xb9, 0xf4, 0xa5, 0xf1, 0xb6, 0xea, 0x2b}};

class MidweekExpiryPolicy final : public ICertPolicy2 {
public:
    MidweekExpiryPolicy();

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;

    // ICertPolicy
    HRESULT STDMETHODCALLTYPE Initialize(const BSTR strConfig) override;
    HRESULT STDMETHODCALLTYPE VerifyRequest(
        const BSTR strConfig,
        LONG Context,
        LONG bNewRequest,
        LONG Flags,
        LONG* pDisposition) override;
    HRESULT STDMETHODCALLTYPE GetDescription(BSTR* pstrDescription) override;
    HRESULT STDMETHODCALLTYPE ShutDown() override;

    // ICertPolicy2
    HRESULT STDMETHODCALLTYPE GetManageModule(
        ICertManageModule** ppManageModule) override;

private:
    ~MidweekExpiryPolicy();

    static bool IsMondayOrTuesday(const SYSTEMTIME& st);
    static FILETIME AdjustNotAfterToMidweek(const FILETIME& notAfter);
    static std::wstring FileTimeToIso8601(const FILETIME& ft);

    HRESULT SetCertificateValidity(ICertServerPolicy* pServer, const FILETIME& adjustedNotAfter) const;

    volatile LONG m_refCount;
    std::wstring m_config;
};
