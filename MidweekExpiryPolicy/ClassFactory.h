#pragma once

#include <windows.h>

class PolicyClassFactory final : public IClassFactory {
public:
    PolicyClassFactory();

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;

    // IClassFactory
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override;
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override;

private:
    ~PolicyClassFactory() = default;
    volatile LONG m_refCount;
};

extern volatile LONG g_serverLocks;
