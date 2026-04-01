#include "ClassFactory.h"

#include "PolicyModule.h"

#include <new>

volatile LONG g_serverLocks = 0;

PolicyClassFactory::PolicyClassFactory()
    : m_refCount(1) {}

HRESULT PolicyClassFactory::QueryInterface(REFIID riid, void** ppvObject) {
    if (ppvObject == nullptr) {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == IID_IUnknown || riid == IID_IClassFactory) {
        *ppvObject = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG PolicyClassFactory::AddRef() {
    return InterlockedIncrement(&m_refCount);
}

ULONG PolicyClassFactory::Release() {
    const LONG count = InterlockedDecrement(&m_refCount);
    if (count == 0) {
        delete this;
    }

    return static_cast<ULONG>(count);
}

HRESULT PolicyClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) {
    if (ppvObject == nullptr) {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (pUnkOuter != nullptr) {
        return CLASS_E_NOAGGREGATION;
    }

    MidweekExpiryPolicy* module = new (std::nothrow) MidweekExpiryPolicy();
    if (module == nullptr) {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = module->QueryInterface(riid, ppvObject);
    module->Release();
    return hr;
}

HRESULT PolicyClassFactory::LockServer(BOOL fLock) {
    if (fLock) {
        InterlockedIncrement(&g_serverLocks);
    } else {
        InterlockedDecrement(&g_serverLocks);
    }

    return S_OK;
}
