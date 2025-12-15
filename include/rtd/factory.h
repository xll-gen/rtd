#ifndef RTD_FACTORY_H
#define RTD_FACTORY_H

#include <windows.h>
#include <ole2.h>
#include "module.h"

namespace rtd {

    template <typename ServerClass>
    class ClassFactory : public IClassFactory {
        long m_refCount;
    public:
        ClassFactory() : m_refCount(1) {
            GlobalModule::Lock();
        }

        virtual ~ClassFactory() {
            GlobalModule::Unlock();
        }

        HRESULT __stdcall QueryInterface(REFIID riid, void** ppv) override {
            if (riid == IID_IUnknown || riid == IID_IClassFactory) {
                *ppv = static_cast<IClassFactory*>(this);
                AddRef(); return S_OK;
            }
            *ppv = nullptr; return E_NOINTERFACE;
        }

        ULONG __stdcall AddRef() override {
            return InterlockedIncrement(&m_refCount);
        }

        ULONG __stdcall Release() override {
            ULONG res = InterlockedDecrement(&m_refCount);
            if (res == 0) delete this; return res;
        }

        HRESULT __stdcall CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) override {
            if (pUnkOuter) return CLASS_E_NOAGGREGATION;
            ServerClass* p = new ServerClass();
            HRESULT hr = p->QueryInterface(riid, ppv);
            p->Release();
            return hr;
        }

        HRESULT __stdcall LockServer(BOOL fLock) override {
            if (fLock) {
                GlobalModule::Lock();
            } else {
                GlobalModule::Unlock();
            }
            return S_OK;
        }
    };

} // namespace rtd

#endif // RTD_FACTORY_H
