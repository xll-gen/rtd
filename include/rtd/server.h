#ifndef RTD_SERVER_H
#define RTD_SERVER_H

#include "defs.h"
#include "module.h"

namespace rtd {

    /**
     * @brief Base class for implementing an RTD Server.
     * Handles IUnknown and IDispatch (Stub) details.
     */
    class RtdServerBase : public IRtdServer {
    private:
        long m_refCount;

    protected:
        IRTDUpdateEvent* m_callback;

    public:
        RtdServerBase() : m_refCount(1), m_callback(nullptr) {
            GlobalModule::Lock();
        }
        virtual ~RtdServerBase() {
            if (m_callback) m_callback->Release();
            GlobalModule::Unlock();
        }

        // --- IUnknown ---
        HRESULT __stdcall QueryInterface(REFIID riid, void** ppv) override {
            // Check for IUnknown, IDispatch, or our specific IRtdServer IID
            if (riid == IID_IUnknown || riid == IID_IDispatch || IsEqualGUID(riid, IID_IRtdServer)) {
                *ppv = static_cast<IRtdServer*>(this);
            }
            else {
                *ppv = nullptr;
                return E_NOINTERFACE;
            }
            AddRef();
            return S_OK;
        }

        ULONG __stdcall AddRef() override {
            return InterlockedIncrement(&m_refCount);
        }

        ULONG __stdcall Release() override {
            ULONG res = InterlockedDecrement(&m_refCount);
            if (res == 0) delete this;
            return res;
        }

        // --- IDispatch (Stub) ---
        HRESULT __stdcall GetTypeInfoCount(UINT*) override { return E_NOTIMPL; }
        HRESULT __stdcall GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
        HRESULT __stdcall GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
        HRESULT __stdcall Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) override { return E_NOTIMPL; }

        // --- IRtdServer Default Implementations ---
        HRESULT __stdcall ServerStart(IRTDUpdateEvent* Callback, long* pfRes) override {
            m_callback = Callback;
            if (m_callback) m_callback->AddRef();
            *pfRes = 1;
            return S_OK;
        }

        HRESULT __stdcall ServerTerminate() override {
            if (m_callback) {
                m_callback->Release();
                m_callback = nullptr;
            }
            return S_OK;
        }

        HRESULT __stdcall DisconnectData(long TopicID) override { return S_OK; }
        HRESULT __stdcall Heartbeat(long* pfRes) override { *pfRes = 1; return S_OK; }

        // User must implement:
        // ConnectData
        // RefreshData
    };

} // namespace rtd

#endif // RTD_SERVER_H
