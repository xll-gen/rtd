#ifndef RTD_SERVER_H
#define RTD_SERVER_H

#include <mutex>
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
        std::mutex m_mutex;

    public:
        RtdServerBase() : m_refCount(1), m_callback(nullptr) {
            GlobalModule::Lock();
        }
        virtual ~RtdServerBase() {
            // No need to lock m_mutex; object is being destroyed (RefCount=0)
            if (m_callback) m_callback->Release();
            GlobalModule::Unlock();
        }

        // --- IUnknown ---
        HRESULT __stdcall QueryInterface(REFIID riid, void** ppv) override {
            if (!ppv) return E_POINTER;
            *ppv = nullptr;

            // Check for IUnknown, IDispatch, or our specific IRtdServer IID
            // Note: IID_IRtdServer is defined in defs.h and matches {EC0E6191-DB51-11D3-8F3E-00C04F3651B8}
            if (IsEqualGUID(riid, IID_IUnknown) || IsEqualGUID(riid, IID_IDispatch) || IsEqualGUID(riid, IID_IRtdServer)) {
                *ppv = static_cast<IRtdServer*>(this);
                AddRef();
                return S_OK;
            }
            return E_NOINTERFACE;
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
        HRESULT __stdcall GetTypeInfoCount(UINT* pctinfo) override {
            if (!pctinfo) return E_POINTER;
            *pctinfo = 0;
            return S_OK;
        }
        HRESULT __stdcall GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
        HRESULT __stdcall GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
        HRESULT __stdcall Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) override { return E_NOTIMPL; }

        // --- IRtdServer Default Implementations ---
        HRESULT __stdcall ServerStart(IRTDUpdateEvent* Callback, long* pfRes) override {
            if (!pfRes) return E_POINTER;
            std::lock_guard<std::mutex> lock(m_mutex);
            if (m_callback) m_callback->Release();
            m_callback = Callback;
            if (m_callback) m_callback->AddRef();
            *pfRes = 1;
            return S_OK;
        }

        HRESULT __stdcall ServerTerminate() override {
            std::lock_guard<std::mutex> lock(m_mutex);
            if (m_callback) {
                m_callback->Release();
                m_callback = nullptr;
            }
            return S_OK;
        }

        /**
         * @brief Thread-safe helper to notify Excel of updates.
         */
        void NotifyUpdate() {
            IRTDUpdateEvent* tempCallback = nullptr;
            {
                std::lock_guard<std::mutex> lock(m_mutex);
                if (m_callback) {
                    tempCallback = m_callback;
                    tempCallback->AddRef();
                }
            }
            if (tempCallback) {
                tempCallback->UpdateNotify();
                tempCallback->Release();
            }
        }

        /**
         * @brief Helper to create the standard 2D SafeArray for RefreshData.
         * The array is [2][topicCount].
         * Row 0: Topic IDs.
         * Row 1: Values.
         */
        static HRESULT CreateRefreshDataArray(long topicCount, SAFEARRAY** ppArray) {
            if (!ppArray) return E_POINTER;
            if (topicCount < 0) return E_INVALIDARG;

            if (topicCount == 0) {
                 *ppArray = nullptr;
                 return S_OK;
            }

            SAFEARRAYBOUND bounds[2];
            // Dimension 1: Rows (0=TopicID, 1=Value) - Left-most (rgsabound[1])
            bounds[1].cElements = 2;
            bounds[1].lLbound = 0;
            // Dimension 2: Columns (Topics) - Right-most (rgsabound[0])
            bounds[0].cElements = topicCount;
            bounds[0].lLbound = 0;

            *ppArray = SafeArrayCreate(VT_VARIANT, 2, bounds);
            if (!*ppArray) return E_OUTOFMEMORY;
            return S_OK;
        }

        HRESULT __stdcall DisconnectData(long TopicID) override { return S_OK; }
        HRESULT __stdcall Heartbeat(long* pfRes) override {
            if (!pfRes) return E_POINTER;
            *pfRes = 1;
            return S_OK;
        }

        // User must implement:
        // ConnectData
        // RefreshData
        virtual HRESULT __stdcall ConnectData(long TopicID, SAFEARRAY** Strings, VARIANT_BOOL* GetNewValues, VARIANT* pvarOut) override = 0;
    };

} // namespace rtd

#endif // RTD_SERVER_H
