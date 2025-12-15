#ifndef MY_HYBRID_SERVER_IMPL_H
#define MY_HYBRID_SERVER_IMPL_H

#include <rtd/rtd.h>
#include <map>
#include <vector>
#include <thread>
#include <mutex>
#include <atomic>
#include <chrono>

// 1. Define Identity
// GUID: {AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}
static const GUID CLSID_MyRtdServer =
{ 0xAAAAAAAA, 0xBBBB, 0xCCCC, { 0xDD, 0xDD, 0xEE, 0xEE, 0xEE, 0xEE, 0xEE, 0xEE } };

const wchar_t* g_szProgID = L"My.Hybrid.Server";
const wchar_t* g_szFriendlyName = L"MinGW Hybrid RTD Server";

// 2. Implement Server Logic
class MyRtdServer : public rtd::RtdServerBase {
private:
    std::map<long, std::chrono::steady_clock::time_point> m_pendingTopics;
    std::vector<long> m_readyTopics;
    std::mutex m_mutex;
    std::thread m_workerThread;
    std::atomic<bool> m_running;

    void Worker() {
        while (m_running) {
            std::this_thread::sleep_for(std::chrono::milliseconds(500));

            bool notify = false;
            {
                std::lock_guard<std::mutex> lock(m_mutex);
                auto now = std::chrono::steady_clock::now();

                for (auto it = m_pendingTopics.begin(); it != m_pendingTopics.end(); ) {
                    auto elapsed = std::chrono::duration_cast<std::chrono::seconds>(now - it->second).count();
                    if (elapsed >= 2) {
                        m_readyTopics.push_back(it->first);
                        it = m_pendingTopics.erase(it); // Remove from pending, now ready
                        notify = true;
                    } else {
                        ++it;
                    }
                }
            }

            // Important: Call UpdateNotify outside the lock if possible,
            // though standard COM usually marshals this.
            // Check m_callback validity under lock? m_callback is protected by base class rules usually.
            // RtdServerBase implementation: m_callback is set in ServerStart.
            if (notify && m_callback) {
                m_callback->UpdateNotify();
            }
        }
    }

public:
    MyRtdServer() : m_running(false) {}

    virtual ~MyRtdServer() {
        ServerTerminate();
    }

    HRESULT __stdcall ServerStart(rtd::IRTDUpdateEvent* Callback, long* pfRes) override {
        HRESULT hr = RtdServerBase::ServerStart(Callback, pfRes);
        if (FAILED(hr)) return hr;

        m_running = true;
        m_workerThread = std::thread(&MyRtdServer::Worker, this);

        *pfRes = 1;
        return S_OK;
    }

    HRESULT __stdcall ServerTerminate() override {
        m_running = false;
        if (m_workerThread.joinable()) {
            m_workerThread.join();
        }
        return RtdServerBase::ServerTerminate();
    }

    HRESULT __stdcall ConnectData(long TopicID, SAFEARRAY** Strings, bool* GetNewValues, VARIANT* pvarOut) override {
        {
            std::lock_guard<std::mutex> lock(m_mutex);
            m_pendingTopics[TopicID] = std::chrono::steady_clock::now();
        }

        VariantInit(pvarOut);
        pvarOut->vt = VT_ERROR;
        pvarOut->scode = 2043; // xlErrGettingData
        return S_OK;
    }

    HRESULT __stdcall RefreshData(long* TopicCount, SAFEARRAY** parrayOut) override {
        std::lock_guard<std::mutex> lock(m_mutex);

        *TopicCount = static_cast<long>(m_readyTopics.size());

        if (*TopicCount == 0) {
            // Nothing to update
            return S_OK;
        }

        // Create 2D SafeArray: 2 Rows x N Columns
        // bounds[0] is the Right-Most dimension (Columns/Topics)
        // bounds[1] is the Left-Most dimension (Rows)
        SAFEARRAYBOUND bounds[2];
        bounds[0].cElements = *TopicCount;
        bounds[0].lLbound = 0;
        bounds[1].cElements = 2;
        bounds[1].lLbound = 0;

        *parrayOut = SafeArrayCreate(VT_VARIANT, 2, bounds);
        if (!*parrayOut) return E_OUTOFMEMORY;

        long indices[2];

        for (long i = 0; i < *TopicCount; ++i) {
            long topicID = m_readyTopics[i];

            // Row 0: TopicID
            // indices[0] is Rightmost dimension (0..N-1)
            // indices[1] is Leftmost dimension (0..1)
            indices[0] = i; indices[1] = 0;
            VARIANT vID; VariantInit(&vID);
            vID.vt = VT_I4; vID.lVal = topicID;
            SafeArrayPutElement(*parrayOut, indices, &vID);

            // Row 1: Value
            indices[0] = i; indices[1] = 1;
            VARIANT vVal; VariantInit(&vVal);
            vVal.vt = VT_BSTR;
            vVal.bstrVal = SysAllocString(L"Hello World!");
            SafeArrayPutElement(*parrayOut, indices, &vVal);
            SysFreeString(vVal.bstrVal);
        }

        m_readyTopics.clear();
        return S_OK;
    }

    HRESULT __stdcall DisconnectData(long TopicID) override {
        std::lock_guard<std::mutex> lock(m_mutex);
        m_pendingTopics.erase(TopicID);
        for (auto it = m_readyTopics.begin(); it != m_readyTopics.end(); ++it) {
            if (*it == TopicID) {
                m_readyTopics.erase(it);
                break;
            }
        }
        return S_OK;
    }
};

#endif // MY_HYBRID_SERVER_IMPL_H
