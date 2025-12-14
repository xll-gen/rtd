#include <iostream>
#include <vector>
#include <string>
#include <thread>
#include <chrono>

// Include the implementation directly to test logic without COM overhead
#include "../examples/simple/server_impl.h"
#include <rtd/module.h> // Ensure we can access GlobalModule

// Mock IRTDUpdateEvent for ServerStart
struct MockUpdateEvent : public rtd::IRTDUpdateEvent {
    long m_refCount = 1;
    HRESULT __stdcall UpdateNotify() override { return S_OK; }
    long __stdcall get_HeartbeatInterval() override { return 1000; }
    HRESULT __stdcall put_HeartbeatInterval(long value) override { return S_OK; }
    HRESULT __stdcall Disconnect() override { return S_OK; }

    // IUnknown
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppv) override {
        *ppv = nullptr; return E_NOINTERFACE;
    }
    ULONG __stdcall AddRef() override { return ++m_refCount; }
    ULONG __stdcall Release() override { return --m_refCount; }
};

void Assert(bool condition, const std::string& message) {
    if (!condition) {
        std::cerr << "FAILED: " << message << std::endl;
        exit(1);
    } else {
        std::cout << "PASSED: " << message << std::endl;
    }
}

int main() {
    std::cout << "Running Unit Tests..." << std::endl;

    long initialCount = rtd::GlobalModule::GetLockCount();
    Assert(initialCount == 0, "Initial Global Lock Count should be 0");

    {
        MyRtdServer server;
        long countAfterCreate = rtd::GlobalModule::GetLockCount();
        Assert(countAfterCreate == initialCount + 1, "Global Lock Count should increment after Server creation");

        // Test 0: QueryInterface Validation
        {
            void* ppv = nullptr;
            HRESULT hr = server.QueryInterface(rtd::IID_IRtdServer, &ppv);
            Assert(hr == S_OK, "QueryInterface(IID_IRtdServer) should return S_OK");
            Assert(ppv != nullptr, "ppv should not be null");
            if (ppv) static_cast<IUnknown*>(ppv)->Release();

            hr = server.QueryInterface(IID_IDispatch, &ppv);
            Assert(hr == S_OK, "QueryInterface(IID_IDispatch) should return S_OK");
            if (ppv) static_cast<IUnknown*>(ppv)->Release();

            GUID IID_Random = { 0x12345678, 0x1234, 0x1234, { 0x12, 0x34, 0x12, 0x34, 0x12, 0x34, 0x12, 0x34 } };
            hr = server.QueryInterface(IID_Random, &ppv);
            Assert(hr == E_NOINTERFACE, "QueryInterface(IID_Random) should return E_NOINTERFACE");
        }

        // Test 1: ConnectData
        long topicID = 1;
        bool getNewValues = false;
        VARIANT result;
        VariantInit(&result);

        HRESULT hr = server.ConnectData(topicID, nullptr, &getNewValues, &result);

        Assert(hr == S_OK, "ConnectData should return S_OK");
        Assert(result.vt == VT_ERROR, "Result type should be VT_ERROR");

        if (result.vt == VT_ERROR) {
            bool match = (result.scode == 2043);
            Assert(match, "Result should be Error 2043 (xlErrGettingData)");
        }

        // Test 2: RefreshData (Initial)
        long topicCount = -1;
        SAFEARRAY* outputArray = nullptr;

        hr = server.RefreshData(&topicCount, &outputArray);

        Assert(hr == S_OK, "RefreshData should return S_OK");
        Assert(topicCount == 0, "TopicCount should be 0");
    }

    long finalCount = rtd::GlobalModule::GetLockCount();
    Assert(finalCount == initialCount, "Global Lock Count should return to initial value after Server destruction");

    // Test 3: Multiple Topics (Verify Fix for Crash)
    std::cout << "Test 3: Multiple Topics (>2)..." << std::endl;

    {
        MyRtdServer server;
        MockUpdateEvent* pMockCallback = new MockUpdateEvent();
        long res = 0;
        server.ServerStart(pMockCallback, &res);

        bool getNewValues = false;
        VARIANT result;
        VariantInit(&result);

        // Connect 3 topics
        for (long id = 10; id < 13; ++id) {
             server.ConnectData(id, nullptr, &getNewValues, &result);
        }

        std::cout << "Waiting for topics to become ready..." << std::endl;
        std::this_thread::sleep_for(std::chrono::milliseconds(2500));

        long topicCount = -1;
        SAFEARRAY* outputArray = nullptr;
        HRESULT hr = server.RefreshData(&topicCount, &outputArray);
        Assert(hr == S_OK, "RefreshData should return S_OK with multiple topics");
        Assert(topicCount == 3, "TopicCount should be 3");

        if (outputArray) {
            long ubound1, ubound2;
            SafeArrayGetUBound(outputArray, 1, &ubound1);
            SafeArrayGetUBound(outputArray, 2, &ubound2);
            SafeArrayDestroy(outputArray);
        }

        server.ServerTerminate();
        delete pMockCallback;
    }

    std::cout << "All Unit Tests Passed." << std::endl;
    return 0;
}
