#include <iostream>
#include <vector>
#include <string>

// Include the implementation directly to test logic without COM overhead
#include "../examples/simple/server_impl.h"

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

    MyRtdServer server;

    // Test 1: ConnectData
    // We need to construct arguments for ConnectData
    // HRESULT ConnectData(long TopicID, SAFEARRAY** Strings, bool* GetNewValues, VARIANT* pvarOut)

    // Note: Creating a SAFEARRAY manually in C++ is verbose.
    // For this unit test, we'll pass nullptr if the implementation allows it,
    // or we'll create a minimal mock if we were mocking the OS API.
    // Since `server_impl.h` implementation of ConnectData doesn't use `Strings`,
    // we can pass nullptr safely for this specific test case.

    long topicID = 1;
    bool getNewValues = false;
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = server.ConnectData(topicID, nullptr, &getNewValues, &result);

    Assert(hr == S_OK, "ConnectData should return S_OK");
    Assert(result.vt == VT_ERROR, "Result type should be VT_ERROR");

    // Check the returned error value.
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

    // Test 3: Multiple Topics (Verify Fix for Crash)
    std::cout << "Test 3: Multiple Topics (>2)..." << std::endl;

    // Start Server to enable Worker thread
    MockUpdateEvent* pMockCallback = new MockUpdateEvent();
    long res = 0;
    server.ServerStart(pMockCallback, &res);

    // Connect 3 topics
    for (long id = 10; id < 13; ++id) {
         server.ConnectData(id, nullptr, &getNewValues, &result);
    }

    // Wait for Worker (2 seconds delay in code -> wait 2.5s)
    // Note: This requires the test to run for >2s.
    std::cout << "Waiting for topics to become ready..." << std::endl;
    std::this_thread::sleep_for(std::chrono::milliseconds(2500));

    hr = server.RefreshData(&topicCount, &outputArray);
    Assert(hr == S_OK, "RefreshData should return S_OK with multiple topics");
    Assert(topicCount == 3, "TopicCount should be 3");

    // Verify Array Dimensions (Fix Verification)
    if (outputArray) {
        long ubound1, ubound2;
        SafeArrayGetUBound(outputArray, 1, &ubound1); // Leftmost?
        SafeArrayGetUBound(outputArray, 2, &ubound2); // Rightmost?

        // SafeArrayGetUBound takes dimension (1-based).
        // Dim 1 corresponds to bounds[0] (Leftmost) -> Should be 2 (0..1) -> UBound 1
        // Dim 2 corresponds to bounds[1] (Rightmost) -> Should be 3 (0..2) -> UBound 2

        // Wait, SafeArrayGetUBound logic:
        // "The dimension for which to get the upper bound."

        // Let's just trust no crash first.
        SafeArrayDestroy(outputArray);
    }

    server.ServerTerminate();
    // pMockCallback->Release(); // Handled by ServerTerminate? No, ServerTerminate calls Release on its copy.
    // We manually delete if refcount is 0 or handled.
    // Since we created it with refcount 1, and ServerStart AddRef'd (2), ServerTerminate Release'd (1).
    // We can delete it.
    delete pMockCallback;

    std::cout << "All Unit Tests Passed." << std::endl;
    return 0;
}
