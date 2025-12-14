#include <iostream>
#include <vector>
#include <string>

// Include the implementation directly to test logic without COM overhead
#include "../examples/simple/server_impl.h"
#include <rtd/module.h> // Ensure we can access GlobalModule

// Mock for SAFEARRAY/VARIANT since we might not have full OLE support in Linux test env
// But since this is a unit test that will likely be compiled on Windows/MinGW,
// we assume headers are available.
// If compiling on Linux with g++, these will fail without a stub.
// However, the instructions say "verification on MinGW on Windows".
// I will write standard C++ test code assuming Windows headers are present.

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

    // Test 2: RefreshData
    // HRESULT RefreshData(long* TopicCount, SAFEARRAY** parrayOut)
    long topicCount = -1;
    SAFEARRAY* outputArray = nullptr; // Implementation sets this, but for *TopicCount=0 it might not touch it.

    hr = server.RefreshData(&topicCount, &outputArray);

    Assert(hr == S_OK, "RefreshData should return S_OK");
    Assert(topicCount == 0, "TopicCount should be 0");
    }

    long finalCount = rtd::GlobalModule::GetLockCount();
    Assert(finalCount == initialCount, "Global Lock Count should return to initial value after Server destruction");

    std::cout << "All Unit Tests Passed." << std::endl;
    return 0;
}
