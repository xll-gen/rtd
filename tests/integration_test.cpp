#include <windows.h>
#include <iostream>
#include <rtd/rtd.h>

// GUID from server_impl.h
static const GUID CLSID_MyRtdServer =
{ 0xAAAAAAAA, 0xBBBB, 0xCCCC, { 0xDD, 0xDD, 0xEE, 0xEE, 0xEE, 0xEE, 0xEE, 0xEE } };

void Assert(bool condition, const std::string& message) {
    if (!condition) {
        std::cerr << "FAILED: " << message << std::endl;
        exit(1);
    } else {
        std::cout << "PASSED: " << message << std::endl;
    }
}

int main() {
    std::cout << "Running Integration Test (COM Loading)..." << std::endl;

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        std::cerr << "CoInitialize failed" << std::endl;
        return 1;
    }

    rtd::IRtdServer* pServer = nullptr;
    static const GUID IID_IRtdServer = { 0xEC0E6191, 0xDB51, 0x11D3, { 0x8F, 0x3E, 0x00, 0xC0, 0x4F, 0x36, 0x51, 0xB8 } };
    hr = CoCreateInstance(CLSID_MyRtdServer, NULL, CLSCTX_INPROC_SERVER, IID_IRtdServer, (void**)&pServer);

    if (FAILED(hr)) {
        std::cerr << "CoCreateInstance failed. HRESULT: " << std::hex << hr << std::endl;
        std::cerr << "Ensure the DLL is registered using 'regsvr32 MyHybridServer.xll'" << std::endl;
        CoUninitialize();
        return 1;
    }

    Assert(pServer != nullptr, "Server instance created successfully");

    // Basic Interface Test
    // Call ServerStart
    long res = 0;
    hr = pServer->ServerStart(nullptr, &res); // Callback can be null for this test
    Assert(hr == S_OK || hr == 0, "ServerStart returned success");

    // Clean up
    pServer->Release();
    CoUninitialize();

    std::cout << "Integration Test Passed." << std::endl;
    return 0;
}
