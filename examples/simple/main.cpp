#include <rtd/rtd.h>

// 1. Define Identity
// GUID: {AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}
static const GUID CLSID_MyRtdServer =
{ 0xAAAAAAAA, 0xBBBB, 0xCCCC, { 0xDD, 0xDD, 0xEE, 0xEE, 0xEE, 0xEE, 0xEE, 0xEE } };

const wchar_t* g_szProgID = L"My.Hybrid.Server";
const wchar_t* g_szFriendlyName = L"MinGW Hybrid RTD Server";

// 2. Implement Server Logic
class MyRtdServer : public rtd::RtdServerBase {
public:
    HRESULT __stdcall ConnectData(long TopicID, SAFEARRAY** Strings, bool* GetNewValues, VARIANT* pvarOut) override {
        // Simple Echo/Test Logic
        VariantInit(pvarOut);
        pvarOut->vt = VT_BSTR;
        pvarOut->bstrVal = SysAllocString(L"Connecting...");
        return S_OK;
    }

    HRESULT __stdcall RefreshData(long* TopicCount, SAFEARRAY** parrayOut) override {
        *TopicCount = 0;
        return S_OK;
    }
};

// 3. Define Entry Points
RTD_DEFINE_DLL_ENTRY(MyRtdServer, CLSID_MyRtdServer, g_szProgID, g_szFriendlyName)

// 4. XLL Entry Point (Dummy for now, to prove hybrid capability)
extern "C" __declspec(dllexport) void WINAPI xlAutoOpen() {
    // XLL initialization code would go here
}
