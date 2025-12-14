#include "server_impl.h"

// 3. Define Entry Points
RTD_DEFINE_DLL_ENTRY(MyRtdServer, CLSID_MyRtdServer, g_szProgID, g_szFriendlyName)

// 4. XLL Entry Point (Dummy for now, to prove hybrid capability)
extern "C" __declspec(dllexport) void WINAPI xlAutoOpen() {
    // XLL initialization code would go here
}
