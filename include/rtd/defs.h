#ifndef RTD_DEFS_H
#define RTD_DEFS_H

#include <windows.h>
#include <ole2.h>
#include <olectl.h>
#include <string>

namespace rtd {

    // IRTDUpdateEvent Interface
    struct IRTDUpdateEvent : public IUnknown {
        virtual HRESULT __stdcall UpdateNotify() = 0;
        virtual long __stdcall get_HeartbeatInterval() = 0;
        virtual HRESULT __stdcall put_HeartbeatInterval(long value) = 0;
        virtual HRESULT __stdcall Disconnect() = 0;
    };

    // IRtdServer Interface
    struct IRtdServer : public IDispatch {
        virtual HRESULT __stdcall ServerStart(IRTDUpdateEvent* Callback, long* pfRes) = 0;
        virtual HRESULT __stdcall ConnectData(long TopicID, SAFEARRAY** Strings, bool* GetNewValues, VARIANT* pvarOut) = 0;
        virtual HRESULT __stdcall RefreshData(long* TopicCount, SAFEARRAY** parrayOut) = 0;
        virtual HRESULT __stdcall DisconnectData(long TopicID) = 0;
        virtual HRESULT __stdcall Heartbeat(long* pfRes) = 0;
        virtual HRESULT __stdcall ServerTerminate() = 0;
    };

    // Helper to get IID for IRtdServer since __uuidof might not work in all MinGW setups or requires explicit decl.
    // However, usually we use the IID passed to QueryInterface.
    // We will define the IID for IRtdServer if needed, but often we just check against the user provided GUID or IDispatch.

} // namespace rtd

#endif // RTD_DEFS_H
