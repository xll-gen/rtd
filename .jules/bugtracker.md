# Bug Tracker

| ID | Issue | Severity | User Opinion | Status |
|----|-------|----------|--------------|--------|
| BUG-001 | **Memory Leak in ServerStart**: `RtdServerBase::ServerStart` overwrites `m_callback` without calling `Release()` on the existing pointer. This causes a reference leak if the method is called multiple times. | Medium | Fix Approved | Fixed |
| SEC-001 | **Potential Registry Wipe**: `DeleteKeyUser` (registry.h) does not validate if `szKey` is empty. If called with an empty string (e.g., empty ProgID), it targets `Software\Classes`, recursively deleting all user COM registrations. | Critical | Fix Approved | Fixed |
| SEC-002 | **Buffer Safety**: `RegisterServer` uses a fixed `MAX_PATH` buffer for `GetModuleFileNameW`. May fail on systems with long paths enabled if the library is deeply nested. Also lacks return value checking. | Low | Fix Approved | Fixed |
| BUG-002 | **Missing Null Check**: `ServerStart` and `Heartbeat` do not check if the output parameter `pfRes` is null before writing to it, potentially causing a crash (E_POINTER). | Low | Fix Approved | Fixed |
| BUG-003 | **Interface Signature Mismatch**: `IRTDUpdateEvent::get_HeartbeatInterval` in `defs.h` returns `long` directly, whereas the standard COM binary interface requires returning `HRESULT` and passing the value via an out pointer (`long*`). This causes a binary incompatibility with Excel, likely leading to crashes or stack corruption. | Critical | Fix Approved | Fixed |
| SEC-003 | **Excessive Registry Privileges**: `SetKeyAndValueUser` in `registry.h` requests `KEY_ALL_ACCESS` permission. While in HKCU this is usually allowed, best practice (Least Privilege) suggests requesting only `KEY_SET_VALUE` or `KEY_WRITE` to minimize potential damage or permission errors. | Low | Pending | Open |
