# Bug Tracker

| ID | Issue | Severity | User Opinion | Status |
|----|-------|----------|--------------|--------|
| BUG-001 | **Memory Leak in ServerStart**: `RtdServerBase::ServerStart` overwrites `m_callback` without calling `Release()` on the existing pointer. This causes a reference leak if the method is called multiple times. | Medium | Fix Approved | Fixed |
| SEC-001 | **Potential Registry Wipe**: `DeleteKeyUser` (registry.h) does not validate if `szKey` is empty. If called with an empty string (e.g., empty ProgID), it targets `Software\Classes`, recursively deleting all user COM registrations. | Critical | Fix Approved | Fixed |
| SEC-002 | **Buffer Safety**: `RegisterServer` uses a fixed `MAX_PATH` buffer for `GetModuleFileNameW`. May fail on systems with long paths enabled if the library is deeply nested. Also lacks return value checking. | Low | Fix Approved | Fixed |
| BUG-002 | **Missing Null Check**: `ServerStart` and `Heartbeat` do not check if the output parameter `pfRes` is null before writing to it, potentially causing a crash (E_POINTER). | Low | Fix Approved | Fixed |
