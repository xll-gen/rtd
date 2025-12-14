#ifndef RTD_REGISTRY_H
#define RTD_REGISTRY_H

#include <windows.h>
#include <string>
#include <vector>

namespace rtd {

    /**
     * @brief Helper to delete registry keys in HKCU (Current User).
     */
    inline long DeleteKeyUser(const wchar_t* szKey) {
        wchar_t szFullKey[1024];
        wsprintfW(szFullKey, L"Software\\Classes\\%s", szKey);

        // RegDeleteTreeW is available since Vista, which is safe to assume.
        // It recursively deletes the key and all subkeys.
        return RegDeleteTreeW(HKEY_CURRENT_USER, szFullKey);
    }

    /**
     * @brief Helper to set registry keys in HKCU (Current User).
     * This avoids the need for Administrator privileges.
     */
    inline long SetKeyAndValueUser(const wchar_t* szKey, const wchar_t* szSubkey, const wchar_t* szValue) {
        HKEY hKey;
        wchar_t szFullKey[1024];
        if (szSubkey) {
            wsprintfW(szFullKey, L"Software\\Classes\\%s\\%s", szKey, szSubkey);
        } else {
            wsprintfW(szFullKey, L"Software\\Classes\\%s", szKey);
        }

        if (RegCreateKeyExW(HKEY_CURRENT_USER, szFullKey, 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
            return E_FAIL;

        if (szValue) {
            RegSetValueExW(hKey, nullptr, 0, REG_SZ, (const BYTE*)szValue, (wcslen(szValue) + 1) * sizeof(wchar_t));
        }
        RegCloseKey(hKey);
        return S_OK;
    }

    /**
     * @brief Helper to register the COM server.
     */
    inline HRESULT RegisterServer(HMODULE hModule, const GUID& clsid, const wchar_t* progID, const wchar_t* friendlyName) {
        wchar_t szModule[MAX_PATH];
        GetModuleFileNameW(hModule, szModule, ARRAYSIZE(szModule));

        LPOLESTR pszCLSID;
        StringFromCLSID(clsid, &pszCLSID);

        wchar_t szCLSIDString[64];
        wcscpy(szCLSIDString, pszCLSID);
        CoTaskMemFree(pszCLSID);

        wchar_t szCLSIDKey[128];
        wsprintfW(szCLSIDKey, L"CLSID\\%s", szCLSIDString);

        // 1. ProgID -> CLSID
        if (FAILED(SetKeyAndValueUser(progID, nullptr, friendlyName))) return E_FAIL;
        if (FAILED(SetKeyAndValueUser(progID, L"CLSID", szCLSIDString))) return E_FAIL;

        // 2. CLSID -> DLL Path
        if (FAILED(SetKeyAndValueUser(szCLSIDKey, nullptr, friendlyName))) return E_FAIL;
        if (FAILED(SetKeyAndValueUser(szCLSIDKey, L"ProgID", progID))) return E_FAIL;
        if (FAILED(SetKeyAndValueUser(szCLSIDKey, L"InprocServer32", szModule))) return E_FAIL;

        // ThreadingModel = Apartment is crucial for Excel RTD
        wchar_t szInprocKey[256];
        wsprintfW(szInprocKey, L"%s\\InprocServer32", szCLSIDKey);

        // Re-open InprocServer32 key to set ThreadingModel
        HKEY hKey;
        wchar_t szFullKey[1024];
        wsprintfW(szFullKey, L"Software\\Classes\\%s", szInprocKey);
        if (RegOpenKeyExW(HKEY_CURRENT_USER, szFullKey, 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS) {
            const wchar_t* threading = L"Apartment";
            RegSetValueExW(hKey, L"ThreadingModel", 0, REG_SZ, (const BYTE*)threading, (wcslen(threading) + 1) * sizeof(wchar_t));
            RegCloseKey(hKey);
        }

        return S_OK;
    }

    /**
     * @brief Unregister the server (Clean up).
     */
    inline HRESULT UnregisterServer(const GUID& clsid, const wchar_t* progID) {
        LPOLESTR pszCLSID;
        StringFromCLSID(clsid, &pszCLSID);

        wchar_t szCLSIDString[64];
        wcscpy(szCLSIDString, pszCLSID);
        CoTaskMemFree(pszCLSID);

        wchar_t szCLSIDKey[128];
        wsprintfW(szCLSIDKey, L"CLSID\\%s", szCLSIDString);

        // 1. Delete ProgID
        DeleteKeyUser(progID);

        // 2. Delete CLSID
        DeleteKeyUser(szCLSIDKey);

        return S_OK;
    }

} // namespace rtd

#endif // RTD_REGISTRY_H
