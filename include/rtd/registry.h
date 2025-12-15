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
        std::wstring szFullKey = L"Software\\Classes\\";
        szFullKey += szKey;

        // RegDeleteTreeW is available since Vista, which is safe to assume.
        // It recursively deletes the key and all subkeys.
        return RegDeleteTreeW(HKEY_CURRENT_USER, szFullKey.c_str());
    }

    /**
     * @brief Helper to set registry keys in HKCU (Current User).
     * This avoids the need for Administrator privileges.
     */
    inline long SetKeyAndValueUser(const wchar_t* szKey, const wchar_t* szSubkey, const wchar_t* szValue) {
        HKEY hKey;
        std::wstring szFullKey = L"Software\\Classes\\";
        szFullKey += szKey;
        if (szSubkey) {
            szFullKey += L"\\";
            szFullKey += szSubkey;
        }

        if (RegCreateKeyExW(HKEY_CURRENT_USER, szFullKey.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
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

        std::wstring szCLSIDKey = L"CLSID\\";
        szCLSIDKey += szCLSIDString;

        // 1. ProgID -> CLSID
        if (FAILED(SetKeyAndValueUser(progID, nullptr, friendlyName))) return E_FAIL;
        if (FAILED(SetKeyAndValueUser(progID, L"CLSID", szCLSIDString))) return E_FAIL;

        // 2. CLSID -> DLL Path
        if (FAILED(SetKeyAndValueUser(szCLSIDKey.c_str(), nullptr, friendlyName))) return E_FAIL;
        if (FAILED(SetKeyAndValueUser(szCLSIDKey.c_str(), L"ProgID", progID))) return E_FAIL;
        if (FAILED(SetKeyAndValueUser(szCLSIDKey.c_str(), L"InprocServer32", szModule))) return E_FAIL;

        // ThreadingModel = Apartment is crucial for Excel RTD
        std::wstring szInprocKey = szCLSIDKey + L"\\InprocServer32";

        // Re-open InprocServer32 key to set ThreadingModel
        HKEY hKey;
        std::wstring szFullKey = L"Software\\Classes\\";
        szFullKey += szInprocKey;

        if (RegOpenKeyExW(HKEY_CURRENT_USER, szFullKey.c_str(), 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS) {
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

        std::wstring szCLSIDKey = L"CLSID\\";
        szCLSIDKey += szCLSIDString;

        // 1. Delete ProgID
        DeleteKeyUser(progID);

        // 2. Delete CLSID
        DeleteKeyUser(szCLSIDKey.c_str());

        return S_OK;
    }

} // namespace rtd

#endif // RTD_REGISTRY_H
