# xll-gen-rtd

A C++ header-only library to create **Hybrid XLL/RTD Servers** for Microsoft Excel using **MinGW-w64**.

This project implements "Raw COM" (no ATL/MFC) to allow a single binary (`.xll`) to function as both an Excel Add-in (XLL) and a Real-Time Data (RTD) Server.

## Requirements

*   **Windows OS** (for execution)
*   **MinGW-w64** (GCC for Windows)
    *   On Linux: `sudo apt install mingw-w64`
    *   On Windows: Install MSYS2 or MinGW-w64 distros.
*   **CMake** (3.10+)

## Building

1.  Create a build directory:
    ```bash
    mkdir build
    cd build
    ```

2.  Configure with CMake (using MinGW toolchain):
    *   **On Linux (Cross-compile):**
        ```bash
        cmake -DCMAKE_CXX_COMPILER=x86_64-w64-mingw32-g++ -DCMAKE_RC_COMPILER=x86_64-w64-mingw32-windres -DMINGW=ON ..
        ```
    *   **On Windows (MinGW):**
        ```bash
        cmake -G "MinGW Makefiles" ..
        ```

3.  Build:
    ```bash
    cmake --build .
    ```
    This will generate `MyHybridServer.xll` (or `libMyHybridServer.xll` with MinGW defaults) and test executables.

## Registration

The generated `.xll` is a standard COM server. You must register it before Excel can use the RTD function.

**To Register:**
```cmd
regsvr32 MyHybridServer.xll
:: Or if built with MinGW defaults:
:: regsvr32 libMyHybridServer.xll
```
*   This registers the server in `HKEY_CURRENT_USER` (HKCU), so **no Administrator privileges** are required.

**To Unregister:**
```cmd
regsvr32 /u MyHybridServer.xll
:: Or:
:: regsvr32 /u libMyHybridServer.xll
```

## Running Tests

### 1. Unit Tests (C++)
The unit tests verify the internal logic of the C++ class without needing Excel.
```bash
# Inside build directory
# On Windows:
./unit_test.exe
# On Linux (with Wine):
wine unit_test.exe
```

### 2. Integration Tests (C++)
This test attempts to load the COM server via `CoCreateInstance`. This verifies that:
1.  The DLL is valid.
2.  Registration was successful.
```bash
# Inside build directory (After running regsvr32)
# On Windows:
./integration_test.exe
# On Linux (with Wine):
wine integration_test.exe
```

### 3. End-to-End Test (PowerShell + Excel)
A PowerShell script is provided to verify the RTD server inside a real Excel instance.

**Steps:**
1.  Open PowerShell.
2.  Navigate to the `tests` directory.
3.  Run the script:
    ```powershell
    .\test_rtd_excel.ps1
    ```

This script will:
1.  Register the DLL.
2.  Open Excel.
3.  Insert `=RTD("My.Hybrid.Server",, "Topic1")`.
4.  Verify the cell updates to "Connecting...".
5.  Cleanup (Close Excel and Unregister).

## Project Structure

*   `include/rtd/`: The core header-only library.
*   `examples/simple/`: A minimal example of a hybrid server.
*   `tests/`: Unit and integration tests.
