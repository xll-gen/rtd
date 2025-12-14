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
        cmake -DCMAKE_TOOLCHAIN_FILE=../cmake/mingw-w64-x86_64.cmake ..
        # Or manually specify the compiler if no toolchain file exists:
        cmake -DCMAKE_CXX_COMPILER=x86_64-w64-mingw32-g++ ..
        ```
    *   **On Windows (MinGW):**
        ```bash
        cmake -G "MinGW Makefiles" ..
        ```

3.  Build:
    ```bash
    cmake --build .
    ```
    This will generate `MyHybridServer.xll` (and test executables).

## Registration

The generated `.xll` is a standard COM server. You must register it before Excel can use the RTD function.

**To Register:**
```cmd
regsvr32 MyHybridServer.xll
```
*   This registers the server in `HKEY_CURRENT_USER` (HKCU), so **no Administrator privileges** are required.

**To Unregister:**
```cmd
regsvr32 /u MyHybridServer.xll
```

## Running Tests

### 1. Unit Tests (C++)
The unit tests verify the internal logic of the C++ class without needing Excel.
```bash
# Inside build directory
./unit_test.exe
```

### 2. Integration Tests (C++)
This test attempts to load the COM server via `CoCreateInstance`. This verifies that:
1.  The DLL is valid.
2.  Registration was successful.
```bash
# Inside build directory (After running regsvr32)
./integration_test.exe
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
