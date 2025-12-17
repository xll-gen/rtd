# Simple Hybrid XLL/RTD Server Example

This example demonstrates how to implement a Hybrid XLL/RTD server using the `xll-gen` library.
It acts as both an Excel Add-in (XLL) and a COM RTD Server.

## Features
*   **XLL Interface:** Exports `MyHello` function.
*   **RTD Server:** Implements `IRtdServer` with a background thread updating topics.
*   **Automatic Registration:** `xlAutoOpen` registers the COM server in HKCU.

## Building

### Prerequisites
*   CMake
*   MinGW-w64 toolchain (for Windows cross-compilation)
*   Wine (for testing on Linux)

### Steps
1.  Navigate to the project root.
2.  Create a build directory:
    ```bash
    mkdir build && cd build
    ```
3.  Configure CMake (assuming standard MinGW cross-compiler names):
    ```bash
    cmake -DCMAKE_CXX_COMPILER=x86_64-w64-mingw32-g++ -DCMAKE_RC_COMPILER=x86_64-w64-mingw32-windres -DMINGW=ON ..
    ```
4.  Build:
    ```bash
    cmake --build .
    ```

## Testing

### Unit Tests
Runs the server logic in isolation (mocking COM).
```bash
wine unit_test.exe
```

### Integration Tests
Tests the full COM instantiation and interface calls.
1.  **Register the Server:**
    ```bash
    wine regsvr32 libMyHybridServer.xll
    ```
2.  **Run the Test:**
    ```bash
    wine integration_test.exe
    ```
