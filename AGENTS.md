# AI Agent Instructions for xll-gen

This file contains instructions for AI Agents working on this repository.

## **Codebase Authority**

*   **MinGW Compatibility:** The project must compile and run using GCC (MinGW-w64) on Windows.
*   **Raw COM:** Do not use ATL, MFC, or other MSVC-specific libraries. Implement COM interfaces (`IUnknown`, `IDispatch`, `IRtdServer`) using standard C++.
*   **Hybrid XLL/RTD:** The architecture follows the "xlOil" pattern: a single binary acting as both XLL and COM Server.
*   **Versioning:** Patch version updates (x.y.Z) must maintain API compatibility.

## **COM Compliance & Best Practices**

Follow the [Microsoft COM Documentation](https://learn.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal) rigorously.

1.  **IUnknown Implementation:**
    *   **QueryInterface:**
        *   Must support `IID_IUnknown`.
        *   Must check `ppv` for `nullptr` and return `E_POINTER` if invalid.
        *   Must initialize `*ppv = nullptr` before checking IIDs.
        *   Must use `IsEqualGUID(riid, ...)` for IID comparison (avoid `operator==` for cross-compiler safety).
        *   Must return the **same pointer identity** when queried for `IID_IUnknown`.
    *   **AddRef/Release:**
        *   Use `InterlockedIncrement` / `InterlockedDecrement` for thread safety.

2.  **Class Factory:**
    *   `CreateInstance` must validate parameters (`ppv`, `pUnkOuter`).
    *   Aggregations are generally not supported (`CLASS_E_NOAGGREGATION`).
    *   Use `new` to allocate servers; initial reference count is 1.

3.  **Registration:**
    *   **No Admin:** Register only in `HKEY_CURRENT_USER` (HKCU).
    *   **Safe Strings:** Use `std::wstring` or safe string functions.
    *   **ThreadingModel:** Must be set to `Apartment` for Excel RTD servers.

## **Project Structure**

*   **include/rtd/**: C++ Header-only library. **Do not create .cpp files for the library.**
*   **examples/**: Example usage and usage tests.
*   **CMakeLists.txt**: Build configuration.

## **Coding Standards**

*   **C++:**
    *   Header-only architecture.
    *   Use `doxygen` style comments (`/** ... */`).
    *   **Cross-Platform:** Must compile on MinGW-w64 (Windows).
    *   **Namespace:** Use `rtd` namespace.
    *   **Style:** Follow standard C++ practices while adhering to COM requirements (explicit `__stdcall`, `HRESULT`, etc.).

*   **Comments:**
    *   **Self-Documenting Code:** Avoid verbose or redundant comments.
    *   **Doc Comments:** Maintain Public API documentation.

## **Verification**

*   **Build:** Use `cmake` with MinGW toolchain to verify the build.
*   **Pre-commit:** Always verify your changes with `read_file` or `ls` before submitting.

## **General Rules**

*   **Do not** delete this file.
*   **Do not** create `src/` directory for the library (use `include/rtd/`).
*   **Do not** use admin-required registry keys (`HKCR`). Use `HKCU` (Current User) for COM registration.
*   All documentation, code comments, commit messages, and other project-related text must be in English.
