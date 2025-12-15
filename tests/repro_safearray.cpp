#include <windows.h>
#include <oleauto.h>
#include <iostream>

// Standalone test to verify SafeArray Indexing behavior
// Compile with: x86_64-w64-mingw32-g++ repro_safearray.cpp -o repro.exe -loleaut32 -static

int main() {
    std::cout << "Verifying SafeArray Behavior..." << std::endl;

    // Simulation of server_impl.h logic
    // bounds[0] = 2 (intended for Row Count?)
    // bounds[1] = 3 (intended for Topic Count?)
    SAFEARRAYBOUND bounds[2];
    bounds[0].cElements = 2; // Index 0 (Right-Most in Windows API)
    bounds[0].lLbound = 0;
    bounds[1].cElements = 3; // Index 1 (Left-Most in Windows API)
    bounds[1].lLbound = 0;

    SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 2, bounds);
    if (!psa) {
        std::cerr << "Failed to create SafeArray" << std::endl;
        return 1;
    }

    long indices[2];
    HRESULT hr;

    // Test 1: Accessing indices[0] = 2 (Topic Index 2)
    // If bounds[0] corresponds to indices[0], this should FAIL (size 2).
    indices[0] = 2; // "Topic 2"
    indices[1] = 0; // "Row 0"

    VARIANT v; VariantInit(&v);
    v.vt = VT_I4; v.lVal = 123;

    hr = SafeArrayPutElement(psa, indices, &v);

    if (FAILED(hr)) {
        std::cout << "Test Passed: SafeArrayPutElement failed as expected with index 2 on dimension 0 (size 2)." << std::endl;
        std::cout << "Error Code: " << std::hex << hr << std::endl;
    } else {
        std::cerr << "Test Failed: SafeArrayPutElement succeeded! This means indices[0] does NOT map to bounds[0] or bounds behavior is non-standard." << std::endl;
    }

    SafeArrayDestroy(psa);
    return 0;
}
