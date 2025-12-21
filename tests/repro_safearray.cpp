#include <windows.h>
#include <oleauto.h>
#include <iostream>

int main() {
    // Probe SafeArray Dimension Ordering
    // We want to create A[2][5] -> Left=2, Right=5.

    // Attempt 1: "Natural Order" (bounds[0]=Left, bounds[1]=Right)
    // If we set bounds[0]=2, bounds[1]=5.
    {
        std::cout << "--- Attempt 1: bounds[0]=2, bounds[1]=5 ---" << std::endl;
        SAFEARRAYBOUND bounds[2];
        bounds[0].cElements = 2;
        bounds[0].lLbound = 0;
        bounds[1].cElements = 5;
        bounds[1].lLbound = 0;

        SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 2, bounds);
        if (!psa) {
            std::cout << "Failed to create SafeArray" << std::endl;
            return 1;
        }

        long ub1, ub2;
        SafeArrayGetUBound(psa, 1, &ub1); // Left-Most
        SafeArrayGetUBound(psa, 2, &ub2); // Right-Most

        std::cout << "UBound(1) [Left-Most]: " << ub1 << " (Size " << ub1 + 1 << ")" << std::endl;
        std::cout << "UBound(2) [Right-Most]: " << ub2 << " (Size " << ub2 + 1 << ")" << std::endl;

        // Print interpretation
        if (ub1 == 1 && ub2 == 4) std::cout << "Result: A[2][5] (Matches Excel Requirement)" << std::endl;
        else if (ub1 == 4 && ub2 == 1) std::cout << "Result: A[5][2] (Inverted)" << std::endl;
        else std::cout << "Result: Unknown" << std::endl;

        SafeArrayDestroy(psa);
    }

    std::cout << std::endl;

    // Attempt 2: "Reverse Order" (bounds[0]=Right, bounds[1]=Left)
    // If we want A[2][5], we set bounds[0]=5 (Right), bounds[1]=2 (Left).
    {
        std::cout << "--- Attempt 2: bounds[0]=5, bounds[1]=2 ---" << std::endl;
        SAFEARRAYBOUND bounds[2];
        bounds[0].cElements = 5;
        bounds[0].lLbound = 0;
        bounds[1].cElements = 2;
        bounds[1].lLbound = 0;

        SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 2, bounds);
        if (!psa) {
            std::cout << "Failed to create SafeArray" << std::endl;
            return 1;
        }

        long ub1, ub2;
        SafeArrayGetUBound(psa, 1, &ub1); // Left-Most
        SafeArrayGetUBound(psa, 2, &ub2); // Right-Most

        std::cout << "UBound(1) [Left-Most]: " << ub1 << " (Size " << ub1 + 1 << ")" << std::endl;
        std::cout << "UBound(2) [Right-Most]: " << ub2 << " (Size " << ub2 + 1 << ")" << std::endl;

        // Print interpretation
        if (ub1 == 1 && ub2 == 4) std::cout << "Result: A[2][5] (Matches Excel Requirement)" << std::endl;
        else if (ub1 == 4 && ub2 == 1) std::cout << "Result: A[5][2] (Inverted)" << std::endl;
        else std::cout << "Result: Unknown" << std::endl;

        SafeArrayDestroy(psa);
    }

    return 0;
}
