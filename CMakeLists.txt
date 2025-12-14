cmake_minimum_required(VERSION 3.10)
project(xll-gen-rtd)

# Define the header-only library
add_library(rtd INTERFACE)
target_include_directories(rtd INTERFACE include)

# --- Example: Simple Hybrid Server ---
# XLL is essentially a DLL, so we use add_library with SHARED
add_library(MyHybridServer SHARED examples/simple/main.cpp examples/simple/MyHybrid.def)

# Link against standard Windows libraries
if(MINGW)
    target_link_libraries(MyHybridServer PRIVATE ole32 oleaut32 uuid)

    # Static linking to avoid dependencies
    target_link_options(MyHybridServer PRIVATE -static -static-libgcc -static-libstdc++)
else()
    # On Linux/Non-MinGW, we can't link real Windows libs.
    # But we want to allow compilation checks if possible, or just skip linking.
    # For now, we leave it as is, assuming the user will build with MinGW.
    target_link_libraries(MyHybridServer PRIVATE ole32 oleaut32 uuid)
endif()

# Link the RTD library
target_link_libraries(MyHybridServer PRIVATE rtd)
target_include_directories(MyHybridServer PRIVATE examples/simple) # To find server_impl.h

# Rename output to .xll for Excel
set_target_properties(MyHybridServer PROPERTIES SUFFIX ".xll")

# --- Tests ---
enable_testing()

# Unit Test (C++ logic only)
add_executable(unit_test tests/unit_test.cpp)
target_link_libraries(unit_test PRIVATE rtd ole32 oleaut32 uuid)
target_include_directories(unit_test PRIVATE examples/simple)
add_test(NAME UnitTest COMMAND unit_test)

# Integration Test (Requires Windows/COM)
add_executable(integration_test tests/integration_test.cpp)
target_link_libraries(integration_test PRIVATE rtd ole32 oleaut32 uuid)
add_test(NAME IntegrationTest COMMAND integration_test)
