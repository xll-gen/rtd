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
    target_link_libraries(MyHybridServer PRIVATE ole32 oleaut32 uuid)
endif()

# Link the RTD library
target_link_libraries(MyHybridServer PRIVATE rtd)

# Rename output to .xll for Excel
set_target_properties(MyHybridServer PROPERTIES SUFFIX ".xll")
