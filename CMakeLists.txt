# CMakeList.txt : Top-level CMake project file, do global configuration
# and include sub-projects here.
#
cmake_minimum_required (VERSION 3.8)

# Here because why not ? CMAKE_ASM_NASM_OBJECT_FORMAT is checked when project is define.
# Hence early definition
set(CMAKE_ASM_NASM_OBJECT_FORMAT win64)

# Project name
project(DynamicWrapperEx VERSION 1.0 LANGUAGES CXX ASM_NASM)

# Define C++ standard
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Make makefile verbose to display command lines 
set(CMAKE_VERBOSE_MAKEFILE ON)

# Enable NASM
enable_language(ASM_NASM)
string(APPEND CMAKE_ASM_NASM_FLAGS "-g -I ${CMAKE_CURRENT_SOURCE_DIR}/asm")

# Compile code
add_library(DynamicWrapperEx MODULE
	# Definition file
	"exports.def"

	# NASM files
	"asm/DynamicCall.asm"
	"asm/DynamicCall.inc"
	
	# C++ files
	"src/dllmain.cpp"
	"src/Util.cpp"
	"src/DynamicMethod.cpp"
	"src/AutomationFactory.cpp"
	"src/IDynamicWrapperEx.cpp"
	"src/CDynamicWrapperEx.cpp"
)

# Make the headers accessible everywhere
target_include_directories(DynamicWrapperEx PRIVATE "${CMAKE_CURRENT_SOURCE_DIR}/inc")

# Add library for COM util
target_link_libraries(DynamicWrapperEx PRIVATE comsuppw.lib)