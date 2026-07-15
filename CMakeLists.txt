cmake_minimum_required(VERSION 3.15)
project(ExcelSheetSwitcher)

# 使用静态运行时库，避免目标机器缺少 vcruntime
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

set(SOURCES
    src/main.cpp
    src/ExcelCom.cpp
)

add_executable(ExcelSheetSwitcher WIN32 ${SOURCES})

target_compile_options(ExcelSheetSwitcher PRIVATE /O2 /MT)

# 确保编译为 Release 优化版本
target_compile_definitions(ExcelSheetSwitcher PRIVATE $<$<CONFIG:Release>:NDEBUG>)
