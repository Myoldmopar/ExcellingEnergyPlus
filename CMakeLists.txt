cmake_minimum_required(VERSION 3.10)
project(Trampoline C)

if(NOT WIN32)
    message(FATAL_ERROR "This project can only be built on Windows.")
endif()

add_executable(trampoline trampoline.c)

target_link_libraries(MyWindowsApp PRIVATE
    user32
    kernel32
)
