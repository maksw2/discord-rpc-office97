cmake_minimum_required (VERSION 3.15.0)

if(NOT CMAKE_GENERATOR_PLATFORM STREQUAL "Win32")
    message(FATAL_ERROR "Please configure with -A Win32!")
endif()

set(CMAKE_SKIP_INSTALL_RULES ON)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

if(NOT DEFINED CMAKE_GENERATOR_PLATFORM)
    set(CMAKE_GENERATOR_PLATFORM Win32 CACHE STRING "Platform to build for" FORCE)
endif()

project (DiscordRPC)

# thirdparty stuff
execute_process(
    COMMAND mkdir ${CMAKE_CURRENT_SOURCE_DIR}/thirdparty
    ERROR_QUIET
)

find_file(RAPIDJSONTEST NAMES rapidjson rapidjson-1.1.0 PATHS ${CMAKE_CURRENT_SOURCE_DIR}/thirdparty CMAKE_FIND_ROOT_PATH_BOTH)
if (NOT RAPIDJSONTEST)
    message("no rapidjson, download")
    set(RJ_TAR_FILE ${CMAKE_CURRENT_SOURCE_DIR}/thirdparty/v1.1.0.tar.gz)
    file(DOWNLOAD https://github.com/miloyip/rapidjson/archive/v1.1.0.tar.gz ${RJ_TAR_FILE})
    execute_process(
        COMMAND ${CMAKE_COMMAND} -E tar xzf ${RJ_TAR_FILE}
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/thirdparty
    )
    file(REMOVE ${RJ_TAR_FILE})
endif(NOT RAPIDJSONTEST)

find_file(RAPIDJSON NAMES rapidjson rapidjson-1.1.0 PATHS ${CMAKE_CURRENT_SOURCE_DIR}/thirdparty CMAKE_FIND_ROOT_PATH_BOTH)

add_library(rapidjson STATIC IMPORTED ${RAPIDJSON})

add_subdirectory(src)
add_subdirectory(office97)
