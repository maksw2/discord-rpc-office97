include_directories(${PROJECT_SOURCE_DIR}/include)
add_library(office97 SHARED
    excel.c
    word.c
    shared.c
    richpresence.c
    richpresence.h
    exports.def
)

target_link_libraries(office97 discord-rpc)

target_compile_options(office97 PRIVATE /std:clatest)
