cmake_minimum_required(VERSION 3.20)
project(text_to_ppt CXX)

# ⚠️ Aspose.Slides FOSS 要求 C++20
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# === 关键：告诉 CMake 使用本地的依赖，不要去 GitHub 下载 ===
set(FETCHCONTENT_SOURCE_DIR_PUGIXML "${CMAKE_CURRENT_SOURCE_DIR}/third_party/pugixml-1.16" CACHE PATH "Local pugixml")
set(FETCHCONTENT_SOURCE_DIR_MINIZ "${CMAKE_CURRENT_SOURCE_DIR}/third_party/miniz-3.1.2" CACHE PATH "Local miniz")

# === 引入本地的 Aspose.Slides ===
add_subdirectory(third_party/Aspose.Slides-FOSS-for-Cpp-26.9.0)

# === 你的项目源文件 ===
file(GLOB_RECURSE SOURCES src/*.cpp)
add_executable(text_to_ppt ${SOURCES})
target_include_directories(text_to_ppt PRIVATE src)

# === 【修改这里】链接并显式添加头文件路径 ===
target_link_libraries(text_to_ppt PRIVATE AsposeSlidesFoss)  # 改成日志里的真实目标名

target_include_directories(text_to_ppt PRIVATE
    src
    third_party
    third_party/Aspose.Slides-FOSS-for-Cpp-26.9.0/src     # 头文件所在位置
)