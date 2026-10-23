# Install script for directory: C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/third_party/miniz-3.1.2

# Set the install prefix
if(NOT DEFINED CMAKE_INSTALL_PREFIX)
  set(CMAKE_INSTALL_PREFIX "C:/Program Files (x86)/text_to_ppt")
endif()
string(REGEX REPLACE "/$" "" CMAKE_INSTALL_PREFIX "${CMAKE_INSTALL_PREFIX}")

# Set the install configuration name.
if(NOT DEFINED CMAKE_INSTALL_CONFIG_NAME)
  if(BUILD_TYPE)
    string(REGEX REPLACE "^[^A-Za-z0-9_]+" ""
           CMAKE_INSTALL_CONFIG_NAME "${BUILD_TYPE}")
  else()
    set(CMAKE_INSTALL_CONFIG_NAME "Release")
  endif()
  message(STATUS "Install configuration: \"${CMAKE_INSTALL_CONFIG_NAME}\"")
endif()

# Set the component getting installed.
if(NOT CMAKE_INSTALL_COMPONENT)
  if(COMPONENT)
    message(STATUS "Install component: \"${COMPONENT}\"")
    set(CMAKE_INSTALL_COMPONENT "${COMPONENT}")
  else()
    set(CMAKE_INSTALL_COMPONENT)
  endif()
endif()

# Is this installation the result of a crosscompile?
if(NOT DEFINED CMAKE_CROSSCOMPILING)
  set(CMAKE_CROSSCOMPILING "FALSE")
endif()

# Set path to fallback-tool for dependency-resolution.
if(NOT DEFINED CMAKE_OBJDUMP)
  set(CMAKE_OBJDUMP "C:/msys64/ucrt64/bin/objdump.exe")
endif()

if(CMAKE_INSTALL_COMPONENT STREQUAL "Unspecified" OR NOT CMAKE_INSTALL_COMPONENT)
  file(INSTALL DESTINATION "${CMAKE_INSTALL_PREFIX}/lib/pkgconfig" TYPE FILE FILES "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/build/_deps/miniz-build/miniz.pc")
endif()

if(CMAKE_INSTALL_COMPONENT STREQUAL "Unspecified" OR NOT CMAKE_INSTALL_COMPONENT)
  file(INSTALL DESTINATION "${CMAKE_INSTALL_PREFIX}/lib" TYPE STATIC_LIBRARY FILES "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/build/_deps/miniz-build/libminiz.a")
endif()

if(CMAKE_INSTALL_COMPONENT STREQUAL "Unspecified" OR NOT CMAKE_INSTALL_COMPONENT)
  if(EXISTS "$ENV{DESTDIR}${CMAKE_INSTALL_PREFIX}/lib/cmake/miniz/minizTargets.cmake")
    file(DIFFERENT _cmake_export_file_changed FILES
         "$ENV{DESTDIR}${CMAKE_INSTALL_PREFIX}/lib/cmake/miniz/minizTargets.cmake"
         "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/build/_deps/miniz-build/CMakeFiles/Export/ce466b740c10b52bd493a42de91419ff/minizTargets.cmake")
    if(_cmake_export_file_changed)
      file(GLOB _cmake_old_config_files "$ENV{DESTDIR}${CMAKE_INSTALL_PREFIX}/lib/cmake/miniz/minizTargets-*.cmake")
      if(_cmake_old_config_files)
        string(REPLACE ";" ", " _cmake_old_config_files_text "${_cmake_old_config_files}")
        message(STATUS "Old export file \"$ENV{DESTDIR}${CMAKE_INSTALL_PREFIX}/lib/cmake/miniz/minizTargets.cmake\" will be replaced.  Removing files [${_cmake_old_config_files_text}].")
        unset(_cmake_old_config_files_text)
        file(REMOVE ${_cmake_old_config_files})
      endif()
      unset(_cmake_old_config_files)
    endif()
    unset(_cmake_export_file_changed)
  endif()
  file(INSTALL DESTINATION "${CMAKE_INSTALL_PREFIX}/lib/cmake/miniz" TYPE FILE FILES "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/build/_deps/miniz-build/CMakeFiles/Export/ce466b740c10b52bd493a42de91419ff/minizTargets.cmake")
  if(CMAKE_INSTALL_CONFIG_NAME MATCHES "^([Rr][Ee][Ll][Ee][Aa][Ss][Ee])$")
    file(INSTALL DESTINATION "${CMAKE_INSTALL_PREFIX}/lib/cmake/miniz" TYPE FILE FILES "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/build/_deps/miniz-build/CMakeFiles/Export/ce466b740c10b52bd493a42de91419ff/minizTargets-release.cmake")
  endif()
endif()

if(CMAKE_INSTALL_COMPONENT STREQUAL "Devel" OR NOT CMAKE_INSTALL_COMPONENT)
  file(INSTALL DESTINATION "${CMAKE_INSTALL_PREFIX}/lib/cmake/miniz" TYPE FILE FILES
    "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/build/_deps/miniz-build/miniz/minizConfig.cmake"
    "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/build/_deps/miniz-build/miniz/minizConfigVersion.cmake"
    )
endif()

if(CMAKE_INSTALL_COMPONENT STREQUAL "Unspecified" OR NOT CMAKE_INSTALL_COMPONENT)
  file(INSTALL DESTINATION "${CMAKE_INSTALL_PREFIX}/include/miniz" TYPE FILE FILES
    "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/third_party/miniz-3.1.2/miniz.h"
    "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/third_party/miniz-3.1.2/miniz_common.h"
    "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/third_party/miniz-3.1.2/miniz_tdef.h"
    "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/third_party/miniz-3.1.2/miniz_tinfl.h"
    "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/third_party/miniz-3.1.2/miniz_zip.h"
    "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/build/_deps/miniz-build/miniz_export.h"
    )
endif()

string(REPLACE ";" "\n" CMAKE_INSTALL_MANIFEST_CONTENT
       "${CMAKE_INSTALL_MANIFEST_FILES}")
if(CMAKE_INSTALL_LOCAL_ONLY)
  file(WRITE "C:/Users/54751/Desktop/myCproject/PPTmaker/text_to_ppt/build/_deps/miniz-build/install_local_manifest.txt"
     "${CMAKE_INSTALL_MANIFEST_CONTENT}")
endif()
