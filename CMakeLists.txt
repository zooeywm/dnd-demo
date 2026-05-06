cmake_minimum_required(VERSION 3.16)

project(dnd-demo VERSION 0.1 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Qt6 REQUIRED COMPONENTS Core Network Widgets)

qt_standard_project_setup(REQUIRES 6.5)

qt_add_executable(vfd-linux-client
    src/linux_client.cpp
    src/vfd_protocol.h
)

target_link_libraries(vfd-linux-client
    PRIVATE
        Qt6::Core
        Qt6::Network
        Qt6::Widgets
)

if(WIN32)
    qt_add_executable(vfd-windows-agent
        src/windows_agent.cpp
        src/vfd_protocol.h
    )

    target_link_libraries(vfd-windows-agent
        PRIVATE
            Qt6::Core
            Qt6::Network
            ole32
            user32
    )

    set_target_properties(vfd-windows-agent PROPERTIES
        WIN32_EXECUTABLE FALSE
    )
endif()

include(GNUInstallDirs)
install(TARGETS vfd-linux-client
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
)

if(WIN32)
    install(TARGETS vfd-windows-agent
        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    )
endif()
