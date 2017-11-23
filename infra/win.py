import os
import platform


def win_bits() -> int:
    return 64 if "PROGRAMFILES(X86)" in os.environ else 32


def win_ver():
    ver = platform.version().split('.')
    major_ver = int(ver[0])
    build = int(ver[2])

    str_build = ''
    if build <= 10240:
        str_build = 'th1'
    elif build <= 10586:
        str_build = 'th2'
    elif build <= 14393:
        str_build = 'rs1'
    elif build <= 15063:
        str_build = 'rs2'
    elif build <= 16299:
        str_build = 'rs3'
    elif build <= 17046:
        str_build = 'rs4'

    return major_ver, build, str_build
