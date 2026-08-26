#!/usr/bin/env python3
"""
启动 WPS Excel Skill 桥接服务
支持 Linux/Windows x86+ARM，纯 Python 标准库。
"""

import sys
import os

# 将 bridge 目录加入路径
script_dir = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(script_dir)
bridge_dir = os.path.join(project_dir, "bridge")
sys.path.insert(0, bridge_dir)

# 导入并启动桥接服务
from server import main

if __name__ == "__main__":
    main()
