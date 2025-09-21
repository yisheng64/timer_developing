"""
安装依赖脚本
运行此脚本来安装所需的依赖包
"""

import subprocess
import sys

def install_package(package):
    """安装Python包"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"✓ 成功安装 {package}")
        return True
    except subprocess.CalledProcessError:
        print(f"✗ 安装 {package} 失败")
        return False

def main():
    print("正在安装任务计时器所需的依赖包...")
    print("=" * 50)
    
    packages = [
        "psutil",
        "pywin32"
    ]
    
    success_count = 0
    for package in packages:
        if install_package(package):
            success_count += 1
    
    print("=" * 50)
    if success_count == len(packages):
        print("✓ 所有依赖包安装成功！现在可以运行任务计时器了。")
    else:
        print(f"⚠ 有 {len(packages) - success_count} 个包安装失败，请手动安装。")
        print("手动安装命令：")
        for package in packages:
            print(f"  pip install {package}")

if __name__ == "__main__":
    main()
