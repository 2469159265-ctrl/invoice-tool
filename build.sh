#!/bin/bash
# 构建脚本 - 在 macOS 上交叉编译 Windows exe
# 需要安装 PyInstaller: pip install pyinstaller

set -e

echo "=== 多页小票提取工具 构建脚本 ==="

# 检查是否为 macOS
if [[ "$OSTYPE" == "darwin"* ]]; then
    echo "检测到 macOS，需要使用 Docker 进行交叉编译"
    echo "或者可以直接在 Windows 上运行: pip install -r requirements.txt && pyinstaller receipt_extractor.spec"
    exit 1
fi

# 创建虚拟环境（如果需要）
if [ ! -d "venv" ]; then
    echo "创建虚拟环境..."
    python -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
echo "安装依赖..."
pip install -r requirements.txt

# 清理旧构建
echo "清理旧构建..."
rm -rf build dist

# 构建
echo "开始构建..."
pyinstaller receipt_extractor.spec

echo ""
echo "=== 构建完成 ==="
echo "输出文件: dist/多页小票提取工具.exe"
ls -la dist/
