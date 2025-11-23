# 使用官方 Python 基础镜像（稳定版）
FROM python:3.10-slim

# 1. 安装系统依赖：包括 Rust 所需的构建工具、curl 和 LangChain/PDF 处理所需的库
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# 2. 安装 Rust 工具链
# 使用 rustup 安装最新的 stable Rust 版本
# 注意：Streamlit Cloud 的基础环境通常提供旧版 Rust。
# 在 Dockerfile 中安装最新的 Rust 可以绕过该限制。
ENV RUSTUP_HOME=/usr/local/rustup \
    CARGO_HOME=/usr/local/cargo \
    PATH=/usr/local/cargo/bin:$PATH

RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y --default-toolchain stable

# 3. 复制代码和安装 Python 依赖
WORKDIR /app

# 复制 requirements.txt 并安装，利用新的 Rust 工具链编译 tiktoken
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 4. 复制应用程序代码
COPY . /app

# 5. 定义启动命令
# 暴露 Streamlit 的端口
EXPOSE 8501

# 启动 Streamlit 应用程序
ENTRYPOINT ["streamlit", "run", "app.py"]
