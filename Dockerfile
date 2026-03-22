# 使用官方 Python 轻量级基础镜像
FROM python:3.10-slim

# 设置工作目录
WORKDIR /app

# 替换 apt 源为阿里云镜像（解决国内网络 apt-get update 失败报错 code 100 的问题）
RUN sed -i 's/deb.debian.org/mirrors.aliyun.com/g' /etc/apt/sources.list.d/debian.sources 2>/dev/null || sed -i 's/deb.debian.org/mirrors.aliyun.com/g' /etc/apt/sources.list 2>/dev/null || true

# 安装系统级依赖库 (ONNX 和 PyMuPDF 在 Linux 下需要这些 C++ 运行库)
RUN apt-get clean && apt-get update --fix-missing && apt-get install -y \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖列表并安装
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 复制项目所有文件到工作目录
COPY . .

# 暴露端口 (云平台通常映射这个端口)
EXPOSE 5000

# 使用 Gunicorn 作为生产级 Web 服务器启动，开启 2 个 Worker 进程处理并发
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120", "app:app"]
