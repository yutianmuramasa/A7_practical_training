# 模式切换：'mock' 用本地模拟，'real' 用真实 API
MODE = 'mock'  # 上线时改成 'real'

# 真实大模型 API 地址（假设用通义千问、ChatGLM、文心一言都行）
REAL_API_URL = 'http://127.0.0.1:8000/generate'  # 你本地部署的模型地址

# 如果需要 Key 鉴权，这里填
API_KEY = ''
