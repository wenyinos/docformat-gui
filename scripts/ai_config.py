#!/usr/bin/env python3
"""AI 检查功能配置。

本模块集中存放 AI 逻辑/病句检查相关的可配置项。
后续接入真实大模型 API 时，接口地址、密钥来源、模型名、超时等
可能变化的内容都优先在这里调整，避免散落到业务代码中。
"""

import json
import os
from pathlib import Path


# 支持的厂商预设。
# 这里使用 OpenAI 兼容的 chat/completions 完整 URL，因为 ai_checker._call_api
# 直接向 API_BASE_URL 发 POST。若厂商文档更新，优先改这里或在本地配置文件里覆盖。
PROVIDER_PRESETS = {
    "kimi": {
        "api_base_url": "https://api.moonshot.cn/v1/chat/completions",
        "model_name": "moonshot-v1-8k",
    },
    "deepseek": {
        "api_base_url": "https://api.deepseek.com/chat/completions",
        "model_name": "deepseek-v4-pro",
    },
}


def _get_config_dir():
    """获取 AI 本地配置目录。

    按用户习惯放在项目根目录，便于直接找到并编辑。
    注意：ai_settings.json 里会保存 API Key，因此应加入 .gitignore，避免误提交。
    """
    return Path(__file__).resolve().parent.parent


CONFIG_FILE = _get_config_dir() / "ai_settings.json"
GLOSSARY_FILE = _get_config_dir() / "ai_glossary.txt"


def _load_local_config():
    """读取本地 AI 配置文件；读取失败时返回空配置，不影响程序启动。"""
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}


def _to_bool(value, default=False):
    """把配置文件或环境变量中的布尔值转换成 bool。"""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on", "是"}
    return default


_LOCAL_CONFIG = _load_local_config()

# 配置优先级：默认值 < 本地 ai_settings.json < 环境变量。
# 环境变量适合临时调试；长期使用建议写入本地配置文件，避免每次启动都 export。
PROVIDER = os.environ.get("DOCFORMAT_AI_PROVIDER", _LOCAL_CONFIG.get("provider", "")).strip().lower()
_PROVIDER_CONFIG = PROVIDER_PRESETS.get(PROVIDER, {})

# 大模型接口地址。
# 可通过本地配置 api_base_url 覆盖；也可通过 DOCFORMAT_AI_BASE_URL 临时覆盖。
API_BASE_URL = os.environ.get(
    "DOCFORMAT_AI_BASE_URL",
    _LOCAL_CONFIG.get("api_base_url") or _PROVIDER_CONFIG.get("api_base_url") or "https://example.com/v1/chat/completions",
)

# API 密钥。
# 优先使用环境变量 DOCFORMAT_AI_KEY；没有环境变量时读取本地配置文件的 api_key。
# 后续如果要做 GUI 设置页或系统钥匙串，也只需要在这里扩展来源。
API_KEY = os.environ.get("DOCFORMAT_AI_KEY", _LOCAL_CONFIG.get("api_key", ""))

# 模型名。
# 可通过本地配置 model_name 覆盖；也可通过 DOCFORMAT_AI_MODEL 临时覆盖。
MODEL_NAME = os.environ.get(
    "DOCFORMAT_AI_MODEL",
    _LOCAL_CONFIG.get("model_name") or _PROVIDER_CONFIG.get("model_name") or "placeholder-model",
)

# 单次发送给模型的最大字符数。
# 先使用保守值，后续可根据真实模型上下文长度、费用和响应稳定性调整。
MAX_CHARS_PER_CHUNK = int(_LOCAL_CONFIG.get("max_chars_per_chunk", 2000) or 2000)

# 请求超时秒数。
# 真实联网调用时用于避免界面长时间无响应。
REQUEST_TIMEOUT = int(_LOCAL_CONFIG.get("request_timeout", 60) or 60)

# 是否强制使用假实现。
# True 时不会联网，check_document 会返回占位问题，方便没有 API Key 时验证 GUI/批注流程。
USE_FAKE = _to_bool(
    os.environ.get("DOCFORMAT_AI_USE_FAKE", _LOCAL_CONFIG.get("use_fake", False)),
    default=False,
)

# v1.8.2 稳定版默认隐藏 AI GUI 入口，避免实验功能影响 Win7 热修版。
# 开发/内测时可设置环境变量 DOCFORMAT_ENABLE_EXPERIMENTAL_AI=1，或在 ai_settings.json
# 中写入 "enable_experimental_ai": true 后显示 AI 相关入口。
EXPERIMENTAL_AI_UI = _to_bool(
    os.environ.get("DOCFORMAT_ENABLE_EXPERIMENTAL_AI", _LOCAL_CONFIG.get("enable_experimental_ai", False)),
    default=False,
)

# AI 功能是否可用。
# 有真实 API Key 时可用；开启 USE_FAKE 时也视为可用，方便 GUI 入口测试完整流程。
ENABLED = bool(API_KEY) or USE_FAKE


def is_ai_available():
    """返回 AI 功能当前是否可用。"""
    return ENABLED


def is_ai_ui_enabled():
    """返回是否在 GUI 中显示实验性 AI 功能入口。"""
    return EXPERIMENTAL_AI_UI


def get_config_file_path():
    """返回本地 AI 配置文件路径，便于 GUI 或命令行提示用户。"""
    return CONFIG_FILE


def get_glossary_path():
    """返回本单位术语表路径，便于 GUI 或 README 提示用户。"""
    return GLOSSARY_FILE


def load_glossary():
    """读取本单位术语表文本。

    过滤规则：
    - 忽略纯空行；
    - 忽略去掉前置空格后以 # 开头的注释行；
    - 其他内容原样保留，包括【】标题行和中文标点。

    每次调用都重新读取文件，方便用户编辑 ai_glossary.txt 后立即生效。
    任何读取异常、编码错误、文件不存在或无有效内容时，都返回空字符串。
    """
    try:
        with open(GLOSSARY_FILE, "r", encoding="utf-8") as f:
            lines = []
            for line in f.read().splitlines():
                if not line.strip():
                    continue
                if line.lstrip().startswith("#"):
                    continue
                lines.append(line)
        return "\n".join(lines)
    except Exception:
        return ""
