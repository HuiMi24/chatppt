# -*- coding: utf-8 -*-
"""Configuration loader for ChatPPT.

This module loads API keys, base URLs, and other settings from a .env file.
"""
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# API Keys
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

# API Base URLs
OPENAI_API_BASE = os.getenv("OPENAI_API_BASE")
GROQ_API_BASE = os.getenv("GROQ_API_BASE") # Example: "https://api.groq.com/openai/v1"

# Default Model Names (can be overridden by user)
# These are examples; users can set them in their .env
# OLLAMA_DEFAULT_MODEL = os.getenv("OLLAMA_DEFAULT_MODEL", "llama3")
# ANTHROPIC_DEFAULT_MODEL = os.getenv("ANTHROPIC_DEFAULT_MODEL", "claude-3-haiku-20240307")
# OPENAI_DEFAULT_MODEL = os.getenv("OPENAI_DEFAULT_MODEL", "gpt-3.5-turbo")
GROQ_DEFAULT_MODEL = os.getenv("GROQ_DEFAULT_MODEL") # e.g., "mixtral-8x7b-32768"

# You can add a general default model if no provider-specific default is set
# DEFAULT_MODEL_NAME = os.getenv("DEFAULT_MODEL_NAME", "gpt-3.5-turbo")

# UI specific configurations (optional)
# HIDE_API_KEYS_IF_SET = os.getenv("HIDE_API_KEYS_IF_SET", "False").lower() in ("true", "1", "t")
# HIDE_PROVIDER_SELECTION_IF_ONLY_ONE_CONFIGURED_STR = os.getenv(
#     "HIDE_PROVIDER_SELECTION_IF_ONLY_ONE_CONFIGURED", "False"
# )
# HIDE_PROVIDER_SELECTION_IF_ONLY_ONE_CONFIGURED = (
#     HIDE_PROVIDER_SELECTION_IF_ONLY_ONE_CONFIGURED_STR.lower()
#     in ("true", "1", "t")  # Wrapped for length
# )

# Add other configurations as needed in future steps
