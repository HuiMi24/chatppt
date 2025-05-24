# ChatPPT

ChatPPT is a versatile tool that leverages multiple Large Language Model (LLM) providers to help you generate PowerPoint presentations (`.pptx` files). It supports content generation in English and Chinese and allows for customization through templates and audience-specific tailoring.

## Table of Contents

- [Features](#features)
- [What is ChatPPT](#what-is-chatppt)
- [Configuration via `.env` File](#configuration-via-env-file)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
  - [Command-Line Interface (CLI)](#command-line-interface-cli)
  - [Streamlit UI](#streamlit-ui)
- [Contributing](#contributing)
- [License](#license)

## Features

ChatPPT has been significantly updated with a focus on flexibility and user experience:

*   **Multiple LLM Provider Support**:
    *   **OpenAI**: Utilize models like GPT-3.5-turbo, GPT-4, GPT-4o, etc.
    *   **Ollama**: Connect to your local Ollama instance and use any of your downloaded models.
    *   **Anthropic**: Leverage Claude models for content generation.
    *   **Groq**: Access fast inference on models like Llama3 and Mixtral via Groq's OpenAI-compatible API.
*   **Configuration via `.env` File**: Securely manage API keys, default models, and API base URLs using an environment file, keeping sensitive information out of your command line and UI inputs.
*   **Custom OpenAI API Endpoint**: Support for specifying a custom base URL for OpenAI-compatible APIs (e.g., for local LLMs or proxies).
*   **Audience-Specific Content**: Tailor your presentation's tone and content by specifying the target audience.
*   **Enhanced UI for Model Selection**:
    *   **Provider Choice**: Easily switch between OpenAI, Ollama, Anthropic, and Groq.
    *   **OpenAI**: Select specific models from a dropdown. Input field for API key and custom API base URL (hidden if set in `.env`).
    *   **Ollama**: Dynamically fetches and lists available models from your connected Ollama instance.
    *   **Anthropic**: Select specific models from a dropdown. API key input hidden if set in `.env`.
    *   **Groq**: Input for model name (e.g., `mixtral-8x7b-32768`). API key input hidden if set in `.env`.
*   **PowerPoint Template Upload**: Users can upload their own `.pptx` file to be used as a template.
*   **Custom Prompt Instructions**: Provide additional instructions to the LLM for more fine-tuned results.
*   **Streamlit Web Interface**: An intuitive UI for easy presentation generation.
*   **Command-Line Interface**: A powerful CLI for automation and advanced users.

<!-- TODO: Update screenshots to reflect new UI features -->

## What is ChatPPT

ChatPPT is powered by various leading LLMs. It's designed to simplify the creation of presentations by generating slide outlines and content based on your topic, language, and target audience.

<!-- TODO: Update screenshots to reflect new UI features -->

## Configuration via `.env` File

ChatPPT supports configuration of API keys and other settings through an environment file named `.env` located in the project root. This method is recommended for managing sensitive information like API keys and for setting persistent defaults.

1.  **Create your `.env` file**: Copy the provided sample file:
    ```bash
    cp sample.env .env
    ```
2.  **Edit `.env`**: Open the `.env` file in a text editor and fill in your actual values. Remove the `#` from the beginning of lines you wish to activate and set the appropriate values.

**Supported Environment Variables:**

*   `OPENAI_API_KEY="sk-..."`: Your OpenAI API Key.
*   `OPENAI_API_BASE="https://api.example.com/v1"`: Optional custom base URL for OpenAI-compatible services.
*   `ANTHROPIC_API_KEY="sk-ant-..."`: Your Anthropic API Key.
*   `GROQ_API_KEY="gsk_..."`: Your Groq API Key.
*   `GROQ_API_BASE="https://api.groq.com/openai/v1"`: Optional Groq API base URL (defaults to the standard if not set).
*   `GROQ_DEFAULT_MODEL="mixtral-8x7b-32768"`: Default model to use for Groq if not specified by the user.

**Benefits of using `.env`:**

*   **Security**: Keeps API keys out of your shell history and command-line arguments.
*   **Convenience**: Avoids repetitive input of keys or URLs in the UI.
*   **UI Integration**: If an API key or custom base URL is set in `.env`, the corresponding input field in the Streamlit UI will be hidden, and a message will indicate that the value is loaded from the environment.

## Requirements

*   Python 3.8 or higher.
*   Dependencies listed in `requirements.txt` (install with `pip install -r requirements.txt`).
*   **API Keys (if not using Ollama exclusively)**:
    *   **OpenAI**: Required if using the OpenAI provider.
    *   **Anthropic**: Required if using the Anthropic provider.
    *   **Groq**: Required if using the Groq provider.
    *   These can be set in the `.env` file or provided via CLI/UI.
*   **Ollama**: Needs to be installed and running if using the Ollama provider.

## Installation

1.  **Clone the repository (if you haven't already):**
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Install Python dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Set up your `.env` file (Recommended):**
    ```bash
    cp sample.env .env
    ```
    Then edit `.env` to add your API keys and other configurations.

4.  **Set up LLM Providers (if not using `.env` for keys):**
    *   **Ollama**: Follow the [official guide](https://ollama.com/) to install Ollama and download models (e.g., `ollama pull llama3`).
    *   **OpenAI**: Get your API key from <https://platform.openai.com/account/api-keys>.
    *   **Anthropic**: Get your API key from the [Anthropic Console](https://console.anthropic.com/).
    *   **Groq**: Get your API key from <https://console.groq.com/keys>.

## Usage

ChatPPT can be used via its Streamlit web interface or through the command line.

### Command-Line Interface (CLI)

The CLI allows for detailed control over presentation generation. For the most up-to-date list of commands and arguments, run:
```bash
python chatppt.py -h
```

**Current Help Message (as of last update):**
```bash
usage: chatppt.py [-h] [-m {openai,ollama,anthropic,groq}] [-n MODEL_NAME] -t
                  TOPIC [--audience AUDIENCE] [-k API_KEY]
                  [--anthropic_api_key ANTHROPIC_API_KEY]
                  [--groq_api_key GROQ_API_KEY]
                  [--openai_api_base OPENAI_API_BASE] [-u OLLAMA_URL]
                  [-p PAGES] [-l {cn,en}]

I am your PPT assistant, I can help to you generate PPT.

options:
  -h, --help            show this help message and exit
  -m {openai,ollama,anthropic,groq}, --model_provider {openai,ollama,anthropic,groq}
                        Select the model provider
  -n MODEL_NAME, --model_name MODEL_NAME
                        Specify the model name to use (e.g., gpt-3.5-turbo,
                        llama3, claude-3-opus, mixtral-8x7b-32768). Required
                        unless provider has a default (e.g. Groq).
  -t TOPIC, --topic TOPIC
                        Your topic name
  --audience AUDIENCE   Specify the target audience for the presentation
                        (e.g., 'students', 'technical experts').
  -k API_KEY, --api_key API_KEY
                        Your OpenAI API key or file path. Overrides .env.
  --anthropic_api_key ANTHROPIC_API_KEY
                        Your Anthropic API key or file path. Overrides .env.
  --groq_api_key GROQ_API_KEY
                        Your Groq API key. Overrides .env.
  --openai_api_base OPENAI_API_BASE
                        Custom base URL for OpenAI API. Overrides .env.
  -u OLLAMA_URL, --ollama_url OLLAMA_URL
                        Your ollama url
  -p PAGES, --pages PAGES
                        How many slides to generate
  -l {cn,en}, --language {cn,en}
                        Output language
```
*   The `--openai_api_base <URL>` argument allows specifying a custom endpoint for OpenAI-compatible APIs.
*   The `--audience "<description>"` argument helps tailor the content (e.g., `--audience "High school students"`).

**CLI Examples:**

*   **OpenAI (Standard):**
    ```bash
    python chatppt.py --model_provider openai --model_name gpt-4o --topic "The Future of AI" --pages 7 --audience "Tech Enthusiasts"
    ```
    (Set `OPENAI_API_KEY` in `.env` or use `--api_key YOUR_KEY`.)

*   **OpenAI (Custom API Base URL):**
    ```bash
    python chatppt.py --model_provider openai --model_name your_model --openai_api_base http://localhost:8000/v1 --topic "Local LLM Test"
    ```
    (Set API key in `.env` or use `--api_key YOUR_KEY`.)

*   **Ollama:**
    ```bash
    python chatppt.py --model_provider ollama --model_name llama3 --ollama_url http://localhost:11434 --topic "Introduction to Ollama"
    ```

*   **Anthropic:**
    ```bash
    python chatppt.py --model_provider anthropic --model_name claude-3-opus-20240229 --topic "AI Ethics"
    ```
    (Set `ANTHROPIC_API_KEY` in `.env` or use `--anthropic_api_key YOUR_KEY`.)

*   **Groq:**
    ```bash
    # Model name can be defaulted from .env (GROQ_DEFAULT_MODEL) or uses mixtral-8x7b-32768 if not specified
    python chatppt.py --model_provider groq --topic "Fast Inference with Groq"
    # Specify model explicitly
    python chatppt.py --model_provider groq --model_name llama3-70b-8192 --topic "Large Models on Groq"
    ```
    (Set `GROQ_API_KEY` in `.env` or use `--groq_api_key YOUR_KEY`.)

### Streamlit UI

The Streamlit UI provides an easy-to-use interface for all features.

1.  **Start the Streamlit application:**
    ```bash
    streamlit run chatppt_ui.py
    ```

2.  **Open the URL provided by Streamlit in your browser (usually `http://localhost:8501`).**

3.  **Using the UI:**
    *   **Select Model Provider**: Choose between "openai", "ollama", "anthropic", or "groq".
    *   **API Keys & Base URLs**:
        *   Input fields for API keys (OpenAI, Anthropic, Groq) and OpenAI API Base URL will be shown *only if* the corresponding values are not set in your `.env` file.
        *   If set in `.env`, a message will confirm it's loaded from the environment.
    *   **Model Selection**:
        *   **OpenAI/Anthropic**: Select specific models from a dropdown.
        *   **Ollama**: Enter your Ollama URL; available models are then fetched and listed in a dropdown. A text input fallback is provided.
        *   **Groq**: Enter the model name (e.g., `mixtral-8x7b-32768`). Defaults to `GROQ_DEFAULT_MODEL` from `.env` or `mixtral-8x7b-32768`.
    *   **Enter Topic**: Provide the topic for your presentation.
    *   **Number of Pages**: Use the slider to set the desired number of slides.
    *   **Select Language**: Choose between "en" (English) or "cn" (Chinese).
    *   **Select Target Audience**: Choose from predefined options ("General", "Student", "Software Engineer", "Kids") or select "Custom" to enter a specific audience description.
    *   **Custom Instructions (Optional)**: Add any specific instructions or context for the AI.
    *   **Upload Template (Optional)**: Upload a `.pptx` file to use as a template.
    *   **Generate Slide**: Click the "Generate Slide" button.

<!-- TODO: Update screenshots to reflect new UI features -->

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues to improve ChatPPT.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
