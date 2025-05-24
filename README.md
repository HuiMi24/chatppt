# ChatPPT

ChatPPT is a versatile tool that leverages multiple Large Language Model (LLM) providers to help you generate PowerPoint presentations (`.pptx` files). It supports content generation in English and Chinese and allows for customization through templates.

## Table of Contents

- [Features](#features)
- [What is ChatPPT](#what-is-chatppt)
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
*   **Custom OpenAI API Endpoint**: Support for specifying a custom base URL for OpenAI-compatible APIs (e.g., for local LLMs or proxies).
*   **Enhanced UI for Model Selection**:
    *   **Provider Choice**: Easily switch between OpenAI, Ollama, and Anthropic.
    *   **OpenAI**: Select specific models (e.g., `gpt-3.5-turbo`, `gpt-4`, `gpt-4o`) from a dropdown. Input field for custom OpenAI API base URL.
    *   **Ollama**: Dynamically fetches and lists available models from your connected Ollama instance in a dropdown. Provides a text input fallback if models can't be fetched.
    *   **Anthropic**: Input your API key and select from available Claude models (e.g., `claude-3-opus-20240229`, `claude-3-sonnet-20240229`) via a dropdown.
*   **PowerPoint Template Upload**:
    *   Users can upload their own `.pptx` file, which ChatPPT will use as a template for the generated presentation, preserving layouts and styles.
*   **Streamlit Web Interface**: An intuitive UI for easy presentation generation without needing CLI commands for most users.
*   **Command-Line Interface**: Retains a powerful CLI for users who prefer or need to automate presentation generation.

<!-- TODO: Update screenshots to reflect new UI features -->
<!-- Existing screenshots like ui_demo_1.png, ui_demo_2.png might be outdated -->

## What is ChatPPT

ChatPPT is powered by various leading LLMs. It's designed to simplify the creation of presentations by generating slide outlines and content based on your topic. It supports output in English and Chinese.

<!-- TODO: Update screenshots to reflect new UI features -->
<!-- Existing demo screenshots like demo1.png, demo2.png might be outdated -->

## Requirements

*   Python 3.8 or higher.
*   Dependencies listed in `requirements.txt` (includes `openai`, `ollama`, `anthropic`, `python-pptx`, `streamlit`).
*   **API Keys**:
    *   An API key is required for models from **OpenAI**.
    *   An API key is required for models from **Anthropic**.
    *   Ollama runs locally and typically does not require an API key, but needs to be installed and running.

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

3.  **Set up LLM Providers:**
    *   **Ollama**: Follow the [official guide](https://ollama.com/) to install Ollama and download your desired models (e.g., `ollama pull llama3`).
    *   **OpenAI**: Generate your OpenAI API key at <https://platform.openai.com/account/api-keys>. You can provide this key directly in the UI/CLI or store it in a file (e.g., `.token`) and provide the file path. If using a custom OpenAI-compatible API, you might use a different key and will need the base URL.
    *   **Anthropic**: Obtain your API key from the [Anthropic Console](https://console.anthropic.com/). Provide this key directly in the UI/CLI or store it in a file.

## Usage

ChatPPT can be used via its Streamlit web interface or through the command line.

### Command-Line Interface (CLI)

The CLI has been updated to support the new model providers and options.

**Help Message:**
```bash
python chatppt.py -h
usage: chatppt.py [-h] [-m {openai,ollama,anthropic}] -n MODEL_NAME -t TOPIC
                  [-k API_KEY] [--anthropic_api_key ANTHROPIC_API_KEY]
                  [--openai_api_base OPENAI_API_BASE] [-u OLLAMA_URL]
                  [-p PAGES] [-l {cn,en}]

I am your PPT assistant, I can help to you generate PPT.

options:
  -h, --help            show this help message and exit
  -m {openai,ollama,anthropic}, --model_provider {openai,ollama,anthropic}
                        Select the model provider (e.g., openai, ollama,
                        anthropic)
  -n MODEL_NAME, --model_name MODEL_NAME
                        Specify the model name to use (e.g., gpt-3.5-turbo,
                        llama3, claude-3-opus-20240229)
  -t TOPIC, --topic TOPIC
                        Your topic name
  -k API_KEY, --api_key API_KEY
                        Your OpenAI API key or file path
  --anthropic_api_key ANTHROPIC_API_KEY
                        Your Anthropic API key or file path
  --openai_api_base OPENAI_API_BASE
                        Optional custom base URL for the OpenAI API. (e.g.,
                        http://localhost:8000/v1)
  -u OLLAMA_URL, --ollama_url OLLAMA_URL
                        Your ollama url
  -p PAGES, --pages PAGES
                        How many slides to generate
  -l {cn,en}, --language {cn,en}
                        Output language
```
The `--openai_api_base <URL>` argument allows you to specify a custom endpoint for the OpenAI API. This is useful if you are using a proxy, a local LLM server that mimics the OpenAI API (like LocalAI or vLLM's OpenAI-compatible server), or any other service that provides an OpenAI-compatible API.

**CLI Examples:**

*   **OpenAI (Standard):**
    ```bash
    python chatppt.py --model_provider openai --model_name gpt-4o --api_key YOUR_OPENAI_KEY --topic "The Future of AI" --pages 7
    ```
    (Replace `YOUR_OPENAI_KEY` with your actual key or a path to a file containing the key.)

*   **OpenAI (Custom API Base URL):**
    ```bash
    # Example using a custom OpenAI-compatible API endpoint
    python chatppt.py --model_provider openai \
                     --model_name your_compatible_model_name \
                     --api_key your_api_key_for_custom_endpoint \
                     --openai_api_base http://localhost:8000/v1 \
                     --topic "My Local LLM Presentation"
    ```
    (Replace `your_compatible_model_name` and `your_api_key_for_custom_endpoint` accordingly. The API key might be optional or different depending on your custom endpoint's configuration.)

*   **Ollama:**
    ```bash
    python chatppt.py --model_provider ollama --model_name llama3 --ollama_url http://localhost:11434 --topic "Introduction to Ollama"
    ```
    (Ensure your Ollama instance is running at the specified URL and the model `llama3` is available.)

*   **Anthropic:**
    ```bash
    python chatppt.py --model_provider anthropic --model_name claude-3-opus-20240229 --anthropic_api_key YOUR_ANTHROPIC_KEY --topic "Advanced Language Models by Anthropic"
    ```
    (Replace `YOUR_ANTHROPIC_KEY` with your actual key or a path to a file containing the key.)

### Streamlit UI

The Streamlit UI provides an easy-to-use interface for all features.

1.  **Start the Streamlit application:**
    ```bash
    streamlit run chatppt_ui.py
    ```

2.  **Open the URL provided by Streamlit in your browser (usually `http://localhost:8501`).**

3.  **Using the UI:**
    *   **Select Model Provider**: Choose between "openai", "ollama", or "anthropic" from the main dropdown.
    *   **Conditional Inputs**:
        *   **If OpenAI is selected**:
            *   Enter your OpenAI API Key.
            *   Select a specific OpenAI model (e.g., `gpt-3.5-turbo`, `gpt-4`, `gpt-4o`) from the dropdown.
            *   **OpenAI API Base URL (Optional)**: If you are using an OpenAI-compatible proxy or a local LLM server, enter its base URL here (e.g., `http://localhost:8000/v1`). Leave blank to use the default OpenAI API.
        *   **If Ollama is selected**:
            *   Enter the URL for your running Ollama instance (defaults to `http://localhost:11434`).
            *   Available models will be dynamically fetched and displayed in a dropdown. If fetching fails or the URL is not provided, a text input field will appear to manually enter the Ollama model name.
        *   **If Anthropic is selected**:
            *   Enter your Anthropic API Key.
            *   Select a specific Anthropic Claude model (e.g., `claude-3-opus-20240229`) from the dropdown.
    *   **Enter Topic**: Provide the topic for your presentation.
    *   **Number of Pages**: Use the slider to set the desired number of slides.
    *   **Select Language**: Choose between "en" (English) or "cn" (Chinese).
    *   **Custom Instructions (Optional)**: Add any specific instructions or context for the AI in the text area.
    *   **Upload Template (Optional)**: Click the "Browse files" button to upload a `.pptx` file to be used as a template for your presentation.
    *   **Generate Slide**: Click the "Generate Slide" button.

<!-- TODO: Update screenshots to reflect new UI features -->
<!-- Existing screenshot ui.png might be outdated -->

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues to improve ChatPPT.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
