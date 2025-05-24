# ChatPPT

ChatPPT is a versatile tool that leverages multiple Large Language Model (LLM) providers to help you generate PowerPoint presentations (`.pptx` files). It supports content generation in English and Chinese and allows for customization through templates, audience-specific tailoring, and in-UI page content editing.

## Table of Contents

- [Features](#features)
- [What is ChatPPT](#what-is-chatppt)
- [Configuration via `.env` File](#configuration-via-env-file)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
  - [Command-Line Interface (CLI)](#command-line-interface-cli)
  - [Streamlit UI](#streamlit-ui)
    - [In-UI Page Content Editing](#in-ui-page-content-editing)
- [Contributing](#contributing)
- [License](#license)

## Features

ChatPPT has been significantly updated with a focus on flexibility and user experience:

*   **Multiple LLM Provider Support**: OpenAI, Ollama, Anthropic, and Groq.
*   **Configuration via `.env` File**: Securely manage API keys, default models, and API base URLs. UI inputs for these are hidden if set in `.env`.
*   **Custom OpenAI API Endpoint**: Support for specifying a custom base URL for OpenAI-compatible APIs.
*   **Audience-Specific Content**: Tailor your presentation's tone and content by specifying the target audience.
*   **In-UI Text-Based Page Editing**: After initial content generation, select individual pages to edit their titles and bullet points directly within the UI and regenerate them with new instructions.
*   **Enhanced UI for Model Selection**: Provider-specific model selection and configuration.
*   **PowerPoint Template Upload**: Users can upload their own `.pptx` file as a template.
*   **Custom Prompt Instructions**: Provide additional global instructions to the LLM.
*   **Streamlined Streamlit Web Interface**: Global configurations are now neatly organized in a sidebar, with the main area dedicated to content viewing and editing.
*   **Command-Line Interface**: Retains a powerful CLI for automation and advanced users.

<!-- TODO: Update screenshots to reflect new UI features, especially the sidebar and page editing form. -->

## What is ChatPPT

ChatPPT is powered by various leading LLMs. It's designed to simplify the creation of presentations by generating slide outlines and content based on your topic, language, and target audience. You can further refine content by editing individual pages within the application.

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
*   **API Keys (if not using Ollama exclusively)**: OpenAI, Anthropic, Groq.
*   **Ollama**: Needs to be installed and running if using the Ollama provider.

## Installation

1.  **Clone the repository.**
2.  **Install Python dependencies:** `pip install -r requirements.txt`
3.  **Set up your `.env` file (Recommended):** `cp sample.env .env` and edit.
4.  **Set up LLM Providers (if not using `.env` for keys).**

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
*   The `--audience "<description>"` argument helps tailor the content.

**CLI Examples:** (Refer to previous README version for detailed examples if needed, structure is similar)

### Streamlit UI

The Streamlit UI provides an easy-to-use interface for all features.

1.  **Start the Streamlit application:**
    ```bash
    streamlit run chatppt_ui.py
    ```
2.  **Open the URL provided by Streamlit in your browser (usually `http://localhost:8501`).**

3.  **Using the UI:**
    *   **Layout**: Global configuration options (like model selection, API keys, topic, language, audience, etc.) are located in a collapsible sidebar on the left. The main area is dedicated to displaying the generated presentation content and the page editing interface.
    *   **Configuration (in Sidebar)**:
        *   **Select Model Provider**: Choose between "openai", "ollama", "anthropic", or "groq".
        *   **API Keys & Base URLs**: Input fields for API keys (OpenAI, Anthropic, Groq) and OpenAI API Base URL will be shown *only if* the corresponding values are not set in your `.env` file. If set in `.env`, a message confirms it's loaded.
        *   **Model Selection**: Configure the model for the chosen provider.
        *   **Topic, Pages, Language, Audience, Custom Instructions, Template**: Set these parameters for your presentation.
        *   Click **"Generate Slide Content"** to generate the initial presentation outline.
    *   **Viewing Content (Main Area)**: Once content is generated, the main area will display the presentation title and an overview of all pages.
    *   **Downloading**: A "Generate and Download PPTX File" button is available to save your presentation.

#### In-UI Page Content Editing

After generating the initial presentation content, you can refine individual pages directly within the UI:

1.  **Select a Page**: Use the dropdown menu (labeled "Select a page to view or edit its details:") above the content outline in the main area to choose a specific page.
2.  **Edit Content**: The selected page's title and bullet points (both titles and descriptions) will appear in editable text fields within a form.
3.  **Provide Instructions**: You can also add page-specific instructions in the "Instructions to refine this page (optional):" text area to guide the LLM's regeneration for that particular page.
4.  **Submit Changes**: Click the **"Prepare Page Update"** button within the form.
5.  **Automatic Regeneration**: The application will then use the LLM to regenerate only that page's content based on your edits and instructions.
6.  **View Updated Content**: The main presentation outline will automatically update to reflect the changes. The downloaded `.pptx` file will include all such modifications.

**Note**: This feature currently supports text-based editing of existing page elements (title and bullet points). Adding or removing bullet points is not yet directly supported through this editing interface (the LLM might do it based on instructions, but there are no UI buttons for it).

<!-- TODO: Update screenshots to reflect new UI features, including the sidebar and the page editing form. -->

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues to improve ChatPPT.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
