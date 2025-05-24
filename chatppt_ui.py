# -*- coding: utf-8 -*-
# pylint: disable=invalid-name  # Allow lowercase for module-level Streamlit UI variables
"""
ChatPPT Streamlit User Interface.

This module provides a web-based user interface for ChatPPT, allowing users to
generate PowerPoint presentations using various LLM providers. It handles user input for
configuration, topic, content generation parameters, template uploads, and page editing.
"""

import json # Used by get_ollama_models for response.json()
import requests
import streamlit as st

from chatppt import ChatPPT
from config import (
    OPENAI_API_KEY, ANTHROPIC_API_KEY, GROQ_API_KEY,
    OPENAI_API_BASE, GROQ_DEFAULT_MODEL
    # GROQ_API_BASE is not directly used in UI logic, only passed to ChatPPT via config
)


def get_ollama_models(ollama_url: str | None) -> list[str]:
    """
    Fetches the list of available models from an Ollama instance.

    Args:
        ollama_url: The base URL of the Ollama API.

    Returns:
        A list of model names available on the Ollama instance.
        Returns an empty list if fetching fails or ollama_url is None.

    Handles:
        requests.exceptions.Timeout: If the request times out.
        requests.exceptions.RequestException: For other network or HTTP errors.
        json.JSONDecodeError: If the response from Ollama is not valid JSON.
        Exception: For any other unexpected errors during the process.
    """
    if not ollama_url:
        return []
    try:
        tags_url = ollama_url.strip('/') + "/api/tags"
        response = requests.get(tags_url, timeout=5)
        response.raise_for_status()
        data = response.json()
        return [model['name'] for model in data.get('models', []) if 'name' in model]
    except requests.exceptions.Timeout:
        st.warning(
            f"Timeout while trying to connect to {tags_url}. "
            "Please check the Ollama URL and ensure Ollama is running."
        )
    except requests.exceptions.RequestException as e:
        st.warning(
            f"Error fetching Ollama models from {tags_url}: {e}. "
            "Please ensure Ollama is running and the URL is correct."
        )
    except json.JSONDecodeError:
        st.warning(
            f"Error parsing JSON response from {tags_url}. "
            "The response was not valid JSON."
        )
    except Exception as e:  # pylint: disable=broad-except
        # Catching general exception here to prevent UI crash for unexpected issues
        st.warning(f"An unexpected error occurred while fetching Ollama models: {e}")
    return []


# --- Session State Initializations ---
if 'ppt_content' not in st.session_state:
    st.session_state.ppt_content = None
if 'selected_page_index' not in st.session_state:
    st.session_state.selected_page_index = None
if 'edited_page_data_for_update' not in st.session_state:
    st.session_state.edited_page_data_for_update = None
if 'current_topic' not in st.session_state:
    st.session_state.current_topic = ""
if 'current_language_code' not in st.session_state:
    st.session_state.current_language_code = "en"

# Session state for provider settings to be used by regeneration
default_provider_settings = {
    'model_provider_for_regen': "openai",
    'selected_model_name_for_regen': "gpt-3.5-turbo",
    'openai_api_key_for_regen': OPENAI_API_KEY,
    'anthropic_api_key_for_regen': ANTHROPIC_API_KEY,
    'groq_api_key_for_regen': GROQ_API_KEY,
    'ollama_url_for_regen': "http://localhost:11434",
    'openai_api_base_for_regen': OPENAI_API_BASE
}
for key, value in default_provider_settings.items():
    if key not in st.session_state:
        st.session_state[key] = value


# --- Language Mapping ---
LANGUAGE_MAP = {"cn": "Chinese", "en": "English"}

# --- Sidebar for Configuration ---
st.sidebar.header("⚙️ Configuration")

# These variables will hold the current UI selections from the sidebar
# and will be used to update session state or pass to ChatPPT
sidebar_openai_api_key_to_use = OPENAI_API_KEY
sidebar_openai_api_base_to_use = OPENAI_API_BASE
sidebar_anthropic_api_key_to_use = ANTHROPIC_API_KEY
sidebar_groq_api_key_to_use = GROQ_API_KEY
sidebar_selected_model_name = None
sidebar_ollama_url = None

model_provider = st.sidebar.selectbox(
    "Select Model Provider",
    ["openai", "ollama", "anthropic", "groq"],
    key="sidebar_model_provider_select"
)

if model_provider == "openai":
    if not OPENAI_API_KEY:
        openai_api_key_input = st.sidebar.text_input(
            "Enter your OpenAI API Key", type="password", key="sidebar_openai_key_input"
        )
        sidebar_openai_api_key_to_use = openai_api_key_input
    else:
        st.sidebar.success("OpenAI API Key loaded from .env.")
    sidebar_selected_model_name = st.sidebar.selectbox(
        "Select OpenAI Model",
        ["gpt-3.5-turbo", "gpt-4", "gpt-4o", "gpt-4-turbo"],
        key="sidebar_openai_model_select"
    )
    if not OPENAI_API_BASE:
        openai_api_base_input = st.sidebar.text_input(
            "OpenAI API Base URL (Optional)",
            placeholder="e.g., http://localhost:8000/v1",
            help="Leave blank for default.",
            key="sidebar_openai_api_base_input"
        )
        sidebar_openai_api_base_to_use = (
            openai_api_base_input if openai_api_base_input and
            openai_api_base_input.strip() else None
        )
    else:
        st.sidebar.info(f"Using OpenAI API Base from .env: {OPENAI_API_BASE}")
    sidebar_anthropic_api_key_to_use = None
    sidebar_groq_api_key_to_use = None
elif model_provider == "ollama":
    sidebar_ollama_url = st.sidebar.text_input(
        "Enter your Ollama URL",
        value=st.session_state.ollama_url_for_regen, # Use session state default
        key="sidebar_ollama_url"
    )
    available_ollama_models = get_ollama_models(sidebar_ollama_url) if sidebar_ollama_url else []
    if available_ollama_models:
        sidebar_selected_model_name = st.sidebar.selectbox(
            "Select Ollama Model", available_ollama_models, key="sidebar_ollama_model_select"
        )
    else:
        sidebar_selected_model_name = st.sidebar.text_input(
            "Enter Ollama Model Name", value="llama3", key="sidebar_ollama_model_text_fallback"
        )
    sidebar_openai_api_key_to_use = None
    sidebar_openai_api_base_to_use = None
    sidebar_anthropic_api_key_to_use = None
    sidebar_groq_api_key_to_use = None
elif model_provider == "anthropic":
    if not ANTHROPIC_API_KEY:
        anthropic_api_key_input = st.sidebar.text_input(
            "Enter your Anthropic API Key", type="password", key="sidebar_anthropic_key_input"
        )
        sidebar_anthropic_api_key_to_use = anthropic_api_key_input
    else:
        st.sidebar.success("Anthropic API Key loaded from .env.")
    sidebar_selected_model_name = st.sidebar.selectbox(
        "Select Anthropic Model",
        ["claude-3-opus-20240229", "claude-3-sonnet-20240229", "claude-3-haiku-20240307"],
        key="sidebar_anthropic_model_select"
    )
    sidebar_openai_api_key_to_use = None
    sidebar_openai_api_base_to_use = None
    sidebar_groq_api_key_to_use = None
elif model_provider == "groq":
    if not GROQ_API_KEY:
        groq_api_key_input = st.sidebar.text_input(
            "Enter your Groq API Key", type="password", key="sidebar_groq_key_input"
        )
        sidebar_groq_api_key_to_use = groq_api_key_input
    else:
        st.sidebar.success("Groq API Key loaded from .env.")
    sidebar_selected_model_name = st.sidebar.text_input(
        "Enter Groq Model Name",
        value=(GROQ_DEFAULT_MODEL or "mixtral-8x7b-32768"),
        key="sidebar_groq_model_name_input"
    )
    sidebar_openai_api_key_to_use = None
    sidebar_openai_api_base_to_use = None
    sidebar_anthropic_api_key_to_use = None

topic_input = st.sidebar.text_input(
    "Enter the topic for the presentation",
    value=st.session_state.current_topic, key="sidebar_topic_input"
)
num_slides_input = st.sidebar.slider(
    "Number of pages", 5, 20, 5, key="sidebar_num_slides_input"
)
language_code_input = st.sidebar.selectbox(
    "Select language", ["en", "cn"],
    index=["en", "cn"].index(st.session_state.current_language_code),
    key="sidebar_language_select"
)

sidebar_audience_options = ["General", "Student", "Software Engineer", "Kids", "Custom"]
sidebar_selected_audience_option = st.sidebar.selectbox(
    "Select Target Audience", sidebar_audience_options, key="sidebar_audience_select"
)
sidebar_final_audience_str = ""
if sidebar_selected_audience_option == "Custom":
    sidebar_final_audience_str = st.sidebar.text_input(
        "Enter Custom Audience Description", key="sidebar_custom_audience_input"
    )
elif sidebar_selected_audience_option != "General":
    sidebar_final_audience_str = sidebar_selected_audience_option
else:
    sidebar_final_audience_str = None

sidebar_custom_prompt_instructions = st.sidebar.text_area(
    "Custom Instructions (Optional)",
    help="Add specific instructions for the AI.",
    key="sidebar_custom_prompt_area"
)
sidebar_template_file = st.sidebar.file_uploader(
    "Upload a PowerPoint template (optional)", type=["pptx"], key="sidebar_template_uploader"
)

if st.sidebar.button("Generate Slide Content", key="sidebar_generate_content_button"):
    is_input_valid = True
    if model_provider == "openai" and not sidebar_openai_api_key_to_use:
        st.error("OpenAI API Key required.")
        is_input_valid = False
    if model_provider == "anthropic" and not sidebar_anthropic_api_key_to_use:
        st.error("Anthropic API Key required.")
        is_input_valid = False
    if model_provider == "groq" and not sidebar_groq_api_key_to_use:
        st.error("Groq API Key required.")
        is_input_valid = False
    if model_provider == "ollama" and \
       (not sidebar_ollama_url or not sidebar_selected_model_name):
        st.error("Ollama URL & Model Name required.")
        is_input_valid = False
    if not topic_input:
        st.error("Topic required.")
        is_input_valid = False
    if not sidebar_selected_model_name and model_provider not in ["ollama"]:
        st.error(f"{model_provider.capitalize()} Model Name required.")
        is_input_valid = False

    if is_input_valid:
        st.session_state.current_topic = topic_input
        st.session_state.current_language_code = language_code_input
        st.session_state.model_provider_for_regen = model_provider
        st.session_state.selected_model_name_for_regen = sidebar_selected_model_name
        st.session_state.openai_api_key_for_regen = sidebar_openai_api_key_to_use
        st.session_state.anthropic_api_key_for_regen = sidebar_anthropic_api_key_to_use
        st.session_state.groq_api_key_for_regen = sidebar_groq_api_key_to_use
        st.session_state.ollama_url_for_regen = sidebar_ollama_url
        st.session_state.openai_api_base_for_regen = (
            sidebar_openai_api_base_to_use if sidebar_openai_api_base_to_use and
            sidebar_openai_api_base_to_use.strip() else None
        )

        with st.spinner("Generating Slide Content..."):
            try:
                # Use sidebar values for initial generation
                chat_ppt_instance = ChatPPT(
                    model_provider=model_provider,
                    api_key=sidebar_openai_api_key_to_use,
                    model_name=sidebar_selected_model_name,
                    ollama_url=sidebar_ollama_url,
                    anthropic_api_key=sidebar_anthropic_api_key_to_use,
                    openai_api_base=(sidebar_openai_api_base_to_use if
                                     sidebar_openai_api_base_to_use and
                                     sidebar_openai_api_base_to_use.strip() else None),
                    groq_api_key=sidebar_groq_api_key_to_use
                )
                generated_content = chat_ppt_instance.chatppt(
                    topic_input, num_slides_input, language_code_input,
                    custom_prompt_instructions=sidebar_custom_prompt_instructions,
                    audience=sidebar_final_audience_str
                )
                st.session_state.ppt_content = generated_content
                st.session_state.selected_page_index = None
                st.session_state.edited_page_data_for_update = None
                st.success("Slide content generated successfully!")
                st.rerun()
            # pylint: disable=broad-except # Catch all backend errors for UI display
            except Exception as e:
                st.error(f"Error generating slide content: {e}")
                st.session_state.ppt_content = None

# --- Process Page Update (remains at the top level of script execution) ---
if st.session_state.edited_page_data_for_update:
    edited_data = st.session_state.edited_page_data_for_update
    page_idx_to_update = edited_data["page_index"]

    with st.spinner(f"Updating content for page {page_idx_to_update + 1}..."):
        try:
            chat_ppt_instance_for_update = ChatPPT(
                model_provider=st.session_state.model_provider_for_regen,
                api_key=st.session_state.openai_api_key_for_regen,
                model_name=st.session_state.selected_model_name_for_regen,
                ollama_url=st.session_state.ollama_url_for_regen,
                anthropic_api_key=st.session_state.anthropic_api_key_for_regen,
                openai_api_base=st.session_state.openai_api_base_for_regen,
                groq_api_key=st.session_state.groq_api_key_for_regen
            )
            page_json_to_edit = {"title": edited_data["title"], "content": edited_data["content"]}
            language_str = LANGUAGE_MAP.get(st.session_state.current_language_code, "English")

            new_page_content_json = chat_ppt_instance_for_update.regenerate_single_page(
                original_topic=st.session_state.current_topic,
                original_language=language_str,
                page_data_to_edit=page_json_to_edit,
                new_instructions_for_page=edited_data["instructions"]
            )
            if st.session_state.ppt_content and \
               'pages' in st.session_state.ppt_content and \
               0 <= page_idx_to_update < len(st.session_state.ppt_content['pages']):
                st.session_state.ppt_content['pages'][page_idx_to_update] = new_page_content_json
                st.success(f"Page {page_idx_to_update + 1} updated successfully!")
            else:
                st.error("Failed to update page: Content structure invalid or page index out of bounds.")
            st.session_state.edited_page_data_for_update = None
            st.rerun()
        # pylint: disable=broad-except # Catch all backend errors for UI display
        except Exception as e:
            st.error(f"Error updating page {page_idx_to_update + 1}: {e}")
            st.session_state.edited_page_data_for_update = None

# --- Main Area for Display and Editing ---
st.title("ChatPPT Generator")
st.write("Generate a slide presentation using AI!")

if st.session_state.ppt_content:
    st.markdown("---")
    st.header(
        f"Presentation Outline: {st.session_state.ppt_content.get('title', 'Untitled Presentation')}"
    )
    pages_data = st.session_state.ppt_content.get('pages', [])
    page_titles = [f"Page {i+1}: {page.get('title', 'Untitled')}" for i, page in enumerate(pages_data)]

    if page_titles:
        selection_options = ["View All / Select a Page"] + page_titles
        selectbox_current_index = 0
        if st.session_state.selected_page_index is not None:
            try:
                if 0 <= st.session_state.selected_page_index < len(page_titles):
                    selectbox_current_index = selection_options.index(
                        page_titles[st.session_state.selected_page_index]
                    )
            except ValueError:
                st.session_state.selected_page_index = None

        selected_page_display_name = st.selectbox(
            "Select a page to view or edit its details:",
            selection_options, index=selectbox_current_index, key="main_page_selector"
        )
        if selected_page_display_name != "View All / Select a Page":
            for i, title_option in enumerate(page_titles):
                if title_option == selected_page_display_name:
                    st.session_state.selected_page_index = i
                    break
        else:
            st.session_state.selected_page_index = None
    else:
        st.write("No pages generated in the content.")

    if st.session_state.selected_page_index is not None:
        page_idx = st.session_state.selected_page_index
        if 0 <= page_idx < len(pages_data):
            current_page_data = pages_data[page_idx]
            st.subheader(f"Editing Page {page_idx + 1}")
            with st.form(key=f"main_edit_page_form_{page_idx}"):
                edited_page_title = st.text_input(
                    "Page Title", value=current_page_data.get('title', ''),
                    key=f"main_title_edit_{page_idx}"
                )
                edited_bullets_inputs_values = []
                for bullet_i, bullet_content in enumerate(current_page_data.get('content', [])):
                    st.markdown(f"--- Editing Bullet {bullet_i+1} ---")
                    bullet_title_val = st.text_input(
                        f"Bullet {bullet_i+1} Title", value=bullet_content.get('title', ''),
                        key=f"main_bullet_title_edit_{page_idx}_{bullet_i}"
                    )
                    bullet_desc_val = st.text_area(
                        f"Bullet {bullet_i+1} Description",
                        value=bullet_content.get('description', ''),
                        key=f"main_bullet_desc_edit_{page_idx}_{bullet_i}"
                    )
                    edited_bullets_inputs_values.append(
                        {"title": bullet_title_val, "description": bullet_desc_val}
                    )
                page_specific_instructions_input = st.text_area(
                    "Instructions to refine this page (optional):",
                    key=f"main_page_instr_edit_{page_idx}"
                )
                submit_page_edit_button = st.form_submit_button("Prepare Page Update")
                if submit_page_edit_button:
                    st.session_state.edited_page_data_for_update = {
                        "page_index": page_idx, "title": edited_page_title,
                        "content": edited_bullets_inputs_values,
                        "instructions": page_specific_instructions_input
                    }
                    st.success(f"Page {page_idx+1} data prepared. Rerunning to apply update...")
                    st.rerun()
    else:
        st.subheader("Full Presentation Outline:")
        if not pages_data:
            st.write("No pages to display.")
        for i, page in enumerate(pages_data):
            st.markdown(f"### Page {i+1}: {page.get('title', 'Untitled')}")
            page_content = page.get('content', [])
            if not page_content:
                st.markdown("_No bullets for this page._")
            for bullet_idx, bullet in enumerate(page_content):
                st.markdown(f"**Bullet {bullet_idx + 1} Title:** {bullet.get('title', '')}")
                st.markdown(
                    f"**Bullet {bullet_idx + 1} Description:** {bullet.get('description', '')}"
                )
                st.markdown("---")
            st.markdown("----")

    st.markdown("---")
    st.subheader("Download Presentation")
    if st.button("Generate and Download PPTX File", key="main_download_button"):
        if st.session_state.ppt_content:
            with st.spinner("Generating PPTX file..."):
                try:
                    # Use session state for regeneration settings when downloading
                    current_openai_api_base = st.session_state.openai_api_base_for_regen
                    dl_chat_ppt_instance = ChatPPT(
                        model_provider=st.session_state.model_provider_for_regen,
                        api_key=st.session_state.openai_api_key_for_regen,
                        model_name=st.session_state.selected_model_name_for_regen,
                        ollama_url=st.session_state.ollama_url_for_regen,
                        anthropic_api_key=st.session_state.anthropic_api_key_for_regen,
                        openai_api_base=current_openai_api_base,
                        groq_api_key=st.session_state.groq_api_key_for_regen
                    )
                    ppt_file_name = dl_chat_ppt_instance.generate_ppt(
                        st.session_state.ppt_content, template=sidebar_template_file
                    )
                    with open(ppt_file_name, "rb") as f:
                        st.download_button(
                            label="Click to Download PPTX", data=f, file_name=ppt_file_name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation" # pylint: disable=line-too-long
                        )
                # pylint: disable=broad-except # Catch all backend errors for UI display
                except Exception as e:
                    st.error(f"Error generating PPTX file: {e}")
        else:
            st.warning("No presentation content available to download. Please generate content first.")
