import streamlit as st
from chatppt import ChatPPT
import requests
import json

# Helper function to get Ollama models
def get_ollama_models(ollama_url):
    if not ollama_url:
        return []
    try:
        tags_url = ollama_url.strip('/') + "/api/tags"
        response = requests.get(tags_url, timeout=5) 
        response.raise_for_status()
        data = response.json()
        models = [model['name'] for model in data.get('models', []) if 'name' in model]
        return models
    except requests.exceptions.Timeout:
        st.warning(f"Timeout while trying to connect to {tags_url}. Please check the Ollama URL and ensure Ollama is running.")
        return []
    except requests.exceptions.RequestException as e:
        st.warning(f"Error fetching Ollama models from {tags_url}: {e}. Please ensure Ollama is running and the URL is correct.")
        return []
    except json.JSONDecodeError:
        st.warning(f"Error parsing JSON response from {tags_url}. The response was not valid JSON.")
        return []
    except Exception as e:
        st.warning(f"An unexpected error occurred while fetching Ollama models: {e}")
        return []

# Set up the Streamlit interface
st.title("ChatPPT Generator")
st.write("Generate a slide presentation using AI!")

# Initialize variables to None
openai_model_name = None
anthropic_model_name = None
ollama_model_name = None
openai_api_key = None
anthropic_api_key = None
ollama_url = None
selected_model_name = None
template_file = None 
custom_prompt_instructions = None
openai_api_base = None # Initialize openai_api_base

# User selects the Model Provider
model_provider = st.selectbox("Select Model Provider", ["openai", "ollama", "anthropic"])

# Depending on the AI model, different inputs are required
if model_provider == "openai":
    openai_api_key = st.text_input("Enter your OpenAI API Key", key="openai_key")
    openai_model_name = st.selectbox(
        "Select OpenAI Model", 
        ["gpt-3.5-turbo", "gpt-4", "gpt-4o", "gpt-4-turbo"], 
        key="openai_model_select"
    )
    openai_api_base = st.text_input( # Added OpenAI API Base URL input
        "OpenAI API Base URL (Optional)",
        placeholder="e.g., http://localhost:8000/v1",
        help="Leave blank to use the default OpenAI API URL.",
        key="openai_api_base_url"
    )
    selected_model_name = openai_model_name
    # Ensure other provider specific params are None
    ollama_url = None 
    anthropic_api_key = None

elif model_provider == "ollama":
    ollama_url = st.text_input("Enter your Ollama URL", value="http://localhost:11434", key="ollama_url")
    
    available_ollama_models = []
    if ollama_url:
        available_ollama_models = get_ollama_models(ollama_url)

    if available_ollama_models:
        ollama_model_name = st.selectbox(
            "Select Ollama Model", 
            available_ollama_models, 
            key="ollama_model_select"
        )
    else:
        ollama_model_name = st.text_input(
            "Enter Ollama Model Name (or specify valid URL above to auto-detect)", 
            value="llama3", 
            key="ollama_model_text_fallback"
        )
    selected_model_name = ollama_model_name
    # Ensure other provider specific params are None
    openai_api_key = None
    anthropic_api_key = None
    openai_api_base = None # Ensure openai_api_base is None for non-OpenAI providers

elif model_provider == "anthropic":
    anthropic_api_key = st.text_input("Enter your Anthropic API Key", key="anthropic_key")
    anthropic_model_name = st.selectbox(
        "Select Anthropic Model", 
        ["claude-3-opus-20240229", "claude-3-sonnet-20240229", "claude-3-haiku-20240307"], 
        key="anthropic_model_select"
    )
    selected_model_name = anthropic_model_name
    # Ensure other provider specific params are None
    openai_api_key = None
    ollama_url = None
    openai_api_base = None # Ensure openai_api_base is None for non-OpenAI providers


# User inputs for the presentation
topic = st.text_input("Enter the topic for the presentation")
num_slides = st.slider("Number of pages", 5, 20, 5) 
language = st.selectbox("Select language", ["en", "cn"])

# Text area for custom prompt instructions
custom_prompt_instructions = st.text_area(
    "Custom Instructions (Optional)", 
    help="Add any specific instructions, context, or constraints for the AI. This will be appended to the main prompt.",
    key="custom_prompt_area"
)

# File uploader for PowerPoint template
template_file = st.file_uploader("Upload a PowerPoint template (optional)", type=["pptx"], key="template_uploader")

# Button to generate the Slide
generate_button = st.button("Generate Slide", disabled=False)

# If the button is clicked, generate the Slide
if generate_button:
    if model_provider == "openai" and not openai_api_key:
        st.error("OpenAI API Key is required.")
    elif model_provider == "ollama" and (not ollama_url or not selected_model_name):
        st.error("Ollama URL and Model Name are required.")
    elif model_provider == "anthropic" and not anthropic_api_key:
        st.error("Anthropic API Key is required.")
    elif not topic:
        st.error("Topic for the presentation is required.")
    elif not selected_model_name and model_provider != "ollama": 
         st.error(f"{model_provider.capitalize()} Model Name is required.")
    else:
        with st.spinner("Generating Slide..."):
            try:
                # If openai_api_base is an empty string, treat it as None
                current_openai_api_base = openai_api_base if openai_api_base and openai_api_base.strip() else None

                chat_ppt = ChatPPT(
                    model_provider=model_provider, 
                    api_key=openai_api_key, 
                    model_name=selected_model_name, 
                    ollama_url=ollama_url, 
                    anthropic_api_key=anthropic_api_key,
                    openai_api_base=current_openai_api_base # Pass the potentially None-ified base URL
                )
                ppt_content = chat_ppt.chatppt(
                    topic, 
                    num_slides, 
                    language, 
                    custom_prompt_instructions=custom_prompt_instructions
                )
            except Exception as e:
                st.error(f"Error generating Slide content: {e}")
                st.stop()

            if ppt_content: 
                title = ppt_content.get("title", "")
                st.title(f"Title: {title}")
                slides = ppt_content.get("pages", [])
                st.subheader(f"Your slide has {len(slides)} pages:")

                for index, slide in enumerate(slides):
                    st.markdown(f"- Slide {index+1}: {slide.get('title','')}")

                try:
                    ppt_file_name = chat_ppt.generate_ppt(ppt_content, template=template_file) 
                except Exception as e:
                    st.error(f"Error generating Slide file: {e}")
                    st.stop()

                st.write("Slide generated!")
                st.write("Download your Slide:")
                with open(ppt_file_name, "rb") as f:
                    st.download_button("Download file", f, file_name=ppt_file_name)
            else:
                st.error("Failed to generate PPT content. Please check the inputs and try again.")
