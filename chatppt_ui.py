import streamlit as st
from chatppt import ChatPPT
import requests
import json
# Import configurations from config.py
from config import (
    OPENAI_API_KEY, ANTHROPIC_API_KEY, GROQ_API_KEY,
    OPENAI_API_BASE, GROQ_API_BASE, GROQ_DEFAULT_MODEL
)

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

# Initialize variables that will hold the actual values passed to ChatPPT
# These are resolved from UI input or .env config
openai_api_key_to_use = OPENAI_API_KEY
openai_api_base_to_use = OPENAI_API_BASE
anthropic_api_key_to_use = ANTHROPIC_API_KEY
groq_api_key_to_use = GROQ_API_KEY
# Groq base URL is handled by chatppt.py using GROQ_API_BASE from config or a hardcoded default

# Initialize UI-specific state variables
selected_model_name = None
ollama_url = None # Ollama URL is always taken from UI input
template_file = None 
custom_prompt_instructions = None
final_audience_str = ""

# User selects the Model Provider
model_provider = st.selectbox("Select Model Provider", ["openai", "ollama", "anthropic", "groq"])

# Depending on the AI model, different inputs are required
if model_provider == "openai":
    if not OPENAI_API_KEY:
        openai_api_key_input = st.text_input("Enter your OpenAI API Key", type="password", key="openai_key_input")
        openai_api_key_to_use = openai_api_key_input
    else:
        st.success("OpenAI API Key loaded from .env file.")

    selected_model_name = st.selectbox(
        "Select OpenAI Model", 
        ["gpt-3.5-turbo", "gpt-4", "gpt-4o", "gpt-4-turbo"], 
        key="openai_model_select"
    )
    
    if not OPENAI_API_BASE:
        openai_api_base_input = st.text_input(
            "OpenAI API Base URL (Optional)",
            placeholder="e.g., http://localhost:8000/v1",
            help="Leave blank to use the default OpenAI API URL.",
            key="openai_api_base_input"
        )
        openai_api_base_to_use = openai_api_base_input if openai_api_base_input and openai_api_base_input.strip() else None
    else:
        st.info(f"Using OpenAI API Base URL from .env: {OPENAI_API_BASE}")
    
    # Ensure other provider keys are None for this branch
    anthropic_api_key_to_use = None
    groq_api_key_to_use = None

elif model_provider == "ollama":
    ollama_url = st.text_input("Enter your Ollama URL", value="http://localhost:11434", key="ollama_url")
    
    available_ollama_models = []
    if ollama_url:
        available_ollama_models = get_ollama_models(ollama_url)

    if available_ollama_models:
        selected_model_name = st.selectbox(
            "Select Ollama Model", 
            available_ollama_models, 
            key="ollama_model_select"
        )
    else:
        selected_model_name = st.text_input(
            "Enter Ollama Model Name (or specify valid URL above to auto-detect)", 
            value="llama3", 
            key="ollama_model_text_fallback"
        )
    # Ensure other provider keys/bases are None
    openai_api_key_to_use = None
    openai_api_base_to_use = None
    anthropic_api_key_to_use = None
    groq_api_key_to_use = None

elif model_provider == "anthropic":
    if not ANTHROPIC_API_KEY:
        anthropic_api_key_input = st.text_input("Enter your Anthropic API Key", type="password", key="anthropic_key_input")
        anthropic_api_key_to_use = anthropic_api_key_input
    else:
        st.success("Anthropic API Key loaded from .env file.")
        
    selected_model_name = st.selectbox(
        "Select Anthropic Model", 
        ["claude-3-opus-20240229", "claude-3-sonnet-20240229", "claude-3-haiku-20240307"], 
        key="anthropic_model_select"
    )
    # Ensure other provider keys/bases are None
    openai_api_key_to_use = None
    openai_api_base_to_use = None
    groq_api_key_to_use = None

elif model_provider == "groq":
    if not GROQ_API_KEY:
        groq_api_key_input = st.text_input("Enter your Groq API Key", type="password", key="groq_key_input")
        groq_api_key_to_use = groq_api_key_input
    else:
        st.success("Groq API Key loaded from .env file.")

    selected_model_name = st.text_input(
        "Enter Groq Model Name", 
        value=(GROQ_DEFAULT_MODEL or "mixtral-8x7b-32768"), 
        help="e.g., mixtral-8x7b-32768, llama3-70b-8192",
        key="groq_model_name_input"
    )
    # Ensure other provider keys/bases are None
    openai_api_key_to_use = None
    openai_api_base_to_use = None
    anthropic_api_key_to_use = None
    # Groq API base is handled by chatppt.py using config.GROQ_API_BASE or hardcoded default

# User inputs for the presentation
topic = st.text_input("Enter the topic for the presentation")
num_slides = st.slider("Number of pages", 5, 20, 5) 
language = st.selectbox("Select language", ["en", "cn"])

# Audience Selection
audience_options = ["General", "Student", "Software Engineer", "Kids", "Custom"]
selected_audience_option = st.selectbox("Select Target Audience", audience_options, key="audience_select")

if selected_audience_option == "Custom":
    final_audience_str = st.text_input("Enter Custom Audience Description", key="custom_audience_input")
elif selected_audience_option != "General":
    final_audience_str = selected_audience_option
else:
    final_audience_str = None # Explicitly None for "General" or if not set

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

if generate_button:
    # Validation logic using the resolved *_to_use variables
    if model_provider == "openai" and not openai_api_key_to_use:
        st.error("OpenAI API Key is required. Please enter it or set OPENAI_API_KEY in your .env file.")
    elif model_provider == "anthropic" and not anthropic_api_key_to_use:
        st.error("Anthropic API Key is required. Please enter it or set ANTHROPIC_API_KEY in your .env file.")
    elif model_provider == "groq" and not groq_api_key_to_use:
        st.error("Groq API Key is required. Please enter it or set GROQ_API_KEY in your .env file.")
    elif model_provider == "ollama" and (not ollama_url or not selected_model_name):
        st.error("Ollama URL and Model Name are required.")
    elif not topic:
        st.error("Topic for the presentation is required.")
    elif not selected_model_name and model_provider != "ollama": 
         st.error(f"{model_provider.capitalize()} Model Name is required.")
    else:
        with st.spinner("Generating Slide..."):
            try:
                # Ensure that an empty string for openai_api_base_to_use is treated as None
                current_openai_api_base_for_call = openai_api_base_to_use if openai_api_base_to_use and openai_api_base_to_use.strip() else None
                
                chat_ppt = ChatPPT(
                    model_provider=model_provider, 
                    api_key=openai_api_key_to_use, # For OpenAI
                    model_name=selected_model_name, 
                    ollama_url=ollama_url, 
                    anthropic_api_key=anthropic_api_key_to_use, 
                    openai_api_base=current_openai_api_base_for_call,
                    groq_api_key=groq_api_key_to_use # For Groq
                )
                ppt_content = chat_ppt.chatppt(
                    topic, 
                    num_slides, 
                    language, 
                    custom_prompt_instructions=custom_prompt_instructions,
                    audience=final_audience_str # Pass the audience string
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
