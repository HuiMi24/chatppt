import streamlit as st
from chatppt import ChatPPT

# Set up the Streamlit interface
st.title("ChatPPT Generator")
st.write("Generate a slide presentation using AI!")

# User selects the AI model
ai_model = st.selectbox("Select AI Model", ["openai", "ollama"])

# Depending on the AI model, different inputs are required
if ai_model == "openai":
    api_key = st.text_input("Enter your OpenAI API Key")
    ollama_url = None
    ollama_model = None
elif ai_model == "ollama":
    ollama_url = st.text_input("Enter your Ollama URL", "http://localhost:11434")
    ollama_model = st.text_input("Enter your Ollama Model", "llama3")
    api_key = None

# User inputs for the presentation
topic = st.text_input("Enter the topic for the presentation")
num_slides = st.slider("Number of slides", 5, 20, 5)
language = st.selectbox("Select language", ["en", "cn"])

# Button to generate the Slide
generate_button = st.button("Generate Slide", disabled=False)

# If the button is clicked, generate the Slide
if generate_button:
    with st.spinner("Generating Slide..."):
        chat_ppt = ChatPPT(ai_model, api_key, ollama_url, ollama_model)
        try:
            ppt_content = chat_ppt.chatppt(topic, num_slides, language)
        except Exception as e:
            st.error(f"Error generating Slide: {e}")
            st.stop()

        # Display the title and number of slides
        title = ppt_content.get("title", "")
        st.title(f"Title: {title}")
        slides = ppt_content.get("pages", [])
        st.subheader(f"Your slide has {len(slides)} pages:")

        # Display the title of each slide
        for index, slide in enumerate(slides):
            st.markdown(f"- Slide {index+1}: {slide.get('title','')}")

        # Generate the Slide file
        try:
            ppt_file_name = chat_ppt.generate_ppt(ppt_content)
        except Exception as e:
            st.error(f"Error generating Slide: {e}")
            st.stop()

        # Provide a download link for the Slide
        st.write("Slide generated!")
        st.write("Download your Slide:")
        with open(ppt_file_name, "rb") as f:
            st.download_button("Download file", f, ppt_file_name)
