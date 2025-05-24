import re
import openai
import json
import argparse
import time
import random
import datetime
import sys 
import anthropic 
from ollama import Client
from pptx import Presentation
from config import OPENAI_API_KEY, ANTHROPIC_API_KEY, OPENAI_API_BASE, \
                   GROQ_API_KEY, GROQ_API_BASE, GROQ_DEFAULT_MODEL

class ChatPPT:
    def __init__(self, model_provider, api_key, model_name, ollama_url=None, 
                 anthropic_api_key=None, openai_api_base=None, groq_api_key=None): 
        self.model_provider = model_provider
        self.api_key = api_key 
        self.model_name = model_name
        self.ollama_url = ollama_url
        self.anthropic_api_key = anthropic_api_key 
        self.openai_api_base = openai_api_base 
        self.groq_api_key = groq_api_key 

    @staticmethod
    def robot_print(text):
        for char in text:
            print(char, end="", flush=True)
            time.sleep(random.randrange(1, 2) / 1000.0)
        print("\r")
        return text

    def chatppt(self, topic, pages, language, custom_prompt_instructions=None, audience=None):
        language_map = {"cn": "Chinese", "en": "English"}
        language_str = language_map.get(language, language) # Use language directly if not in map
        self.robot_print(f"I'm working hard to generate your PPT about {topic}.")
        self.robot_print("It may takes about a few minutes.")
        self.robot_print(f"Your PPT will be generated in {language_str}")
        output_format = self._get_output_format()
        messages = self._get_messages(topic, pages, language_str, output_format, custom_prompt_instructions, audience)
        content = self._get_content(messages)
        return self._parse_content(content)

    def _get_output_format(self): # This is for the whole presentation
        return {
            "title": "example title",
            "pages": [
                {
                    "title": "title for page 1",
                    "content": [
                        { "title": "title for bullet 1", "description": "detail for bullet 1" },
                        { "title": "title for bullet 2", "description": "detail for bullet 2" },
                        { "title": "title for bullet 3", "description": "detail for bullet 3" },
                    ],
                },
                {
                    "title": "title for page 2",
                    "content": [
                        { "title": "title for bullet 1", "description": "detail for bullet 1" },
                        { "title": "title for bullet 2", "description": "detail for bullet 2" },
                    ],
                },
            ],
        }
    
    def _get_single_page_output_format_example(self): # Example format for a single page
        return {
            "title": "page title",
            "content": [
                {"title": "bullet title 1", "description": "bullet description 1"},
                {"title": "bullet title 2", "description": "bullet description 2"}
            ]
        }

    def _get_messages(self, topic, pages, language_str, output_format, custom_prompt_instructions=None, audience=None): 
        prompt_lines = [
            f"I am preparing a presentation on {topic}.",
            f"Please assist in generating an outline in JSON format, adhering to the specified format {json.dumps(output_format)}.",
            f"The presentation should span {pages} pages, with as many bullet points as possible.",
            f"The content should be returned in {language_str}.",
            "You must add content for each slide.",
            "For each slide, you must add at least 4 bullet.",
            "Please ensure the output is valid JSON match the RFC-8295 specification."
        ]
        if audience and audience.strip():
            prompt_lines.append(f"The target audience for this presentation is: {audience.strip()}. Please tailor the content accordingly.")
        if custom_prompt_instructions and custom_prompt_instructions.strip():
            prompt_lines.append(f"\nAdditional Custom Instructions:\n{custom_prompt_instructions.strip()}")
        full_prompt = "\n".join(prompt_lines)
        return [{"role": "user", "content": full_prompt}]

    def regenerate_single_page(self, original_topic, original_language, page_data_to_edit, new_instructions_for_page=None):
        page_content_str = json.dumps(page_data_to_edit, ensure_ascii=False, indent=2)
        single_page_format_example = json.dumps(self._get_single_page_output_format_example(), ensure_ascii=False, indent=2)

        prompt_lines = [
            f"I am working on a presentation about '{original_topic}'.",
            f"I need to refine a single slide. The language for the output should be {original_language}.",
            f"The current content of the slide is (in JSON format):\n{page_content_str}\n"
        ]

        if new_instructions_for_page and new_instructions_for_page.strip():
            prompt_lines.append(f"Please refine this slide based on the following instructions: {new_instructions_for_page.strip()}\n")
        else:
            prompt_lines.append("Please review and refine the content of this slide, ensuring clarity, accuracy, and completeness based on its current title and bullet points.\n")

        prompt_lines.extend([
            "You MUST return ONLY the complete JSON object for this single refined slide, including its 'title' and 'content' (with 'title' and 'description' for each bullet point).",
            "Do not return any other text, explanations, or markdown formatting around the JSON.",
            f"The JSON structure for the page should be exactly like this example (though with different actual content): {single_page_format_example}."
        ])
        
        prompt = "\n".join(prompt_lines)
        messages = [{"role": "user", "content": prompt}]

        self.robot_print(f"Refining page: '{page_data_to_edit.get('title', 'Untitled Page')}'...")
        llm_response_str = self._get_content(messages)

        try:
            match = re.search(r"(\{.*\})", llm_response_str, re.DOTALL)
            if match:
                json_str = match.group(0)
            else:
                json_str = llm_response_str 
            
            new_page_json = json.loads(json_str.strip())

            if not isinstance(new_page_json, dict) or \
               "title" not in new_page_json or \
               "content" not in new_page_json or \
               not isinstance(new_page_json["content"], list):
                raise ValueError("LLM returned JSON but not in the expected page format (missing title, content, or content is not a list).")
            
            for item in new_page_json["content"]:
                if not isinstance(item, dict) or "title" not in item or "description" not in item:
                    raise ValueError("LLM returned JSON with malformed bullet points (missing title or description in a bullet).")
            
            self.robot_print(f"Page '{new_page_json.get('title', 'Untitled Page')}' refined successfully.")
            return new_page_json
        except Exception as e:
            error_message = f"Error parsing regenerated page content: {e}. Raw response snippet: {llm_response_str[:500]}..."
            print(f"\n{error_message}") # Print for CLI/log visibility
            raise Exception(error_message) from e


    def _get_content(self, messages):
        # ... (existing _get_content method remains unchanged)
        if self.model_provider == "openai":
            if not self.api_key:
                raise Exception("OpenAI API key is required but not provided/configured.")
            openai.api_key = self.api_key 
            
            original_api_base = openai.api_base 
            api_base_changed = False
            if self.openai_api_base and self.openai_api_base.strip():
                if openai.api_base != self.openai_api_base:
                    openai.api_base = self.openai_api_base
                    api_base_changed = True
            elif openai.api_base is None: 
                openai.api_base = "https://api.openai.com/v1"

            try:
                completion = openai.ChatCompletion.create(model=self.model_name, messages=messages)
                return completion.choices[0].message.content
            finally:
                if api_base_changed or (self.openai_api_base and self.openai_api_base.strip() and original_api_base is None):
                     openai.api_base = original_api_base
                elif original_api_base is None and not (self.openai_api_base and self.openai_api_base.strip()):
                    if openai.api_base == "https://api.openai.com/v1" and not (self.openai_api_base and self.openai_api_base.strip()) and original_api_base is None:
                         openai.api_base = None
        
        elif self.model_provider == "groq":
            if not self.groq_api_key:
                raise Exception("Groq API Key is required for 'groq' provider.")
            
            original_openai_api_key = openai.api_key
            original_openai_api_base = openai.api_base
            
            openai.api_key = self.groq_api_key
            current_groq_api_base = GROQ_API_BASE if GROQ_API_BASE and GROQ_API_BASE.strip() else "https://api.groq.com/openai/v1"
            openai.api_base = current_groq_api_base
            
            try:
                completion = openai.ChatCompletion.create(model=self.model_name, messages=messages)
                return completion.choices[0].message.content
            finally:
                openai.api_key = original_openai_api_key
                openai.api_base = original_openai_api_base

        elif self.model_provider == "ollama":
            if self.ollama_url is None:
                raise Exception("Ollama URL is required when model_provider is 'ollama'")
            client = Client(host=self.ollama_url)
            response = client.chat(model=self.model_name, messages=messages)
            return response["message"]["content"]
        elif self.model_provider == "anthropic": 
            if not self.anthropic_api_key:
                raise Exception("Anthropic API key is required but not provided/configured.")
            client = anthropic.Anthropic(api_key=self.anthropic_api_key) 
            response = client.messages.create(model=self.model_name, max_tokens=2048, messages=messages)
            return response.content[0].text


    def _parse_content(self, content): # This is for the whole presentation JSON
        try:
            match = re.search(r"(\{.*\})", content, re.DOTALL)
            if match: content = match.groups()[0]
            parsed_json = json.loads(content.strip())
            # Basic validation for full presentation structure
            if not isinstance(parsed_json, dict) or "title" not in parsed_json or "pages" not in parsed_json or not isinstance(parsed_json["pages"], list):
                raise ValueError("LLM response for full presentation is not in the expected root format (missing title or pages list).")
            return parsed_json
        except Exception as e:
            print(f"The response is not a valid JSON format for full presentation: {e}\nRaw content from LLM: {content}")
            raise Exception("The LLM return invalid result for full presentation, please retry later..")

    def generate_ppt(self, content, template=None):
        ppt = Presentation(template) if template else Presentation()
        self.create_slides(ppt, content)
        ppt_name = self._get_ppt_name(content)
        ppt.save(ppt_name)
        self.robot_print(f"Generate done, enjoy!\nYour PPT: {ppt_name}")
        return ppt_name

    def create_slides(self, presentation, content):
        self.create_title_slide(presentation, content)
        pages = content.get("pages", [])
        self.robot_print(f"Your PPT has {len(pages)} pages.")
        for index, page in enumerate(pages): self.create_content_slide(presentation, page, index)

    def create_title_slide(self, presentation, content):
        title_slide_layout = presentation.slide_layouts[0]
        title_slide = presentation.slides.add_slide(title_slide_layout)
        title_slide.shapes.title.text = content.get("title", "")
        title_slide.placeholders[1].text = "Generated by ChatPPT"

    def create_content_slide(self, presentation, page, index):
        page_title = page.get("title", "")
        self.robot_print(f"Page {index+1}: {page_title}")
        bullet_slide_layout = presentation.slide_layouts[1]
        bullet_slide = presentation.slides.add_slide(bullet_slide_layout)
        bullet_slide.shapes.title.text = page_title
        self.add_bullets_to_slide(bullet_slide, page)

    def add_bullets_to_slide(self, slide, page):
        body_shape = slide.shapes.placeholders[1]
        for bullet in page.get("content", []): self.add_bullet(body_shape, bullet)

    def add_bullet(self, body_shape, bullet):
        paragraph = body_shape.text_frame.add_paragraph()
        paragraph.text = bullet.get("title", "")
        paragraph.level = 1
        paragraph = body_shape.text_frame.add_paragraph()
        paragraph.text = bullet.get("description", "")
        paragraph.level = 2
        
    def _get_ppt_name(self, content):
        ppt_name = re.sub(r'[\\/:*?"<>|]', "", content.get("title", ""))
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        return f"{ppt_name}_{timestamp}.pptx"

def args_parser():
    parser = argparse.ArgumentParser(description="I am your PPT assistant, I can help to you generate PPT.")
    parser.add_argument("-m", "--model_provider", choices=["openai", "ollama", "anthropic", "groq"], default="openai", help="Select the model provider")
    parser.add_argument("-n", "--model_name", type=str, default=None, help="Specify the model name to use (e.g., gpt-3.5-turbo, llama3, claude-3-opus, mixtral-8x7b-32768). Required unless provider has a default (e.g. Groq).")
    parser.add_argument("-t", "--topic", type=str, required=True, help="Your topic name")
    parser.add_argument("--audience", type=str, default=None, help="Specify the target audience for the presentation (e.g., 'students', 'technical experts').") 
    
    parser.add_argument("-k", "--api_key", type=str, default=None, help="Your OpenAI API key or file path. Overrides .env.")
    parser.add_argument("--anthropic_api_key", type=str, default=None, help="Your Anthropic API key or file path. Overrides .env.")
    parser.add_argument("--groq_api_key", type=str, default=None, help="Your Groq API key. Overrides .env.")
    parser.add_argument("--openai_api_base", type=str, default=None, help="Custom base URL for OpenAI API. Overrides .env.")
    
    parser.add_argument("-u", "--ollama_url", type=str, default="http://localhost:11434", help="Your ollama url")
    parser.add_argument("-p", "--pages", type=int, default=5, help="How many slides to generate")
    parser.add_argument("-l", "--language", choices=["cn", "en"], default="en", help="Output language")
    
    args = parser.parse_args()
    return args

def main():
    args = args_parser()

    resolved_openai_api_key = args.api_key if args.api_key is not None else OPENAI_API_KEY
    resolved_anthropic_api_key = args.anthropic_api_key if args.anthropic_api_key is not None else ANTHROPIC_API_KEY
    resolved_groq_api_key = args.groq_api_key if args.groq_api_key is not None else GROQ_API_KEY
    resolved_openai_api_base = args.openai_api_base if args.openai_api_base is not None else OPENAI_API_BASE
    
    model_name_to_pass = args.model_name
    if args.model_provider == "groq" and not model_name_to_pass:
        model_name_to_pass = GROQ_DEFAULT_MODEL
        if not model_name_to_pass:
            model_name_to_pass = "mixtral-8x7b-32768" 
            print("Warning: No model name provided for Groq via --model_name and no GROQ_DEFAULT_MODEL in .env. Using 'mixtral-8x7b-32768'.")
    
    if not model_name_to_pass and args.model_provider != "ollama": 
         print(f"Error: Model name is required for provider '{args.model_provider}'. Use --model_name or set a default in .env for Groq.")
         sys.exit(1)

    if args.model_provider == "openai" and resolved_openai_api_key is None and args.api_key is None:
        try:
            with open(".token", 'r') as f: 
                resolved_openai_api_key = f.read().strip()
        except FileNotFoundError:
            if OPENAI_API_KEY is None:
                 print("Error: OpenAI API key is required. Provide it via --api_key, .env (OPENAI_API_KEY), or a .token file.")
                 sys.exit(1)
        except Exception as e:
            print(f"Error reading .token file: {e}")
            sys.exit(1)
    
    if args.model_provider == "openai" and not resolved_openai_api_key:
        print("Error: OpenAI API key is required but not found in CLI arguments, .env, or .token file.")
        sys.exit(1)
    if args.model_provider == "anthropic" and not resolved_anthropic_api_key:
        print("Error: Anthropic API key is required but not found in CLI arguments or .env.")
        sys.exit(1)
    if args.model_provider == "groq" and not resolved_groq_api_key:
        print("Error: Groq API key is required but not found in CLI arguments or .env.")
        sys.exit(1)

    chat_ppt = ChatPPT(
        model_provider=args.model_provider, 
        api_key=resolved_openai_api_key, 
        model_name=model_name_to_pass, 
        ollama_url=args.ollama_url, 
        anthropic_api_key=resolved_anthropic_api_key,
        openai_api_base=resolved_openai_api_base,
        groq_api_key=resolved_groq_api_key
    )
    chat_ppt.robot_print("Hi, I am your PPT assistant.")
    ppt_content = chat_ppt.chatppt(
        args.topic, 
        args.pages, 
        args.language, 
        custom_prompt_instructions=None, 
        audience=args.audience 
    )
    chat_ppt.generate_ppt(ppt_content)

if __name__ == "__main__":
    main()
