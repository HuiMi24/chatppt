import re
import openai
import json
import argparse
import time
import random
import datetime
import anthropic 
from ollama import Client
from pptx import Presentation


class ChatPPT:
    def __init__(self, model_provider, api_key, model_name, ollama_url=None, anthropic_api_key=None):
        self.model_provider = model_provider
        self.api_key = api_key
        self.model_name = model_name
        self.ollama_url = ollama_url
        self.anthropic_api_key = anthropic_api_key

    @staticmethod
    def robot_print(text):
        for char in text:
            print(char, end="", flush=True)
            time.sleep(random.randrange(1, 2) / 1000.0)
        print("\r")
        return text

    def chatppt(self, topic, pages, language, custom_prompt_instructions=None): # Added custom_prompt_instructions
        language_map = {"cn": "Chinese", "en": "English"}
        language_str = language_map[language]
        self.robot_print(f"I'm working hard to generate your PPT about {topic}.")
        self.robot_print("It may takes about a few minutes.")
        self.robot_print(f"Your PPT will be generated in {language_str}")
        output_format = self._get_output_format()
        # Pass custom_prompt_instructions to _get_messages
        messages = self._get_messages(topic, pages, language_str, output_format, custom_prompt_instructions)
        content = self._get_content(messages)
        return self._parse_content(content)

    def _get_output_format(self):
        return {
            "title": "example title",
            "pages": [
                {
                    "title": "title for page 1",
                    "content": [
                        {
                            "title": "title for bullet 1",
                            "description": "detail for bullet 1",
                        },
                        {
                            "title": "title for bullet 2",
                            "description": "detail for bullet 2",
                        },
                        {
                            "title": "title for bullet 3",
                            "description": "detail for bullet 3",
                        },
                    ],
                },
                {
                    "title": "title for page 2",
                    "content": [
                        {
                            "title": "title for bullet 1",
                            "description": "detail for bullet 1",
                        },
                        {
                            "title": "title for bullet 2",
                            "description": "detail for bullet 2",
                        },
                    ],
                },
            ],
        }

    def _get_messages(self, topic, pages, language_str, output_format, custom_prompt_instructions=None): # Added custom_prompt_instructions
        base_prompt = f"""I am preparing a presentation on {topic}. Please assist in generating an outline in JSON format, adhering to the specified format {json.dumps(output_format)}. The presentation should span {pages} pages, with as many bullet points as possible. The content should be returned in {language_str}. You must add content for each slide. For each slide, you must add at least 4 bullet. Please ensure the output is valid JSON match the RFC-8295 specification. Don't return any other message"""
        
        if custom_prompt_instructions and custom_prompt_instructions.strip():
            full_prompt = f"{base_prompt}\n\nAdditional Custom Instructions:\n{custom_prompt_instructions.strip()}"
        else:
            full_prompt = base_prompt
            
        return [{"role": "user", "content": full_prompt}]


    def _get_content(self, messages):
        if self.model_provider == "openai":
            openai.api_key = self.api_key
            completion = openai.ChatCompletion.create(
                model=self.model_name, messages=messages
            )
            return completion.choices[0].message.content
        elif self.model_provider == "ollama":
            if self.ollama_url is None:
                raise Exception("Ollama URL is required when model_provider is 'ollama'")
            client = Client(host=self.ollama_url)
            response = client.chat(model=self.model_name, messages=messages)
            return response["message"]["content"]
        elif self.model_provider == "anthropic": 
            if self.anthropic_api_key is None:
                raise Exception("Anthropic API key is required when model_provider is 'anthropic'")
            client = anthropic.Anthropic(api_key=self.anthropic_api_key)
            
            response = client.messages.create(
                model=self.model_name,
                max_tokens=2048, 
                messages=messages
            )
            return response.content[0].text

    def _parse_content(self, content):
        try:
            match = re.search(r"(\{.*\})", content, re.DOTALL)
            if match:
                content = match.groups()[0]
            return json.loads(content.strip())
        except Exception as e:
            print(f"The response is not a valid JSON format: {e}")
            print(f"Raw content from LLM: {content}") 
            print("I'm a PPT assistant, your PPT generate failed, please retry later..")
            raise Exception("The LLM return invalid result, please retry later..")
            exit(1)

    def generate_ppt(self, content, template=None):
        ppt = Presentation()
        if template:
            ppt = Presentation(template)
        self.create_slides(ppt, content)
        ppt_name = self._get_ppt_name(content)
        ppt.save(ppt_name)
        self.robot_print("Generate done, enjoy!")
        self.robot_print(f"Your PPT: {ppt_name}")
        return ppt_name

    def create_slides(self, presentation, content):
        self.create_title_slide(presentation, content)
        pages = content.get("pages", [])
        self.robot_print(f"Your PPT has {len(pages)} pages.")
        for index, page in enumerate(pages):
            self.create_content_slide(presentation, page, index)

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
        for bullet in page.get("content", []):
            self.add_bullet(body_shape, bullet)

    def add_bullet(self, body_shape, bullet):
        bullet_title = bullet.get("title", "")
        bullet_description = bullet.get("description", "")
        self.add_paragraph(body_shape, bullet_title, level=1)
        self.add_paragraph(body_shape, bullet_description, level=2)

    def add_paragraph(self, body_shape, text, level):
        paragraph = body_shape.text_frame.add_paragraph()
        paragraph.text = text
        paragraph.level = level

    def _get_ppt_name(self, content):
        ppt_name = content.get("title", "")
        ppt_name = re.sub(r'[\\/:*?"<>|]', "", ppt_name)
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        return f"{ppt_name}_{timestamp}.pptx"


def main():
    args = args_parser()
    chat_ppt = ChatPPT(args.model_provider, args.api_key, args.model_name, args.ollama_url, args.anthropic_api_key)
    chat_ppt.robot_print("Hi, I am your PPT assistant.")
    # CLI does not yet support custom_prompt_instructions, so pass None
    ppt_content = chat_ppt.chatppt(args.topic, args.pages, args.language, custom_prompt_instructions=None)
    chat_ppt.generate_ppt(ppt_content)


def args_parser():
    parser = argparse.ArgumentParser(
        description="I am your PPT assistant, I can help to you generate PPT."
    )
    parser.add_argument(
        "-m",
        "--model_provider",
        choices=["openai", "ollama", "anthropic"], 
        default="openai",
        help="Select the model provider (e.g., openai, ollama, anthropic)",
    )
    parser.add_argument(
        "-n",
        "--model_name",
        type=str,
        required=True,
        help="Specify the model name to use (e.g., gpt-3.5-turbo, llama3, claude-3-opus-20240229)",
    )
    parser.add_argument(
        "-t", "--topic", type=str, required=True, help="Your topic name"
    )
    parser.add_argument( 
        "-k", "--api_key", type=str, default=".token", help="Your OpenAI API key or file path"
    )
    parser.add_argument( 
        "--anthropic_api_key", type=str, default=None, help="Your Anthropic API key or file path"
    )
    parser.add_argument(
        "-u",
        "--ollama_url",
        type=str,
        default="http://localhost:11434",
        help="Your ollama url",
    )
    parser.add_argument(
        "-p",
        "--pages",
        type=int,
        required=False,
        default=5,
        help="How many slides to generate",
    )
    parser.add_argument(
        "-l",
        "--language",
        choices=["cn", "en"],
        default="en",
        required=False,
        help="Output language",
    )
    args = parser.parse_args()
    if args.model_provider == "openai" and args.api_key == ".token":
        try:
            with open(args.api_key, 'r') as f:
                args.api_key = f.read().strip()
        except FileNotFoundError:
             parser.error("--api_key file not found or direct key not provided when model_provider is 'openai'")
        except Exception as e:
            parser.error(f"Error reading --api_key for openai: {e}")

    if args.model_provider == "anthropic":
        if args.anthropic_api_key is None:
            parser.error("--anthropic_api_key is required when model_provider is 'anthropic'")
        else:
            try:
                with open(args.anthropic_api_key, 'r') as f:
                    args.anthropic_api_key = f.read().strip()
            except FileNotFoundError:
                pass 
            except Exception as e:
                 parser.error(f"Error reading --anthropic_api_key: {e}")

    if args.model_provider == "openai" and (args.api_key is None or args.api_key == ".token"):
         parser.error("--api_key is required when model_provider is 'openai'")

    return args

if __name__ == "__main__":
    main()
