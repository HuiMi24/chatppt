# -*- coding: utf-8 -*-
"""
ChatPPT: AI-Powered Presentation Generator.

This module provides the core logic for ChatPPT, a tool that leverages
Large Language Models (LLMs) from various providers (OpenAI, Anthropic, Groq, Ollama)
to generate PowerPoint presentation outlines and content based on user input.
It supports customization via templates, audience targeting, and custom instructions.
"""

import argparse
import datetime
import json
import random
import re
import sys
import time

import anthropic
import openai # type: ignore
from ollama import Client
from pptx import Presentation
from pptx.shapes.shapetree import SlideShapes

from config import (ANTHROPIC_API_KEY, GROQ_API_BASE, GROQ_API_KEY,
                    GROQ_DEFAULT_MODEL, OPENAI_API_BASE, OPENAI_API_KEY)


# pylint: disable=too-many-arguments
class ChatPPT:
    """
    Handles the generation of presentation content using various LLM providers.

    This class encapsulates the logic for interacting with different LLMs,
    formatting prompts, parsing responses, and generating the final .pptx file.
    Note: Pylint's too-many-arguments is disabled for __init__ as it's
    a central configuration point and explicit arguments are preferred here.
    """

    def __init__(self, model_provider: str, api_key: str | None, model_name: str,
                 ollama_url: str | None = None,
                 anthropic_api_key: str | None = None,
                 openai_api_base: str | None = None,
                 groq_api_key: str | None = None):
        """
        Initialize the ChatPPT instance.

        Args:
            model_provider: The LLM provider to use (e.g., "openai", "ollama").
            api_key: The API key for OpenAI (used if model_provider is "openai").
            model_name: The specific model name to use for the selected provider.
            ollama_url: The URL for the Ollama service (if model_provider is "ollama").
            anthropic_api_key: API key for Anthropic (if model_provider is "anthropic").
            openai_api_base: Custom base URL for OpenAI-compatible APIs.
            groq_api_key: API key for Groq (if model_provider is "groq").
        """
        self.model_provider = model_provider
        self.api_key = api_key  # For OpenAI
        self.model_name = model_name
        self.ollama_url = ollama_url
        self.anthropic_api_key = anthropic_api_key  # For Anthropic
        self.openai_api_base = openai_api_base  # For OpenAI
        self.groq_api_key = groq_api_key  # For Groq

    @staticmethod
    def robot_print(text: str) -> str:
        """
        Prints text with a robot-like typing effect.

        Args:
            text: The string to print.

        Returns:
            The original text string.
        """
        for char_val in text:
            print(char_val, end="", flush=True)
            time.sleep(random.randrange(1, 2) / 1000.0)
        print("\r")
        return text

    # pylint: disable=too-many-arguments
    # Note: 5 arguments is acceptable; disabled as parameters are distinct and clear.
    def chatppt(self, topic: str, pages: int, language: str,
                custom_prompt_instructions: str | None = None,
                audience: str | None = None) -> dict:
        """
        Generates the full presentation content as a JSON-like dictionary.

        Args:
            topic: The main topic of the presentation.
            pages: The desired number of pages (slides).
            language: The language for the presentation content (e.g., "en", "cn").
            custom_prompt_instructions: Optional custom instructions for the LLM.
            audience: Optional target audience for the presentation.

        Returns:
            A dictionary representing the structured presentation content.
        """
        language_map = {"cn": "Chinese", "en": "English"}
        language_str = language_map.get(language, language)
        self.robot_print(f"I'm working hard to generate your PPT about {topic}.")
        self.robot_print("It may takes about a few minutes.")
        self.robot_print(f"Your PPT will be generated in {language_str}")

        messages = self._get_messages(
            topic, pages, language_str,
            custom_prompt_instructions, audience
        )
        content = self._get_content(messages)
        return self._parse_content(content)

    def _get_output_format(self) -> dict:
        """
        Returns the example JSON structure for the entire presentation outline.
        This structure is used in prompts to guide the LLM's output format.
        Returns:
            A dictionary representing the example JSON structure.
        """
        return {
            "title": "example title",
            "pages": [
                {
                    "title": "title for page 1",
                    "content": [
                        {"title": "title for bullet 1", "description": "detail for bullet 1"},
                        {"title": "title for bullet 2", "description": "detail for bullet 2"},
                        {"title": "title for bullet 3", "description": "detail for bullet 3"},
                    ],
                },
                {
                    "title": "title for page 2",
                    "content": [
                        {"title": "title for bullet 1", "description": "detail for bullet 1"},
                        {"title": "title for bullet 2", "description": "detail for bullet 2"},
                    ],
                },
            ],
        }

    def _get_single_page_output_format_example(self) -> dict:
        """
        Returns the example JSON structure for a single page.
        This is used in prompts when regenerating individual pages.
        Returns:
            A dictionary representing the example JSON structure for a single page.
        """
        return {
            "title": "page title",
            "content": [
                {"title": "bullet title 1", "description": "bullet description 1"},
                {"title": "bullet title 2", "description": "bullet description 2"}
            ]
        }

    # pylint: disable=too-many-arguments
    # Note: output_format removed; 5 args is acceptable as parameters are distinct.
    def _get_messages(self, topic: str, pages: int, language_str: str,
                      custom_prompt_instructions: str | None = None,
                      audience: str | None = None) -> list[dict[str, str]]:
        """
        Constructs the prompt messages for the LLM for full presentation generation.

        Args:
            topic: The presentation topic.
            pages: The number of pages.
            language_str: The language for the content (e.g., "English").
            custom_prompt_instructions: Optional custom instructions.
            audience: Optional target audience.

        Returns:
            A list containing a single dictionary, representing the user prompt.
        """
        output_format = self._get_output_format() # Fetched internally
        prompt_lines = [
            f"I am preparing a presentation on {topic}.",
            "Please assist in generating an outline in JSON format, adhering to the specified format "
            f"{json.dumps(output_format)}.",
            f"The presentation should span {pages} pages, with as many bullet points as possible.",
            f"The content should be returned in {language_str}.",
            "You must add content for each slide.",
            "For each slide, you must add at least 4 bullet.",
            "Please ensure the output is valid JSON match the RFC-8295 specification."
        ]
        if audience and audience.strip():
            prompt_lines.append(
                f"The target audience for this presentation is: {audience.strip()}. "
                "Please tailor the content accordingly."
            )
        if custom_prompt_instructions and custom_prompt_instructions.strip():
            prompt_lines.append(
                f"\nAdditional Custom Instructions:\n{custom_prompt_instructions.strip()}"
            )
        full_prompt = "\n".join(prompt_lines)
        return [{"role": "user", "content": full_prompt}]

    # pylint: disable=too-many-arguments
    # Note: 4 arguments is fine, Pylint might be overly strict here.
    def regenerate_single_page(self, original_topic: str, original_language: str,
                               page_data_to_edit: dict,
                               new_instructions_for_page: str | None = None) -> dict:
        """
        Regenerates content for a single page based on edits and new instructions.

        Args:
            original_topic: The main topic of the overall presentation.
            original_language: The language for the content (e.g., "English").
            page_data_to_edit: The current JSON data for the page to be edited.
            new_instructions_for_page: Optional new instructions for refining this page.

        Returns:
            A dictionary representing the JSON for the regenerated page.

        Raises:
            ValueError: If the LLM response cannot be parsed or validated.
        """
        page_content_str = json.dumps(page_data_to_edit, ensure_ascii=False, indent=2)
        single_page_format_example_str = json.dumps(
            self._get_single_page_output_format_example(), ensure_ascii=False, indent=2
        )

        prompt_lines = [
            f"I am working on a presentation about '{original_topic}'.",
            f"I need to refine a single slide. The language for the output should be {original_language}.",
            f"The current content of the slide is (in JSON format):\n{page_content_str}\n"
        ]

        if new_instructions_for_page and new_instructions_for_page.strip():
            prompt_lines.append(
                "Please refine this slide based on the following instructions: "
                f"{new_instructions_for_page.strip()}\n"
            )
        else:
            prompt_lines.append(
                "Please review and refine the content of this slide, ensuring clarity, "
                "accuracy, and completeness based on its current title and bullet points.\n"
            )

        prompt_lines.extend([
            "You MUST return ONLY the complete JSON object for this single refined slide, "
            "including its 'title' and 'content' (with 'title' and 'description' "
            "for each bullet point).",
            "Do not return any other text, explanations, or markdown formatting around the JSON.",
            "The JSON structure for the page should be exactly like this example "
            f"(though with different actual content): {single_page_format_example_str}."
        ])

        prompt = "\n".join(prompt_lines)
        messages = [{"role": "user", "content": prompt}]

        self.robot_print(f"Refining page: '{page_data_to_edit.get('title', 'Untitled Page')}'...")
        llm_response_str = self._get_content(messages)

        try:
            match = re.search(r"(\{.*\})", llm_response_str, re.DOTALL)
            json_str = match.group(0) if match else llm_response_str
            new_page_json = json.loads(json_str.strip())

            if not isinstance(new_page_json, dict) or \
               "title" not in new_page_json or \
               "content" not in new_page_json or \
               not isinstance(new_page_json["content"], list):
                raise ValueError(
                    "LLM returned JSON but not in the expected page format "
                    "(missing title, content, or content is not a list)."
                )
            for item in new_page_json["content"]:
                if not isinstance(item, dict) or "title" not in item or "description" not in item:
                    raise ValueError(
                        "LLM returned JSON with malformed bullet points "
                        "(missing title or description in a bullet)."
                    )
            self.robot_print(f"Page '{new_page_json.get('title', 'Untitled Page')}' "
                             "refined successfully.")
            return new_page_json
        except (json.JSONDecodeError, ValueError) as e:
            error_message = (
                f"Error parsing regenerated page content: {e}. "
                f"Raw response snippet: {llm_response_str[:500]}..."
            )
            print(f"\n{error_message}")
            raise ValueError(error_message) from e

    # pylint: disable=too-many-branches
    # Note: Multiple providers handled, structure is clear for targeted functionality.
    def _get_content(self, messages: list[dict[str, str]]) -> str:
        """
        Calls the appropriate LLM provider to get content based on the prompt.
        Args:
            messages: A list of message dictionaries for the LLM prompt.
        Returns:
            The string content received from the LLM.
        Raises:
            ValueError: If a required API key or URL is missing, or provider is unknown.
        """
        if self.model_provider == "openai":
            if not self.api_key:
                raise ValueError("OpenAI API key is required.")
            openai.api_key = self.api_key
            original_api_base = openai.api_base
            api_base_changed = False
            if self.openai_api_base and self.openai_api_base.strip():
                if openai.api_base != self.openai_api_base:
                    openai.api_base = self.openai_api_base
                    api_base_changed = True
            elif openai.api_base is None: # Ensure default if not set and no custom base
                openai.api_base = "https://api.openai.com/v1"
            try:
                completion = openai.ChatCompletion.create(model=self.model_name, messages=messages)
                return str(completion.choices[0].message.content)
            finally: # Reset openai.api_base to its original state
                if api_base_changed or (self.openai_api_base and self.openai_api_base.strip() and original_api_base is None):
                    openai.api_base = original_api_base
                elif original_api_base is None and not (self.openai_api_base and self.openai_api_base.strip()):
                     if openai.api_base == "https://api.openai.com/v1": # If we set it to default
                         openai.api_base = None # Reset to None if it was originally None

        elif self.model_provider == "groq":
            if not self.groq_api_key:
                raise ValueError("Groq API Key is required.")
            original_openai_api_key = openai.api_key
            original_openai_api_base = openai.api_base
            openai.api_key = self.groq_api_key
            current_groq_api_base = (GROQ_API_BASE if GROQ_API_BASE and GROQ_API_BASE.strip()
                                     else "https://api.groq.com/openai/v1")
            openai.api_base = current_groq_api_base
            try:
                completion = openai.ChatCompletion.create(model=self.model_name, messages=messages)
                return str(completion.choices[0].message.content)
            finally:
                openai.api_key = original_openai_api_key
                openai.api_base = original_openai_api_base
        elif self.model_provider == "ollama":
            if self.ollama_url is None:
                raise ValueError("Ollama URL is required.")
            client = Client(host=self.ollama_url)
            response = client.chat(model=self.model_name, messages=messages)
            return str(response["message"]["content"])
        elif self.model_provider == "anthropic":
            if not self.anthropic_api_key:
                raise ValueError("Anthropic API key is required.")
            client = anthropic.Anthropic(api_key=self.anthropic_api_key)
            response = client.messages.create(model=self.model_name,
                                              max_tokens=2048, messages=messages)
            return str(response.content[0].text)

        raise ValueError(f"Unknown model provider: {self.model_provider}")


    def _parse_content(self, content: str) -> dict:
        """
        Parses the LLM's string response to extract the main JSON content for the presentation.
        Args:
            content: The raw string response from the LLM.
        Returns:
            A dictionary representing the parsed presentation structure.
        Raises:
            ValueError: If the content cannot be parsed into valid JSON or
                        if the JSON structure is not as expected for a full presentation.
        """
        try:
            match = re.search(r"(\{.*\})", content, re.DOTALL)
            json_str = match.group(0) if match else content
            parsed_json = json.loads(json_str.strip())

            if not isinstance(parsed_json, dict) or \
               "title" not in parsed_json or \
               "pages" not in parsed_json or \
               not isinstance(parsed_json["pages"], list):
                raise ValueError(
                    "LLM response for full presentation is not in the expected root format "
                    "(missing title or pages list)."
                )
            return parsed_json
        except (json.JSONDecodeError, ValueError) as e:
            error_msg = (
                f"The response is not a valid JSON format for full presentation: {e}\n"
                f"Raw content from LLM: {content[:500]}..."
            )
            print(error_msg)
            raise ValueError(
                "The LLM returned an invalid result for the full presentation, please retry later."
            ) from e

    def generate_ppt(self, content: dict, template: str | None = None) -> str:
        """
        Generates a .pptx file from the structured presentation content.
        Args:
            content: A dictionary representing the presentation structure.
            template: Optional path to a .pptx template file.
        Returns:
            The filename of the generated .pptx file.
        """
        ppt = Presentation(template) if template else Presentation()
        self.create_slides(ppt, content)
        ppt_name = self._get_ppt_name(content)
        ppt.save(ppt_name)
        self.robot_print(f"Generate done, enjoy!\nYour PPT: {ppt_name}")
        return ppt_name

    def create_slides(self, presentation: Presentation, content: dict) -> None:
        """
        Adds slides to the presentation object based on the content.
        Args:
            presentation: The python-pptx Presentation object.
            content: The dictionary containing slide data.
        """
        self.create_title_slide(presentation, content)
        pages = content.get("pages", [])
        self.robot_print(f"Your PPT has {len(pages)} pages.")
        for index, page in enumerate(pages):
            self.create_content_slide(presentation, page, index)

    def create_title_slide(self, presentation: Presentation, content: dict) -> None:
        """
        Creates the title slide.
        Args:
            presentation: The python-pptx Presentation object.
            content: The dictionary containing the main title.
        """
        title_slide_layout = presentation.slide_layouts[0]
        title_slide = presentation.slides.add_slide(title_slide_layout)
        if title_slide.shapes.title: # Check if title placeholder exists
            title_slide.shapes.title.text = content.get("title", "")
        # Check if subtitle placeholder exists (usually placeholders[1])
        if title_slide.placeholders and len(title_slide.placeholders) > 1:
            if title_slide.placeholders[1]:
                title_slide.placeholders[1].text = "Generated by ChatPPT"

    def create_content_slide(self, presentation: Presentation, page: dict, index: int) -> None:
        """
        Creates a content slide with a title and bullet points.
        Args:
            presentation: The python-pptx Presentation object.
            page: A dictionary containing data for the page.
            index: The page number (0-indexed).
        """
        page_title = page.get("title", "")
        self.robot_print(f"Page {index + 1}: {page_title}")
        bullet_slide_layout = presentation.slide_layouts[1] # Assuming layout 1 for content
        bullet_slide = presentation.slides.add_slide(bullet_slide_layout)
        if bullet_slide.shapes.title: # Check if title placeholder exists
            bullet_slide.shapes.title.text = page_title
        # Check if body placeholder exists (usually placeholders[1])
        if bullet_slide.placeholders and len(bullet_slide.placeholders) > 1:
            self.add_bullets_to_slide(bullet_slide, page)
        else:
            print(f"Warning: Slide layout for page {index+1} (title: '{page_title}') "
                  "might not have a body placeholder for bullets.")


    def add_bullets_to_slide(self, slide: SlideShapes, page: dict) -> None:
        """
        Adds bullet points to a content slide.
        Args:
            slide: The python-pptx Slide object.
            page: A dictionary containing bullet point data for the page.
        """
        try:
            body_shape = slide.placeholders[1] # Assuming placeholder 1 for body
        except IndexError:
            slide_title = slide.shapes.title.text if slide.shapes.title else "Untitled"
            print(f"Warning: No body placeholder found for slide titled '{slide_title}'. "
                  "Bullets cannot be added.")
            return

        for bullet in page.get("content", []):
            self.add_bullet(body_shape, bullet)

    def add_bullet(self, body_shape, bullet: dict) -> None:
        """
        Adds a single bullet point (title and description) to a shape.
        Args:
            body_shape: The shape to add the bullet point to.
            bullet: A dictionary containing the bullet's title and description.
        """
        paragraph = body_shape.text_frame.add_paragraph()
        paragraph.text = bullet.get("title", "")
        paragraph.level = 1
        paragraph = body_shape.text_frame.add_paragraph()
        paragraph.text = bullet.get("description", "")
        paragraph.level = 2

    def _get_ppt_name(self, content: dict) -> str:
        """
        Generates a filename for the PPT based on the title and timestamp.
        Args:
            content: The dictionary containing the presentation title.
        Returns:
            A sanitized filename string.
        """
        ppt_name_base = content.get("title", "Untitled_Presentation")
        ppt_name_sanitized = re.sub(r'[\\/:*?"<>|]', "", ppt_name_base)
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        return f"{ppt_name_sanitized}_{timestamp}.pptx"

def args_parser() -> argparse.Namespace:
    """Parses command-line arguments for ChatPPT."""
    parser = argparse.ArgumentParser(
        description="I am your PPT assistant, I can help to you generate PPT."
    )
    parser.add_argument(
        "-m", "--model_provider",
        choices=["openai", "ollama", "anthropic", "groq"],
        default="openai", help="Select the model provider"
    )
    parser.add_argument(
        "-n", "--model_name", type=str, default=None,
        help="Specify the model name to use (e.g., gpt-3.5-turbo, llama3, "
             "claude-3-opus, mixtral-8x7b-32768). Required unless provider "
             "has a default (e.g. Groq)."
    )
    parser.add_argument("-t", "--topic", type=str, required=True, help="Your topic name")
    parser.add_argument(
        "--audience", type=str, default=None,
        help="Specify the target audience for the presentation "
             "(e.g., 'students', 'technical experts')."
    )
    parser.add_argument(
        "-k", "--api_key", type=str, default=None,
        help="Your OpenAI API key or file path. Overrides .env."
    )
    parser.add_argument(
        "--anthropic_api_key", type=str, default=None,
        help="Your Anthropic API key or file path. Overrides .env."
    )
    parser.add_argument(
        "--groq_api_key", type=str, default=None,
        help="Your Groq API key. Overrides .env."
    )
    parser.add_argument(
        "--openai_api_base", type=str, default=None,
        help="Custom base URL for OpenAI API. Overrides .env."
    )
    parser.add_argument(
        "-u", "--ollama_url", type=str, default="http://localhost:11434",
        help="Your ollama url"
    )
    parser.add_argument(
        "-p", "--pages", type=int, default=5,
        help="How many slides to generate"
    )
    parser.add_argument(
        "-l", "--language", choices=["cn", "en"], default="en",
        help="Output language"
    )
    return parser.parse_args()

# pylint: disable=too-many-locals
def _resolve_main_configs(args: argparse.Namespace) -> dict:
    """
    Resolves configurations for main execution, prioritizing CLI then .env.

    Args:
        args: Parsed command-line arguments.

    Returns:
        A dictionary of resolved configuration values.
    """
    resolved_configs = {
        "openai_api_key": args.api_key if args.api_key is not None else OPENAI_API_KEY,
        "anthropic_api_key": (args.anthropic_api_key
                              if args.anthropic_api_key is not None
                              else ANTHROPIC_API_KEY),
        "groq_api_key": args.groq_api_key if args.groq_api_key is not None else GROQ_API_KEY,
        "openai_api_base": (args.openai_api_base
                            if args.openai_api_base is not None
                            else OPENAI_API_BASE),
        "model_name": args.model_name
    }

    if args.model_provider == "groq" and not resolved_configs["model_name"]:
        resolved_configs["model_name"] = GROQ_DEFAULT_MODEL
        if not resolved_configs["model_name"]:
            resolved_configs["model_name"] = "mixtral-8x7b-32768"
            print("Warning: No model name for Groq via --model_name or .env. "
                  "Using 'mixtral-8x7b-32768'.")

    # Fallback to .token file for OpenAI API key if still not resolved
    if args.model_provider == "openai" and not resolved_configs["openai_api_key"]:
        try:
            with open(".token", 'r', encoding='utf-8') as f:
                resolved_configs["openai_api_key"] = f.read().strip()
        except FileNotFoundError:
            # This case is handled by _validate_resolved_configs if OPENAI_API_KEY was also None
            pass
        except IOError as e:
            print(f"Error reading .token file: {e}")
            sys.exit(1)
    return resolved_configs

def _validate_resolved_configs(args: argparse.Namespace, resolved_configs: dict) -> None:
    """
    Validates that necessary resolved configurations are present, exiting if not.

    Args:
        args: Parsed command-line arguments.
        resolved_configs: Dictionary of resolved configuration values.
    """
    if not resolved_configs["model_name"] and args.model_provider not in ["ollama"]:
        # Ollama model can be selected in UI, so CLI might not have it.
        print(f"Error: Model name is required for provider '{args.model_provider}'. "
              "Use --model_name or set a default in .env for Groq.")
        sys.exit(1)
    if args.model_provider == "openai" and not resolved_configs["openai_api_key"]:
        print("Error: OpenAI API key not found in CLI, .env, or .token file.")
        sys.exit(1)
    if args.model_provider == "anthropic" and not resolved_configs["anthropic_api_key"]:
        print("Error: Anthropic API key not found in CLI or .env.")
        sys.exit(1)
    if args.model_provider == "groq" and not resolved_configs["groq_api_key"]:
        print("Error: Groq API key not found in CLI or .env.")
        sys.exit(1)

def main() -> None:
    """Main function to run ChatPPT from the command line."""
    args = args_parser()
    resolved_configs = _resolve_main_configs(args)
    _validate_resolved_configs(args, resolved_configs)

    chat_ppt_instance = ChatPPT(
        model_provider=args.model_provider,
        api_key=resolved_configs["openai_api_key"],
        model_name=str(resolved_configs["model_name"]), # Ensure model_name is string
        ollama_url=args.ollama_url,
        anthropic_api_key=resolved_configs["anthropic_api_key"],
        openai_api_base=resolved_configs["openai_api_base"],
        groq_api_key=resolved_configs["groq_api_key"]
    )
    chat_ppt_instance.robot_print("Hi, I am your PPT assistant.")
    ppt_content_data = chat_ppt_instance.chatppt(
        args.topic,
        args.pages,
        args.language,
        custom_prompt_instructions=None, # CLI doesn't support this yet
        audience=args.audience
    )
    chat_ppt_instance.generate_ppt(ppt_content_data)

if __name__ == "__main__":
    main()
