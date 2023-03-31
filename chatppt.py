import openai
import json
import argparse
import time
import random
from pptx import Presentation


def robot_print(text):
    for char in text:
        print(char, end="", flush=True)
        time.sleep(random.randrange(1, 5) / 100.0)
    print("\r")


def chatppt(topic: str, pages: int, api_key: str, language: str):
    language_map = {"cn": "Chinese", "en": "English"}
    language = language_map[language]

    output_format = {
        "title": "example title",
        "pages": [
            {
                "title": "title for page 1",
                # "subtitle": "subtitle for page 1",
                "content": [
                    {
                        "title": "title for bullet 1",
                        "desctription": "detail for bullet 1",
                    },
                    {
                        "title": "title for bullet 2",
                        "desctription": "detail for bullet 2",
                    },
                ],
            },
            {
                "title": "title for page 2",
                # "subtitle": "subtitle for page 2",
                "content": [
                    {
                        "title": "title for bullet 1",
                        "desctription": "detail for bullet 1",
                    },
                    {
                        "title": "title for bullet 2",
                        "desctription": "detail for bullet 2",
                    },
                ],
            },
        ],
    }

    messages = [
        {
            "role": "user",
            "content": f"I'm going to prepare a presentation about {topic}, please help to outline detailed about this topic, output with JSON language with follow in format {output_format}, please help to generate {pages} pages, the bullet for each as much as possible, please only return JSON format and use double quotes, please return the content in {language}",
        },
    ]

    robot_print(f"I'm working hard to generate your PPT about {topic}.")
    robot_print("It may takes about a few minutes.")
    robot_print(f"Your PPT will be generated in {language}")

    openai.api_key = api_key
    completion = openai.ChatCompletion.create(model="gpt-3.5-turbo", messages=messages)
    try:
        content = completion.choices[0].message.content
        # just replace ' to " is not a good soluation
        # print(content)
        content = json.loads(content.strip())
        return content
    except Exception as e:
        print("I'm a PPT assistant, your PPT generate failed, please retry later..")
        exit(1)
        # raise Exception("I'm PPT generate assistant, some error happened, please retry")


def generate_ppt(content: str, template=None):
    ppt = Presentation()
    if template:
        ppt = Presentation(template)

    # Creating slide layout
    first_slide_layout = ppt.slide_layouts[0]

    # """ Ref for slide types:
    # 0 ->  title and subtitle
    # 1 ->  title and content
    # 2 ->  section header
    # 3 ->  two content
    # 4 ->  Comparison
    # 5 ->  Title only
    # 6 ->  Blank
    # 7 ->  Content with caption
    # 8 ->  Pic with caption
    # """

    slide = ppt.slides.add_slide(first_slide_layout)
    slide.shapes.title.text = content.get("title", "")
    slide.placeholders[1].text = "Generate by ChatPPT"

    pages = content.get("pages", [])
    robot_print(f"Your PPT have {len(pages)} pages.")
    for i, page in enumerate(pages):
        page_title = page.get("title", "")
        robot_print(f"page {i+1}: {page_title}")
        bullet_layout = ppt.slide_layouts[1]
        bullet_slide = ppt.slides.add_slide(bullet_layout)
        bullet_spahe = bullet_slide.shapes
        bullet_slide.shapes.title.text = page_title

        body_shape = bullet_spahe.placeholders[1]
        for bullet in page.get("content", []):
            paragraph = body_shape.text_frame.add_paragraph()
            paragraph.text = bullet.get("title", "")
            paragraph.level = 1

            paragraph = body_shape.text_frame.add_paragraph()
            paragraph.text = bullet.get("description", "")
            paragraph.level = 2

    ppt_name = content.get("title", "")
    ppt_name = f"{ppt_name}.pptx"
    ppt.save(ppt_name)
    robot_print("Generate done, enjoy!")
    robot_print(f"Your PPT: {ppt_name}")


def main(topic: str, pages: int, api_key: str, language: str, template_path=None):
    robot_print("Hi, I am your PPT assistant.")
    robot_print("I am powerby ChatGPT")
    # robot_print("If you have any issue, please contact hui_mi@dell.com")
    ppt_content = chatppt(topic, pages, api_key, language)
    generate_ppt(ppt_content)


if __name__ == "__main__":
    # create the parser
    parser = argparse.ArgumentParser(
        description="I am your PPT assistant, I can help to you generate PPT."
    )

    # add arguments
    parser.add_argument(
        "-t", "--topic", type=str, required=True, help="Your topic name"
    )
    parser.add_argument(
        "-k",
        "--api_key",
        type=str,
        required=True,
        default=".token",
        help="Your api key file path",
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

    # parse the arguments
    args = parser.parse_args()

    main(args.topic, args.pages, args.api_key, args.language)
