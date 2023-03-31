# What is ChatPPT
ChatPPT is powered by chatgpt, it could help you to generate PPT/slide. It supports output in English and Chinese

![What is GPT | 600](demo1.png)

![什么是AWS](demo2.png)

# Requirements
Python 3.8.10 +

# Help message
```
$ python chatppt.py --help
usage: chatppt.py [-h] -t TOPIC -k API_KEY [-p PAGES] [-l {cn,en}]

I am your PPT assistant, I can help to you generate PPT.

optional arguments:
  -h, --help            show this help message and exit
  -t TOPIC, --topic TOPIC
                        Your topic name
  -k API_KEY, --api_key API_KEY
                        Your api key file path
  -p PAGES, --pages PAGES
                        How many slides to generate
  -l {cn,en}, --language {cn,en}
                        Output language
```

# How to use it

1. Generate your openai API key https://platform.openai.com/account/api-keys

2. Install requirements
    ```
    $ pip install -r requirements
    ```
3. Generate your PPT

    ```
    $ python chatppt.py -t "What is GPT" -k <your api key> -p 5 -l en

    Hi, I am your PPT assistant.
    I am powerby ChatGPT
    I'm working hard to generate your PPT about What is GPT.
    It may takes about a few minutes.
    Your PPT will be generated in English
    Your PPT have 5 pages.
    page 1: Overview
    page 2: History
    page 3: Applications
    page 4: Limitations
    page 5: Future
    Generate done, enjoy!
    Your PPT: What is GPT.pptx
    ```

