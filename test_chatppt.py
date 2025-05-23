import unittest
import sys
from unittest.mock import patch, mock_open
import json # Added for _get_output_format

from chatppt import args_parser, ChatPPT

class TestArgsParser(unittest.TestCase):
    def setUp(self):
        self.original_argv = sys.argv

    def tearDown(self):
        sys.argv = self.original_argv

    def test_openai_args_direct_key(self):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-3.5-turbo',
            '--api_key', 'test_openai_key_direct', # Direct key
            '--topic', 'OpenAI Direct Key Test',
        ]
        args = args_parser()
        self.assertEqual(args.model_provider, 'openai')
        self.assertEqual(args.model_name, 'gpt-3.5-turbo')
        self.assertEqual(args.api_key, 'test_openai_key_direct')
        self.assertEqual(args.topic, 'OpenAI Direct Key Test')
        self.assertIsNone(args.anthropic_api_key)

    @patch("builtins.open", new_callable=mock_open, read_data="test_openai_key_from_default_dot_token_file")
    def test_openai_args_key_from_default_dot_token_file(self, mock_file):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-4',
            '--topic', 'OpenAI Default .token File Key Test'
        ]
        args = args_parser()
        self.assertEqual(args.model_provider, 'openai')
        self.assertEqual(args.model_name, 'gpt-4')
        self.assertEqual(args.api_key, 'test_openai_key_from_default_dot_token_file')
        self.assertEqual(args.topic, 'OpenAI Default .token File Key Test')

    def test_openai_args_key_as_custom_path_not_default_dot_token(self):
        custom_path_key = 'path/to/my_openai_key.txt'
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-4',
            '--api_key', custom_path_key, 
            '--topic', 'OpenAI Custom Path As Key Test'
        ]
        args = args_parser()
        self.assertEqual(args.model_provider, 'openai')
        self.assertEqual(args.model_name, 'gpt-4')
        self.assertEqual(args.api_key, custom_path_key) 
        self.assertEqual(args.topic, 'OpenAI Custom Path As Key Test')

    def test_ollama_args(self):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'ollama',
            '--model_name', 'llama3',
            '--ollama_url', 'http://localhost:1111',
            '--topic', 'Ollama Test Topic'
        ]
        args = args_parser()
        self.assertEqual(args.model_provider, 'ollama')
        self.assertEqual(args.model_name, 'llama3')
        self.assertEqual(args.ollama_url, 'http://localhost:1111')
        self.assertEqual(args.topic, 'Ollama Test Topic')
        self.assertEqual(args.api_key, ".token") 
        self.assertIsNone(args.anthropic_api_key)

    def test_anthropic_args_direct_key(self):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'anthropic',
            '--model_name', 'claude-3-opus-20240229',
            '--anthropic_api_key', 'test_anthropic_key_direct',
            '--topic', 'Anthropic Direct Key Test'
        ]
        args = args_parser()
        self.assertEqual(args.model_provider, 'anthropic')
        self.assertEqual(args.model_name, 'claude-3-opus-20240229')
        self.assertEqual(args.anthropic_api_key, 'test_anthropic_key_direct')
        self.assertEqual(args.topic, 'Anthropic Direct Key Test')
        self.assertEqual(args.api_key, ".token")

    @patch("builtins.open", new_callable=mock_open, read_data="test_anthropic_key_from_file")
    def test_anthropic_args_key_from_file(self, mock_file):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'anthropic',
            '--model_name', 'claude-3-sonnet-20240229',
            '--anthropic_api_key', 'path/to/fake_anthropic_key.txt', 
            '--topic', 'Anthropic File Key Test'
        ]
        args = args_parser()
        self.assertEqual(args.model_provider, 'anthropic')
        self.assertEqual(args.model_name, 'claude-3-sonnet-20240229')
        self.assertEqual(args.anthropic_api_key, 'test_anthropic_key_from_file')
        self.assertEqual(args.topic, 'Anthropic File Key Test')

    def test_error_openai_missing_api_key_file_not_found(self):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-3.5-turbo',
            '--topic', 'Test Topic'
        ]
        with patch("builtins.open", side_effect=FileNotFoundError):
            with self.assertRaises(SystemExit):
                args_parser()
    
    def test_error_anthropic_missing_api_key(self):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'anthropic',
            '--model_name', 'claude-opus',
            '--topic', 'Test Topic'
        ]
        with self.assertRaises(SystemExit):
            args_parser()

    def test_error_missing_topic(self):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-3.5-turbo',
            '--api_key', 'fake_key'
        ]
        with self.assertRaises(SystemExit):
            args_parser()

    def test_error_missing_model_name(self):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--api_key', 'fake_key',
            '--topic', 'Test Topic'
        ]
        with self.assertRaises(SystemExit):
            args_parser()

class TestChatPPTInitialization(unittest.TestCase):
    def test_init_openai(self):
        chat_instance = ChatPPT(model_provider="openai", api_key="fake_openai_key", model_name="gpt-4")
        self.assertEqual(chat_instance.model_provider, "openai")
        self.assertEqual(chat_instance.model_name, "gpt-4")
        self.assertEqual(chat_instance.api_key, "fake_openai_key")
        self.assertIsNone(chat_instance.ollama_url)
        self.assertIsNone(chat_instance.anthropic_api_key)

    def test_init_ollama(self):
        chat_instance = ChatPPT(model_provider="ollama", api_key=None, model_name="llama3", ollama_url="http://fakeollamaurl:11434")
        self.assertEqual(chat_instance.model_provider, "ollama")
        self.assertEqual(chat_instance.model_name, "llama3")
        self.assertEqual(chat_instance.ollama_url, "http://fakeollamaurl:11434")
        self.assertIsNone(chat_instance.api_key)
        self.assertIsNone(chat_instance.anthropic_api_key)

    def test_init_anthropic(self):
        chat_instance = ChatPPT(model_provider="anthropic", api_key=None, model_name="claude-3-haiku", anthropic_api_key="fake_anthropic_key")
        self.assertEqual(chat_instance.model_provider, "anthropic")
        self.assertEqual(chat_instance.model_name, "claude-3-haiku")
        self.assertEqual(chat_instance.anthropic_api_key, "fake_anthropic_key")
        self.assertIsNone(chat_instance.api_key)
        self.assertIsNone(chat_instance.ollama_url)

    def test_init_all_params_for_openai(self):
        chat_instance = ChatPPT(
            model_provider="openai", 
            api_key="relevant_openai_key",
            model_name="gpt-4o", 
            ollama_url="http://irrelevant_ollama", 
            anthropic_api_key="irrelevant_anthropic"
        )
        self.assertEqual(chat_instance.model_provider, "openai")
        self.assertEqual(chat_instance.model_name, "gpt-4o")
        self.assertEqual(chat_instance.api_key, "relevant_openai_key")
        self.assertEqual(chat_instance.ollama_url, "http://irrelevant_ollama") 
        self.assertEqual(chat_instance.anthropic_api_key, "irrelevant_anthropic")

class TestChatPPTPromptCustomization(unittest.TestCase):
    def setUp(self):
        # api_key is a required positional argument for ChatPPT constructor
        self.chat_instance = ChatPPT(model_provider="test_provider", api_key="dummy_key_for_test", model_name="test_model")
        self.output_format = self.chat_instance._get_output_format() 

    def test_get_messages_no_custom_instructions(self):
        messages = self.chat_instance._get_messages(
            topic="Test Topic",
            pages=3,
            language_str="English",
            output_format=self.output_format,
            custom_prompt_instructions=None
        )
        self.assertIsInstance(messages, list)
        self.assertEqual(len(messages), 1)
        self.assertNotIn("Additional Custom Instructions:", messages[0]["content"])
        self.assertIn("I am preparing a presentation on Test Topic", messages[0]["content"])
        # Check that the output_format (JSON string) is in the prompt
        self.assertIn(json.dumps(self.output_format), messages[0]["content"])

    def test_get_messages_with_custom_instructions(self):
        custom_text = "Ensure all content is suitable for a young audience."
        messages = self.chat_instance._get_messages(
            topic="Test Topic",
            pages=3,
            language_str="English",
            output_format=self.output_format,
            custom_prompt_instructions=custom_text
        )
        self.assertIsInstance(messages, list)
        self.assertEqual(len(messages), 1)
        self.assertIn("Additional Custom Instructions:", messages[0]["content"])
        self.assertIn(custom_text, messages[0]["content"])
        # Check that the custom text is appended after the "Additional Custom Instructions:" line
        expected_ending = f"Additional Custom Instructions:\n{custom_text}"
        self.assertTrue(messages[0]["content"].strip().endswith(custom_text.strip())) # More robust check
        self.assertIn(expected_ending, messages[0]["content"])
        # Check for some part of the base prompt
        self.assertIn("I am preparing a presentation on Test Topic", messages[0]["content"])

    def test_get_messages_empty_custom_instructions(self):
        custom_text = "" # Empty string
        messages = self.chat_instance._get_messages(
            topic="Test Topic",
            pages=3,
            language_str="English",
            output_format=self.output_format,
            custom_prompt_instructions=custom_text
        )
        self.assertIsInstance(messages, list)
        self.assertEqual(len(messages), 1)
        self.assertNotIn("Additional Custom Instructions:", messages[0]["content"])
        # Check for some part of the base prompt
        self.assertIn("I am preparing a presentation on Test Topic", messages[0]["content"])

    def test_get_messages_whitespace_custom_instructions(self):
        custom_text = "   " # Whitespace only
        messages = self.chat_instance._get_messages(
            topic="Test Topic",
            pages=3,
            language_str="English",
            output_format=self.output_format,
            custom_prompt_instructions=custom_text
        )
        self.assertIsInstance(messages, list)
        self.assertEqual(len(messages), 1)
        self.assertNotIn("Additional Custom Instructions:", messages[0]["content"])
        # Check for some part of the base prompt
        self.assertIn("I am preparing a presentation on Test Topic", messages[0]["content"])

if __name__ == '__main__':
    unittest.main()
