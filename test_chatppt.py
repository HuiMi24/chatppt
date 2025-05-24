import unittest
import sys
from unittest.mock import patch, mock_open, Mock
import json 
import openai # Added for testing openai.api_base

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
            '--api_key', 'test_openai_key_direct',
            '--topic', 'OpenAI Direct Key Test',
        ]
        args = args_parser()
        self.assertEqual(args.model_provider, 'openai')
        self.assertEqual(args.model_name, 'gpt-3.5-turbo')
        self.assertEqual(args.api_key, 'test_openai_key_direct')
        self.assertEqual(args.topic, 'OpenAI Direct Key Test')
        self.assertIsNone(args.anthropic_api_key)
        self.assertIsNone(args.openai_api_base) # Check default

    @patch("builtins.open", new_callable=mock_open, read_data="test_openai_key_from_default_dot_token_file")
    def test_openai_args_key_from_default_dot_token_file(self, mock_file):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-4',
            '--topic', 'OpenAI Default .token File Key Test'
        ]
        args = args_parser()
        self.assertEqual(args.api_key, 'test_openai_key_from_default_dot_token_file')

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
        self.assertEqual(args.api_key, custom_path_key) 

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
        self.assertIsNone(args.openai_api_base)

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
        self.assertIsNone(args.openai_api_base)

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
        self.assertEqual(args.anthropic_api_key, 'test_anthropic_key_from_file')

    def test_error_openai_missing_api_key_file_not_found(self):
        sys.argv = ['chatppt.py', '--model_provider', 'openai', '--model_name', 'gpt-3.5-turbo', '--topic', 'Test']
        with patch("builtins.open", side_effect=FileNotFoundError):
            with self.assertRaises(SystemExit):
                args_parser()
    
    def test_error_anthropic_missing_api_key(self):
        sys.argv = ['chatppt.py', '--model_provider', 'anthropic', '--model_name', 'claude-opus', '--topic', 'Test']
        with self.assertRaises(SystemExit):
            args_parser()

    def test_error_missing_topic(self):
        sys.argv = ['chatppt.py', '--model_provider', 'openai', '--model_name', 'gpt-3.5-turbo', '--api_key', 'fake']
        with self.assertRaises(SystemExit):
            args_parser()

    def test_error_missing_model_name(self):
        sys.argv = ['chatppt.py', '--model_provider', 'openai', '--api_key', 'fake', '--topic', 'Test']
        with self.assertRaises(SystemExit):
            args_parser()

    def test_args_parser_with_openai_api_base(self):
        custom_base_url = "http://localhost:1234/v1"
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-3.5-turbo',
            '--api_key', 'test_key',
            '--topic', 'Test Topic',
            '--openai_api_base', custom_base_url
        ]
        args = args_parser()
        self.assertEqual(args.openai_api_base, custom_base_url)

    def test_args_parser_without_openai_api_base(self):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-3.5-turbo',
            '--api_key', 'test_key',
            '--topic', 'Test Topic'
        ]
        args = args_parser()
        self.assertIsNone(args.openai_api_base)


class TestChatPPTInitialization(unittest.TestCase):
    def test_init_openai(self):
        chat_instance = ChatPPT(model_provider="openai", api_key="fake_openai_key", model_name="gpt-4")
        self.assertEqual(chat_instance.api_key, "fake_openai_key")
        self.assertIsNone(chat_instance.openai_api_base) # Check default

    def test_init_ollama(self):
        chat_instance = ChatPPT(model_provider="ollama", api_key=None, model_name="llama3", ollama_url="http://fakeurl:11434")
        self.assertIsNone(chat_instance.openai_api_base)

    def test_init_anthropic(self):
        chat_instance = ChatPPT(model_provider="anthropic", api_key=None, model_name="claude-3", anthropic_api_key="fake_key")
        self.assertIsNone(chat_instance.openai_api_base)
        
    def test_init_with_openai_api_base(self):
        custom_base_url = "http://custom/v1"
        chat_instance = ChatPPT(
            model_provider="openai", 
            api_key="fake_key", 
            model_name="gpt-test", 
            openai_api_base=custom_base_url
        )
        self.assertEqual(chat_instance.openai_api_base, custom_base_url)

    def test_init_without_openai_api_base(self):
        chat_instance = ChatPPT(
            model_provider="openai", 
            api_key="fake_key", 
            model_name="gpt-test"
            # openai_api_base is not provided, should default to None in constructor
        )
        self.assertIsNone(chat_instance.openai_api_base)


class TestChatPPTPromptCustomization(unittest.TestCase):
    def setUp(self):
        self.chat_instance = ChatPPT(model_provider="test", api_key="dummy", model_name="test_model")
        self.output_format = self.chat_instance._get_output_format() 

    def test_get_messages_no_custom_instructions(self):
        messages = self.chat_instance._get_messages("Topic", 3, "English", self.output_format, None)
        self.assertNotIn("Additional Custom Instructions:", messages[0]["content"])

    def test_get_messages_with_custom_instructions(self):
        custom_text = "Young audience."
        messages = self.chat_instance._get_messages("Topic", 3, "English", self.output_format, custom_text)
        self.assertIn(f"Additional Custom Instructions:\n{custom_text}", messages[0]["content"])

    def test_get_messages_empty_custom_instructions(self):
        messages = self.chat_instance._get_messages("Topic", 3, "English", self.output_format, "")
        self.assertNotIn("Additional Custom Instructions:", messages[0]["content"])

    def test_get_messages_whitespace_custom_instructions(self):
        messages = self.chat_instance._get_messages("Topic", 3, "English", self.output_format, "   ")
        self.assertNotIn("Additional Custom Instructions:", messages[0]["content"])


class TestOpenAICustomBaseURL(unittest.TestCase):
    def setUp(self):
        self.original_openai_api_base = openai.api_base

    def tearDown(self):
        openai.api_base = self.original_openai_api_base

    @patch('openai.ChatCompletion.create')
    def test_custom_base_url_is_set_during_call_and_reset(self, mock_create):
        custom_url = "http://mycustomopenai/v1"
        
        # Configure the mock to check openai.api_base when it's called
        def side_effect_check_api_base(*args, **kwargs):
            self.assertEqual(openai.api_base, custom_url)
            mock_response = Mock()
            mock_response.choices = [Mock()]
            mock_response.choices[0].message.content = "mocked response"
            return mock_response
        mock_create.side_effect = side_effect_check_api_base

        chat_instance = ChatPPT(
            model_provider="openai",
            model_name="test-model",
            api_key="test_key",
            openai_api_base=custom_url
        )
        chat_instance._get_content(messages=[{"role": "user", "content": "Hello"}])
        
        self.assertEqual(openai.api_base, self.original_openai_api_base)
        mock_create.assert_called_once()

    @patch('openai.ChatCompletion.create')
    def test_no_custom_base_url_uses_default_or_original_and_resets(self, mock_create):
        # Store the very initial openai.api_base at the beginning of this test
        initial_api_base_for_test = openai.api_base
        
        # Expected base URL during the call if no custom one is provided
        # The code sets it to "https://api.openai.com/v1" if self.openai_api_base is None AND original openai.api_base was None
        # Otherwise, it uses the existing openai.api_base
        expected_during_call = "https://api.openai.com/v1" if initial_api_base_for_test is None else initial_api_base_for_test

        def side_effect_check_api_base(*args, **kwargs):
            self.assertEqual(openai.api_base, expected_during_call)
            mock_response = Mock()
            mock_response.choices = [Mock()]
            mock_response.choices[0].message.content = "mocked response"
            return mock_response
        mock_create.side_effect = side_effect_check_api_base

        chat_instance = ChatPPT(
            model_provider="openai",
            model_name="test-model",
            api_key="test_key",
            openai_api_base=None # No custom base URL
        )
        chat_instance._get_content(messages=[{"role": "user", "content": "Hello"}])
        
        # Should be reset to what it was at the start of this specific test
        self.assertEqual(openai.api_base, initial_api_base_for_test)
        mock_create.assert_called_once()

    @patch('openai.ChatCompletion.create')
    def test_custom_base_url_is_empty_string(self, mock_create):
        # Store the very initial openai.api_base
        initial_api_base_for_test = openai.api_base
        expected_during_call = "https://api.openai.com/v1" if initial_api_base_for_test is None else initial_api_base_for_test

        def side_effect_check_api_base(*args, **kwargs):
            self.assertEqual(openai.api_base, expected_during_call)
            mock_response = Mock()
            mock_response.choices = [Mock()]
            mock_response.choices[0].message.content = "mocked response"
            return mock_response
        mock_create.side_effect = side_effect_check_api_base

        chat_instance = ChatPPT(
            model_provider="openai",
            model_name="test-model",
            api_key="test_key",
            openai_api_base="   " # Empty string (after strip)
        )
        chat_instance._get_content(messages=[{"role": "user", "content": "Hello"}])
        
        self.assertEqual(openai.api_base, initial_api_base_for_test)
        mock_create.assert_called_once()


if __name__ == '__main__':
    unittest.main()
