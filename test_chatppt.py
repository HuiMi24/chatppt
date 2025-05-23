import unittest
import sys
from unittest.mock import patch, mock_open

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

    # Test for when --api_key is NOT provided, so it defaults to ".token"
    # and args_parser tries to read this file.
    @patch("builtins.open", new_callable=mock_open, read_data="test_openai_key_from_default_dot_token_file")
    def test_openai_args_key_from_default_dot_token_file(self, mock_file):
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-4',
            # No --api_key provided, so it defaults to ".token"
            '--topic', 'OpenAI Default .token File Key Test'
        ]
        args = args_parser()
        self.assertEqual(args.model_provider, 'openai')
        self.assertEqual(args.model_name, 'gpt-4')
        # The args_parser logic should try to open ".token" and replace args.api_key
        self.assertEqual(args.api_key, 'test_openai_key_from_default_dot_token_file')
        self.assertEqual(args.topic, 'OpenAI Default .token File Key Test')

    # Test for when --api_key IS provided with a specific path (NOT ".token")
    # In this case, the current args_parser logic should treat the path itself as the key.
    def test_openai_args_key_as_custom_path_not_default_dot_token(self):
        custom_path_key = 'path/to/my_openai_key.txt'
        sys.argv = [
            'chatppt.py',
            '--model_provider', 'openai',
            '--model_name', 'gpt-4',
            '--api_key', custom_path_key, # Specific path provided
            '--topic', 'OpenAI Custom Path As Key Test'
        ]
        # No mock for open here, as it shouldn't be called to read the file
        # if the provided api_key is not ".token"
        args = args_parser()
        self.assertEqual(args.model_provider, 'openai')
        self.assertEqual(args.model_name, 'gpt-4')
        # According to current args_parser, if api_key is not ".token", it's used directly.
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
            '--anthropic_api_key', 'path/to/fake_anthropic_key.txt', # This path will be "opened"
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
            # API key defaults to '.token'
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

if __name__ == '__main__':
    unittest.main()
