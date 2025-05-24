# -*- coding: utf-8 -*-
"""
Unit tests for ChatPPT application.

This module contains unit tests for various components of the ChatPPT application,
including argument parsing, class initialization, prompt generation logic,
and interaction with external services (mocked).
"""

import importlib
import json
import os
import sys
import unittest
from unittest.mock import Mock, mock_open, patch

import openai

# Assuming chatppt.py and config.py are in the same directory or accessible
import config  # Import config to be reloaded
from chatppt import ChatPPT, args_parser


class TestConfigLoading(unittest.TestCase):
    """Tests for the loading of configurations from environment variables."""

    @patch.dict(os.environ, {
        "OPENAI_API_KEY": "env_openai_key",
        "ANTHROPIC_API_KEY": "env_anthropic_key",
        "GROQ_API_KEY": "env_groq_key",
        "OPENAI_API_BASE": "env_openai_base",
        "GROQ_API_BASE": "env_groq_base",
        "GROQ_DEFAULT_MODEL": "env_groq_model"
    })
    def test_load_config_from_env(self):
        """Verify that config variables are correctly loaded from mocked environment."""
        importlib.reload(config)  # Reload config to pick up mocked env vars
        self.assertEqual(config.OPENAI_API_KEY, "env_openai_key")
        self.assertEqual(config.ANTHROPIC_API_KEY, "env_anthropic_key")
        self.assertEqual(config.GROQ_API_KEY, "env_groq_key")
        self.assertEqual(config.OPENAI_API_BASE, "env_openai_base")
        self.assertEqual(config.GROQ_API_BASE, "env_groq_base")
        self.assertEqual(config.GROQ_DEFAULT_MODEL, "env_groq_model")

    @patch.dict(os.environ, {}, clear=True)  # Clear all env vars for this test
    def test_load_config_no_env(self):
        """Verify that config variables are None when no environment variables are set."""
        importlib.reload(config)
        self.assertIsNone(config.OPENAI_API_KEY)
        self.assertIsNone(config.ANTHROPIC_API_KEY)
        self.assertIsNone(config.GROQ_API_KEY)
        self.assertIsNone(config.OPENAI_API_BASE)
        self.assertIsNone(config.GROQ_API_BASE)
        self.assertIsNone(config.GROQ_DEFAULT_MODEL)


class TestArgsParser(unittest.TestCase):
    """Tests for the command-line argument parsing logic."""

    def setUp(self):
        """Set up test environment before each test."""
        self.original_argv = sys.argv
        # Ensure config is reloaded with no env vars before each argparser test
        # to avoid interference from TestConfigLoading's mocks
        with patch.dict(os.environ, {}, clear=True):
            importlib.reload(config)

    def tearDown(self):
        """Clean up test environment after each test."""
        sys.argv = self.original_argv
        # Restore config to its original state after tests
        with patch.dict(os.environ, {}, clear=True):
            importlib.reload(config)

    def test_openai_args_direct_key(self):
        """Test parsing OpenAI arguments with a direct API key."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'openai', '--model_name', 'gpt-3.5-turbo',
            '--api_key', 'test_openai_key_direct', '--topic', 'Test'
        ]
        args = args_parser()
        self.assertEqual(args.api_key, 'test_openai_key_direct')
        self.assertIsNone(args.anthropic_api_key)
        self.assertIsNone(args.groq_api_key)
        self.assertIsNone(args.openai_api_base)
        self.assertIsNone(args.audience)

    def test_groq_args_with_key_and_model(self):
        """Test parsing Groq arguments with API key and model name."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'groq', '--model_name', 'gmodel-test',
            '--groq_api_key', 'gkey-test', '--topic', 'Groq Test'
        ]
        args = args_parser()
        self.assertEqual(args.groq_api_key, 'gkey-test')
        self.assertEqual(args.model_name, 'gmodel-test')

    def test_groq_args_without_key(self):
        """Test parsing Groq arguments without API key (should be None)."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'groq',
            '--model_name', 'gmodel-test', '--topic', 'Groq Test No Key'
        ]
        args = args_parser()
        self.assertIsNone(args.groq_api_key)

    def test_groq_args_without_model_name(self):
        """Test parsing Groq arguments without model name (should be None)."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'groq',
            '--groq_api_key', 'gkey-test', '--topic', 'Groq Test No Model'
        ]
        args = args_parser()
        self.assertIsNone(args.model_name)  # main() will handle default

    def test_args_with_audience(self):
        """Test parsing the --audience argument."""
        sys.argv = ['chatppt.py', '-t', 'Topic', '-n', 'model', '--audience', 'students']
        args = args_parser()
        self.assertEqual(args.audience, 'students')

    def test_args_without_audience(self):
        """Test that args.audience is None when the argument is not provided."""
        sys.argv = ['chatppt.py', '-t', 'Topic', '-n', 'model']
        args = args_parser()
        self.assertIsNone(args.audience)

    @patch("builtins.open", new_callable=mock_open,
           read_data="test_openai_key_from_default_dot_token_file")
    def test_openai_args_key_from_default_dot_token_file(self, _mock_file):
        """Test OpenAI args when API key is not provided (expecting .token logic in main)."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'openai',
            '--model_name', 'gpt-4', '--topic', 'Test'
        ]
        args = args_parser()  # --api_key is not provided, defaults to None in argparse
        self.assertIsNone(args.api_key)

    def test_openai_args_key_as_custom_path_not_default_dot_token(self):
        """Test OpenAI args when API key is a custom path (not default .token)."""
        custom_path_key = 'path/to/my_openai_key.txt'
        sys.argv = [
            'chatppt.py', '--model_provider', 'openai', '--model_name', 'gpt-4',
            '--api_key', custom_path_key, '--topic', 'Test'
        ]
        args = args_parser()
        self.assertEqual(args.api_key, custom_path_key)

    def test_ollama_args(self):
        """Test parsing Ollama arguments."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'ollama', '--model_name', 'llama3',
            '--ollama_url', 'http://localhost:1111', '--topic', 'Test'
        ]
        args = args_parser()
        self.assertIsNone(args.api_key)

    def test_anthropic_args_direct_key(self):
        """Test parsing Anthropic arguments with a direct API key."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'anthropic', '--model_name', 'claude-3',
            '--anthropic_api_key', 'key_direct', '--topic', 'Test'
        ]
        args = args_parser()
        self.assertEqual(args.anthropic_api_key, 'key_direct')
        self.assertIsNone(args.api_key)

    @patch("builtins.open", new_callable=mock_open, read_data="test_anthropic_key_from_file")
    def test_anthropic_args_key_from_file(self, _mock_file):
        """Test Anthropic args when API key is a path (main handles actual reading)."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'anthropic', '--model_name', 'claude-3',
            '--anthropic_api_key', 'path/file.txt', '--topic', 'Test'
        ]
        args = args_parser()
        self.assertEqual(args.anthropic_api_key, 'path/file.txt')

    def test_error_missing_topic(self):
        """Test that SystemExit is raised if the required --topic argument is missing."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'openai',
            '--model_name', 'gpt-3.5-turbo', '--api_key', 'fake'
        ]
        with self.assertRaises(SystemExit):
            args_parser()

    def test_missing_model_name_argparse_level(self):
        """Test that model_name is None at argparse level if not provided."""
        sys.argv = [
            'chatppt.py', '--model_provider', 'openai',
            '--api_key', 'fake', '--topic', 'Test'
        ]
        args = args_parser()
        self.assertIsNone(args.model_name) # main() handles requirement based on provider


class TestChatPPTInitialization(unittest.TestCase):
    """Tests for the initialization of the ChatPPT class."""

    def test_init_groq_provider(self):
        """Test ChatPPT initialization with Groq provider details."""
        inst = ChatPPT("groq", None, "gmodel", groq_api_key="gkey")
        self.assertEqual(inst.groq_api_key, "gkey")
        self.assertEqual(inst.model_provider, "groq")

    def test_init_openai(self):
        """Test ChatPPT initialization for OpenAI."""
        inst = ChatPPT("openai", "fake_key", "gpt-4")
        self.assertEqual(inst.api_key, "fake_key")
        self.assertIsNone(inst.openai_api_base)

    def test_init_ollama(self):
        """Test ChatPPT initialization for Ollama."""
        inst = ChatPPT("ollama", None, "llama3", ollama_url="http://fakeurl:11434")
        self.assertEqual(inst.ollama_url, "http://fakeurl:11434")
        self.assertIsNone(inst.openai_api_base)

    def test_init_anthropic(self):
        """Test ChatPPT initialization for Anthropic."""
        inst = ChatPPT("anthropic", None, "claude-3", anthropic_api_key="fake_key")
        self.assertEqual(inst.anthropic_api_key, "fake_key")
        self.assertIsNone(inst.openai_api_base)

    def test_init_with_openai_api_base(self):
        """Test ChatPPT initialization with a custom OpenAI API base URL."""
        custom_base = "http://custom/v1"
        inst = ChatPPT("openai", "fake_key", "gpt-test", openai_api_base=custom_base)
        self.assertEqual(inst.openai_api_base, custom_base)

    def test_init_without_openai_api_base(self):
        """Test ChatPPT initialization without a custom OpenAI API base URL."""
        inst = ChatPPT("openai", "fake_key", "gpt-test")
        self.assertIsNone(inst.openai_api_base)


class TestChatPPTPromptCustomization(unittest.TestCase):
    """Tests for prompt generation logic, including audience and custom instructions."""

    def setUp(self):
        """Set up a ChatPPT instance for prompt testing."""
        self.chat_instance = ChatPPT("test_provider", "dummy_api_key", "test_model")

    def test_get_messages_no_custom_no_audience(self):
        """Test _get_messages with no custom instructions or audience."""
        # pylint: disable=protected-access
        messages = self.chat_instance._get_messages(
            "Topic", 3, "English", None, None
        )
        content = messages[0]["content"]
        self.assertNotIn("Additional Custom Instructions:", content)
        self.assertNotIn("The target audience for this presentation is:", content)

    def test_get_messages_audience_only(self):
        """Test _get_messages with audience specified but no custom instructions."""
        audience = "Software Engineers"
        # pylint: disable=protected-access
        messages = self.chat_instance._get_messages(
            "Topic", 3, "English", None, audience
        )
        content = messages[0]["content"]
        self.assertIn(f"The target audience for this presentation is: {audience}", content)
        self.assertNotIn("Additional Custom Instructions:", content)

    def test_get_messages_custom_only(self):
        """Test _get_messages with custom instructions but no audience."""
        custom_text = "Focus on practical examples."
        # pylint: disable=protected-access
        messages = self.chat_instance._get_messages(
            "Topic", 3, "English", custom_text, None
        )
        content = messages[0]["content"]
        self.assertIn(f"Additional Custom Instructions:\n{custom_text}", content)
        self.assertNotIn("target audience for this presentation is:", content)

    def test_get_messages_audience_and_custom(self):
        """Test _get_messages with both audience and custom instructions."""
        audience = "University Students"
        custom_text = "Include a quiz at the end."
        # pylint: disable=protected-access
        messages = self.chat_instance._get_messages(
            "Topic", 3, "English", custom_text, audience
        )
        content = messages[0]["content"]
        self.assertIn(f"The target audience for this presentation is: {audience}", content)
        self.assertIn(f"Additional Custom Instructions:\n{custom_text}", content)
        self.assertTrue(content.find(audience) < content.find(custom_text))

    def test_get_messages_empty_audience_or_custom(self):
        """Test _get_messages with empty strings for audience and custom instructions."""
        # pylint: disable=protected-access
        messages = self.chat_instance._get_messages(
            "Topic", 3, "English", "", ""
        )
        content = messages[0]["content"]
        self.assertNotIn("Additional Custom Instructions:", content)
        self.assertNotIn("The target audience for this presentation is:", content)

    def test_get_messages_whitespace_audience_or_custom(self):
        """Test _get_messages with whitespace-only audience and custom instructions."""
        # pylint: disable=protected-access
        messages = self.chat_instance._get_messages(
            "Topic", 3, "English", "   ", "   "
        )
        content = messages[0]["content"]
        self.assertNotIn("Additional Custom Instructions:", content)
        self.assertNotIn("The target audience for this presentation is:", content)


class TestOpenAICustomBaseURL(unittest.TestCase):
    """Tests for custom OpenAI API base URL functionality."""

    def setUp(self):
        """Save original openai.api_base."""
        self.original_openai_api_base = openai.api_base

    def tearDown(self):
        """Restore original openai.api_base."""
        openai.api_base = self.original_openai_api_base

    @patch('openai.ChatCompletion.create')
    def test_custom_base_url_set_and_reset(self, mock_create: Mock):
        """Verify custom base URL is used during API call and then reset."""
        custom_url = "http://mycustomopenai/v1"
        # pylint: disable=protected-access
        def side_effect_check_api_base(*_args, **_kwargs):
            self.assertEqual(openai.api_base, custom_url)
            mock_response = Mock()
            mock_response.choices = [Mock(message=Mock(content="response"))]
            return mock_response
        mock_create.side_effect = side_effect_check_api_base

        chat_ppt = ChatPPT("openai", "test_key", "model", openai_api_base=custom_url)
        chat_ppt._get_content([{"role": "user", "content": "Hi"}])
        self.assertEqual(openai.api_base, self.original_openai_api_base)
        mock_create.assert_called_once()

    @patch('openai.ChatCompletion.create')
    def test_no_custom_base_url_default_behavior(self, mock_create: Mock):
        """Verify default/original base URL behavior when no custom URL is set."""
        initial_api_base_for_test = openai.api_base
        expected_during_call = ("https://api.openai.com/v1"
                                if initial_api_base_for_test is None
                                else initial_api_base_for_test)
        # pylint: disable=protected-access
        def side_effect_check_api_base(*_args, **_kwargs):
            self.assertEqual(openai.api_base, expected_during_call)
            mock_response = Mock()
            mock_response.choices = [Mock(message=Mock(content="response"))]
            return mock_response
        mock_create.side_effect = side_effect_check_api_base

        chat_ppt = ChatPPT("openai", "test_key", "model", openai_api_base=None)
        chat_ppt._get_content([{"role": "user", "content": "Hi"}])
        self.assertEqual(openai.api_base, initial_api_base_for_test)

    @patch('openai.ChatCompletion.create')
    def test_empty_custom_base_url_default_behavior(self, mock_create: Mock):
        """Verify behavior when custom base URL is an empty string."""
        initial_api_base_for_test = openai.api_base
        expected_during_call = ("https://api.openai.com/v1"
                                if initial_api_base_for_test is None
                                else initial_api_base_for_test)
        # pylint: disable=protected-access
        def side_effect_check_api_base(*_args, **_kwargs):
            self.assertEqual(openai.api_base, expected_during_call)
            mock_response = Mock()
            mock_response.choices = [Mock(message=Mock(content="response"))]
            return mock_response
        mock_create.side_effect = side_effect_check_api_base
        chat_ppt = ChatPPT("openai", "test_key", "model", openai_api_base="   ")
        chat_ppt._get_content([{"role": "user", "content": "Hi"}])
        self.assertEqual(openai.api_base, initial_api_base_for_test)


class TestGroqContentGeneration(unittest.TestCase):
    """Tests specific to Groq provider content generation."""

    def setUp(self):
        """Set up mocks and save original OpenAI API details."""
        self.original_openai_api_key = openai.api_key
        self.original_openai_api_base = openai.api_base
        # Patch chatppt.GROQ_API_BASE directly as it's imported at module level
        self.groq_base_patcher = patch('chatppt.GROQ_API_BASE', None)
        self.mock_groq_base = self.groq_base_patcher.start()


    def tearDown(self):
        """Restore original OpenAI API details and stop patches."""
        openai.api_key = self.original_openai_api_key
        openai.api_base = self.original_openai_api_base
        self.groq_base_patcher.stop()

    @patch('openai.ChatCompletion.create')
    def test_groq_uses_key_and_default_base(self, mock_chat_completion: Mock):
        """Verify Groq uses its key and the default Groq base URL."""
        test_groq_key = "fake_groq_key"
        test_model = "mixtral-test"
        expected_groq_base = "https://api.groq.com/openai/v1" # Default in chatppt.py

        # pylint: disable=protected-access
        def side_effect_check_api_details(*_args, **_kwargs):
            self.assertEqual(openai.api_key, test_groq_key)
            self.assertEqual(openai.api_base, expected_groq_base)
            mock_response = Mock()
            mock_response.choices = [Mock(message=Mock(content="Groq response"))]
            return mock_response
        mock_chat_completion.side_effect = side_effect_check_api_details

        chat_instance = ChatPPT(
            "groq", None, test_model, groq_api_key=test_groq_key
        )
        chat_instance._get_content([{"role": "user", "content": "Hello Groq"}])
        mock_chat_completion.assert_called_once_with(
            model=test_model, messages=[{"role": "user", "content": "Hello Groq"}]
        )
        self.assertEqual(openai.api_key, self.original_openai_api_key)
        self.assertEqual(openai.api_base, self.original_openai_api_base)

    @patch('openai.ChatCompletion.create')
    def test_groq_uses_key_and_config_base(self, mock_chat_completion: Mock):
        """Verify Groq uses its key and a Groq base URL from config."""
        test_groq_key = "fake_groq_key_2"
        test_model = "llama3-test"
        custom_groq_base_from_config = "http://custom.groq.api/v1"

        with patch('chatppt.GROQ_API_BASE', custom_groq_base_from_config):
            # pylint: disable=protected-access
            def side_effect_check_api_details(*_args, **_kwargs):
                self.assertEqual(openai.api_key, test_groq_key)
                self.assertEqual(openai.api_base, custom_groq_base_from_config)
                mock_response = Mock()
                mock_response.choices = [Mock(message=Mock(content="Groq custom base response"))]
                return mock_response
            mock_chat_completion.side_effect = side_effect_check_api_details

            chat_instance = ChatPPT(
                "groq", None, test_model, groq_api_key=test_groq_key
            )
            chat_instance._get_content(
                [{"role": "user", "content": "Hello Groq Custom Base"}]
            )
        self.assertEqual(openai.api_key, self.original_openai_api_key)
        self.assertEqual(openai.api_base, self.original_openai_api_base)


class TestRegenerateSinglePage(unittest.TestCase):
    """Tests for the regenerate_single_page method of ChatPPT."""

    def setUp(self):
        """Set up common test data for single page regeneration."""
        self.chat_instance = ChatPPT("test_provider", "dummy_key", "test_model")
        self.original_topic = "Mars Colonization"
        self.original_language = "English"
        self.page_data_to_edit = {
            "title": "Old Page Title",
            "content": [{"title": "Old Bullet 1", "description": "Old Desc 1"}]
        }
        # pylint: disable=protected-access
        self.single_page_format_example_str = json.dumps(
            self.chat_instance._get_single_page_output_format_example(), indent=2
        )

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_basic_flow(self, mock_get_content: Mock):
        """Test basic flow of regenerating a page with new instructions."""
        mock_llm_response = ('{"title": "Refined Title", "content": '
                             '[{"title": "New Bullet", "description": "New Desc"}]}')
        mock_get_content.return_value = mock_llm_response
        new_instructions = "Make it more concise."

        result = self.chat_instance.regenerate_single_page(
            self.original_topic, self.original_language,
            self.page_data_to_edit, new_instructions
        )
        mock_get_content.assert_called_once()
        prompt_content = mock_get_content.call_args[0][0][0]['content']

        self.assertIn(self.original_topic, prompt_content)
        self.assertIn(self.original_language, prompt_content)
        self.assertIn(json.dumps(self.page_data_to_edit, indent=2), prompt_content)
        self.assertIn(new_instructions, prompt_content)
        self.assertIn("Please refine this slide based on the following instructions:", prompt_content)
        self.assertIn(self.single_page_format_example_str, prompt_content)
        self.assertEqual(result, json.loads(mock_llm_response))

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_no_instructions(self, mock_get_content: Mock):
        """Test page regeneration when no new instructions are provided."""
        mock_get_content.return_value = ('{"title": "Reviewed Title", "content": []}')
        self.chat_instance.regenerate_single_page(
            self.original_topic, self.original_language, self.page_data_to_edit, None
        )
        prompt_content = mock_get_content.call_args[0][0][0]['content']
        self.assertIn("Please review and refine the content of this slide", prompt_content)
        self.assertNotIn("Please refine this slide based on the following instructions:", prompt_content)

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_empty_instructions(self, mock_get_content: Mock):
        """Test page regeneration with empty string instructions."""
        mock_get_content.return_value = ('{"title": "Reviewed Title 2", "content": []}')
        self.chat_instance.regenerate_single_page(
            self.original_topic, self.original_language, self.page_data_to_edit, "   "
        )
        prompt_content = mock_get_content.call_args[0][0][0]['content']
        self.assertIn("Please review and refine the content of this slide", prompt_content)

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_parsing_error_malformed_json(self, mock_get_content: Mock):
        """Test error handling for malformed JSON response from LLM."""
        mock_get_content.return_value = '{"title": "Bad JSON", "content": [}'
        with self.assertRaisesRegex(ValueError, "Error parsing regenerated page content"):
            self.chat_instance.regenerate_single_page(
                self.original_topic, self.original_language, self.page_data_to_edit
            )

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_validation_error_wrong_structure(self, mock_get_content: Mock):
        """Test error handling for JSON response with incorrect structure."""
        mock_get_content.return_value = '{"page_title": "Wrong Key", "bullets": []}'
        with self.assertRaisesRegex(ValueError, "LLM returned JSON but not in the expected page format"):
            self.chat_instance.regenerate_single_page(
                self.original_topic, self.original_language, self.page_data_to_edit
            )

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_validation_error_malformed_bullets(self, mock_get_content: Mock):
        """Test error handling for JSON response with malformed bullet points."""
        mock_get_content.return_value = ('{"title": "Page", "content": '
                                         '[{"name": "Missing title/desc"}]}')
        with self.assertRaisesRegex(ValueError, "LLM returned JSON with malformed bullet points"):
            self.chat_instance.regenerate_single_page(
                self.original_topic, self.original_language, self.page_data_to_edit
            )


if __name__ == '__main__':
    unittest.main()
