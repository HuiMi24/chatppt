import unittest
import sys
from unittest.mock import patch, mock_open, Mock
import json 
import openai 
import importlib # Added for reloading config
import os # For patch.dict(os.environ)

# Assuming chatppt.py and config.py are in the same directory or accessible
from chatppt import args_parser, ChatPPT
import config # Import config to be reloaded

class TestConfigLoading(unittest.TestCase):
    @patch.dict(os.environ, {
        "OPENAI_API_KEY": "env_openai_key",
        "ANTHROPIC_API_KEY": "env_anthropic_key",
        "GROQ_API_KEY": "env_groq_key",
        "OPENAI_API_BASE": "env_openai_base",
        "GROQ_API_BASE": "env_groq_base",
        "GROQ_DEFAULT_MODEL": "env_groq_model"
    })
    def test_load_config_from_env(self):
        importlib.reload(config) # Reload config to pick up mocked env vars
        self.assertEqual(config.OPENAI_API_KEY, "env_openai_key")
        self.assertEqual(config.ANTHROPIC_API_KEY, "env_anthropic_key")
        self.assertEqual(config.GROQ_API_KEY, "env_groq_key")
        self.assertEqual(config.OPENAI_API_BASE, "env_openai_base")
        self.assertEqual(config.GROQ_API_BASE, "env_groq_base")
        self.assertEqual(config.GROQ_DEFAULT_MODEL, "env_groq_model")

    @patch.dict(os.environ, {}, clear=True) # Clear all env vars for this test
    def test_load_config_no_env(self):
        importlib.reload(config)
        self.assertIsNone(config.OPENAI_API_KEY)
        self.assertIsNone(config.ANTHROPIC_API_KEY)
        self.assertIsNone(config.GROQ_API_KEY)
        self.assertIsNone(config.OPENAI_API_BASE)
        self.assertIsNone(config.GROQ_API_BASE)
        self.assertIsNone(config.GROQ_DEFAULT_MODEL)


class TestArgsParser(unittest.TestCase):
    def setUp(self):
        self.original_argv = sys.argv
        with patch.dict(os.environ, {}, clear=True):
            importlib.reload(config)

    def tearDown(self):
        sys.argv = self.original_argv
        with patch.dict(os.environ, {}, clear=True):
             importlib.reload(config)

    def test_openai_args_direct_key(self):
        sys.argv = ['chatppt.py', '--model_provider', 'openai', '--model_name', 'gpt-3.5-turbo', '--api_key', 'test_openai_key_direct','--topic', 'Test']
        args = args_parser()
        self.assertEqual(args.api_key, 'test_openai_key_direct')
        self.assertIsNone(args.anthropic_api_key)
        self.assertIsNone(args.groq_api_key)
        self.assertIsNone(args.openai_api_base)
        self.assertIsNone(args.audience)

    def test_groq_args_with_key_and_model(self):
        sys.argv = ['chatppt.py', '--model_provider', 'groq', '--model_name', 'gmodel-test', '--groq_api_key', 'gkey-test', '--topic', 'Groq Test']
        args = args_parser()
        self.assertEqual(args.groq_api_key, 'gkey-test')

    def test_groq_args_without_key(self):
        sys.argv = ['chatppt.py', '--model_provider', 'groq', '--model_name', 'gmodel-test', '--topic', 'Groq Test No Key']
        args = args_parser()
        self.assertIsNone(args.groq_api_key)

    def test_groq_args_without_model_name(self):
        sys.argv = ['chatppt.py', '--model_provider', 'groq', '--groq_api_key', 'gkey-test', '--topic', 'Groq Test No Model']
        args = args_parser()
        self.assertIsNone(args.model_name)

    def test_args_with_audience(self):
        sys.argv = ['chatppt.py', '-t', 'Topic', '-n', 'model', '--audience', 'students']
        args = args_parser()
        self.assertEqual(args.audience, 'students')

    def test_args_without_audience(self):
        sys.argv = ['chatppt.py', '-t', 'Topic', '-n', 'model']
        args = args_parser()
        self.assertIsNone(args.audience)
        
    @patch("builtins.open", new_callable=mock_open, read_data="test_openai_key_from_default_dot_token_file")
    def test_openai_args_key_from_default_dot_token_file(self, mock_file):
        sys.argv = ['chatppt.py', '--model_provider', 'openai', '--model_name', 'gpt-4', '--topic', 'Test']
        args = args_parser() # --api_key is not provided, defaults to None in argparse
        # The resolution of .token file happens in main(), not args_parser()
        self.assertIsNone(args.api_key) 

    def test_openai_args_key_as_custom_path_not_default_dot_token(self):
        custom_path_key = 'path/to/my_openai_key.txt'
        sys.argv = ['chatppt.py', '--model_provider', 'openai', '--model_name', 'gpt-4', '--api_key', custom_path_key, '--topic', 'Test']
        args = args_parser()
        self.assertEqual(args.api_key, custom_path_key) 

    def test_ollama_args(self):
        sys.argv = ['chatppt.py', '--model_provider', 'ollama', '--model_name', 'llama3', '--ollama_url', 'http://localhost:1111', '--topic', 'Test']
        args = args_parser()
        self.assertIsNone(args.api_key) 

    def test_anthropic_args_direct_key(self):
        sys.argv = ['chatppt.py', '--model_provider', 'anthropic', '--model_name', 'claude-3', '--anthropic_api_key', 'key_direct', '--topic', 'Test']
        args = args_parser()
        self.assertEqual(args.anthropic_api_key, 'key_direct')
        self.assertIsNone(args.api_key) 

    @patch("builtins.open", new_callable=mock_open, read_data="test_anthropic_key_from_file")
    def test_anthropic_args_key_from_file(self, mock_file):
        sys.argv = ['chatppt.py', '--model_provider', 'anthropic', '--model_name', 'claude-3', '--anthropic_api_key', 'path/file.txt', '--topic', 'Test']
        args = args_parser()
        self.assertEqual(args.anthropic_api_key, 'path/file.txt') # main() handles reading

    def test_error_missing_topic(self): # This is still raised by argparse
        sys.argv = ['chatppt.py', '--model_provider', 'openai', '--model_name', 'gpt-3.5-turbo', '--api_key', 'fake']
        with self.assertRaises(SystemExit): args_parser()

    # model_name is now not required by argparse, but by main()
    def test_missing_model_name_argparse_level(self):
        sys.argv = ['chatppt.py', '--model_provider', 'openai', '--api_key', 'fake', '--topic', 'Test']
        args = args_parser()
        self.assertIsNone(args.model_name)

class TestChatPPTInitialization(unittest.TestCase):
    def test_init_groq_provider(self):
        inst = ChatPPT("groq", None, "gmodel", groq_api_key="gkey")
        self.assertEqual(inst.groq_api_key, "gkey")
    def test_init_openai(self):
        inst = ChatPPT("openai", "fake_key", "gpt-4")
        self.assertEqual(inst.api_key, "fake_key")
        self.assertIsNone(inst.openai_api_base) 
    def test_init_ollama(self):
        inst = ChatPPT("ollama", None, "llama3", ollama_url="http://fakeurl:11434")
        self.assertIsNone(inst.openai_api_base)
    def test_init_anthropic(self):
        inst = ChatPPT("anthropic", None, "claude-3", anthropic_api_key="fake_key")
        self.assertIsNone(inst.openai_api_base)
    def test_init_with_openai_api_base(self):
        inst = ChatPPT("openai", "fake_key", "gpt-test", openai_api_base="http://custom/v1")
        self.assertEqual(inst.openai_api_base, "http://custom/v1")
    def test_init_without_openai_api_base(self):
        inst = ChatPPT("openai", "fake_key", "gpt-test")
        self.assertIsNone(inst.openai_api_base)

class TestChatPPTPromptCustomization(unittest.TestCase):
    def setUp(self):
        self.chat_instance = ChatPPT("test", "dummy_api_key", "test_model")
        self.output_format = self.chat_instance._get_output_format() 
    def test_get_messages_no_custom_no_audience(self):
        m = self.chat_instance._get_messages("T", 3, "L", self.output_format, None, None)[0]["content"]
        self.assertNotIn("Additional Custom Instructions:", m)
        self.assertNotIn("The target audience for this presentation is:", m)
    def test_get_messages_audience_only(self):
        m = self.chat_instance._get_messages("T", 3, "L", self.output_format, None, "A")[0]["content"]
        self.assertIn("target audience for this presentation is: A", m)
        self.assertNotIn("Additional Custom Instructions:", m)
    def test_get_messages_custom_only(self):
        m = self.chat_instance._get_messages("T", 3, "L", self.output_format, "C", None)[0]["content"]
        self.assertIn("Additional Custom Instructions:\nC", m)
        self.assertNotIn("target audience for this presentation is:", m)
    def test_get_messages_audience_and_custom(self):
        m = self.chat_instance._get_messages("T", 3, "L", self.output_format, "C", "A")[0]["content"]
        self.assertIn("target audience for this presentation is: A", m)
        self.assertIn("Additional Custom Instructions:\nC", m)
        self.assertTrue(m.find("target audience") < m.find("Additional Custom Instructions:"))
    def test_get_messages_empty_audience_or_custom(self):
        m = self.chat_instance._get_messages("T", 3, "L", self.output_format, "", "")[0]["content"]
        self.assertNotIn("Additional Custom Instructions:", m)
        self.assertNotIn("The target audience for this presentation is:", m)
    def test_get_messages_whitespace_audience_or_custom(self):
        m = self.chat_instance._get_messages("T", 3, "L", self.output_format, "   ", "   ")[0]["content"]
        self.assertNotIn("Additional Custom Instructions:", m)
        self.assertNotIn("The target audience for this presentation is:", m)

class TestOpenAICustomBaseURL(unittest.TestCase): 
    def setUp(self): self.original_openai_api_base = openai.api_base
    def tearDown(self): openai.api_base = self.original_openai_api_base
    @patch('openai.ChatCompletion.create')
    def test_custom_base_url_set_and_reset(self, mock_create):
        custom_url = "http://mycustomopenai/v1"
        mock_create.side_effect = lambda *a, **kw: (self.assertEqual(openai.api_base, custom_url), Mock(choices=[Mock(message=Mock(content="r"))]))[1]
        ChatPPT("openai", "k", "m", openai_api_base=custom_url)._get_content([{"role":"u","content":"h"}])
        self.assertEqual(openai.api_base, self.original_openai_api_base)
    @patch('openai.ChatCompletion.create')
    def test_no_custom_base_url_default_behavior(self, mock_create):
        initial = openai.api_base
        expected = "https://api.openai.com/v1" if initial is None else initial
        mock_create.side_effect = lambda *a, **kw: (self.assertEqual(openai.api_base, expected), Mock(choices=[Mock(message=Mock(content="r"))]))[1]
        ChatPPT("openai", "k", "m", openai_api_base=None)._get_content([{"role":"u","content":"h"}])
        self.assertEqual(openai.api_base, initial)
    @patch('openai.ChatCompletion.create')
    def test_empty_custom_base_url_default_behavior(self, mock_create):
        initial = openai.api_base
        expected = "https://api.openai.com/v1" if initial is None else initial
        mock_create.side_effect = lambda *a, **kw: (self.assertEqual(openai.api_base, expected), Mock(choices=[Mock(message=Mock(content="r"))]))[1]
        ChatPPT("openai", "k", "m", openai_api_base="   ")._get_content([{"role":"u","content":"h"}])
        self.assertEqual(openai.api_base, initial)

class TestGroqContentGeneration(unittest.TestCase):
    def setUp(self):
        self.original_openai_api_key = openai.api_key
        self.original_openai_api_base = openai.api_base
        self.config_patcher = patch.multiple(config, GROQ_API_BASE=None)
        self.mocked_config = self.config_patcher.start()
    def tearDown(self):
        openai.api_key = self.original_openai_api_key
        openai.api_base = self.original_openai_api_base
        self.config_patcher.stop()
    @patch('openai.ChatCompletion.create')
    def test_groq_uses_key_and_default_base(self, mock_cc):
        key, model, base = "gkey", "m", "https://api.groq.com/openai/v1"
        mock_cc.side_effect = lambda *a, **kw: (self.assertEqual(openai.api_key, key), self.assertEqual(openai.api_base, base), Mock(choices=[Mock(message=Mock(content="r"))]))[1]
        ChatPPT("groq", None, model, groq_api_key=key)._get_content([{"role":"u","content":"h"}])
        mock_cc.assert_called_once_with(model=model, messages=[{"role":"u","content":"h"}])
        self.assertEqual(openai.api_key, self.original_openai_api_key)
        self.assertEqual(openai.api_base, self.original_openai_api_base)
    @patch('openai.ChatCompletion.create')
    def test_groq_uses_key_and_config_base(self, mock_cc):
        key, model, custom_base = "gkey2", "l3", "http://custom.groq/v1"
        self.mocked_config['GROQ_API_BASE'] = custom_base
        mock_cc.side_effect = lambda *a, **kw: (self.assertEqual(openai.api_key, key), self.assertEqual(openai.api_base, custom_base), Mock(choices=[Mock(message=Mock(content="r"))]))[1]
        ChatPPT("groq", None, model, groq_api_key=key)._get_content([{"role":"u","content":"h"}])
        self.assertEqual(openai.api_key, self.original_openai_api_key)
        self.assertEqual(openai.api_base, self.original_openai_api_base)

class TestRegenerateSinglePage(unittest.TestCase):
    def setUp(self):
        self.chat_instance = ChatPPT(model_provider="test_provider", api_key="dummy_key", model_name="test_model")
        self.original_topic = "Mars Colonization"
        self.original_language = "English"
        self.page_data_to_edit = {"title": "Old Page Title", "content": [{"title": "Old Bullet 1", "description": "Old Desc 1"}]}
        self.single_page_format_example_str = json.dumps(self.chat_instance._get_single_page_output_format_example(), indent=2)


    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_basic_flow(self, mock_get_content):
        mock_llm_response_str = '{"title": "Refined Page Title", "content": [{"title": "Concise Bullet 1", "description": "Concise Desc 1"}]}'
        mock_get_content.return_value = mock_llm_response_str
        new_instructions = "Make it more concise."

        result = self.chat_instance.regenerate_single_page(
            self.original_topic, self.original_language, self.page_data_to_edit, new_instructions
        )
        mock_get_content.assert_called_once()
        args_passed = mock_get_content.call_args[0][0]
        prompt_content = args_passed[0]['content']

        self.assertIn(self.original_topic, prompt_content)
        self.assertIn(self.original_language, prompt_content)
        self.assertIn(json.dumps(self.page_data_to_edit, indent=2), prompt_content)
        self.assertIn(new_instructions, prompt_content)
        self.assertIn("Please refine this slide based on the following instructions:", prompt_content)
        self.assertIn("MUST return ONLY the complete JSON object for this single refined slide", prompt_content)
        self.assertIn(self.single_page_format_example_str, prompt_content)
        
        expected_output = json.loads(mock_llm_response_str)
        self.assertEqual(result, expected_output)

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_no_instructions(self, mock_get_content):
        mock_llm_response_str = '{"title": "Reviewed Page Title", "content": [{"title": "Reviewed Bullet 1", "description": "Reviewed Desc 1"}]}'
        mock_get_content.return_value = mock_llm_response_str

        result = self.chat_instance.regenerate_single_page(
            self.original_topic, self.original_language, self.page_data_to_edit, new_instructions_for_page=None
        )
        mock_get_content.assert_called_once()
        prompt_content = mock_get_content.call_args[0][0][0]['content']
        self.assertIn("Please review and refine the content of this slide", prompt_content)
        self.assertNotIn("Please refine this slide based on the following instructions:", prompt_content)
        self.assertEqual(result, json.loads(mock_llm_response_str))

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_empty_instructions(self, mock_get_content):
        mock_llm_response_str = '{"title": "Reviewed Page Title 2", "content": []}'
        mock_get_content.return_value = mock_llm_response_str
        result = self.chat_instance.regenerate_single_page(
            self.original_topic, self.original_language, self.page_data_to_edit, new_instructions_for_page="   " # Whitespace
        )
        prompt_content = mock_get_content.call_args[0][0][0]['content']
        self.assertIn("Please review and refine the content of this slide", prompt_content)
        self.assertEqual(result, json.loads(mock_llm_response_str))

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_parsing_error_malformed_json(self, mock_get_content):
        mock_get_content.return_value = '{"title": "Bad JSON", "content": [}' # Malformed
        with self.assertRaisesRegex(Exception, "Error parsing regenerated page content.*Raw response snippet:.*Bad JSON"):
            self.chat_instance.regenerate_single_page(self.original_topic, self.original_language, self.page_data_to_edit)

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_validation_error_wrong_structure(self, mock_get_content):
        mock_get_content.return_value = '{"page_title": "Wrong Key", "bullets": []}' # Valid JSON, wrong structure
        with self.assertRaisesRegex(Exception, "LLM returned JSON but not in the expected page format"):
            self.chat_instance.regenerate_single_page(self.original_topic, self.original_language, self.page_data_to_edit)

    @patch.object(ChatPPT, '_get_content')
    def test_regenerate_page_validation_error_malformed_bullets(self, mock_get_content):
        mock_get_content.return_value = '{"title": "Page With Bad Bullet", "content": [{"name": "Missing title/desc"}]}'
        with self.assertRaisesRegex(Exception, "LLM returned JSON with malformed bullet points"):
            self.chat_instance.regenerate_single_page(self.original_topic, self.original_language, self.page_data_to_edit)


if __name__ == '__main__':
    unittest.main()
