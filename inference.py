# inference.py
#
# Initialize the model and send the data to the model
#
# @author: Dung Tran
# @date: 2025-10-17

import json
import os
from openai import OpenAI
from dotenv import load_dotenv
import ast

import regex as re

# Load environment variables from .env file
load_dotenv()

def initialize_model(model_name, base_url, api_key, config={"temperature": 0.2, "max_tokens": 4096}):
    """
    Initialize a model instance for text generation.
    
    Args:
        model_name (str): The name of the model to use (e.g., 'gpt-3.5-turbo', 'gpt-4', 'claude-3-sonnet')
        api_key (str): The API key for authentication
        config (dict, optional): Configuration parameters for the model
        
    Returns:
        OpenAI client instance configured for the specified model
    """
    if config is None:
        config = {}
    
    # Initialize OpenAI client with the provided API key
    client = OpenAI(api_key=api_key, base_url=base_url)
    
    # Store model configuration for later use
    client.model_name = model_name
    client.model_config = config
    
    return client


def generate_data(model, data, prompt):
    """
    Generate a response using the provided model, data, and prompt.
    
    Args:
        model: The initialized model instance (from initialize_model)
        data (list): A list of strings containing the data to process
        prompt (str): The prompt/instruction for the model
        
    Returns:
        str: The generated response from the model
    """
    try:
        # Combine the prompt with the data
        data_text = "\n".join(data) if data else ""
        full_prompt = f"{prompt}\n\nData:\n{data_text}"
        
        # Get model configuration if available
        config = getattr(model, 'model_config', {})
        model_name = getattr(model, 'model_name', 'gpt-3.5-turbo')
        
        # Create the chat completion request
        response = model.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "user", "content": full_prompt}
            ],
            **config  # Apply any additional configuration parameters
        )
        
        # Extract and return the response content
        return response.choices[0].message.content
        
    except Exception as e:
        print(f"Error generating response: {e}")
        return None


def parse_response(response, expected_length=None):
    """
    Parse the generated response as a JSON string that should be a list.
    Cleans markdown syntax including JSON code blocks before parsing.
    
    Args:
        response (str): The generated response from the model (should be JSON)
        expected_length (int, optional): Expected length of the parsed list for validation
        
    Returns:
        list: The parsed JSON list, or None if parsing fails
    """
    if response is None:
        print("Error: Response is None")
        return None
    
    try:
        # Clean the response by removing markdown syntax
        cleaned_response = clean_markdown(response)
        
        # Parse the JSON response
        parsed_data = ast.literal_eval(cleaned_response)
        
        # Validate that it's a list
        if not isinstance(parsed_data, list):
            print(f"Error: Expected list, got {type(parsed_data)}")
            return None
        
        # Validate length if expected_length is provided
        if expected_length is not None and len(parsed_data) != expected_length:
            print(f"Error: Expected {expected_length} items, got {len(parsed_data)}")
            return None
        
        return parsed_data
        
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
        print(f"Cleaned response content: {cleaned_response}")
        return cleaned_response
    except Exception as e:
        print(f"Error parsing response: {e}")
        return None


def clean_markdown(text):
    """
    Clean markdown syntax from text, especially JSON code blocks.
    Extract data bounded by square brackets [ and ].
    
    Args:
        text (str): The text to clean
        
    Returns:
        str: The cleaned text
    """
    if not text:
        return text
    
    # First, try to find content bounded by square brackets
    bracket_matches = re.findall(r'\[(.*?)\]', text, flags=re.DOTALL)
    if bracket_matches:
        # If we found content in brackets, use the first match
        # This handles cases like: "Here is the result: ["item1", "item2"]"
        # Return the last match to handle cases like: ["item1", "item2"] -> ["item1", "item2"]
        text = '[' + bracket_matches[-1] + ']'
        # clean the text
        text = text.strip()
    else:
        # Fallback to original markdown cleaning if no brackets found
        # Remove JSON code blocks (```json ... ``` or ``` ... ```)
        text = re.sub(r'```(?:json)?\s*\n?(.*?)\n?```', r'\1', text, flags=re.DOTALL | re.IGNORECASE)
        
        # Remove other markdown syntax
        # Remove bold/italic markers
        text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        text = re.sub(r'\*(.*?)\*', r'\1', text)
        text = re.sub(r'__(.*?)__', r'\1', text)
        text = re.sub(r'_(.*?)_', r'\1', text)
        
        # Remove code inline markers
        text = re.sub(r'`(.*?)`', r'\1', text)
        
        # Remove headers
        text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
        
        # Remove horizontal rules
        text = re.sub(r'^[-*_]{3,}$', '', text, flags=re.MULTILINE)
        
        # Remove list markers
        text = re.sub(r'^[\s]*[-*+]\s+', '', text, flags=re.MULTILINE)
        text = re.sub(r'^[\s]*\d+\.\s+', '', text, flags=re.MULTILINE)
        
        # Remove blockquotes
        text = re.sub(r'^>\s*', '', text, flags=re.MULTILINE)
        
        # Clean up extra whitespace
        text = re.sub(r'\n\s*\n', '\n', text)
        text = text.strip()
    
    return text


def generate_prompt(data):
    return f"""Bạn là chuyên gia dịch thuật từ tiếng Trung giản thể sang tiếng Việt chuyên dành cho các dự án xây dựng và thiết kế. Hãy dịch chính xác với ngữ cảnh và đúng từ chuyên ngành.\n"
        Ngôn ngữ nguồn: tiếng Trung giản thể\n
        Ngôn ngữ đích: tiếng Việt\n
        - Chỉ dịch phần văn bản; nếu là công thức Excel hoặc tham chiếu ô/sheet thì bỏ qua.\n
        - Giữ nguyên số, mã SKU, ký hiệu đặc biệt nếu không cần dịch.\n
        - Bảo toàn xuống dòng và khoảng trắng quan trọng.\n
        - Trả KẾT QUẢ DUY NHẤT là một JSON array các chuỗi chứa từ gốc ở vị trí thứ nhất và từ được dịch ở vị trí thứ 2. Các từ được trả theo đúng thứ tự input, không thêm giải thích. Ví dụ: bạn nhận được [\"瓷砖楼梯\", "1"], bạn trả lời [(\'瓷砖楼梯\', \'cầu thang đá men\'), (\'1\', \'1\')].
        """ + "Dữ liệu cần dịch: [" + ",".join(data) + "]"


def translate_data(data):
    """
    Translate the data using the provided model and prompt.
    
    Args:
        data (list): A list of strings containing the data to process
        
    Returns:
        list: The translated data
    """
    # Get environment variables
    model_name = os.getenv("MODEL_NAME")
    base_url = os.getenv("BASE_URL")
    api_key = os.getenv("API_KEY")
    
    # Check if required environment variables are set
    if not api_key:
        raise ValueError("API_KEY not found in environment variables. Please check your .env file.")
    if not model_name:
        raise ValueError("MODEL_NAME not found in environment variables. Please check your .env file.")
    if not base_url:
        raise ValueError("BASE_URL not found in environment variables. Please check your .env file.")
    
    model = initialize_model(model_name, base_url, api_key)
    prompt = generate_prompt(data)
    response = generate_data(model, data, prompt)
    return parse_response(response)


if __name__ == "__main__":
    data = ["瓷砖楼梯", "瓷砖楼梯", "瓷砖楼梯"]
    translated_data = translate_data(data)
    print(translated_data)
