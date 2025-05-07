import os
import argparse
import pandas as pd
from PIL import Image
import io
import base64
import glob
import json
import requests
from google.oauth2 import service_account
import google.generativeai as genai

def read_prompt_file(prompt_file):
    """Read the prompt from the specified file."""
    try:
        with open(prompt_file, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception as e:
        print(f"Error reading prompt file: {e}")
        return None

def process_image(image_path, api_key, model, custom_prompt):
    """Process a single image using Gemini API."""
    try:
        # Configure the API
        genai.configure(api_key=api_key)
        
        # Open and encode the image
        with open(image_path, 'rb') as f:
            image_bytes = f.read()
        
        # Get the model
        model_instance = genai.GenerativeModel(model)
        
        # Create a prompt that includes OCR instructions and any custom prompt
        prompt = f"""
        Please perform OCR on this image. 
        
        {custom_prompt}
        
        Return the results in a structured JSON format that can be converted to Excel.
        """
        
        # Generate content
        response = model_instance.generate_content([prompt, {"mime_type": "image/jpeg", "data": image_bytes}])
        
        # Extract the response text
        response_text = response.text
        
        # Try to parse the response as JSON
        try:
            # Find JSON in the response (sometimes the model wraps JSON in markdown code blocks)
            import re
            json_match = re.search(r'```json\n(.*?)\n```', response_text, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                json_str = response_text
            
            data = json.loads(json_str)
            return data
        except json.JSONDecodeError:
            # If the response is not valid JSON, just return the raw text
            return {"raw_text": response_text}
            
    except Exception as e:
        print(f"Error processing image {image_path}: {e}")
        return {"error": str(e)}

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Process images using Google Gemini API and convert to Excel')
    parser.add_argument('--api_key', required=True, help='Google API Key')
    parser.add_argument('--model', default='gemini-2.0-flash', help='Gemini model name')
    parser.add_argument('--photo_dir', default='Photo', help='Directory containing photos')
    parser.add_argument('--output_path', default='output.xlsx', help='Output Excel file path')
    parser.add_argument('--prompt_file', default='prompt.txt', help='File containing custom prompt')
    
    args = parser.parse_args()
    
    # Read the custom prompt
    custom_prompt = read_prompt_file(args.prompt_file)
    if not custom_prompt:
        print("Warning: Couldn't read custom prompt. Using default OCR instructions.")
        custom_prompt = "Extract all text from the image and organize it into structured data."
    
    # Get list of image files
    image_extensions = ['*.jpg', '*.jpeg', '*.png', '*.bmp', '*.gif']
    image_files = []
    for ext in image_extensions:
        image_files.extend(glob.glob(os.path.join(args.photo_dir, ext)))
    
    if not image_files:
        print(f"No image files found in directory: {args.photo_dir}")
        return
    
    print(f"Found {len(image_files)} image files.")
    
    # Process each image
    all_results = []
    for image_path in image_files:
        print(f"Processing {image_path}...")
        result = process_image(image_path, args.api_key, args.model, custom_prompt)
        
        # Add the image file name to the result
        if isinstance(result, dict):
            result['image_file'] = os.path.basename(image_path)
            all_results.append(result)
        else:
            all_results.append({
                'image_file': os.path.basename(image_path),
                'error': 'Unexpected result format'
            })
    
    # Convert results to DataFrame
    try:
        # Handle case where results have different schemas
        if all_results:
            df = pd.json_normalize(all_results)
            
            # Create the output directory if it doesn't exist
            output_dir = os.path.dirname(args.output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Save to Excel
            df.to_excel(args.output_path, index=False)
            print(f"Results saved to {args.output_path}")
        else:
            print("No results to save.")
    except Exception as e:
        print(f"Error saving results to Excel: {e}")
        # Save raw results as JSON as a fallback
        with open('results.json', 'w') as f:
            json.dump(all_results, f, indent=2)
        print("Results saved as results.json instead.")

if __name__ == "__main__":
    main()