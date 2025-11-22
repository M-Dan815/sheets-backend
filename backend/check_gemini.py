"""
Check available Gemini models with your API key
"""
import google.generativeai as genai
import json

# Load your API key
with open('config.json') as f:
    config = json.load(f)

genai.configure(api_key=config['gemini_api_key'])

print("Checking available Gemini models...\n")
print("="*60)

try:
    # List all models
    models = genai.list_models()
    
    print("\nALL AVAILABLE MODELS:")
    print("-" * 60)
    for model in models:
        print(f"\nModel: {model.name}")
        print(f"  Display Name: {model.display_name if hasattr(model, 'display_name') else 'N/A'}")
        print(f"  Description: {model.description if hasattr(model, 'description') else 'N/A'}")
        
        # Check supported methods
        if hasattr(model, 'supported_generation_methods'):
            print(f"  Supported Methods: {model.supported_generation_methods}")
        
    print("\n" + "="*60)
    print("\nMODELS THAT SUPPORT generateContent:")
    print("-" * 60)
    
    for model in models:
        if hasattr(model, 'supported_generation_methods'):
            if 'generateContent' in model.supported_generation_methods:
                print(f"✓ {model.name}")
        
except Exception as e:
    print(f"Error listing models: {e}")
    print("\nTrying alternative method...")
    
    # Try using specific models directly
    test_models = [
        'gemini-pro',
        'gemini-1.5-flash',
        'gemini-1.5-pro',
        'gemini-1.0-pro',
        'models/gemini-pro',
        'models/gemini-1.5-flash',
        'models/gemini-1.5-pro'
    ]
    
    print("\nTesting specific model names:")
    print("-" * 60)
    
    for model_name in test_models:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content("Say 'test'")
            print(f"✓ {model_name} - WORKS!")
        except Exception as e:
            print(f"✗ {model_name} - Failed: {str(e)[:50]}")