import os
from groq import Groq
from dotenv import load_dotenv

class CoverLetterGenerator:
    def __init__(self, env_path='src/credentials/groq.env'):
        load_dotenv(dotenv_path=env_path)
        
        api_key = os.getenv('COLDFLOW_GROQ_API_KEY')
        
        if not api_key:
            raise ValueError("API key not found.")
        
        self.client = Groq(api_key=api_key)
        
    def generate_cover_letter_first_time(self, prompt):
        prompt_first_time = f"{prompt}"
        
        response = self.client.chat.completions.create(
            messages=[{"role": "user", "content": prompt_first_time}],
            model="llama-3.3-70b-versatile"
        )
        cover_letter_first_time = response.choices[0].message.content.strip()
        return cover_letter_first_time
    
    