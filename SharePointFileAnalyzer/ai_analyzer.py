"""
AI-powered document analyzer using OpenAI
"""
from typing import Dict, Optional
from openai import OpenAI, AzureOpenAI
from config import Config


class AIAnalyzer:
    """Analyzer for extracting document information using AI"""

    def __init__(self):
        """Initialize AI client"""
        if Config.is_using_azure_openai():
            # Use Azure OpenAI
            self.client = AzureOpenAI(
                api_key=Config.AZURE_OPENAI_API_KEY,
                api_version=Config.AZURE_OPENAI_API_VERSION,
                azure_endpoint=Config.AZURE_OPENAI_ENDPOINT
            )
            self.model = Config.AZURE_OPENAI_DEPLOYMENT
            print("Initialized Azure OpenAI client")
        else:
            # Use OpenAI
            self.client = OpenAI(api_key=Config.OPENAI_API_KEY)
            self.model = Config.OPENAI_MODEL
            print("Initialized OpenAI client")

    def analyze_document(self, filename: str, content: str, max_content_length: int = 8000) -> Dict[str, str]:
        """
        Analyze a document and extract title and summary

        Args:
            filename: Name of the file
            content: Text content of the file
            max_content_length: Maximum content length to send to AI (to manage costs)

        Returns:
            Dictionary with 'title' and 'summary' keys
        """
        # Truncate content if too long
        if len(content) > max_content_length:
            content = content[:max_content_length] + "\n... [content truncated]"

        # Create prompt for AI
        prompt = f"""Analyze the following document and provide:
1. A concise title that best represents the document's content
2. A one-paragraph summary (3-5 sentences) describing the key points and purpose of the document

Document filename: {filename}

Document content:
{content}

Please respond in the following JSON format:
{{
  "title": "Your extracted or generated title here",
  "summary": "Your one-paragraph summary here"
}}"""

        try:
            # Call AI API
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": "You are a helpful assistant that analyzes documents and extracts key information. Always respond with valid JSON."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.3,
                max_tokens=500
            )

            # Extract response
            response_text = response.choices[0].message.content.strip()

            # Parse JSON response
            import json
            # Remove markdown code blocks if present
            if response_text.startswith('```'):
                response_text = response_text.split('```')[1]
                if response_text.startswith('json'):
                    response_text = response_text[4:]
                response_text = response_text.strip()

            result = json.loads(response_text)

            return {
                'title': result.get('title', filename),
                'summary': result.get('summary', 'Unable to generate summary')
            }

        except Exception as e:
            print(f"Error analyzing document '{filename}': {str(e)}")
            return {
                'title': filename,
                'summary': f"Error during analysis: {str(e)}"
            }

    def analyze_file_metadata(self, filename: str, extension: str) -> Dict[str, str]:
        """
        Analyze file when content cannot be extracted

        Args:
            filename: Name of the file
            extension: File extension

        Returns:
            Dictionary with 'title' and 'summary' keys based on metadata
        """
        return {
            'title': filename,
            'summary': f"File type: {extension}. Content analysis not available for this file type."
        }
