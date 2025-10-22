"""
Configuration module for SharePoint File Analyzer
"""
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


class Config:
    """Application configuration"""

    # SharePoint Configuration
    SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL', '')
    SHAREPOINT_USERNAME = os.getenv('SHAREPOINT_USERNAME', '')
    SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD', '')

    # Azure AD Authentication (Alternative)
    SHAREPOINT_CLIENT_ID = os.getenv('SHAREPOINT_CLIENT_ID', '')
    SHAREPOINT_CLIENT_SECRET = os.getenv('SHAREPOINT_CLIENT_SECRET', '')
    SHAREPOINT_TENANT_ID = os.getenv('SHAREPOINT_TENANT_ID', '')

    # OpenAI Configuration
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY', '')
    OPENAI_MODEL = os.getenv('OPENAI_MODEL', 'gpt-4o-mini')

    # Azure OpenAI Configuration (Alternative)
    AZURE_OPENAI_ENDPOINT = os.getenv('AZURE_OPENAI_ENDPOINT', '')
    AZURE_OPENAI_API_KEY = os.getenv('AZURE_OPENAI_API_KEY', '')
    AZURE_OPENAI_DEPLOYMENT = os.getenv('AZURE_OPENAI_DEPLOYMENT', '')
    AZURE_OPENAI_API_VERSION = os.getenv('AZURE_OPENAI_API_VERSION', '2024-02-01')

    # Application Settings
    MAX_FILE_SIZE_MB = int(os.getenv('MAX_FILE_SIZE_MB', '10'))
    OUTPUT_FILENAME = os.getenv('OUTPUT_FILENAME', 'sharepoint_file_analysis.xlsx')

    # Supported file types for analysis
    SUPPORTED_FILE_TYPES = [
        '.txt', '.pdf', '.docx', '.doc',
        '.xlsx', '.xls', '.pptx', '.ppt',
        '.md', '.csv', '.json', '.xml'
    ]

    @classmethod
    def validate(cls):
        """Validate required configuration"""
        errors = []

        # Check SharePoint configuration
        if not cls.SHAREPOINT_SITE_URL:
            errors.append("SHAREPOINT_SITE_URL is not set")

        # Check authentication method
        has_basic_auth = cls.SHAREPOINT_USERNAME and cls.SHAREPOINT_PASSWORD
        has_azure_auth = (cls.SHAREPOINT_CLIENT_ID and
                         cls.SHAREPOINT_CLIENT_SECRET and
                         cls.SHAREPOINT_TENANT_ID)

        if not (has_basic_auth or has_azure_auth):
            errors.append("SharePoint authentication credentials not set (username/password or Azure AD)")

        # Check AI configuration
        has_openai = cls.OPENAI_API_KEY
        has_azure_openai = (cls.AZURE_OPENAI_ENDPOINT and
                           cls.AZURE_OPENAI_API_KEY and
                           cls.AZURE_OPENAI_DEPLOYMENT)

        if not (has_openai or has_azure_openai):
            errors.append("AI API credentials not set (OpenAI or Azure OpenAI)")

        return errors

    @classmethod
    def is_using_azure_openai(cls):
        """Check if using Azure OpenAI instead of OpenAI"""
        return bool(cls.AZURE_OPENAI_ENDPOINT and
                   cls.AZURE_OPENAI_API_KEY and
                   cls.AZURE_OPENAI_DEPLOYMENT)
