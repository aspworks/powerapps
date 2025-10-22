"""
SharePoint client for file operations
"""
import os
import io
from typing import List, Dict, Optional
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from config import Config


class SharePointClient:
    """Client for interacting with SharePoint"""

    def __init__(self, site_url: Optional[str] = None):
        """
        Initialize SharePoint client

        Args:
            site_url: SharePoint site URL (optional, uses config if not provided)
        """
        self.site_url = site_url or Config.SHAREPOINT_SITE_URL
        self.ctx = None

    def connect(self, username: Optional[str] = None, password: Optional[str] = None):
        """
        Connect to SharePoint site

        Args:
            username: SharePoint username (optional, uses config if not provided)
            password: SharePoint password (optional, uses config if not provided)
        """
        username = username or Config.SHAREPOINT_USERNAME
        password = password or Config.SHAREPOINT_PASSWORD

        if username and password:
            # Use basic authentication
            credentials = UserCredential(username, password)
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)
            print(f"Connected to SharePoint site: {self.site_url}")
        elif Config.SHAREPOINT_CLIENT_ID and Config.SHAREPOINT_CLIENT_SECRET:
            # Use Azure AD authentication
            credentials = ClientCredential(
                Config.SHAREPOINT_CLIENT_ID,
                Config.SHAREPOINT_CLIENT_SECRET
            )
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)
            print(f"Connected to SharePoint site using Azure AD: {self.site_url}")
        else:
            raise ValueError("No valid SharePoint credentials provided")

    def list_files(self, folder_path: str) -> List[Dict]:
        """
        List all files in a SharePoint folder

        Args:
            folder_path: Relative path to the SharePoint folder

        Returns:
            List of file information dictionaries
        """
        if not self.ctx:
            raise RuntimeError("Not connected to SharePoint. Call connect() first.")

        # Ensure folder path doesn't start with /
        folder_path = folder_path.lstrip('/')

        try:
            # Get the folder
            folder = self.ctx.web.get_folder_by_server_relative_path(folder_path)
            files = folder.files
            self.ctx.load(files)
            self.ctx.execute_query()

            file_list = []
            for file in files:
                file_info = {
                    'name': file.properties['Name'],
                    'server_relative_url': file.properties['ServerRelativeUrl'],
                    'size': file.properties['Length'],
                    'time_created': file.properties.get('TimeCreated', ''),
                    'time_modified': file.properties.get('TimeLastModified', ''),
                    'extension': os.path.splitext(file.properties['Name'])[1].lower()
                }
                file_list.append(file_info)

            print(f"Found {len(file_list)} files in folder: {folder_path}")
            return file_list

        except Exception as e:
            print(f"Error listing files in folder '{folder_path}': {str(e)}")
            raise

    def download_file(self, server_relative_url: str) -> bytes:
        """
        Download a file from SharePoint

        Args:
            server_relative_url: Server-relative URL of the file

        Returns:
            File content as bytes
        """
        if not self.ctx:
            raise RuntimeError("Not connected to SharePoint. Call connect() first.")

        try:
            # Get file content
            file = self.ctx.web.get_file_by_server_relative_path(server_relative_url)
            content = file.read()
            self.ctx.execute_query()

            return content

        except Exception as e:
            print(f"Error downloading file '{server_relative_url}': {str(e)}")
            raise

    def get_file_content_as_text(self, server_relative_url: str, file_extension: str) -> str:
        """
        Download and extract text content from a file

        Args:
            server_relative_url: Server-relative URL of the file
            file_extension: File extension (e.g., '.txt', '.pdf', '.docx')

        Returns:
            Extracted text content
        """
        content_bytes = self.download_file(server_relative_url)

        if file_extension == '.txt':
            return content_bytes.decode('utf-8', errors='ignore')

        elif file_extension == '.pdf':
            return self._extract_pdf_text(content_bytes)

        elif file_extension in ['.docx', '.doc']:
            return self._extract_docx_text(content_bytes)

        else:
            # For other file types, try to decode as text
            try:
                return content_bytes.decode('utf-8', errors='ignore')
            except:
                return f"[Unable to extract text from {file_extension} file]"

    def _extract_pdf_text(self, content_bytes: bytes) -> str:
        """Extract text from PDF file"""
        try:
            from PyPDF2 import PdfReader

            pdf_file = io.BytesIO(content_bytes)
            pdf_reader = PdfReader(pdf_file)

            text_parts = []
            for page in pdf_reader.pages:
                text_parts.append(page.extract_text())

            return '\n'.join(text_parts)
        except Exception as e:
            return f"[Error extracting PDF text: {str(e)}]"

    def _extract_docx_text(self, content_bytes: bytes) -> str:
        """Extract text from DOCX file"""
        try:
            from docx import Document

            docx_file = io.BytesIO(content_bytes)
            doc = Document(docx_file)

            text_parts = []
            for paragraph in doc.paragraphs:
                text_parts.append(paragraph.text)

            return '\n'.join(text_parts)
        except Exception as e:
            return f"[Error extracting DOCX text: {str(e)}]"
