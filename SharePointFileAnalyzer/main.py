#!/usr/bin/env python3
"""
SharePoint File Analyzer with AI
Main application entry point
"""
import sys
from typing import List, Dict
from tqdm import tqdm
from config import Config
from sharepoint_client import SharePointClient
from ai_analyzer import AIAnalyzer
from excel_generator import ExcelGenerator


def print_banner():
    """Print application banner"""
    print("=" * 70)
    print("  SharePoint File Analyzer with AI")
    print("  Automated document analysis and summarization")
    print("=" * 70)
    print()


def get_user_input(prompt: str, default: str = None) -> str:
    """
    Get input from user with optional default value

    Args:
        prompt: Prompt to display
        default: Default value if user presses enter

    Returns:
        User input or default value
    """
    if default:
        user_input = input(f"{prompt} [{default}]: ").strip()
        return user_input if user_input else default
    else:
        user_input = input(f"{prompt}: ").strip()
        while not user_input:
            print("This field is required.")
            user_input = input(f"{prompt}: ").strip()
        return user_input


def validate_configuration():
    """Validate configuration and print errors if any"""
    errors = Config.validate()
    if errors:
        print("\nConfiguration Errors:")
        for error in errors:
            print(f"  - {error}")
        print("\nPlease check your .env file or environment variables.")
        print("See .env.example for reference.")
        return False
    return True


def process_files(sp_client: SharePointClient, ai_analyzer: AIAnalyzer,
                 files: List[Dict]) -> tuple[List[Dict], List[Dict]]:
    """
    Process files and analyze them with AI

    Args:
        sp_client: SharePoint client
        ai_analyzer: AI analyzer
        files: List of file information

    Returns:
        Tuple of (analysis_results, errors)
    """
    analysis_results = []
    errors = []

    # Filter files by supported types and size
    supported_files = [
        f for f in files
        if f['extension'] in Config.SUPPORTED_FILE_TYPES
        and f['size'] <= Config.MAX_FILE_SIZE_MB * 1024 * 1024
    ]

    print(f"\nProcessing {len(supported_files)} files (filtered from {len(files)} total files)")
    print(f"Supported types: {', '.join(Config.SUPPORTED_FILE_TYPES)}")
    print(f"Max file size: {Config.MAX_FILE_SIZE_MB} MB\n")

    # Process each file with progress bar
    for file_info in tqdm(supported_files, desc="Analyzing files", unit="file"):
        try:
            filename = file_info['name']
            extension = file_info['extension']

            # Try to extract text content
            try:
                content = sp_client.get_file_content_as_text(
                    file_info['server_relative_url'],
                    extension
                )
            except Exception as e:
                # If content extraction fails, use metadata only
                content = None
                errors.append({
                    'filename': filename,
                    'error_type': 'Content Extraction Failed',
                    'error_message': str(e)
                })

            # Analyze with AI
            if content and len(content.strip()) > 0:
                analysis = ai_analyzer.analyze_document(filename, content)
            else:
                analysis = ai_analyzer.analyze_file_metadata(filename, extension)

            # Add to results
            analysis_results.append({
                'filename': filename,
                'title': analysis['title'],
                'summary': analysis['summary'],
                'size': file_info['size'],
                'time_modified': file_info['time_modified'],
                'extension': extension
            })

        except Exception as e:
            errors.append({
                'filename': file_info['name'],
                'error_type': 'Processing Failed',
                'error_message': str(e)
            })
            print(f"\nError processing {file_info['name']}: {str(e)}")

    return analysis_results, errors


def main():
    """Main application logic"""
    print_banner()

    # Check if using environment variables or prompting user
    use_env_config = Config.SHAREPOINT_SITE_URL and (
        Config.SHAREPOINT_USERNAME or Config.SHAREPOINT_CLIENT_ID
    )

    if not use_env_config:
        print("No configuration found in environment variables.")
        print("You will be prompted for SharePoint connection details.\n")

    # Validate AI configuration (required from env)
    if not validate_configuration():
        sys.exit(1)

    try:
        # Step 1: Get SharePoint site URL
        if use_env_config:
            sharepoint_url = Config.SHAREPOINT_SITE_URL
            print(f"Using SharePoint site: {sharepoint_url}")
        else:
            sharepoint_url = get_user_input(
                "Enter SharePoint site URL",
                "https://yourdomain.sharepoint.com/sites/yoursite"
            )

        # Step 2: Get SharePoint folder path
        folder_path = get_user_input(
            "Enter SharePoint folder path",
            "Shared Documents"
        )

        # Get credentials if not in environment
        username = None
        password = None
        if not use_env_config:
            username = get_user_input("Enter SharePoint username/email")
            password = get_user_input("Enter SharePoint password")

        print("\n" + "=" * 70)
        print("Starting analysis...")
        print("=" * 70 + "\n")

        # Initialize clients
        sp_client = SharePointClient(sharepoint_url)
        sp_client.connect(username, password)

        ai_analyzer = AIAnalyzer()

        # Step 3: List files in folder
        print(f"\nScanning folder: {folder_path}")
        files = sp_client.list_files(folder_path)

        if not files:
            print("No files found in the specified folder.")
            sys.exit(0)

        # Step 4: Process files with AI
        analysis_results, errors = process_files(sp_client, ai_analyzer, files)

        if not analysis_results:
            print("\nNo files were successfully processed.")
            sys.exit(1)

        # Step 5: Generate Excel report
        print(f"\nGenerating Excel report...")
        excel_gen = ExcelGenerator(Config.OUTPUT_FILENAME)
        excel_gen.create_report(analysis_results, sharepoint_url, folder_path)

        if errors:
            excel_gen.add_error_sheet(errors)
            print(f"\nNote: {len(errors)} files had processing errors. See 'Errors' sheet in the report.")

        if excel_gen.save():
            print("\n" + "=" * 70)
            print(f"  Analysis complete!")
            print(f"  Processed: {len(analysis_results)} files")
            print(f"  Report: {Config.OUTPUT_FILENAME}")
            print("=" * 70 + "\n")
        else:
            print("\nFailed to save Excel report.")
            sys.exit(1)

    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
