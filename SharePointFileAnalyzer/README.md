# SharePoint File Analyzer with AI

An automated tool that scans SharePoint folders, analyzes documents using AI, and generates comprehensive Excel reports with document titles and summaries.

## Features

- **SharePoint Integration**: Securely connect to SharePoint Online sites
- **Automated File Scanning**: List and process all files in a specified folder
- **AI-Powered Analysis**: Use OpenAI or Azure OpenAI to extract titles and generate summaries
- **Excel Report Generation**: Create professional reports with detailed analysis results
- **Multiple File Format Support**: Process TXT, PDF, DOCX, XLSX, PPTX, Markdown, and more
- **Error Handling**: Comprehensive error tracking and reporting
- **Progress Tracking**: Visual progress indicators during processing

## Prerequisites

- Python 3.8 or higher
- SharePoint Online account with access permissions
- OpenAI API key or Azure OpenAI service credentials

## Installation

1. **Clone or download this repository**

```bash
cd SharePointFileAnalyzer
```

2. **Create a virtual environment (recommended)**

```bash
python -m venv venv

# On Windows
venv\Scripts\activate

# On macOS/Linux
source venv/bin/activate
```

3. **Install dependencies**

```bash
pip install -r requirements.txt
```

4. **Configure environment variables**

Copy the example environment file:

```bash
cp .env.example .env
```

Edit `.env` and add your credentials:

```env
# SharePoint Configuration
SHAREPOINT_SITE_URL=https://yourdomain.sharepoint.com/sites/yoursite
SHAREPOINT_USERNAME=your.email@domain.com
SHAREPOINT_PASSWORD=your_password

# OpenAI Configuration
OPENAI_API_KEY=your_openai_api_key
OPENAI_MODEL=gpt-4o-mini
```

## Configuration

### SharePoint Authentication

**Option 1: Username/Password (Basic Authentication)**
```env
SHAREPOINT_USERNAME=your.email@domain.com
SHAREPOINT_PASSWORD=your_password
```

**Option 2: Azure AD Application (Recommended for production)**
```env
SHAREPOINT_CLIENT_ID=your_client_id
SHAREPOINT_CLIENT_SECRET=your_client_secret
SHAREPOINT_TENANT_ID=your_tenant_id
```

### AI Configuration

**Option 1: OpenAI**
```env
OPENAI_API_KEY=your_openai_api_key
OPENAI_MODEL=gpt-4o-mini
```

**Option 2: Azure OpenAI**
```env
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_OPENAI_API_KEY=your_azure_openai_key
AZURE_OPENAI_DEPLOYMENT=your_deployment_name
AZURE_OPENAI_API_VERSION=2024-02-01
```

### Application Settings

```env
MAX_FILE_SIZE_MB=10                              # Maximum file size to process
OUTPUT_FILENAME=sharepoint_file_analysis.xlsx    # Output Excel filename
```

## Usage

### Running the Application

```bash
python main.py
```

### Interactive Prompts

The application will prompt you for:

1. **SharePoint site URL** (if not configured in .env)
   - Example: `https://yourdomain.sharepoint.com/sites/yoursite`

2. **SharePoint folder path**
   - Example: `Shared Documents/Reports`
   - Example: `Documents/2024`

3. **Credentials** (if not configured in .env)
   - Username/email
   - Password

### Using Configuration File

If you've configured the `.env` file completely, the application will:
- Use the SharePoint URL from configuration
- Automatically authenticate
- Only prompt for the folder path

### Example Run

```
======================================================================
  SharePoint File Analyzer with AI
  Automated document analysis and summarization
======================================================================

Using SharePoint site: https://yourdomain.sharepoint.com/sites/yoursite
Enter SharePoint folder path [Shared Documents]: Reports/Q4

======================================================================
Starting analysis...
======================================================================

Connected to SharePoint site: https://yourdomain.sharepoint.com/sites/yoursite
Initialized OpenAI client

Scanning folder: Reports/Q4
Found 15 files in folder: Reports/Q4

Processing 12 files (filtered from 15 total files)
Supported types: .txt, .pdf, .docx, .doc, .xlsx, .xls, .pptx, .ppt, .md, .csv, .json, .xml

Analyzing files: 100%|████████████████████| 12/12 [00:45<00:00,  3.8s/file]

Generating Excel report...

Excel report saved successfully: sharepoint_file_analysis.xlsx

======================================================================
  Analysis complete!
  Processed: 12 files
  Report: sharepoint_file_analysis.xlsx
======================================================================
```

## Output

The application generates an Excel file (`sharepoint_file_analysis.xlsx` by default) with:

### Main Sheet: "File Analysis"
- **#**: Row number
- **Filename**: Original filename
- **Title**: AI-extracted or generated title
- **Summary**: One-paragraph summary of the document
- **File Size (KB)**: File size in kilobytes
- **Last Modified**: Last modification timestamp

### Errors Sheet (if applicable)
- **Filename**: Files that encountered errors
- **Error Type**: Type of error
- **Error Message**: Detailed error description

## Supported File Types

- **Text Files**: .txt, .md, .csv, .json, .xml
- **Documents**: .docx, .doc
- **PDFs**: .pdf
- **Spreadsheets**: .xlsx, .xls
- **Presentations**: .pptx, .ppt

## Architecture

```
SharePointFileAnalyzer/
├── main.py                 # Main application entry point
├── config.py               # Configuration management
├── sharepoint_client.py    # SharePoint integration
├── ai_analyzer.py          # AI document analysis
├── excel_generator.py      # Excel report generation
├── requirements.txt        # Python dependencies
├── .env.example           # Example environment variables
├── .gitignore             # Git ignore patterns
└── README.md              # This file
```

## Troubleshooting

### Authentication Issues

**Problem**: "No valid SharePoint credentials provided"
- **Solution**: Ensure either username/password or Azure AD credentials are set in `.env`

**Problem**: Authentication fails with username/password
- **Solution**: Check if your organization requires MFA. If so, use Azure AD app authentication

### SharePoint Access Issues

**Problem**: "Error listing files in folder"
- **Solution**: Verify the folder path is correct and you have read permissions
- **Solution**: Ensure the folder path is relative (don't include the site URL)

### AI API Issues

**Problem**: "AI API credentials not set"
- **Solution**: Ensure OPENAI_API_KEY or Azure OpenAI credentials are set in `.env`

**Problem**: API rate limits or timeout errors
- **Solution**: Add delays between requests or reduce the number of files processed
- **Solution**: Ensure your API subscription has sufficient quota

### File Processing Issues

**Problem**: Some files are skipped
- **Solution**: Check the file size (default max is 10MB)
- **Solution**: Verify the file type is in the supported list

## Security Best Practices

1. **Never commit the `.env` file** to version control
2. **Use Azure AD app authentication** for production deployments
3. **Store credentials securely** using environment variables or secret management services
4. **Limit API keys** to minimum required permissions
5. **Rotate credentials regularly**

## Cost Considerations

- **OpenAI API**: Charges are based on tokens used (input + output)
- **Estimated cost per file**: $0.001 - $0.01 (depending on file size and model)
- **Use gpt-4o-mini**: More cost-effective for document analysis
- **Set MAX_FILE_SIZE_MB**: Control content sent to AI API

## License

This project is free to use, modify, and distribute. No credits needed.

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests.

## Support

For issues, questions, or feature requests, please open an issue on the repository.

## Changelog

### Version 1.0.0 (2025-10-22)
- Initial release
- SharePoint Online integration
- OpenAI and Azure OpenAI support
- Excel report generation
- Support for multiple file formats
- Interactive CLI interface
- Comprehensive error handling
