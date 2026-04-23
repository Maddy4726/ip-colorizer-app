# 📊 Excel IP Colorizer

A Streamlit web application that highlights IP addresses in Excel files based on a reference IP list. Process multiple files at once and download results as a ZIP archive.

## Features

✨ **Key Capabilities:**
- 📁 Upload multiple Excel files simultaneously
- 🎯 Match IPs against a reference IP list (TXT, CSV, or XLSX)
- 🎨 Color-coded highlighting:
  - 🟢 **Green**: All IPs in cell match the reference list
  - 🔴 **Red**: No IPs match the reference list
  - 🟡 **Yellow**: Mixed results (some match, some don't)
- 💬 Automatic comments on mixed cells showing matched and unmatched IPs
- 📦 Batch download all processed files as a ZIP archive
- 📊 Detailed statistics for each processed file

## Tech Stack

- **Framework**: Streamlit
- **Excel Processing**: openpyxl
- **Data Processing**: pandas
- **Language**: Python 3

## Installation

### Prerequisites
- Python 3.7+
- pip

### Setup

1. Clone the repository:
```bash
git clone https://github.com/Maddy4726/ip-colorizer-app.git
cd ip-colorizer-app
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Running the App

Start the Streamlit application:
```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`

### How to Use

1. **Upload Excel Files**: Select one or more Excel files (.xlsx, .xls, .xlsm)
2. **Upload IP List**: Provide a reference IP list in one of these formats:
   - **TXT**: One IP per line
   - **CSV**: IPs in the first column
   - **XLSX**: IPs in the first column
3. **Process**: Click the "🚀 Process Files" button
4. **Download**: Download all processed files as a single ZIP archive

### Input File Formats

**Excel Files:**
- Supported formats: .xlsx, .xls, .xlsm
- All sheets will be processed
- IP addresses can be embedded in text cells

**IP List File:**
- **TXT Format**: 
  ```
  192.168.1.1
  10.0.0.1
  172.16.0.1
  ```
- **CSV Format**: IPs in the first column
- **XLSX Format**: IPs in the first cell of each row

## Output

The processed Excel file(s) include:
- Original data preserved
- Color-coded IP cells
- Comments on mixed cells (Yellow) showing matched/unmatched IPs
- Statistics displayed in the app UI

## Project Structure

```
ip-colorizer-app/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
├── README.md          # This file
└── tests/
    └── test_ip.py     # Test suite
```

## IP Address Regex

The app uses RFC-compliant regex to detect IP addresses:
```regex
(?:25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(?:\.(?:25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}
```

This pattern validates that each octet is between 0-255.

## Development & Testing

Run tests with:
```bash
python -m pytest tests/
```

## License

MIT License - feel free to use and modify this project.

## Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

## Support

For issues, feature requests, or questions, please open an issue on the [GitHub repository](https://github.com/Maddy4726/ip-colorizer-app/issues).
