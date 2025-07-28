# ASUS BIOS Downloader GUI

## Description
A Python GUI tool to automate downloading and extracting the latest BIOS files for ASUS devices. It uses Selenium to interact with the ASUS support website, downloads BIOS ZIP files, and extracts them to a specified folder. The GUI allows configuration of paths, model lists, and provides progress and logging.

## Features
- Download latest BIOS for multiple ASUS models automatically
- Extract BIOS ZIP files after download
- GUI for configuration, progress, and logs
- Model list management
- Activity logs and statistics

## Requirements
- Python 3.7+
- Google Chrome browser
- ChromeDriver (matching your Chrome version)
- Selenium (`pip install selenium`)
- Tkinter (usually included with Python)
- Other standard Python libraries

## Installation

1. Clone the repository:
    ```bash
    git clone <repository-url>
    cd <project-directory>
    ```

2. Install dependencies:
    ```bash
    pip install selenium
    ```

3. Download ChromeDriver and place it in the specified path (default: `.\chromedriver\chromedriver.exe`).

4. Prepare a model list file (default: `model_list.txt`) with ASUS model names.

## Usage

1. Run the GUI:
    ```bash
    python bios_gui.py
    ```

2. Configure paths for ChromeDriver, model list, and download folder in the GUI.

3. Load your model list.

4. Click "Start Download" to begin downloading BIOS files.

5. Monitor progress and logs in the GUI.

## Configuration

- You can use a `config.json` file to specify default paths:
    ```json
    {
      "chromedriver": ".\\chromedriver\\chromedriver.exe",
      "model_list": ".\\model_list.txt",
      "logs": ".\\logs",
      "download_path": "E:\\BIOS"
    }
    ```

## Contributing
Issues and pull requests are welcome.

## License
Specify your project's license here.