# Windows Theme based OpenRGB Colors

This script listens for changes in the Windows theme and sets the colors of the OpenRGB devices accordingly. 
It uses the Windows API to get the current theme and the OpenRGB SDK to set the colors.
Currently just sets the colors of all devices to the same color with a static effect.

## Requirements

- [OpenRGB](https://gitlab.com/CalcProgrammer1/OpenRGB)
- [pywin32](https://pypi.org/project/pywin32/)
- [openrgb-python](https://pypi.org/project/openrgb-python/)

## Setup

1. Install OpenRGB
2. Go to the SDK Server tab in OpenRGB and enable the SDK server
3. Optional: Create a virtual environment

    ```bash
    python -m .venv venv
    .\.venv\Scripts\activate
    ```

4.  Install the requirements

    ```bash
    pip install -r requirements.txt
    ```

5. Run the script

    ```bash
    python main.py
    ```

## Usage

1. Run the script
2. Change the Windows accent color
