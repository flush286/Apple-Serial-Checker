# PyAppleSerialChecker

## Overview

**PyAppleSerialChecker** is a Python script designed to check the warranty status of Apple products using their serial numbers. This tool loads serial numbers from an Excel file, fetches captcha images, and retrieves warranty information from the Apple coverage check website. The results are saved to an Excel file for easy review.

## Features

- Load serial numbers from Excel files
- Fetch and solve captchas automatically
- Retrieve warranty information from Apple’s website
- Save results in an Excel file

## Requirements

Before running the script, ensure you have the following installed:

- Python 3.x
- Required Python libraries:
  - `requests`
  - `pandas`
  - `easyocr`
  - `openpyxl`

You can install the necessary libraries using:

```bash
pip install requests pandas easyocr openpyxl
```

## Usage

1. Clone this repository or download the script file.
2. Prepare an Excel file containing the serial numbers.
3. Run the script:

   ```bash
   python pyapple_serial_checker.py
   ```

4. Follow the on-screen prompts to enter the path of your Excel file and the desired output filename.
5. Check the output Excel file for the warranty status of your Apple products.


## Contributing

Contributions are welcome! If you have suggestions or improvements, feel free to open an issue or submit a [Pull requests](pulls)!

> [!TIP]
> This project is licensed under the [MIT License](LICENSE). Feel free to fork or modify the script, but please mention my name in any derivative works.

> [!IMPORTANT]  
> A GUI executable version of this tool will be available soon! Stay tuned for updates.
