# pylint: disable=missing-module-docstring
"""
PyAppleSerialChecker

This script provides functionality to check the warranty status of Apple products
using their serial numbers. It includes functions to load serial numbers from an
Excel file, fetch captcha images, and retrieve warranty information from the Apple
coverage check website. The results are saved to an Excel file.

Classes:
    Colors: ANSI escape sequences for colored output.
    Head: Contains headers for the serial number check results.

Functions:
    print_credits: Print the credits in colored ASCII art.
    load_serial_numbers: Load serial numbers from an Excel file.
    get_new_user_agent: Generate a new user agent string.
    get_auth_token: Retrieve the authentication token from the Apple coverage check website.
    fetch_captcha: Fetch a captcha image from the Apple coverage check website.
    exponential_backoff: Perform an exponential backoff with jitter.
"""
import base64
import json
import sys
import time
import os
import random
import re
import requests
import easyocr
import pandas as pd


class Colors:
    """
    ANSI escape sequences for colored output.
    """

    RED = "\033[31m"
    GREEN = "\033[32m"
    ORANGE = "\033[93m"
    YELLOW = "\033[33m"
    BLUE = "\033[34m"
    RESET = "\033[0m"
    CYAN = "\033[36m"
    MAGENTA = "\033[35m"

    @staticmethod
    def get_color(color_name):
        """
        Retrieve the ANSI color code for the given color name.

        Args:
            color_name (str): The name of the color.

        Returns:
            str: The ANSI color code for the given color name, or the reset code if not found.
        """
        return getattr(Colors, color_name.upper(), Colors.RESET)

    @staticmethod
    def list_colors():
        """
        List all available ANSI color codes.

        Returns:
            list: A list of color names.
        """
        return [
            attr
            for attr in dir(Colors)
            if not callable(getattr(Colors, attr)) and not attr.startswith("__")
        ]

    @staticmethod
    def print_colored_text(text, color_name):
        """
        Print text in the specified color.

        Args:
            text (str): The text to print.
            color_name (str): The name of the color.
        """
        color = Colors.get_color(color_name)
        print(f"{color}{text}{Colors.RESET}")

    @staticmethod
    def print_all_colors():
        """
        Print a sample text in all available colors.
        """
        for color in Colors.list_colors():
            Colors.print_colored_text(f"Sample text in {color}", color)


class Head:
    """
    Contains headers for the serial number check results.
    """

    serial_number = "Serial Number"
    product_name = "Product Name"
    purchase_date = "Purchase Date"
    coverage_expiry = "Coverage Expiry"
    status = "Status"

    @staticmethod
    def get_headers():
        """
        Retrieve the headers for the serial number check results.

        Returns:
            list: A list of header names.
        """
        return [
            Head.serial_number,
            Head.product_name,
            Head.purchase_date,
            Head.coverage_expiry,
            Head.status,
        ]

    @staticmethod
    def get_statuses():
        """
        Retrieve the possible statuses for a serial number check.

        Returns:
            list: A list of possible statuses.
        """
        return [Head.Status.not_found, Head.Status.valid]

    class Status:
        """
        Represents the status of a serial number check.
        """

        not_found = "NOT_FOUND"
        valid = "VALID"

        @staticmethod
        def get_headers():
            """
            Retrieve the headers for the serial number check results.

            Returns:
                list: A list of header names.
            """
            return [
                Head.serial_number,
                Head.product_name,
                Head.purchase_date,
                Head.coverage_expiry,
                Head.status,
            ]

        @staticmethod
        def get_statuses():
            """
            Retrieve the possible statuses for a serial number check.

            Returns:
                list: A list of possible statuses.
            """
            return [Head.Status.not_found, Head.Status.valid]


def print_credits():
    """
    Print the credits in colored ASCII art.
    """
    header_credits = [
        " ______   ______     ______     ______    ",
        " /\\  == \\ /\\  __ \\   /\\  ___\\   /\\  ___\\   ",
        "  \\ \\  _-/ \\ \\  __ \\  \\ \\___  \\  \\ \\ \\____  ",
        "    \\ \\_\\    \\ \\_\\ \\_\\  \\/\\_____\\  \\ \\_____\\ ",
        "      \\/_/     \\/_/\\/_/   \\/_____/   \\/_____/  ",
        "                                           ",
        "                by Flush                   ",
    ]

    colors = [
        Colors.RED,
        Colors.GREEN,
        Colors.YELLOW,
        Colors.BLUE,
        Colors.CYAN,
        Colors.MAGENTA,
    ]

    for i, line in enumerate(header_credits):
        color = colors[i % len(colors)]
        print(f"{color}{line}{Colors.RESET}")
        time.sleep(0.1)  # Slight delay for fade effect


# Print credits at the start
print_credits()


def load_serial_numbers(filename):
    """
    Load serial numbers from an Excel file.

    Args:
        filename (str): The path to the Excel file.

    Returns:
        list: A list of serial numbers.
    """
    df = pd.read_excel(filename)
    serial_column = [col for col in df.columns if "serial" in col.lower()]

    if not serial_column:
        print("No column containing 'serial' found in the Excel file.")
        return []

    return df[serial_column[0]].dropna().tolist()


def get_new_user_agent():
    """
    Generate a new user agent string.

    Returns:
        str: A new user agent string.
    """
    return (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_5) AppleWebKit/605.1.15 (KHTML, like Gecko) "
        "Version/16.5 Safari/605.1.15"
    )


def get_auth_token():
    """
    Retrieve the authentication token from the Apple coverage check website.

    Returns:
        str: The authentication token, or None if it cannot be retrieved.
    """
    auth_token_url = "https://checkcoverage.apple.com"
    auth_token_headers = {"User-Agent": get_new_user_agent()}
    auth_token_response = requests.get(
        auth_token_url, headers=auth_token_headers, timeout=10
    )
    try:
        return auth_token_response.headers["X-APPLE-AUTH-TOKEN"]
    except KeyError:
        print(
            f"{Colors.RED}[~] "
            f"Unable to retrieve the authentication token due to rate limit!"
            f"{Colors.RESET}"
        )
        return None


reader = easyocr.Reader(["en"])


def fetch_captcha(captcha_auth_token):
    """
    Fetch a captcha image from the Apple coverage check website.

    Args:
        captcha_auth_token (str): The authentication token.

    Returns:
        requests.Response: The response containing the captcha image.
    """
    captcha_url = "https://checkcoverage.apple.com/api/v1/facade/captcha?type=image"
    captcha_headers = {
        "X-Apple-Auth-Token": captcha_auth_token,
        "User-Agent": get_new_user_agent(),
        "Accept": "application/json",
    }
    captcha_response = requests.get(captcha_url, headers=captcha_headers, timeout=10)
    return captcha_response


def exponential_backoff(retry_attempt):
    """
    Perform an exponential backoff with jitter.

    Args:
        retry_attempt (int): The current retry attempt number.
    """
    wait_time = min(60, 2**retry_attempt)
    random_wait = random.uniform(0.5, 1.5) * wait_time
    time.sleep(random_wait)


# Prompt user for the Excel file path
file_path = input(
    "Please enter the path to the Excel file containing the serial numbers: "
)

# Load serial numbers from Excel
serial_numbers = load_serial_numbers(file_path)

if not serial_numbers:
    print("No valid serial numbers found. Exiting...")
    sys.exit(0)

# Ask user for the output Excel filename
output_filename = (
    input("Please enter the output Excel filename (without extension): ") + ".xlsx"
)

# Initialize the Excel file with headers if it doesn't exist
if not os.path.exists(output_filename):
    headers = [
        Head.serial_number,
        Head.product_name,
        Head.purchase_date,
        Head.coverage_expiry,
        Head.status,
    ]
    pd.DataFrame(columns=headers).to_excel(output_filename, index=False)

for serial_number in serial_numbers:
    print(f"{Colors.BLUE}[°]{Colors.RESET} Processing serial number: {serial_number}")

    auth_token = get_auth_token()
    if auth_token is None:
        sys.exit(0)

    CAPTCHA_INVALID_COUNT = 0  # Counter for invalid captcha attempts

    while True:
        CAPTCHA_FETCHED = False
        CAPTCHA_ANSWER = None
        RESULT_ENTRY = None

        for attempt in range(3):  # Allow up to 3 attempts
            response = fetch_captcha(auth_token)

            if b"but we are currently unable to process" in response.content:
                print(
                    f"{Colors.BLUE}[~]{Colors.RESET} "
                    f"Rate limit reached while fetching captcha, "
                    f"waiting..."
                )
                exponential_backoff(1)
                continue

            response_json = json.loads(response.content)

            if "binaryValue" in response_json:
                captcha = base64.b64decode(response_json["binaryValue"])
                with open("captcha.png", "wb") as f:
                    f.write(captcha)

                result = reader.readtext("captcha.png")

                if result:
                    CAPTCHA_ANSWER = result[0][1]
                    confidence = result[0][2]
                    CAPTCHA_FETCHED = True
                    break
                print(f"{Colors.RED}[?]{Colors.RESET} Failed to detect captcha text.")
                print(
                    f"{Colors.RED}[?]{Colors.RESET} "
                    f"Failed to retrieve captcha, trying again..."
                )

        if not CAPTCHA_FETCHED:
            print(
                f"{Colors.BLUE}[°]{Colors.RESET} "
                f"Maximum captcha attempts reached. Refreshing "
                f"the connection..."
            )
            auth_token = get_auth_token()
            if auth_token is None:
                sys.exit(0)
            continue

        # Get serial status
        URL = "https://checkcoverage.apple.com/api/v1/facade/coverage"
        headers = {"X-Apple-Auth-Token": auth_token, "User-Agent": get_new_user_agent()}
        json_data = {
            "captchaAnswer": CAPTCHA_ANSWER,
            "captchaType": "image",
            "serialNumber": serial_number,
        }

        response = requests.post(URL, headers=headers, json=json_data, timeout=10)

        if b"process your request." in response.content:
            print(
                f"{Colors.BLUE}[~]{Colors.RESET} "
                f"Rate limit reached while checking serial number status, "
                f"waiting..."
            )
            exponential_backoff(1)
            continue

        if b"The code you entered does not match" in response.content:
            CAPTCHA_INVALID_COUNT += 1
            print(
                f"{Colors.ORANGE}[!]{Colors.RESET} "
                f"Invalid captcha, trying again... ({CAPTCHA_INVALID_COUNT}/6)"
            )
            if CAPTCHA_INVALID_COUNT >= 6:
                print(
                    f"{Colors.BLUE}[°]{Colors.RESET} "
                    f"Maximum number of invalid captcha attempts reached. "
                    f"Refreshing the connection..."
                )
                auth_token = get_auth_token()
                if auth_token is None:
                    sys.exit(0)
                CAPTCHA_INVALID_COUNT = 0  # Reset the counter
            continue

        if b"Please enter a valid serial number." in response.content:
            print("The serial number is invalid, can be safely used!")
            RESULT_ENTRY = {
                Head.serial_number: serial_number,
                Head.product_name: "N/A",
                Head.purchase_date: "N/A",
                Head.coverage_expiry: "N/A",
                Head.status: "Invalid",
            }
            break

        if b"Sign in to update purchase date" in response.content:
            print(
                "The serial number is: "
                "Unable to verify purchase date, please regenerate!"
            )

            RESULT_ENTRY = {
                Head.serial_number: serial_number,
                Head.product_name: "N/A",
                Head.purchase_date: "N/A",
                Head.coverage_expiry: "N/A",
                Head.status: "Cannot verify purchase date",
            }
            break

        if (
            b"Your coverage includes the following benefits" in response.content
            or b"Coverage Expired" in response.content
        ):
            print("The serial number is: Fully valid, please regenerate!")

            RESULT_ENTRY = {
                Head.serial_number: serial_number,
                Head.product_name: "N/A",
                Head.purchase_date: "N/A",
                Head.coverage_expiry: "N/A",
                Head.status: "Fully valid",
            }
            break

        if b"We cannot process your request at this time." in response.content:
            print("AWWW MEEEEN")
            RESULT_ENTRY = {
                Head.serial_number: serial_number,
                Head.product_name: "N/A",
                Head.purchase_date: "N/A",
                Head.coverage_expiry: "N/A",
                Head.status: "Cannot process request",
            }
            break

        if b"Apple coverage for your product" in response.content:
            print("[=] Success!")
            decoded_data = response.content.decode("utf-8", errors="ignore")

            PRODUCT_NAME_PATTERN = r"MacBook\s+[^\)]*\s*\([^\)]*\)"
            SERIAL_NUMBER_PATTERN = r"([A-Z]{1,2}\d{10})"
            PURCHASE_DATE_PATTERN = r"(\w+\s\d{4})"
            COVERAGE_EXPIRY_PATTERN = r'Expires on\s*:\s*([^"]+)'

            product_name_match = re.search(PRODUCT_NAME_PATTERN, decoded_data)
            serial_number_match = re.search(SERIAL_NUMBER_PATTERN, decoded_data)
            purchase_date = re.search(PURCHASE_DATE_PATTERN, decoded_data)
            coverage_expiry = re.search(COVERAGE_EXPIRY_PATTERN, decoded_data)

            def get_match_or_default(match, default):
                """
                Retrieve the matched string or return a default value if no match is found.

                Args:
                    match (re.Match or None): The match object from a regular expression
                        search.
                    default (str): The default value to return if no match is found.

                Returns:
                    str: The matched string, stripped of leading and trailing whitespace,
                        or the default value.
                """
                return match.group(0).strip() if match else default

            RESULT_ENTRY = {
                Head.serial_number: serial_number,
                Head.product_name: get_match_or_default(
                    product_name_match, Head.Status.not_found
                ),
                Head.purchase_date: get_match_or_default(
                    purchase_date, Head.Status.not_found
                ),
                Head.coverage_expiry: get_match_or_default(
                    coverage_expiry, Head.Status.not_found
                ),
                Head.status: Head.Status.valid,
            }
            break

        print("An unknown error occurred!")
        print(response.content)
        RESULT_ENTRY = {
            Head.serial_number: serial_number,
            Head.product_name: "N/A",
            Head.purchase_date: "N/A",
            Head.coverage_expiry: "N/A",
            Head.status: "Unknown",
        }
        break

    # Append result to Excel file
    result_df = pd.DataFrame([RESULT_ENTRY])
    with pd.ExcelWriter(
        output_filename, engine="openpyxl", mode="a", if_sheet_exists="overlay"
    ) as writer:
        result_df.to_excel(
            writer, index=False, header=False, startrow=writer.sheets["Sheet1"].max_row
        )

print(f"{Colors.GREEN}Results have been saved to {output_filename}.{Colors.RESET}")
