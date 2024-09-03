# WhatsApp Bulk Message Automation

This project automates the process of sending bulk messages through WhatsApp Web using Python and Selenium. It also includes options for attaching images and polls to the messages.

## Features

- **Bulk Messaging**: Send messages to multiple contacts listed in a Google Sheet.
- **Image Attachment**: Optionally attach an image to each message.
- **Poll Attachment**: Optionally attach a poll with customizable options.

## Prerequisites

- Python 3.7 or later
- Google Chrome browser
- ChromeDriver (managed automatically by `webdriver_manager`)
- WhatsApp account

## Installation

1. Clone the repository:

   ```bash
   git clone <your-repo-url>
   cd <your-repo-directory>
   ```

2. Install the required Python packages:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Command-Line Arguments

- `sheet_name`: The name of the Google Sheet containing the contacts and messages (required).
- `--image-path`: The full path to the image to be attached (optional).
- `--attach-poll`: The name of the Google Sheet containing poll data (optional).

### Example

```bash
python app.py "YourSheetName" --image-path="H:/HLT.png" --attach-poll="PollSheetName"
```

### Google Sheet Structure

- The Google Sheet should have the following columns:
  - `Contact`: The phone number of the recipient (with or without the country code).
  - `Message`: The message to be sent.

- For polls, the poll sheet should have:
  - `Question`: The poll question.
  - Columns for each poll option.

## Notes

- The script will prompt you to scan the WhatsApp Web QR code upon the first run.
- Ensure that the contacts are in the correct format (e.g., starting with `+91` for India).

## Troubleshooting

- **Image Not Attaching**: Ensure that the image path is correct and accessible.
- **Timeout Errors**: Increase the wait times or ensure a stable internet connection.

## License

This project is licensed under the MIT License.
```
