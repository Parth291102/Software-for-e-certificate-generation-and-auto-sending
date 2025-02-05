# Software for E-Certificate Generation and Auto-Sending

## Overview

This project automates the process of generating and sending e-certificates using Python. It provides a single platform that generates certificates in bulk and automatically emails them to recipients. This is particularly useful for events, workshops, and educational institutions where a large number of certificates need to be processed efficiently.

## Features

- **Bulk Certificate Generation**: Uses predefined templates and dynamically fills in details.
- **Automated Email Sending**: Sends certificates as email attachments.
- **Customizable Templates**: Supports different certificate designs.
- **Efficient Data Handling**: Reads participant details from an Excel sheet.
- **Secure Email Transmission**: Uses SMTP for automated email delivery.

## Technologies Used

- Python
- OpenCV (cv2) for image processing
- OpenPyXL for Excel file handling
- smtplib for email automation
- MIME for email attachments

## Installation

Follow these steps to set up and run the project:

### 1. Clone the repository

```bash
git clone https://github.com/Parth291102/Software-for-e-certificate-generation-and-auto-sending.git
cd Software-for-e-certificate-generation-and-auto-sending
```

### 2. Create a virtual environment (optional but recommended)

```bash
python -m venv venv
source venv/bin/activate   # On macOS/Linux
venv\Scripts\activate     # On Windows
```

### 3. Run the Certificate Generator

```bash
python "Certificate Generator.py"
```

### 4. Run the Email Sender

```bash
python "Email Attachments.py"
```

## Usage

1. **Prepare Data**: Ensure participant details (name, email, etc.) are stored in an Excel file.
2. **Generate Certificates**: Run `Certificate Generator.py` to create certificates.
3. **Send Certificates via Email**: Run `Email Attachments.py` to distribute certificates to recipients automatically.

## Future Scope

- Add GUI for user-friendly operation.
- Support for multiple email providers.
- Enhance security features for email transmission.
- Expand to other document types like mark sheets and transcripts.

## Contributors

- **Parth Parikh**
- **Aditya Agrawal**
- **Dhruvi Patel**
- **Vaishali Dhare**

