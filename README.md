# Email to Route Sheet Automation

## Overview
This application automates the process of creating route sheets for booking agents by monitoring sent emails and automatically extracting relevant venue information into a Google Spreadsheet. It uses Gmail API to monitor outgoing emails and OpenAI's GPT-3.5 to intelligently extract booking-related information.

## Features
- 📧 Monitors sent emails in real-time
- 🤖 Automatically extracts key booking information using AI:
  - Venue name
  - City
  - Requested dates
  - Contact email
- 📊 Updates Google Sheets automatically with extracted data
- 🔄 Continuous monitoring with error handling and logging
- 💾 Maintains processing state to prevent duplicate entries

## Prerequisites
- Python 3.6+
- Google Cloud Project with Gmail and Sheets APIs enabled
- OpenAI API key
- Google OAuth 2.0 credentials

## Setup

1. Clone the repository
2. Install dependencies: `pip install -r requirements.txt`
3. Set up Google Cloud Project and enable Gmail and Sheets APIs
4. Create a `.env` file with your OpenAI API key and Google OAuth credentials
5. Run the script: `python app.py`
