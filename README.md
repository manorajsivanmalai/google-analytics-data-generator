Google Analytics Data Generator

Technologies: Node.js, Google Analytics API (BetaAnalyticsDataClient), Google Auth Library, XlsxPopulate

Project Overview

The Google Analytics Data Generator is an automation tool that fetches Google Analytics data from multiple websites and generates clean Excel reports with diagrams and charts. It eliminates the need to manually log in to dashboards, saving time and reducing errors. Ideal for companies managing multiple websites.

Features

One-click data extraction from single or multiple websites

Excel report generation with structured data and visual diagrams

Scalable: Can handle 100+ websites at once

Automated reporting to save time and increase productivity

Customizable metrics and dimensions for detailed analytics

Secure authentication using Google Auth Library

Installation

Clone the repository:

git clone https://github.com/manorajsivanmalai/google-analytics-data-generator.git


Install dependencies:

npm install


Create a .env file for credentials:

GOOGLE_SERVICE_ACCOUNT_KEY='your_service_account_json_here'

Usage

Run the script:

node index.js


The tool will fetch data from the configured websites and generate an Excel report with charts.

Business Value

Automates analytics reporting for multiple websites

Provides visual insights through Excel diagrams

Reduces manual effort and human errors

Enables faster decision-making by management

Contributing

Feel free to contribute by opening an issue or pull request. Please do not commit any secrets or API keys.

License

MIT License
