Square API Google Sheets Integration
This project provides a Google Apps Script to integrate Square API data with Google Sheets. It automates the retrieval and processing of catalog, inventory, and location data from Square, enabling seamless reporting and analysis within Google Sheets.

Features
Custom Google Sheets Menu
Adds a custom menu to Google Sheets titled Square API with options to configure and execute the integration:

Set API Key: Prompt to securely store your Square API access token.
Set Email Address: Set an email address for status notifications.
Start Processing: Fetches and processes data from Square, updating the sheet.
Set Daily Timer: Schedules a daily automated data refresh.
Data Processing and Reporting
Fetches:

Catalog items and variations
Inventory counts for all locations
Category and location information
Processes the data into a structured format, adding headers and filling in key details such as pricing, availability, and inventory levels.
Progress Tracking and Error Handling
Provides real-time progress updates in a dedicated sheet. Supports manual halting of the process and sends email notifications on success or failure.

Automated Scheduling
Supports daily automation with the Set Daily Timer function, running the data refresh process at a user-defined time.

Installation
Open a Google Sheet.
Go to Extensions > Apps Script.
Copy and paste the provided script into the editor.
Save and close the editor.

Usage
Reload the Google Sheet to see the Square API menu.
Use Set API Key and Set Email Address to configure the integration.
Run Start Processing to manually initiate data refresh.
Optionally, schedule automatic updates using Set Daily Timer.

Requirements
A valid Square API access token.
Permissions to install and run Google Apps Scripts.
Internet access for API communication.

Compliments of JTPets.ca
This project is brought to you by JTPets.ca, your local pet shop dedicated to providing quality products and services. We're passionate about sharing tools that enhance your business operations.

We welcome collaboration! If you have ideas, suggestions, or improvements, feel free to contribute or reach out. Together, we can build better solutions.

License
This project is licensed under the MIT License. See the LICENSE file for details.

Contributions
Contributions, issues, and feature requests are welcome! Feel free to check out the issues page.
