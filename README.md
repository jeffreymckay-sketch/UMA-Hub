# Staff Portal

This Google Apps Script project provides a comprehensive staff portal for managing scheduling, logistics, proctoring, and other administrative tasks.

## Project Summary

The Staff Portal is a web-based application built on Google Apps Script that centralizes various administrative workflows into a single, easy-to-use interface. The application is designed to streamline processes, improve efficiency, and provide a central hub for all staff-related activities.

### Key Features

*   **Dashboard:** Provides a personalized overview of key information and upcoming tasks.
*   **Scheduling & Logistics:** Manage schedules, assignments, and other logistical details.
*   **Proctoring Services:** Tools for managing and monitoring proctoring sessions.
*   **Data Management:** Import, export, and manage data from various Google Sheets.
*   **Zoom Integration (Future):** Planned integration with Zoom for managing online meetings and recordings.
*   **Settings:** Customizable settings for configuring the application to meet your specific needs.

## Setup Instructions

To deploy and use the Staff Portal, follow these steps:

1.  **Clone the Repository:** Clone this repository to your local machine or Google Drive.
2.  **Open in Google Apps Script Editor:** Open the project in the Google Apps Script editor.
3.  **Configure `Config.gs`:** Open the `Config.gs` file and update the `MASTER_DATA_HUB_ID` variable with the ID of your master data Google Sheet.
4.  **Enable APIs:** In the Apps Script editor, go to **Resources > Advanced Google services** and ensure that the **Calendar API** is enabled.
5.  **Deploy as a Web App:**
    *   Go to **Deploy > New deployment**.
    *   Select **Web app** as the deployment type.
    *   Configure the web app settings as needed (e.g., access permissions).
    *   Click **Deploy**.
6.  **Access the Application:** Once deployed, you will be given a URL to access the Staff Portal.

## Additional Notes

*   The frontend of this application is built using HTML, CSS, and JavaScript, with the backend powered by Google Apps Script.
*   The application leverages several Google Workspace services, including Google Sheets, Google Calendar, and Google Drive.
*   For more detailed information on specific features, refer to the inline documentation within the source code.
