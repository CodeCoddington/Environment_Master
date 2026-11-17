# Environment_Master

## Overview
Welcome to Environment_Master! This is a .NET Windows Forms application designed to provide full control over Environment Variables. I originally built this as a personal home project, but it has practical applications for my work team. 

The main execution begins in `Program.cs`, which initializes the application configuration and launches the primary `Form_EnvVars` user interface.

---
## Features
*   **Environment Variable Management:** Grants full control to view and modify environment variables[cite: 1].
*   **Excel Reporting:** Exports environment variable data into formatted Excel reports.

---
## Setup Instructions
*   Clone the repository to your local machine.
*   Open the project in Visual Studio.
*   Build and run the project (Admin privileges are typically not required for standard debugging, aligning with our standard non-elevated Visual Studio workflow).

---
## ⚠️ Known Issues & Upcoming Changes
*   **Excel Export Path:** The output destination for the Excel reports is currently hard-coded to my personal Google Drive synced folder. **This will fail on your machine in its current state.** 
*   **Temporary Workaround:** I plan to update the application to pick a dynamic output destination soon. Until I push that update, you will need to manually change the hard-coded directory path in the source code to a valid local folder on your machine before generating any reports.

---