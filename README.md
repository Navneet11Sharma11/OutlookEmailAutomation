# Outlook Email Automation

## Overview
This VB.NET application processes emails in a specified folder in Microsoft Outlook, extracts the content of each email, and saves it into uniquely named text files. After processing, the emails are moved to another folder.

## Scope of Work
1. Open and connect to Microsoft Outlook.
2. Read emails from the RFQ_ORIGINAL folder.
3. Extract the contents of each email and save them into uniquely named text files.
4. Move processed emails to the RFQ_DONE folder.

## Technology Stack
- **Technology:** VB.NET (.NET Framework or .NET 8 based on your preference)
- **Libraries Used:** Microsoft Outlook Interop (for local Outlook processing)
- **File Handling:** System.IO for text file creation and management.
- **Error Handling:** Implement robust logging for debugging and smooth execution.

## Requirements
- Microsoft Outlook
- VB.NET
- .NET Framework or .NET 8

## Setup Instructions
1. Clone the repository.
2. Open the project in Visual Studio.
3. Restore the NuGet packages.
4. Build the project.

## How to Run
1. Launch Microsoft Outlook.
2. Run the application from Visual Studio.
3. The application will:
   - Connect to Outlook.
   - Read emails from the RFQ_ORIGINAL folder.
   - Save email contents to uniquely named text files.
   - Move processed emails to the RFQ_DONE folder.

## Error Handling
The application includes robust logging to capture and report errors for debugging and smooth execution.

## Contributions
Feel free to fork this repository, create a branch, and make changes. Submit a pull request for review.

## License
This project is licensed under the MIT License.
