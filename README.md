# RPA SAP Daily Control - UiPath RPA

This RPA project automates daily control tasks in SAP, including monitoring, data extraction, processing, and notification. The bot makes extensive use of **VBA macros** for advanced Excel operations, enhancing performance and flexibility when processing SAP data. Built on UiPath and designed as a flowchart-based process, it leverages standard UiPath Excel, Mail, WebAPI, and UIAutomation activities.

## Project Description

The bot performs daily SAP control tasks to reduce manual effort in monitoring and reporting. It interacts with the SAP GUI using UI automation, extracts key control data, processes it using Excel (including VBA macros), and sends notifications through email or other integrated channels.

## Key Features

- Automates daily SAP monitoring & control routines
- Extracts control data from SAP and processes results
- Uses **VBA macros** for advanced Excel data processing
- Sends email notifications (optional)
- Flowchart-based workflow for clarity
- Modular design: easy to maintain & extend

## Project Structure

| Folder/File           | Description                                      |
|------------------------|--------------------------------------------------|
| Flowchart.xaml         | Main workflow (entry point)                     |
| VBA/                   | Folder containing Excel VBA macro files         |
| project.json           | UiPath project metadata                         |
| README.md              | This documentation                              |

## Workflow Overview

### 1. Initialization
- Loads configuration and prepares resources.
- Starts SAP session or checks if already running.

### 2. SAP Control Task
- Navigates SAP transaction codes (T-Codes).
- Captures relevant daily control data.

### 3. Data Processing with VBA
- Processes extracted data in Excel.
- Executes relevant VBA macros stored in `VBA/` folder for:
  - Data cleaning
  - Report formatting
  - Advanced calculations
  - Sheet manipulation

### 4. Notification
- Sends email notifications summarizing results.

### 5. End Process
- Logs completion and closes all resources gracefully.

## How VBA is Used

This bot relies heavily on VBA (Visual Basic for Applications) macros to perform Excel operations that are inefficient or impractical using only UiPath activities.

### VBA Macro Tasks

The bot leverages a set of VBA scripts located in the `VBA/` folder to handle advanced Excel operations, including:

- Cleaning and removing unnecessary rows
- Splitting columns (text-to-column) for multiple sheet types (e.g., BIK, PDDN, MEDICAL, COMMON, TRANSPORT, SAP)
- Copying and organizing categorized rows
- Filling down formulas and lookup values
- Validating result flags (e.g., NA, FALSE)
- Applying formatting (e.g., borders) and string conversion

These VBA scripts allow the bot to efficiently transform SAP-exported data into a structured and validated report.

### Integration:
- VBA scripts are stored in the `VBA/` folder.
- Macros are called using UiPath's **Execute Macro** activity.
- Requires Excel with macros enabled and trusted access to VBA project object model.

## How to Run

1. Open `Flowchart.xaml` in UiPath Studio.
2. Make sure the `VBA/` folder contains all required macro files.
3. Run the workflow manually or publish to Orchestrator for scheduling.
4. Monitor logs and output files as needed.

## Exception Handling

- Handles SAP UI automation failures with retries.
- Catches Excel/VBA execution errors and logs them.
- Sends error notifications if critical failures occur.

## Requirements

- UiPath Studio (Enterprise)
- SAP GUI installed and configured
- Microsoft Excel with macros enabled
- Trusted access to VBA project object model
- Outlook or SMTP (if sending emails)

## Dependencies

This project uses the following UiPath official packages:
- UiPath.Excel.Activities
- UiPath.Mail.Activities
- UiPath.System.Activities
- UiPath.Testing.Activities
- UiPath.UIAutomation.Activities
- UiPath.WebAPI.Activities

## Contact

For questions, improvements, or collaboration:

- Email: fadillah650@gmail.com  
- LinkedIn: [Enrico Naufal Fadilla](https://linkedin.com/in/enrico-naufal-fadilla-54338a256)
