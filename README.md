# FP&A Developer Assesment Statement

This is a Python-based GUI application that allows users to generate financial reports in various formats (Excel and PowerPoint) based on transaction data. The application provides options to filter and analyze data, generate dynamic reports, and create a presentation with insights.

## Features
- Generate Excel Report: Provides a summary of total transactions per client, per currency, and totals by client in USD.
- Generate Excel Report for Latin American Leaders: Filters transactions based on Latin American countries and identifies clients who had transactions with Russia, Cuba, China, or Venezuela in the last quarter of 2024.
- Generate PowerPoint Presentation: Creates an executive deck with analysis charts based on the data.
- Select Data Folder: The GUI allows the user to select a folder containing the transaction data for different years.
- Logs: The application creates logs of the process and errors for debugging.

## Getting Started

### Prerequisites:
- Python 3.8 or higher
- Required Python libraries:
  - tkinter (for GUI)
  - pandas (for data analysis)
  - matplotlib (for chart generation)
  - openpyxl (for Excel file manipulation)
  - requests (for API data fetching)
  - Pillow (for handling image assets in PowerPoint)
  - python-pptx (for generating PowerPoint presentations)
 
You can install these libraries using the following command:
```python
pip install pandas matplotlib openpyxl requests Pillow python-pptx
```

### Installation:

1. Clone the repository:
```bash
git clone https://github.com/andreacontrerasl/assessment-statement-mml.git
cd assessment-statement-mml
```

2. Ensure all necessary Python libraries are installed as mentioned in the prerequisites.

3. Prepare your transaction data files in the appropriate folder structure. The application requires data files to be present in a folder that you will select during execution.

### Usage:

1. Run the main application:
```bash
python main.py
```

2. A GUI window will open with the following sections:
- Select Data Folder: Use the "Browse" button to select the folder containing the transaction data. This folder can have data for different years.
- Select Actions to Perform: Choose one or more of the following options:
  - Option 1: Generate Excel Report (Deliverable 1): Generates an Excel report containing total transactions per client, per currency, and totals by client in USD.
  - Option 2: Create Deliverable 2: Generates an Excel report filtered for Latin American leaders. It includes transactions for Latin American countries and a list of all clients that had transactions with Russia, Cuba, China, or Venezuela in the last quarter of 2024.
  - Option 3: Executive Deck (PowerPoint): Generates a PowerPoint presentation with an advanced analysis of the data.
- Execute Button: Click this button to execute the selected actions. The progress of the execution will be displayed below.

3. Logs: The application will log important steps and errors in the console for debugging purposes.

### Description of Each Option:

1. Generate Excel Report (Deliverable 1): This option corresponds to the Question 1 requirement and Question 4. It creates an Excel file containing:
  - A pivot table showing total transactions per client, per currency.
  - A summary of total transactions in USD per client.
  - Charts for visualizing the top clients and their transaction values.
    
2. Create Deliverable 2: This option corresponds to the Question 2 requirement. It filters the transactions for Latin American countries and lists all clients who had transactions with Russia, Cuba, China, or Venezuela for the last quarter of 2024.

3. Executive Deck (PowerPoint): This option corresponds to Question 3. It generates an executive PowerPoint presentation using charts and analyses based on the data.

4. Select Data Folder: This part of the GUI allows the user to select a folder containing transaction data, enabling the generation of reports for different years. This functionality answers Question 4 and then execute Question 1.

### 
Certainly! Here is an example of a README.md file for your project:

FP&A Report Generator
This is a Python-based GUI application that allows users to generate financial reports in various formats (Excel and PowerPoint) based on transaction data. The application provides options to filter and analyze data, generate dynamic reports, and create a presentation with insights.

Features
Generate Excel Report: Provides a summary of total transactions per client, per currency, and totals by client in USD.
Generate Excel Report for Latin American Leaders: Filters transactions based on Latin American countries and identifies clients who had transactions with Russia, Cuba, China, or Venezuela in the last quarter of 2024.
Generate PowerPoint Presentation: Creates an executive deck with analysis charts based on the data.
Select Data Folder: The GUI allows the user to select a folder containing the transaction data for different years.
Logs: The application creates logs of the process and errors for debugging.
Getting Started
Prerequisites
Python 3.8 or higher
Required Python libraries:
tkinter (for GUI)
pandas (for data analysis)
matplotlib (for chart generation)
openpyxl (for Excel file manipulation)
requests (for API data fetching)
Pillow (for handling image assets in PowerPoint)
python-pptx (for generating PowerPoint presentations)
You can install these libraries using the following command:

bash
Copy code
pip install pandas matplotlib openpyxl requests Pillow python-pptx
Installation
Clone the repository:

bash
Copy code
git clone https://github.com/yourusername/fpa-report-generator.git
cd fpa-report-generator
Ensure all necessary Python libraries are installed as mentioned in the prerequisites.

Prepare your transaction data files in the appropriate folder structure. The application requires data files to be present in a folder that you will select during execution.

Usage
Run the main application:

bash
Copy code
python main.py
A GUI window will open with the following sections:

Select Data Folder: Use the "Browse" button to select the folder containing the transaction data. This folder can have data for different years.
Select Actions to Perform: Choose one or more of the following options:
Option 1: Generate Excel Report (Deliverable 1): Generates an Excel report containing total transactions per client, per currency, and totals by client in USD.
Option 2: Create Deliverable 2: Generates an Excel report filtered for Latin American leaders. It includes transactions for Latin American countries and a list of all clients that had transactions with Russia, Cuba, China, or Venezuela in the last quarter of 2024.
Option 3: Executive Deck (PowerPoint): Generates a PowerPoint presentation with an advanced analysis of the data.
Execute Button: Click this button to execute the selected actions. The progress of the execution will be displayed below.
Logs: The application will log important steps and errors in the console for debugging purposes.

Description of Each Option
Generate Excel Report (Deliverable 1): This option corresponds to the Question 1 requirement. It creates an Excel file containing:

A pivot table showing total transactions per client, per currency.
A summary of total transactions in USD per client.
Charts for visualizing the top clients and their transaction values.
Create Deliverable 2: This option corresponds to the Question 2 requirement. It filters the transactions for Latin American countries and lists all clients who had transactions with Russia, Cuba, China, or Venezuela for the last quarter of 2024.

Executive Deck (PowerPoint): This option corresponds to Question 3. It generates an executive PowerPoint presentation using charts and analyses based on the data.

Select Data Folder: This part of the GUI allows the user to select a folder containing transaction data, enabling the generation of reports for different years. This functionality answers Question 4.

### File Structure:

- main.py: Entry point of the application. Contains the GUI implementation and logic to execute the different reporting functions.
- load_files.py: Handles loading data files from the selected folder.
- generate_reports_deliverable1.py: Contains the functions to generate the Excel report for Option 1.
- generate_reports_deliverable2.py: Contains the functions to generate the filtered report for Option 2.
- generate_pp_derivable3.py: Handles the generation of charts and the PowerPoint presentation for Option 3.
