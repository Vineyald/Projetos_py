# ML Ad Automator
A semi-automated tool to streamline the management of Mercado Livre ad listings, reducing manual effort and ensuring accuracy. Built with Python, Pandas, and xlrd.

### Description
The ML Ad Automator simplifies the process of preparing ad listings by processing Mercado Livre's base files and user-filled spreadsheets. It enables users to quickly generate complete and consistent ad files, including options for premium and kit-based variations.

## Features
- File Handling: <br>
  Reads and processes: <br>
    - A Mercado Livre-generated base file.
    - A user-filled file containing ads already registered.
- Interactive Input: <br>
  Asks the user to identify whether products in the base file have kits and, if so, details about each kit (e.g., kit name and quantity).
- Automation Output: <br>
  Generates a final spreadsheet containing:
    - Classic and premium ad formats for all products.
    - Kit variations with updated titles, quantities, and prices.
- Error Handling:
  Ensures the accuracy of data processing and validates user inputs to avoid inconsistencies.

## How It Works
1. Setup: <br>
   The user provides two files: <br>
   - A Mercado Livre base file.
   - A user-prepared spreadsheet containing details of one-time ads.

2. Interactive Processing: <br>
    The script loops through all products in the input file and: <br>
   - Prompts the user to specify if the product has kits.
   - Collects the name and quantity of each kit variation.

3. Output Generation:
   - Combines all ad details, including kit variations, into a final spreadsheet.
   - Adds titles, quantities, and price adjustments for kits.
   - Creates both classic and premium ad formats.

## Lessons Learned
- Gained expertise in Python scripting for semi-automated workflows.
- Mastered Pandas for efficient data manipulation and processing of Excel files.
- Learned to implement interactive user input for handling dynamic data.
- Developed a deeper understanding of Mercado Livre's ad formats and requirements.
- Enhanced debugging and error-handling skills for robust script performance.
