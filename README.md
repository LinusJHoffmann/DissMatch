# DissMatch
Dissertation Supervision Matcher

This Python script matches students to supervisors based on their preferences for supervision areas, using linear programming to maximise overall satisfaction. It is designed for academic settings where both students and supervisors rank their preferred research areas, and supervisors have a limited workload.

**Features**

- Assigns students to supervisors based on ranked area preferences from both parties.
- Ensures no supervisor is assigned more students than their specified workload.
- Uses a scoring system to maximise the overall satisfaction of both students and supervisors.
- Reads input data from an Excel file and writes the results to a new sheet in the same file.
- Prints assignment results, satisfaction scores, and summary statistics to the console.

**How It Works**

1. Data collection:
     - Establish a fixed list of research areas that your department can supervise (e.g. Constitutional Law, Competition Law, Human Rights Law, etc). Make sure the list is neither too granular, nor too coarse. 
     - Send a form to supervisors, asking them to rank their 5 preferred supervision areas. Using Microsoft Forms is efficient here, because it can directly output an excel file.
     - Send a form to students, asking them to rank their 3 preferred supervision areas. Using Microsoft Forms is efficient here, because it can directly output an excel file.
 
2. Data preparation: 
   - Collate all the data on student and supervisor preferences into the Excel file "dissmatch_data.xlsx" - make sure you insert the data in the correct sheets and columns. Don't change the names of the columns and sheets.
   - All the data is read from and written to this file. You can download the file from this repository. It contains a small testing dataset with fictional data - you can overwrite it with your real data. 
   - Place your Excel file in the same directory as the script (or change the filename in the script). 
   - Content of the `dissmatch_data.xlsx` Excel file:
     - Sheet 1: Supervisors (with columns for name, workload, and area preferences)
     - Sheet 2: Students (with columns for name and area preferences)
     - Sheet 3: Matches (this sheet is empty and will be automatically filled with the optimal matches)

3. Processing: 
   - Run the code.
     - It cleans and normalises area names to ensure accurate matching.
     - It builds a linear programming model to maximize satisfaction scores based on preferences and constraints.
     - It solves the assignment problem using the PuLP library = it finds all the matches that maximise the aggregate satisfaction = the highest possible choices will be matched.  

4. Output:  
   - The code prints assignments and statistics to the console.
   - It also writes the final matches (with student choice ranking) to a new sheet called `matches` in the Excel file.

**Requirements**

You need to install these to use DissMatch:

- Python 3.x
- pandas
- PuLP
- openpyxl

**Notes**

- You can adjust the satisfaction scoring system in the script to reflect your institution’s priorities. More info in the comments of the code
- The script assumes a fixed number of area preferences for both students and supervisors - the list of research areas that supervisors and students receive must be identical. 

**AI usage**
This project was developed with significant help from GitHub Copilot, which provided AI-powered code suggestions and refinements throughout the process. All code has been reviewed and curated by the project author.

**License**

MIT License 
