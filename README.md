# Social-Responsiveness-Scale-Calculator
The Social Responsiveness Scale Calculator is a python program that helps calculate various social responsiveness scales that helps determine the likelihood of a child to develop autism, it reads various points on an excel sheet and it calculates the score and inputs the scores in each respective column.
The tool automatically processes raw Excel files, extracts item-level responses, and computes standardized subscale scores including:

-AWR (Awareness)

-COG (Cognition)

-COM (Communication)

-MOT (Motivation)

-RRB (Restricted Interests & Repetitive Behaviors)

-SCI (Social Communication Index)

-Total Raw Score

Features:

Excel Parsing: Automatically reads .xlsx survey files and handles complex header structures.

Regex-Driven Column Matching: Flexible parsing even when column names differ slightly.

Automated Subscale Scoring: Implements SRS scoring formulas across multiple domains.

Data Validation: Detects missing or invalid responses.

Extensible: Easy to adapt for related questionnaires or additional scoring rules.

Graphical User Interface: For file selection and user-friendly experience

Technologies:

Python 3.x

pandas (data manipulation)

re (regex) (column parsing)

openpyxl (Excel support)

tkinter (GUI)

Usage:

1.Run the Program:
SRS.py

2.Choose your excel file and via GUI file picker 

3.The program will automatically process the file and output the calculated subscale scores.

Output will include subscale scores and total scores in a CSV or Excel format.

📊 Example Output:
ID	AWR	COG	COM	MOT	RRB	SCI	Total

001	15	22	18	12	9	   55	   76
002	11	19	14	9	  6	   43	   61

Future Improvements:

-Integration with R or SPSS for boarder research compatibility

-Batch scoring for multiple files simultaniously
001	15	22	18	12	9	55	76
002	11	19	14	9	6	43	61
