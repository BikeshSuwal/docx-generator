# docx-generator
Word File Generator using PySimpleGUI and docxtpl
Description
This Python script allows you to generate Word documents using a template provided in the template.docx file. The user is prompted to input various details such as project name, location, population, area, latitude, longitude, and more. The script uses the PySimpleGUI library for creating the user interface and the docxtpl library for rendering the template and generating the Word document.

Usage
Ensure you have Python 3.7 or higher installed.
Install the required libraries using pip:
bash
Copy code
pip install PySimpleGUI docxtpl
Clone or download the repository to your local machine.
Place the template.docx file in the same directory as the script.
Run the script word_file_generator.py:
bash
Copy code
python word_file_generator.py
Follow the prompts in the GUI window to input the necessary details.
Click the "Create File" button to generate the Word document.
The generated Word document will be saved in the same directory with the specified filename.
Dependencies
PySimpleGUI
docxtpl
Author
Bikesh Suwal

Acknowledgments
Inspired by the need to automate the generation of Word documents based on a template with user-provided information.

Note
The script uses the template.docx file as the base template for generating the Word document. Make sure to customize the template according to your needs before running the script.
