# Test-bank-depacker
A python program used to format raw text bank questions to formatted multiple choice. Made with Pandas, Pillow,Docx

Description:
This Python script provides a solution for formatting and processing raw data from an Excel file to create a well-organized question bank document. It utilizes the power of pandas, docx, and PIL libraries to read the data, handle images, and generate a Word document with proper formatting.

Key Features:

Reads data from an Excel file, specifically the MCQs sheet, extracting the necessary columns including question, answer options, and image paths.
Determines the correct answers for each question based on the provided solution columns.
Appends the correct answer column to the DataFrame.
Creates a Word document using the docx library.
Iterates through each row of the data and performs the following steps:
Inserts the first image, specified by the 'Image 1' column, into the document.
Adds the question, answer options, and correct answer as paragraphs to the document.
Inserts the second image, specified by the 'Image 2' column, into the document.
The script supports both JPEG and PNG image formats. If the JPEG format is not supported or the image doesn't exist, it saves the image as PNG and inserts it into the document.
The resulting Word document is saved with a user-specified name.
Usage:

Install the required libraries: pandas, docx, and PIL (Pillow).
Specify the path to the Excel file containing the question bank data.
Update the 'Image 1' and 'Image 2' paths in the Excel file to point to the respective images.
Set the image folder path where the images are located.
Run the script, and it will generate a Word document with the formatted question bank.
Note:

Make sure to have the required images available at the specified image folder path.
Adjust the image width (Inches) according to your document layout.
Customize the name of the output Word document to meet your requirements.
Dependencies:

pandas: https://pandas.pydata.org/
python-docx: https://python-docx.readthedocs.io/
Pillow (PIL): https://python-pillow.org/
Feel free to customize the code as needed and adapt it to your specific use case.
