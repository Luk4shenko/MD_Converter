# Document Converter

This Document Converter program is written in Go and is designed to convert .docx and .xlsx documents to Markdown format while preserving text styles and formatting.

## Features 
- Converts text with styling (bold, italic) to Markdown format
- Handles document headers and formats them as Markdown headings
- Processes lists and formats them accordingly in Markdown
- Converts Excel tables into Markdown format
## Usage
1. Clone or download the repository to your local machine.
2. Make sure you have Go installed on your system.
3. Navigate to the project directory using the command line.
4. Run the following command to compile the program:
```bashgo build -o DocumentConverter```
5. To convert a document, run the compiled executable with the input file path and the output directory as arguments. 
For example:
```bash./DocumentConverter input.docx /output/directory```
6. The converted Markdown file will be saved in the specified output directory.

## How to CompilePrerequisites:
- Go installed on your systemSteps:
    - Download or clone the repository
    - Open a terminal and navigate to the project directory
    - Run the command:
    ```bashgo build -o DocumentConverter```
    - The program will be compiled and you can use it to convert documents to Markdown format.