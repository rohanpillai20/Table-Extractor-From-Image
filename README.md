Table Extractor From Image
==================================
This repository contains the code that extracts a table from an image and exports it to an Excel. To do this, the image is "read" by an OCR which provides a JSON output which is used as the input to the program. The program then arranges the cells row and column-wise as per the JSON input.

NOTE: Still in development mode.

Modules Required
------------
os<br>
copy<br>
pandas==0.22.0<br>
openpyxl==2.4.9<br>
You can also use requirements.txt to install the packages. How? Follow this [link].

Flow
------------
Image -> JSON -> Excel

Steps
------------
1. First of all, install all the import packages specified in the requirements.txt
2. For "reading" an image, use an OCR that converts the format to JSON. 
    -    You can use your own OCR or use [Microsoft Azure Cognitive Services OCR API].
    -    Or you can upload the image at their [text reader demo]. The demo will give you the JSON of the image. Save the JSON to a notepad and run the program.
3. In the program, change the input path and output path according to your requirement.
4. Run the program (JSON-to-Excel.py).
    
Sample Test Case
------------    

Input Image:

![Input Image](https://imgur.com/a/DG1I8WM)

It's Corresponding JSON:

![JSON](https://imgur.com/a/winhl5x)

Excel Output:

![Excel Output](https://imgur.com/a/mAyMxJk)
    
[Microsoft Azure Cognitive Services OCR API]: https://azure.microsoft.com/en-in/services/cognitive-services/computer-vision/
[text reader demo]: https://azure.microsoft.com/en-in/services/cognitive-services/computer-vision/#text
[link]: https://stackoverflow.com/questions/7225900/how-to-install-packages-using-pip-according-to-the-requirements-txt-file-from-a
