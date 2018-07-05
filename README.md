Table-Extractor-From-Image
==================================
This repository contains the code that extracts a table from an image and exports it to an Excel. To do this, the image is "read" by an OCR which provides a JSON output which is used as the input to the program. The program then arranges the cells row and column-wise as per the JSON input.

Modules Required
------------
<!-- To Be Updated Soon -->

Flow
------------
Image -> JSON -> Excel

Steps
------------
1. First of all, install all the import packages specified in the requirements.txt
2. For "reading" an image use an OCR that converts the format to JSON. 
    -    You can use your own OCR or use [Microsoft Azure Cognitive Services OCR API].
    -    Or you can upload the image at their [text reader demo]. The demo will give you the JSON of the image. Save the JSON to a notepad and run the program.
    
    
    
    
[Microsoft Azure Cognitive Services OCR API]: https://azure.microsoft.com/en-in/services/cognitive-services/computer-vision/
[text reader demo]: https://azure.microsoft.com/en-in/services/cognitive-services/computer-vision/#text
