# iLikePDF
You know the amazing website for everything to do with PDFs - <a href = "https://www.ilovepdf.com/">iLovePDF</a>?<br>
Yeah, thought I'd make it open-source! 𓆩♡𓆪<br><br>

Simply use <a href = "https://github.com/adityapathak-cubastion/iLikePDF/tree/main/APIs">the APIs</a> to input a PDF and convert it to DOCX/PPTX.<br>
Or, if you have a directory full of PDFs that you want converted, use the <a href = "https://github.com/adityapathak-cubastion/iLikePDF/tree/main/pipelines-for-bulk-conversion">pipelines for bulk conversion</a>!

## Technologies Used
- Python and it's libraries - python-docx, python-pptx, regex, and more!
- <a href = "https://www.e-iceblue.com/">E-ICEBLUE's</a> incredible Python modules - Spire, Spire.PDF, Spire.Doc, Spire.Presentation
- API with Flask, tested with Postman
- Developed using V.S. Code

###### Note - You must have the dependencies required by Spire, such as LibSkiaSharp.dll.<br>It's advised that you have the C++ Build Tools (<a href = "https://learn.microsoft.com/en-us/cpp/build/vscpp-step-0-installation?view=msvc-170">Use Visual Studio for this!</a>).

## Future Scope
- More APIs for other tasks, such as merging many PDFs/DOCXs/PPTXs/etc.
- Supporting other conversions
- Frontend interface 

## Acknowledgements
Thanks to E-ICEBLUE, Spire for their products.<br>
Special thanks to `scanny@github`, I used some of his find-and-replace logic to remove watermarks for DOCX pipelines.
