# Get-ImageOCRText
Use Built-in Win10 OCR and SnippingTool to get text from images.
Used https://github.com/HumanEquivalentUnit/PowerShell-Misc/blob/master/Get-Win10OcrTextFromImage.ps1 as a base.
Modifications made:
 
 - Starts Snipping Tool (C:\Windows\System32\snippingtool.exe /clip) to grab image from screen
 - Pulls image from clipboard instead of saved image file
 - Pads image to enhance the OCR ability to recognize a screen capture with a small border around text
 - Preforms OCR on text and pushes result to the clipboard for easy pasting
