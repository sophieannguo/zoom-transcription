**Zoom Transcription Program**
Turns a file with auto-transcribed text from Zoom meetings into a more readable Word document.


Before you start:
Run pip install python-docx in your command line or terminal. (only need to do this once)

To clean a file:
Download the file from Zoom (VTT file) onto your computer.
Open the VTT file as a Word Document, and save.

Copy the file path of the Word document. (Ctrl+Shift+C)
Paste the file path on line 67. (set equal to input_path)
Ex. "C:\Users\guophie\Downloads\Filename.vtt.docx"
Make sure you do not remove the r before the quotes.
Write a new file name and path on line 68. (set equal to output_path)
Ex. "C:\Users\guophie\Downloads\Cleaned_Filename.docx"
Make sure you do not remove the r before the quotes.
Run the script.
Check your files to see the new cleaned output file.
