## Zoom Transcription Cleaner ##

This program processes Word documents containing auto-transcribed text from Zoom meetings (exported as `.vtt` files) and converts them into more readable Word documents. The script removes unnecessary timestamps, organizes speaker dialogues, and ensures clean formatting for easier review.

## Features ##

- Reads `.docx` files containing transcription text exported from Zoom meetings.
- Removes unnecessary timestamps and metadata from the transcription.
- Combines speech by the same speaker for better readability.
- Saves the cleaned and formatted text to a new Word document with extra spacing between speakers.

## Setup and Usage ##

### Prerequisites ###
- Python 3.8 or later
- The following Python libraries:
  - `python-docx` (for Word document manipulation)
  - `re` (built-in for text processing)
