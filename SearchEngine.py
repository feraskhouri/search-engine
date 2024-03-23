from pathlib import Path
import win32com.client
import os
# Path settings
input_dir = "D:\school\school\search engine"

# Prompt user to enter the search word
find_str = input("Enter the word to search for: ")

# Find settings
wd_find_wrap = 1  # 2=ask to continue, 1=continue search, 0=end if search range is reached

# Open Word
word_app = win32com.client.DispatchEx("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = False

files_with_word = []

for doc_file in Path(input_dir).rglob("*.doc*"):
    # Open each document and search for the string
    word_app.Documents.Open(str(doc_file))
    # API documentation: https://learn.microsoft.com/en-us/office/vba/api/word.find.execute
    found = word_app.Selection.Find.Execute(
        FindText=find_str,
        Forward=True,
        MatchCase=True,
        MatchWholeWord=False,
        MatchWildcards=True,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Wrap=wd_find_wrap,
        Format=True,
    )

    if found:
        files_with_word.append(doc_file.name)
        print(f"Found '{find_str}' in {doc_file.name}")

    # Close the document without saving changes
    word_app.ActiveDocument.Close(SaveChanges=False)

word_app.Application.Quit()

print("Files containing the word:", files_with_word)
