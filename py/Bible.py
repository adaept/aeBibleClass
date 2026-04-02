import sys
print("Python version")
print(sys.version)
print("Version info.")
print(sys.version_info)

import win32com.client

# Open Word application
word_app = win32com.client.Dispatch("Word.Application")
# Make Word visible (optional)
word_app.Visible = False

# Open the document
doc = word_app.Documents.Open('Peter-USE REFINED English Bible CONTENTS.docx')

# Initialize double space count
double_space_count = 0

# Iterate through all paragraphs in the document
for paragraph in doc.Paragraphs:
    text = paragraph.Range.Text
    # Count double spaces in the paragraph text
    double_space_count += text.count('  ')

# Print the total number of double spaces found
print(f"Total double spaces in the document: {double_space_count}")

# Initialize space count
total_spaces = 0

# Iterate through all shapes in the document
for shape in doc.Shapes:
    if shape.Type == 17:  # Type 17 corresponds to text boxes
        # Extract text from the text box
        text = shape.TextFrame.TextRange.Text
        # Count spaces in the text
        total_spaces += text.count(' ')

# Print the total number of spaces found in text boxes
print(f"Total spaces in text boxes: {total_spaces}")

# Close the document without saving changes
doc.Close(SaveChanges=False)
# Quit the Word application
word_app.Quit()


