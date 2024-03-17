from spire.doc import Document
from spire.doc import FieldType, Color, HyperlinkType

# Load an existing document
doc = Document()
# doc.LoadFromFile('112-2 OC HW3_240314.docx')
doc.LoadFromFile('test.docx')

specific_word = "Tiger"
hyperlink_url = "https://www.google.com"

section = doc.Sections[0]
table = section.Tables[0]
# for row in table.Rows:
    # Iterate through all cells in the row
for i in range(len(table.Rows)):
    paragraph = table.Rows[i].Cells[0].Paragraphs
    if paragraph[0].Text == "ddd":
        final = table.Rows[i].Cells[4].Paragraphs[0]
        final.AppendHyperlink(hyperlink_url, specific_word, HyperlinkType.WebLink)

    # # Iterate through all runs in the paragraph
    # for j in range(len(paragraph[0].Runs)):
    #     run = paragraph[0].Runs[j]
    #     # Check if the specific word is found in the run
    #     if specific_word in run.Text:
    #         # Split the run text into parts before and after the specific word
    #         parts = run.Text.split(specific_word)
    #         # Clear the original run text
    #         run.clear()
    #         # Add the part before the specific word
    #         run.add_text(parts[0])
    #         # Add the hyperlink to the specific word
    #         run.add_hyperlink(hyperlink_url, specific_word)
    #         # Add the part after the specific word
    #         run.add_text(parts[1])

# Iterate through all sections in the document
# for section in doc.Sections:
#     # Iterate through all tables in the section
#     print(section)
#     for table in section.Tables:
#         # Iterate through all rows in the table
#         for row in table.Rows:
#             # Iterate through all cells in the row
#             for cell in row.Cells:
#                 # Iterate through all paragraphs in the cell
#                 for paragraph in cell.Paragraphs:
#                     # Iterate through all runs in the paragraph
#                     for run in paragraph.Runs:
#                         # Check if the specific word is found in the run
#                         if specific_word in run.Text:
#                             # Split the run text into parts before and after the specific word
#                             parts = run.Text.split(specific_word)
#                             # Clear the original run text
#                             run.clear()
#                             # Add the part before the specific word
#                             run.add_text(parts[0])
#                             # Add the hyperlink to the specific word
#                             run.add_hyperlink(hyperlink_url, specific_word)
#                             # Add the part after the specific word
#                             run.add_text(parts[1])
# Get the paragraph where you want to add the hyperlink
# target_paragraph = doc.Sections[0].Paragraphs[0]
# 
# # Add a hyperlink to the paragraph
# field = target_paragraph.AppendField(FieldType.Hyperlink)
# field.link = "https://www.google.com.tw/"
# field.text = "Click here to visit Example"
# 
# # Save the document with the added hyperlink
doc.SaveToFile('test.docx')

print("Hyperlink added successfully.")
# Save to a different file
# doc.SaveToFile("112-2 OC HW3_240314.docx", FileFormat.Docx)
doc.Close()
