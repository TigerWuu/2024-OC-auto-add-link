from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

from spire.doc import Document
from spire.doc import HyperlinkType

from argparse import ArgumentParser
parser = ArgumentParser()
parser.add_argument("-hw", "--homework folder", help="home work number", dest="hw", default="test", type=str)
parser.add_argument("-f", "--file", help="file name", dest="file", default="112-2 OC HW_test.docx", type=str)
args = parser.parse_args()

# Load an existing document
doc = Document()
# doc.LoadFromFile('112-2 OC HW3_240314.docx')
doc.LoadFromFile("./"+args.hw+"/"+args.file)

specific_word = "Link"

# Authenticate and create the PyDrive client.
gauth = GoogleAuth()
gauth.LocalWebserverAuth()  # Creates local webserver and auto handles authentication.
drive = GoogleDrive(gauth)

# Name of the file you want to get the link for.
folder_title = args.hw
folder_id = ''

# Retrieve the folder id - start searching from root
folder_list = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()
for folder in folder_list:
    if(folder['title'] == folder_title):
        folder_id = folder['id']
        break

# print("Folder ID: ", folder_id)
# Search for the file by name.
file_list = drive.ListFile({'q': "\'"+folder_id+"\'" + " in parents and trashed=false"}).GetList()

section = doc.Sections[0]
table = section.Tables[0]
for file in file_list:
    title = file['title']
    link = file['alternateLink']
    # print('title: %s, link: %s' % (file['title'], file['alternateLink']))
    # Print the link to the file.
    # print("Link to the file:", file['alternateLink'])
    for i in range(len(table.Rows)):
        paragraph = table.Rows[i].Cells[0].Paragraphs
        if paragraph[0].Text == title[0:9]:
            Link = table.Rows[i].Cells[2].Paragraphs[0]
            Link.AppendHyperlink(link, specific_word, HyperlinkType.WebLink)
            break
else:
    print("File not found.")


print("Hyperlink added successfully.")
doc.SaveToFile("./"+args.hw+"/"+args.file)
doc.Close()
