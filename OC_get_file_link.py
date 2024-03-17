from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Authenticate and create the PyDrive client.
gauth = GoogleAuth()
gauth.LocalWebserverAuth()  # Creates local webserver and auto handles authentication.
drive = GoogleDrive(gauth)

# Name of the file you want to get the link for.
folder_title = "HW2"
folder_id = ''

# Retrieve the folder id - start searching from root
folder_list = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()
for folder in folder_list:
    if(folder['title'] == folder_title):
        folder_id = folder['id']
        break

print("Folder ID: ", folder_id)
# Search for the file by name.
file_list = drive.ListFile({'q': "\'"+folder_id+"\'" + " in parents and trashed=false"}).GetList()

for file in file_list:
    title = file['title']
    link = file['alternateLink']
    print('title: %s, link: %s' % (file['title'], file['alternateLink']))
    # Print the link to the file.
    # print("Link to the file:", file['alternateLink'])
else:
    print("File not found.")