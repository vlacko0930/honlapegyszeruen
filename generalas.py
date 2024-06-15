# docx kiterjesztesu fajlok listaja az oldalak mappaban rekurzivan

import os
import docx

def getMetaData(file_path):
    doc = docx.Document(file_path)
    metadata = {}
    prop = doc.core_properties
    metadata["author"] = prop.author
    metadata["category"] = prop.category
    metadata["comments"] = prop.comments
    metadata["content_status"] = prop.content_status
    metadata["created"] = prop.created
    metadata["identifier"] = prop.identifier
    metadata["keywords"] = prop.keywords
    metadata["last_modified_by"] = prop.last_modified_by
    metadata["language"] = prop.language
    metadata["modified"] = prop.modified
    metadata["subject"] = prop.subject
    metadata["title"] = prop.title
    metadata["version"] = prop.version
    return metadata

def docx_files():
    docx_files = []
    for root, dirs, files in os.walk('oldalak'):
        for file in files:
            if file.endswith('.docx'):
                docx_files.append(os.path.join(root, file))
    return docx_files

def generateMarkdown():
    for file in docx_files():
        mediaFolderName = file[:-5].replace('/', '_').replace(' ', '_').replace('\\','_')
        os.system(f'pandoc --extract-media ./oldalak/media/{mediaFolderName} -t markdown_mmd "{file}" -o "{file[:-5]}.md"')

        #Replace media folder path in md file
        # Read in the file
        with open(f"{file[:-5]}.md", 'r') as fileOpen:
            filedata = fileOpen.read()

        # Replace the target string
        filedata = filedata.replace(f'./oldalak/media/{mediaFolderName}', f'/media/{mediaFolderName}')

        # Write the file out again
        with open(f"{file[:-5]}.md", 'w') as fileOpen:
            fileOpen.write(filedata)
        
        if file.startswith('oldalak/hírfolyam/hírek'):
            md = getMetaData(file)
        
            with open(f"{file[:-5]}.md", 'r') as fileOpen:
                filedata = fileOpen.read()
            
            with open(f"{file[:-5]}.md", 'w') as fileOpen:
                fileOpen.write(f"---\nauthor: {md['last_modified_by']}\ndate: {md['modified']}\n---\n\n{filedata}")
            

            



print(docx_files())
generateMarkdown()

