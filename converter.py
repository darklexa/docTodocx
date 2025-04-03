import os
import win32com.client

def convert_doc_to_docx(source_folder, target_folder):
    # Create target folder if it doesn't exist
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # Initialize Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # keep Word invisible
    
    # Get list of files in the source folder
    for filename in os.listdir(source_folder):
        # Process only .doc files that are not already .docx
        if filename.lower().endswith(".doc") and not filename.lower().endswith(".docx"):
            source_file = os.path.join(source_folder, filename)
            base_filename = os.path.splitext(filename)[0]
            target_file = os.path.join(target_folder, base_filename + ".docx")
            
            try:
                # Open the .doc file
                doc = word.Documents.Open(source_file)
                # Save as .docx
                # FileFormat=16 corresponds to wdFormatDocumentDefault which creates a .docx file.
                doc.SaveAs(target_file, FileFormat=16)
                doc.Close()
                print(f"{filename} converted to {base_filename + '.docx'} successfully.")
            except Exception as e:
                print(f"{filename} file couldn't be converted. Error: {e}")
                continue
    
    # Quit Word
    word.Quit()

if __name__ == "__main__":
    # Define the source and target directories.
    # Update these paths if needed.
    source_dir = r"C:\source"
    target_dir = r"C:\target"
    
    convert_doc_to_docx(source_dir, target_dir)
