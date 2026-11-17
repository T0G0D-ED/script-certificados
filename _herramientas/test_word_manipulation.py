import win32com.client
import os

def test_word_replace(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File {file_path} not found.")
        return

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path)

        # Attempt to find and replace common patterns found in the CSV
        # For example, replacing "NOVIEMBRE" with "DICIEMBRE"
        # and updating the dates.
        
        print(f"Opening: {file_path}")
        
        # Search for "NOVIEMBRE" and replace with "DICIEMBRE"
        find_text = "NOVIEMBRE"
        replace_text = "DICIEMBRE"
        
        find = doc.Content.Find
        find.Execute(find_text, False, False, False, False, False, True, 1, False, replace_text, 2)
        
        # Update dates - this is more complex as we need to know the exact format
        # Based on CSV: "Desde el 01 de NOVIEMBRE 2026" -> "Desde el 01 de DICIEMBRE 2026"
        # "Hasta el 31 de DICIEMBRE 2026" -> "Hasta el 15 de ENERO 2027"
        
        # Let's try a simple replacement for the date pattern if it exists
        # Note: This is risky without knowing the exact content, but we are testing.
        
        doc.Save()
        doc.Close()
        print("Successfully updated and saved (test).")
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        word.Quit()

if __name__ == "__main__":
    # Use the file found in the previous dir command
    test_file = r"C:\Users\Maxi-\OneDrive\Escritorio\proyecto-certificados\CERTIFICADOS NOVIEMBRE 2026\1 ACONCAGUA FOODS\CERTIFICADOS\CARRIZAL-CHIMBARONGO\2026\01 NOVIEMBRE - CARRIZAL -  2026.doc"
    
    # Create a backup before testing
    backup_file = test_file + ".bak"
    import shutil
    shutil.copy(test_file, backup_file)
    print(f"Backup created at: {backup  backup_file}")
    
    test_word_replace(test_file)
    
    # Restore the original file
    shutil.move(backup_file, test_file)
    print("Original file restored.")
