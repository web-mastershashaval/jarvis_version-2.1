import subprocess
import datetime
import os

def note(text):
    try:
        # Generate a file name with a timestamp
        date = datetime.datetime.now()
        file_name = str(date).replace(":", "-") + "-note.txt"
        
        # Save the note to the file
        with open(file_name, "w") as f:
            f.write(text)
        
        # Open the file with Notepad++
        notepad = "C:\\Windows\\system32\\notepad++.exe"
        subprocess.Popen([notepad, file_name])
        
        print(f"Note saved successfully as {file_name}")
    except Exception as e:
        print(f"An error occurred: {e}")

def delete_note(file_name):
    try:
        # Check if the file exists
        if os.path.isfile(file_name):
            # Delete the file
            os.remove(file_name)
            print(f"Note '{file_name}' deleted successfully.")
        else:
            print(f"File '{file_name}' does not exist.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
if __name__ == "__main__":
    note()
    
    # To delete a note, call the delete_note function with the file name
    # Example: delete_note("2024-09-12 12-30-45-note.txt")
