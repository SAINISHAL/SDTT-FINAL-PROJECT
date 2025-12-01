"""File management utilities for Excel file handling."""
import os
from config import INPUT_DIR, OUTPUT_DIR, REQUIRED_FILES

class FileManager:
    """Manages file operations for the timetable generator."""
    
    INPUT_DIR = INPUT_DIR
    OUTPUT_DIR = OUTPUT_DIR
    REQUIRED_FILES = REQUIRED_FILES
    
    @staticmethod
    def setup_directories():
        """Create input and output directories if they don't exist."""
        os.makedirs(FileManager.INPUT_DIR, exist_ok=True)
        os.makedirs(FileManager.OUTPUT_DIR, exist_ok=True)
        print(f"SUCCESS: Directories created")
        print(f"Input directory: {FileManager.INPUT_DIR}")
        print(f"Output directory: {FileManager.OUTPUT_DIR}")
    
    @staticmethod
    def check_input_files_exist():
        """Check if all required Excel files exist."""
        missing_files = []
        for filename in FileManager.REQUIRED_FILES:
            filepath = os.path.join(FileManager.INPUT_DIR, filename)
            if not os.path.exists(filepath):
                missing_files.append(filename)
        
        if missing_files:
            print("ERROR: Missing Excel files:", missing_files)
            print(f"Please ensure these files are in: {FileManager.INPUT_DIR}")
            return False
        else:
            print("SUCCESS: All required Excel files are present.")
            return True
    
    @staticmethod
    def get_output_path(filename):
        """Returns the full output path for a given filename."""
        return os.path.join(FileManager.OUTPUT_DIR, filename)
    
    @staticmethod
    def list_input_files():
        """List all files in the input directory."""
        if os.path.exists(FileManager.INPUT_DIR):
            files = os.listdir(FileManager.INPUT_DIR)
            print("Files in input directory:")
            for file in files:
                print("  -", file)
            return files
        else:
            print("Input directory does not exist.")
            return []