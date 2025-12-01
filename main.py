"""Main execution module for Excel-based timetable generation."""
import os
import sys
import pandas as pd

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from file_manager import FileManager
from excel_loader import ExcelLoader
from schedule_generator import ScheduleGenerator
from excel_exporter import ExcelExporter
from exam_scheduler import ExamScheduler
from seating_arrangement import SeatingArrangementGenerator
from config import TARGET_SEMESTERS, REQUIRED_FILES, DEPARTMENTS

class TimetableGenerator:
    """Main class to coordinate timetable generation from Excel files."""
    
    def __init__(self):
        self.data_frames = None
        self.schedule_generator = None
        self.excel_exporter = None
        self.exam_scheduler = None
    
    def setup_environment(self):
        """Set up the environment and load Excel data."""
        try:
            import openpyxl
            print("openpyxl is available")
        except ImportError:
            print("Installing openpyxl...")
            os.system('pip install openpyxl -q')
            import openpyxl
        
        FileManager.setup_directories ()
        
        if not FileManager.check_input_files_exist():
            print("ERROR: Required Excel files are missing.")
            print("Please ensure the following Excel files are in the input directory:")
            for file in REQUIRED_FILES:
                print("  -", file)
            print("Input directory:", FileManager.INPUT_DIR)
            FileManager.list_input_files()
            raise Exception("Missing required Excel files")
        
        self.data_frames = ExcelLoader.load_all_data()
        if self.data_frames is None:
            raise Exception("Failed to load data from Excel files")
        
        self.schedule_generator = ScheduleGenerator(self.data_frames)
        self.excel_exporter = ExcelExporter(self.data_frames, self.schedule_generator)
        self.exam_scheduler = ExamScheduler(self.data_frames, self.schedule_generator)
        print("Environment setup completed")
    
    def generate_timetables(self, semesters=None):
        """Generate timetables for specified semesters from Excel data."""
        if semesters is None:
            semesters = TARGET_SEMESTERS
        
        print("\n" + "="*80)
        print("GENERATING TIMETABLES FROM EXCEL FILES")
        print(f"INPUT DIRECTORY: {FileManager.INPUT_DIR}")
        print(f"OUTPUT DIRECTORY: {FileManager.OUTPUT_DIR}")
        print(f"DEPARTMENTS: {', '.join(DEPARTMENTS)}")
        print(f"TARGET SEMESTERS: {semesters}")
        print("SESSIONS: Pre-Mid, Post-Mid")
        print("="*80)
        
        success_count = 0
        for semester in semesters:
            print(f"\n\nPROCESSING SEMESTER {semester}")
            print("-" * 50)
            if self.excel_exporter.export_semester_timetable(semester):
                success_count += 1
                print(f"SUCCESS: Semester {semester} completed successfully")
            else:
                print(f"FAILED: Semester {semester} failed")
        
        return success_count
    
    def print_summary(self, success_count, total_semesters):
        """Print generation summary."""
        print("\n" + "="*80)
        if success_count == total_semesters:
            print("EXPORT COMPLETE!")
        else:
            print("EXPORT PARTIALLY COMPLETE!")
        
        print(f"Generated {success_count}/{total_semesters} timetable files")
        print("\nEach Excel file contains:")
        for dept in DEPARTMENTS:
            print(f"  - {dept}_Pre-Mid")
            print(f"  - {dept}_Post-Mid")
        print("  - Course_Summary sheet")
        print(f"\nFiles saved in: {FileManager.OUTPUT_DIR}")
        print("="*80)
    
    def get_data_summary(self):
        """Print summary of loaded Excel data."""
        if self.data_frames:
            print("\nEXCEL DATA SUMMARY:")
            for key, df in self.data_frames.items():
                print(f"  {key}: {len(df)} records")

            if 'course' in self.data_frames:
                course_df = self.data_frames['course']
                if 'Semester' in course_df.columns:
                    print("\nCOURSES BY SEMESTER:")
                    # Convert semester to numeric for proper sorting
                    course_df = course_df.copy()
                    course_df['Semester'] = pd.to_numeric(course_df['Semester'], errors='coerce')
                    course_df = course_df.dropna(subset=['Semester'])
                    course_df['Semester'] = course_df['Semester'].astype(int)
                    
                    for semester in sorted(course_df['Semester'].unique()):
                        sem_courses = course_df[course_df['Semester'] == semester]
                        print(f"  Semester {semester}: {len(sem_courses)} courses")

def main():
    """Main function to generate timetables from Excel files."""
    generator = TimetableGenerator()
    
    try:
        print("Starting Timetable Generator...")
        generator.setup_environment()
        generator.get_data_summary()
        success_count = generator.generate_timetables()
        
        # Generate 7th semester unified timetable with baskets
        try:
            print("\n" + "="*80)
            print("GENERATING SEMESTER 7 UNIFIED TIMETABLE")
            print("="*80)
            if generator.excel_exporter.export_semester7_timetable():
                print("SUCCESS: Semester 7 timetable generated")
            else:
                print("FAILED: Semester 7 timetable generation failed")
        except Exception as e:
            print(f"ERROR generating Semester 7 timetable: {e}")
            import traceback
            traceback.print_exc()
        
        generator.print_summary(success_count, len(TARGET_SEMESTERS))

        # Generate exam timetable
        try:
            print("\n" + "="*80)
            print("GENERATING EXAM TIMETABLE")
            print("="*80)
            if generator.exam_scheduler.export_exam_timetable():
                print("SUCCESS: Exam timetable generated")
            else:
                print("FAILED: Exam timetable generation failed")
        except Exception as e:
            print(f"ERROR generating exam timetable: {e}")
            import traceback
            traceback.print_exc()

        # Validate room allocation conflicts
        try:
            conflicts = generator.schedule_generator.validate_room_conflicts()
            if conflicts:
                print("\nROOM ALLOCATION CONFLICTS DETECTED:")
                for c in conflicts:
                    sem = c['semester']
                    day = c['day']
                    slot = c['slot']
                    room = c['room']
                    entries = "; ".join([f"{dept}:{course} ({session})" for dept, course, session in c['entries']])
                    print(f"  - {sem} {day} {slot} | Room {room} -> {entries}")
            else:
                print("\nNo room allocation conflicts detected.")
        except Exception as _e:
            print("\nRoom conflict validation could not be completed.")
        
        # Generate seating arrangements
        try:
            print("\n" + "="*80)
            print("GENERATING SEATING ARRANGEMENTS")
            print("="*80)
            seating_gen = SeatingArrangementGenerator(generator.data_frames, generator.schedule_generator)
            if seating_gen.generate_seating_arrangements():
                print("SUCCESS: Seating arrangements generated")
            else:
                print("FAILED: Seating arrangement generation failed")
        except Exception as e:
            print(f"ERROR generating seating arrangements: {e}")
            import traceback
            traceback.print_exc()
        
        return True

    except Exception as e:
        print(f"ERROR: {e}")
        print("\nPlease check:")
        print("1. The course_data.xlsx file exists in the input folder")
        print("2. The Excel file has the correct columns: Course Code, Course Name, Semester, LTPSC, Credits, Combined Class")
        print("3. The Semester values are numeric (1, 3, 5)")
        print("4. There are actual courses for semesters 3 and 5 in your Excel file")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    main()