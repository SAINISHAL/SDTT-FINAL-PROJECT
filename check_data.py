"""Quick diagnostic script to check what data is being loaded."""
import pandas as pd
from excel_loader import ExcelLoader
from config import INPUT_DIR

# Load the Excel file directly
course_file = f"{INPUT_DIR}/course_data.xlsx"
df = pd.read_excel(course_file)

print("="*60)
print("COURSE DATA ANALYSIS")
print("="*60)

print(f"\nTotal courses in Excel file: {len(df)}")

print("\nCourses by Semester:")
sem_counts = df.groupby('Semester').size()
for sem, count in sem_counts.items():
    print(f"  Semester {sem}: {count} courses")

print("\nCourses by Department:")
dept_counts = df.groupby('Department').size()
for dept, count in dept_counts.items():
    print(f"  {dept}: {count} courses")

print("\nChecking for issues:")

# Check for missing LTPSC
missing_ltpsc = df[df['LTPSC'].isna() | (df['LTPSC'] == '')]
print(f"  Courses with missing LTPSC: {len(missing_ltpsc)}")
if len(missing_ltpsc) > 0:
    print("    These courses will get default values based on credits")
    for _, row in missing_ltpsc.iterrows():
        print(f"      - {row['Course Code']}: {row['Course Name']}")

# Check for missing Credits
missing_credits = df[df['Credits'].isna() | (df['Credits'] == '')]
print(f"\n  Courses with missing Credits: {len(missing_credits)}")
if len(missing_credits) > 0:
    for _, row in missing_credits.iterrows():
        print(f"      - {row['Course Code']}: {row['Course Name']}")

print("\n" + "="*60)
print("Now testing with ExcelLoader to see what gets included...")
print("="*60)

# Test ExcelLoader
data_frames = ExcelLoader.load_all_data()
if data_frames:
    for semester in [1, 3, 5]:
        print(f"\nSemester {semester}:")
        sem_courses = ExcelLoader.get_semester_courses(data_frames, semester)
        print(f"  After get_semester_courses: {len(sem_courses)} courses")
        
        parsed_courses = ExcelLoader.parse_ltpsc(sem_courses)
        print(f"  After parse_ltpsc: {len(parsed_courses)} courses")
        
        if len(sem_courses) != len(parsed_courses):
            print(f"  WARNING: {len(sem_courses) - len(parsed_courses)} courses were excluded!")

print("\n" + "="*60)
print("DONE")
print("="*60)

