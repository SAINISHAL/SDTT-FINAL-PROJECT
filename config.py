"""Configuration settings for the timetable generator."""
import os

# Get the directory where this config file is located
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


INPUT_DIR_PATH = "C:\\Users\\saini\\OneDrive\\Desktop\\SDTT FINAL PROJECT\\sdtt_inputs" 
OUTPUT_DIR_PATH ="C:\\Users\\saini\\OneDrive\\Desktop\\SDTT FINAL PROJECT\\output" 

# Use direct path if provided, otherwise use relative path
if INPUT_DIR_PATH:
    INPUT_DIR = INPUT_DIR_PATH
else:
    INPUT_DIR = os.path.join(BASE_DIR, 'sdtt_inputs')

if OUTPUT_DIR_PATH:
    OUTPUT_DIR = OUTPUT_DIR_PATH
else:
    OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

# Required Excel input files (essential for timetable generation)
REQUIRED_FILES = [
    'course_data.xlsx',
    'classroom_data.xlsx'
]

# Optional Excel input files (used for additional features)
OPTIONAL_FILES = [
    'faculty_availability.xlsx',  # Used for invigilation assignments
    'student_data.xlsx',          # Used for enrollment/student information
    'exam_data.xlsx'              # Used for exam-specific data
]

# Departments
DEPARTMENTS = ['CSE-A', 'CSE-B', 'DSAI', 'ECE']

# Target semesters
TARGET_SEMESTERS = [1, 3, 5, 7]
# Session types
PRE_MID = 'Pre-Mid'
POST_MID = 'Post-Mid'

# Minor subject
MINOR_SUBJECT = "Minor"

# Time scheduling configuration
DAYS = ['MON', 'TUE', 'WED', 'THU', 'FRI']

# Teaching time slots in 30-minute increments (07:30 - 17:30)
TEACHING_SLOTS = [
    '07:30-08:00', '08:00-08:30', 
    '09:00-09:30', '09:30-10:00', '10:00-10:30',
    '10:30-11:00', '11:00-11:30', '11:30-12:00',
    '12:00-12:30', '12:30-13:00',
    '13:00-13:30', '13:30-14:00',  # Lunch slots
    '14:00-14:30', '14:30-15:00',
    '15:00-15:30', '15:30-16:00',
    '16:00-16:30', '16:30-17:00',
    '17:00-17:30','17:30-18:00',
]

# Lunch and Minor slot definitions
LUNCH_SLOTS = ['13:00-13:30', '13:30-14:00']
MINOR_SLOTS = ['07:30-08:00', '08:00-08:30']  # 07:30-08:30 represented as two 30-min slots

# Class durations (counted in 30-minute slots)
LECTURE_DURATION = 3    # 1.5 hours = 3 slots
TUTORIAL_DURATION = 2   # 1 hour = 2 slots
LAB_DURATION = 4        # 2 hours = 4 slots (consecutive slots)
MINOR_DURATION = 2      # 1 hour = 2 slots

# Weekly frequency
MINOR_CLASSES_PER_WEEK = 2