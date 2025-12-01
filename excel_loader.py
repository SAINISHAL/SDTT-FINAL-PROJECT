"""Excel file loader for timetable generator."""
import pandas as pd
import os
import re
import math
from config import INPUT_DIR, REQUIRED_FILES, OPTIONAL_FILES, DEPARTMENTS, MINOR_SUBJECT

class ExcelLoader:
    """Handles loading of Excel files for all data inputs."""
    _department_normalization_map = None
    _two_credit_course_session_map = {}

    @staticmethod
    def _get_department_normalization_map():
        """Build a normalization map for department names to ensure consistent matching."""
        if ExcelLoader._department_normalization_map is None:
            normalization_map = {}

            def add_variants(label, target):
                variants = set()
                base = str(label).strip().upper()
                variants.add(base)
                variants.add(base.replace(" ", ""))
                variants.add(base.replace("-", ""))
                variants.add(base.replace("-", " "))
                variants.add(base.replace(" ", "-"))
                for variant in variants:
                    key = re.sub(r"[^A-Z0-9]", "", variant)
                    if key:
                        normalization_map[key] = target

            for dept in DEPARTMENTS:
                add_variants(dept, dept)

            # Add generic CSE base label to ensure base entries are handled
            add_variants("CSE", "CSE")

            ExcelLoader._department_normalization_map = normalization_map

        return ExcelLoader._department_normalization_map

    @staticmethod
    def _normalize_department_label(label):
        """Normalize department names to match configured department labels."""
        if pd.isna(label):
            return ""

        label_str = str(label).strip()
        if label_str == "":
            return ""

        normalization_map = ExcelLoader._get_department_normalization_map()
        lookup_key = re.sub(r"[^A-Z0-9]", "", label_str.upper())
        normalized = normalization_map.get(lookup_key)

        if normalized is None:
            # Try fallbacks: if string starts with CSE, default to base CSE (will split later)
            if lookup_key.startswith("CSE"):
                normalized = "CSE"
            else:
                normalized = label_str.strip()

        return normalized
    
    @staticmethod
    def load_all_data():
        """Loads all Excel files from the input directory into a dictionary of DataFrames."""
        data_frames = {}
        
        # Load required files
        for filename in REQUIRED_FILES:
            filepath = os.path.join(INPUT_DIR, filename)
            try:
                if os.path.exists(filepath):
                    df = pd.read_excel(filepath)
                    # Normalize column names: strip whitespace and handle common variations
                    df.columns = df.columns.str.strip()
                    # Handle specific column name variations
                    column_mapping = {}
                    for col in df.columns:
                        col_upper = str(col).upper()
                        # Normalize "COMBINED CLASS " (with trailing space) to "Combined Class"
                        if 'COMBINED' in col_upper and 'CLASS' in col_upper:
                            column_mapping[col] = 'Combined Class'
                    if column_mapping:
                        df = df.rename(columns=column_mapping)
                    key = filename.replace('_data.xlsx', '').replace('.xlsx', '')
                    data_frames[key] = df
                    print(f"SUCCESS: Loaded {filename} ({len(df)} records)")
                    print(f"Columns: {df.columns.tolist()}")
                    print(f"First few rows:")
                    print(df.head())
                    
                    # Check for additional sheets in course_data.xlsx (e.g., "7th sem ", "Elective Data", "Minor Data" sheets)
                    if filename == 'course_data.xlsx':
                        try:
                            xl_file = pd.ExcelFile(filepath)
                            # Priority sheets to load (Elective Data and Minor Data)
                            priority_sheets = ['Elective Data', 'Minor Data']
                            for sheet_name in xl_file.sheet_names:
                                # Load all additional sheets
                                if sheet_name.lower() not in ['course data', 'sheet1'] and sheet_name.strip():
                                    sheet_df = pd.read_excel(filepath, sheet_name=sheet_name)
                                    # Normalize column names for additional sheets too
                                    sheet_df.columns = sheet_df.columns.str.strip()
                                    column_mapping = {}
                                    for col in sheet_df.columns:
                                        col_upper = str(col).upper()
                                        if 'COMBINED' in col_upper and 'CLASS' in col_upper:
                                            column_mapping[col] = 'Combined Class'
                                    if column_mapping:
                                        sheet_df = sheet_df.rename(columns=column_mapping)
                                    # Store with a key based on sheet name
                                    sheet_key = f"course_{sheet_name.strip().lower().replace(' ', '_')}"
                                    data_frames[sheet_key] = sheet_df
                                    print(f"SUCCESS: Loaded additional sheet '{sheet_name}' from {filename} ({len(sheet_df)} records)")
                        except Exception as e:
                            print(f"INFO: Could not load additional sheets from {filename}: {e}")
                else:
                    print(f"ERROR: File not found: {filepath}")
                    return None
            except Exception as e:
                print(f"ERROR: Could not read {filename}")
                print(f"Error details: {e}")
                return None
        
        # Load additional optional files
        for filename in OPTIONAL_FILES:
            filepath = os.path.join(INPUT_DIR, filename)
            try:
                if os.path.exists(filepath):
                    df = pd.read_excel(filepath)
                    # Normalize column names
                    df.columns = df.columns.str.strip()
                    key = filename.replace('_data.xlsx', '').replace('.xlsx', '').replace('_', '')
                    data_frames[key] = df
                    print(f"SUCCESS: Loaded {filename} ({len(df)} records)")
                else:
                    print(f"INFO: Optional file not found: {filename} (will continue without it)")
            except Exception as e:
                print(f"INFO: Could not load optional file {filename}: {e}")
        
        return data_frames
    
    @staticmethod
    def get_semester_courses(dfs, semester_id):
        """Get ALL courses for a specific semester, handling CSE sections specially."""
        if 'course' not in dfs:
            print("ERROR: 'course' key not found in data frames")
            return pd.DataFrame()
        
        course_df = dfs['course']
        print(f"Total courses loaded: {len(course_df)}")
        
        if course_df.empty:
            print("WARNING: Course dataframe is empty")
            return course_df
        
        # Print available columns for debugging
        print("Available columns in course_data:", course_df.columns.tolist())
        
        # Check if Semester column exists
        if 'Semester' not in course_df.columns:
            print("WARNING: 'Semester' column not found. Using all courses.")
            return course_df
        
        # Convert semester to numeric and filter
        course_df = course_df.copy()
        course_df['Semester'] = pd.to_numeric(course_df['Semester'], errors='coerce')
        
        # Check for NaN values after conversion
        nan_count = course_df['Semester'].isna().sum()
        if nan_count > 0:
            print(f"WARNING: {nan_count} rows have invalid Semester values")
        
        course_df = course_df.dropna(subset=['Semester'])
        course_df['Semester'] = course_df['Semester'].astype(int)
        
        sem_courses = course_df[course_df['Semester'] == semester_id].copy()
        print(f"Courses for semester {semester_id}: {len(sem_courses)}")

        # Normalize department labels for consistent matching/scheduling
        if 'Department' in sem_courses.columns:
            normalized_departments = []
            unknown_departments = set()
            for raw_label in sem_courses['Department']:
                normalized = ExcelLoader._normalize_department_label(raw_label)
                if normalized == "":
                    unknown_departments.add(str(raw_label))
                normalized_departments.append(normalized)

            sem_courses['Department'] = normalized_departments

            if unknown_departments:
                print("WARNING: The following Department labels could not be normalized and will be used as-is:")
                for label in sorted(unknown_departments):
                    print(f"  - '{label}'")
        
        if sem_courses.empty:
            print(f"WARNING: No courses found for semester {semester_id}")
            print("Available semesters:", course_df['Semester'].unique())
        
        # --- REPLACED CSE handling: split base "CSE" between sections, keep explicit A/B unchanged ---
        # Separate non-CSE and CSE-related entries
        sem_courses['Department'] = sem_courses['Department'].astype(str)
        non_cse = sem_courses[~sem_courses['Department'].str.match(r'^CSE', na=False)].copy()
        cse_entries = sem_courses[sem_courses['Department'].str.match(r'^CSE', na=False)].copy()

        # Explicit sectioned entries (CSE-A, CSE-B) remain as-is
        explicit_mask = cse_entries['Department'].isin(['CSE-A', 'CSE-B'])
        explicit_cse = cse_entries[explicit_mask].copy()

        # Base CSE entries (Department exactly 'CSE') will be split between A and B
        base_mask = cse_entries['Department'].str.match(r'^CSE$', na=False)
        base_cse = cse_entries[base_mask].copy().reset_index(drop=True)

        cse_a_assigned = pd.DataFrame()
        cse_b_assigned = pd.DataFrame()
        if not base_cse.empty:
            cse_a_assigned = base_cse.copy()
            cse_b_assigned = base_cse.copy()
            cse_a_assigned['Department'] = 'CSE-A'
            cse_b_assigned['Department'] = 'CSE-B'
            print(f"Duplicated {len(base_cse)} base CSE courses to both CSE-A and CSE-B")

        # Combine all back without duplicating existing explicit A/B entries
        sem_courses = pd.concat([non_cse, explicit_cse, cse_a_assigned, cse_b_assigned], ignore_index=True)
        
        return sem_courses
    
    @staticmethod
    def parse_ltpsc(courses_df):
        """Parse LTPSC format into separate columns for weekly frequency.
        L-T-P-S-C format where:
        - L: Lecture HOURS per week (converted to lecture sessions: each session is 1.5 hours)
        - T: Tutorial HOURS per week (converted to tutorial sessions: each session is 1 hour)
        - P: Lab HOURS per week (converted to lab sessions: each session is 2 hours)
        - S: Session (ignored)
        - C: Credits (ignored)
        Now keeps ALL courses - assigns defaults if LTPSC is missing/invalid."""
        if courses_df.empty:
            print("WARNING: Empty courses dataframe in parse_ltpsc")
            return courses_df
            
        df = courses_df.copy()
        
        # Prepare output columns
        df['Lectures_Per_Week'] = pd.NA
        df['Tutorials_Per_Week'] = pd.NA
        df['Labs_Per_Week'] = pd.NA
        
        if 'LTPSC' in df.columns:
            print("Parsing LTPSC column (assigning defaults for missing/invalid LTPSC)...")
            for idx, row in df.iterrows():
                ltpsc_val = row.get('LTPSC')
                course_code = str(row.get('Course Code', '')).strip()
                course_name = str(row.get('Course Name', '')).strip()
                credits = row.get('Credits', 3)  # Default to 3 credits if missing
                is_minor = (MINOR_SUBJECT.lower() in course_name.lower()) or (MINOR_SUBJECT.lower() in course_code.lower())
                
                # Try to convert credits to numeric
                try:
                    credits = float(credits) if not pd.isna(credits) else 3
                except:
                    credits = 3
                
                if pd.isna(ltpsc_val) or str(ltpsc_val).strip() == '':
                    if is_minor:
                        # Keep minor, set zeros (scheduled separately)
                        df.at[idx, 'Lectures_Per_Week'] = 0
                        df.at[idx, 'Tutorials_Per_Week'] = 0
                        df.at[idx, 'Labs_Per_Week'] = 0
                    else:
                        # CHANGED: Assign default LTPSC based on credits instead of excluding
                        print(f"INFO: Missing LTPSC for {course_code or course_name}; assigning defaults based on {credits} credits")
                        if credits >= 4:
                            # High credit course - assume 3 lectures + 1 lab
                            df.at[idx, 'Lectures_Per_Week'] = 3
                            df.at[idx, 'Tutorials_Per_Week'] = 0
                            df.at[idx, 'Labs_Per_Week'] = 1
                        elif credits >= 3:
                            # Standard course - 3 lectures
                            df.at[idx, 'Lectures_Per_Week'] = 3
                            df.at[idx, 'Tutorials_Per_Week'] = 0
                            df.at[idx, 'Labs_Per_Week'] = 0
                        elif credits >= 2:
                            # 2 credit course - 2 lectures
                            df.at[idx, 'Lectures_Per_Week'] = 2
                            df.at[idx, 'Tutorials_Per_Week'] = 0
                            df.at[idx, 'Labs_Per_Week'] = 0
                        else:
                            # Low credit - 1 lecture
                            df.at[idx, 'Lectures_Per_Week'] = 1
                            df.at[idx, 'Tutorials_Per_Week'] = 0
                            df.at[idx, 'Labs_Per_Week'] = 0
                    continue
                
                parts = str(ltpsc_val).split('-')
                if len(parts) < 3:
                    if is_minor:
                        df.at[idx, 'Lectures_Per_Week'] = 0
                        df.at[idx, 'Tutorials_Per_Week'] = 0
                        df.at[idx, 'Labs_Per_Week'] = 0
                    else:
                        # CHANGED: Assign defaults instead of excluding
                        print(f"INFO: Malformed LTPSC '{ltpsc_val}' for {course_code or course_name}; using defaults")
                        df.at[idx, 'Lectures_Per_Week'] = 3 if credits >= 3 else 2
                        df.at[idx, 'Tutorials_Per_Week'] = 0
                        df.at[idx, 'Labs_Per_Week'] = 0
                    continue
                
                # parse numeric L-T-P (ignore S,C columns)
                # L, T, P represent HOURS per week, not sessions per week
                # Convert hours to sessions by dividing by session duration and rounding up
                try:
                    l = float(parts[0])  # Lecture hours per week
                    t = float(parts[1])  # Tutorial hours per week
                    p = float(parts[2])  # Lab hours per week
                    
                    # Convert lecture hours to lecture sessions (each lecture is 1.5 hours)
                    # If l=3 hours, that means 3/1.5 = 2 lecture sessions per week
                    # If l=2 hours, that means 2/1.5 = 1.33, round up to 2 sessions
                    lecture_hours_per_session = 1.5  # Each lecture is 1.5 hours
                    lecture_sessions = math.ceil(l / lecture_hours_per_session) if l > 0 else 0
                    df.at[idx, 'Lectures_Per_Week'] = int(lecture_sessions)
                    
                    # Convert tutorial hours to tutorial sessions (each tutorial is 1 hour)
                    # If t=1 hour, that means 1/1 = 1 tutorial session per week
                    # If t=2 hours, that means 2/1 = 2 tutorial sessions per week
                    tutorial_hours_per_session = 1.0  # Each tutorial is 1 hour
                    tutorial_sessions = math.ceil(t / tutorial_hours_per_session) if t > 0 else 0
                    df.at[idx, 'Tutorials_Per_Week'] = int(tutorial_sessions)
                    
                    # Convert lab hours to lab sessions (each lab session is 2 hours)
                    # If p=2 hours, that means 2/2 = 1 lab session per week
                    # If p=4 hours, that means 4/2 = 2 lab sessions per week
                    lab_hours_per_session = 2.0  # Each lab session is 2 hours
                    lab_sessions = math.ceil(p / lab_hours_per_session) if p > 0 else 0
                    df.at[idx, 'Labs_Per_Week'] = int(lab_sessions)
                except Exception:
                    if is_minor:
                        df.at[idx, 'Lectures_Per_Week'] = 0
                        df.at[idx, 'Tutorials_Per_Week'] = 0
                        df.at[idx, 'Labs_Per_Week'] = 0
                    else:
                        # CHANGED: Assign defaults instead of excluding
                        print(f"INFO: Non-numeric LTPSC '{ltpsc_val}' for {course_code or course_name}; using defaults")
                        df.at[idx, 'Lectures_Per_Week'] = 3 if credits >= 3 else 2
                        df.at[idx, 'Tutorials_Per_Week'] = 0
                        df.at[idx, 'Labs_Per_Week'] = 0
            
            print(f"LTPSC parsing completed: {len(df)} courses retained")
            return df.reset_index(drop=True)
        else:
            # No LTPSC column - assign defaults for all courses
            print("WARNING: LTPSC column not found. Assigning default values based on credits.")
            for idx, row in df.iterrows():
                credits = row.get('Credits', 3)
                course_name = str(row.get('Course Name', '')).strip()
                is_minor = MINOR_SUBJECT.lower() in course_name.lower()
                
                try:
                    credits = float(credits) if not pd.isna(credits) else 3
                except:
                    credits = 3
                
                if is_minor:
                    df.at[idx, 'Lectures_Per_Week'] = 0
                    df.at[idx, 'Tutorials_Per_Week'] = 0
                    df.at[idx, 'Labs_Per_Week'] = 0
                elif credits >= 3:
                    df.at[idx, 'Lectures_Per_Week'] = 3
                    df.at[idx, 'Tutorials_Per_Week'] = 0
                    df.at[idx, 'Labs_Per_Week'] = 0
                else:
                    df.at[idx, 'Lectures_Per_Week'] = 2
                    df.at[idx, 'Tutorials_Per_Week'] = 0
                    df.at[idx, 'Labs_Per_Week'] = 0
            
            return df.reset_index(drop=True)

    @staticmethod
    def divide_courses_by_session(courses_df, department, all_sem_courses=None):
        """Divide courses between Pre-Mid and Post-Mid sessions, handling electives and shared-course rules.
        RULES:
        - Electives: ALWAYS run FULL semester (included in BOTH Pre-Mid and Post-Mid)
          * This applies regardless of credit value (even if 2 credits or less)
          * Electives are excluded from credit-based division logic
        - Credits > 2: Run FULL semester (included in BOTH Pre-Mid and Post-Mid)
        - Credits <= 2: Run HALF semester (split equally - EITHER Pre-Mid OR Post-Mid)
        - CSE-A and CSE-B: Always have same courses in both sessions
        """
        if courses_df.empty:
            print(f"WARNING: Empty courses dataframe for {department}")
            return pd.DataFrame(), pd.DataFrame()

        print(f"\nDividing courses for {department}")

        # Remove Minor entries from the course list (scheduled separately)
        minor_mask = (courses_df.get('Course Name', '').astype(str).str.contains(MINOR_SUBJECT, case=False, na=False)) | \
                     (courses_df.get('Course Code', '').astype(str).str.contains(MINOR_SUBJECT, case=False, na=False))
        if minor_mask.any():
            print(f"  Excluding {minor_mask.sum()} Minor entries from regular session division")
        courses_df = courses_df[~minor_mask].copy()

        # --- Robust elective detection: accept multiple column name variants ---
        elective_col = None
        for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
            if colname in courses_df.columns:
                elective_col = colname
                break

        if elective_col is None:
            # no elective column -> start with none elective
            elective_mask_local = pd.Series([False]*len(courses_df), index=courses_df.index)
        else:
            elective_mask_local = courses_df.get(elective_col, '').astype(str).str.upper() == 'YES'

        # Override elective detection by patterns:
        # - Anything with 'ELEC' in Course Code/Name is elective (force include)
        # - Anything with 'HSS' is NOT elective (force exclude from elective set)
        pattern_elec = pd.Series([False]*len(courses_df), index=courses_df.index)
        pattern_hss = pd.Series([False]*len(courses_df), index=courses_df.index)
        if 'Course Code' in courses_df.columns:
            pattern_elec = pattern_elec | courses_df['Course Code'].astype(str).str.contains(r'ELEC', case=False, na=False)
            pattern_hss = pattern_hss | courses_df['Course Code'].astype(str).str.contains(r'HSS', case=False, na=False)
        if 'Course Name' in courses_df.columns:
            pattern_elec = pattern_elec | courses_df['Course Name'].astype(str).str.contains(r'ELEC', case=False, na=False)
            pattern_hss = pattern_hss | courses_df['Course Name'].astype(str).str.contains(r'HSS', case=False, na=False)
        # Apply overrides
        elective_mask_local = (elective_mask_local | pattern_elec) & (~pattern_hss)

        elective_courses = courses_df[elective_mask_local].copy()
        regular_courses = courses_df[~elective_mask_local].copy()

        # --- Treat HSS courses as full-semester (present in BOTH sessions), similar to electives ---
        # Detect HSS using robust checks on Course Code and Course Name
        hss_mask = pd.Series([False] * len(regular_courses), index=regular_courses.index)
        if 'Course Code' in regular_courses.columns:
            hss_mask = hss_mask | regular_courses['Course Code'].astype(str).str.contains(r'\bHSS\b', case=False, na=False)
            # Also consider generic patterns like "1-HSS", "HSS-101"
            hss_mask = hss_mask | regular_courses['Course Code'].astype(str).str.contains(r'HSS', case=False, na=False)
        if 'Course Name' in regular_courses.columns:
            hss_mask = hss_mask | regular_courses['Course Name'].astype(str).str.contains(r'\bHSS\b', case=False, na=False)
            hss_mask = hss_mask | regular_courses['Course Name'].astype(str).str.contains(r'HSS', case=False, na=False)

        hss_courses = regular_courses[hss_mask].copy()
        if not hss_courses.empty:
            print(f"  Detected {len(hss_courses)} HSS course(s) to run full semester")
            # Remove HSS from regular course pool so they are not split
            regular_courses = regular_courses[~hss_mask].copy()
        else:
            # If this department has no HSS but the semester contains HSS anywhere,
            # auto-include one HSS for this department so every department has HSS.
            # Only do this when we have the parsed semester-wide view.
            if all_sem_courses is not None and not all_sem_courses.empty:
                # Detect HSS semester-wide
                sem_hss_mask = pd.Series([False] * len(all_sem_courses), index=all_sem_courses.index)
                if 'Course Code' in all_sem_courses.columns:
                    sem_hss_mask = sem_hss_mask | all_sem_courses['Course Code'].astype(str).str.contains(r'\bHSS\b', case=False, na=False)
                    sem_hss_mask = sem_hss_mask | all_sem_courses['Course Code'].astype(str).str.contains(r'HSS', case=False, na=False)
                if 'Course Name' in all_sem_courses.columns:
                    sem_hss_mask = sem_hss_mask | all_sem_courses['Course Name'].astype(str).str.contains(r'\bHSS\b', case=False, na=False)
                    sem_hss_mask = sem_hss_mask | all_sem_courses['Course Name'].astype(str).str.contains(r'HSS', case=False, na=False)
                sem_hss = all_sem_courses[sem_hss_mask].copy()
                if not sem_hss.empty:
                    # Take the first HSS entry as the canonical HSS for this department
                    exemplar = sem_hss.iloc[[0]].copy()
                    exemplar['Department'] = department
                    # Ensure elective flags for HSS are not set
                    for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
                        if colname in exemplar.columns:
                            exemplar[colname] = 'NO'
                    hss_courses = exemplar.reset_index(drop=True)
                    print(f"  Auto-including HSS for {department} (semester-wide HSS detected)")

        # Ensure elective column is explicitly set to 'YES' for all elective courses
        # This ensures the flag is preserved even after merging and concatenation
        if not elective_courses.empty and elective_col is not None:
            elective_courses[elective_col] = 'YES'
        
        print(f"  Found {len(elective_courses)} elective courses (department-local)")
        if not elective_courses.empty and 'Course Code' in elective_courses.columns:
            elective_codes = elective_courses['Course Code'].dropna().astype(str).tolist()
            print(f"    Elective course codes: {', '.join(elective_codes)}")

        # Get semester-wide electives for ALL departments (CSE, DSAI, ECE) to ensure same time slots
        if all_sem_courses is not None and not all_sem_courses.empty:
            # Find elective column in all_sem_courses
            elective_col_all = None
            for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
                if colname in all_sem_courses.columns:
                    elective_col_all = colname
                    break

            if elective_col_all is not None:
                # Get ALL electives from ALL departments in this semester
                mask_elec_all = all_sem_courses.get(elective_col_all, '').astype(str).str.upper() == 'YES'
                elective_from_all = all_sem_courses[mask_elec_all].copy()
                
                if not elective_from_all.empty:
                    # For CSE-A and CSE-B, include base 'CSE' electives
                    if department in ['CSE-A', 'CSE-B']:
                        mask_dept = elective_from_all['Department'].astype(str).isin(['CSE', 'CSE-A', 'CSE-B'])
                        elective_from_all = elective_from_all[mask_dept].copy()
                        print(f"  Using semester-wide CSE elective list for {department}")
                    # For DSAI and ECE, get their department-specific electives
                    elif department in ['DSAI', 'ECE']:
                        mask_dept = elective_from_all['Department'].astype(str) == department
                        elective_from_all = elective_from_all[mask_dept].copy()
                        print(f"  Using semester-wide {department} elective list")
                    else:
                        # For other departments, get their department-specific electives
                        mask_dept = elective_from_all['Department'].astype(str) == department
                        elective_from_all = elective_from_all[mask_dept].copy()
                    
                    if not elective_from_all.empty:
                        # Remove duplicates and combine with department-local electives
                        elective_from_all = elective_from_all.drop_duplicates(subset=['Course Code']).reset_index(drop=True)
                        # Ensure elective column is explicitly set to 'YES' for semester-wide electives
                        # Use the same column name as the local elective_col if it exists, otherwise use the one from all_sem_courses
                        target_elective_col = elective_col if elective_col is not None else elective_col_all
                        if target_elective_col is not None:
                            elective_from_all[target_elective_col] = 'YES'
                            # Also set it in the original column if different
                            if elective_col_all is not None and target_elective_col != elective_col_all:
                                elective_from_all[elective_col_all] = 'YES'
                        # Union with department-local electives
                        # Before concatenating, ensure both have the same elective column
                        if not elective_courses.empty and elective_col is not None and elective_col not in elective_from_all.columns:
                            elective_from_all[elective_col] = 'YES'
                        elective_courses = pd.concat([elective_courses, elective_from_all], ignore_index=True).drop_duplicates(subset=['Course Code']).reset_index(drop=True)
                        # Ensure all elective courses have the elective column set to 'YES'
                        if target_elective_col is not None and target_elective_col in elective_courses.columns:
                            elective_courses[target_elective_col] = 'YES'
                        # Also remove these electives from regular_courses if present
                        regular_courses = regular_courses[~regular_courses['Course Code'].isin(elective_courses['Course Code'])].copy()
                        print(f"  Total electives after semester-wide merge: {len(elective_courses)} electives")
                        if 'Course Code' in elective_courses.columns:
                            all_elective_codes = elective_courses['Course Code'].dropna().astype(str).tolist()
                            print(f"    All elective course codes: {', '.join(all_elective_codes)}")

        # IMPORTANT: Electives are ALWAYS full-semester (in both Pre-Mid and Post-Mid)
        # regardless of their credit value. They are excluded from credit-based division.
        
        # Ensure Credits column handling (only for regular/non-elective courses)
        if 'Credits' not in regular_courses.columns:
            print("  WARNING: 'Credits' column not found. Treating all non-elective courses as half-sem (<=2).")
            regular_courses['Credits'] = 2

        # Convert Credits to numeric to compare (only for regular/non-elective courses)
        regular_courses = regular_courses.copy()
        regular_courses['Credits'] = pd.to_numeric(regular_courses['Credits'], errors='coerce').fillna(2)

        # Full-sem courses: credits > 2 (will be in BOTH sessions)
        # NOTE: Electives are NOT included here - they are handled separately above
        full_sem_courses = regular_courses[regular_courses['Credits'] > 2].copy()
        # Half-sem courses: credits <= 2 (will be split - in EITHER Pre OR Post)
        # NOTE: Electives are NOT included here - they are ALWAYS full-semester
        half_sem_courses = regular_courses[regular_courses['Credits'] <= 2].copy()
        
        # Double-check: Remove any electives that might have slipped into half_sem_courses
        # (This should not happen, but adding as a safety check)
        if not elective_courses.empty and 'Course Code' in elective_courses.columns and 'Course Code' in half_sem_courses.columns:
            elective_codes_set = set(elective_courses['Course Code'].dropna().astype(str))
            half_sem_courses = half_sem_courses[~half_sem_courses['Course Code'].astype(str).isin(elective_codes_set)].copy()
        if not elective_courses.empty and 'Course Code' in elective_courses.columns and 'Course Code' in full_sem_courses.columns:
            elective_codes_set = set(elective_courses['Course Code'].dropna().astype(str))
            full_sem_courses = full_sem_courses[~full_sem_courses['Course Code'].astype(str).isin(elective_codes_set)].copy()

        print(f"  Course breakdown for {department}:")
        print(f"     - Electives (full sem - in both sessions): {len(elective_courses)} courses")
        print(f"     - HSS (full sem - in both sessions): {len(hss_courses)} courses")
        print(f"     - Credits > 2 (full sem - in both sessions): {len(full_sem_courses)} courses")
        print(f"     - Credits <= 2 (half sem - split between sessions): {len(half_sem_courses)} courses")

        # For CSE-A and CSE-B, ensure electives and full-sem courses are present in both sessions
        if department in ['CSE-A', 'CSE-B']:
            # If all_sem_courses provided, ensure half-sem list uses both CSE-A/B entries (shared view)
            if all_sem_courses is not None and not all_sem_courses.empty:
                all_half_sem = all_sem_courses.copy()
                
                # FIRST: Exclude ALL electives before doing any credit-based filtering
                # Electives are ALWAYS full-semester, regardless of credits
                if not elective_courses.empty and 'Course Code' in elective_courses.columns and 'Course Code' in all_half_sem.columns:
                    elective_codes_set = set(elective_courses['Course Code'].dropna().astype(str))
                    all_half_sem = all_half_sem[~all_half_sem['Course Code'].astype(str).isin(elective_codes_set)].copy()
                    print(f"  Excluded {len(elective_codes_set)} electives from half-semester processing for CSE")
                
                # Also exclude electives by checking the elective column directly
                elective_col_all = None
                for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
                    if colname in all_half_sem.columns:
                        elective_col_all = colname
                        break
                if elective_col_all is not None:
                    elective_mask = all_half_sem[elective_col_all].astype(str).str.upper() == 'YES'
                    if elective_mask.any():
                        all_half_sem = all_half_sem[~elective_mask].copy()
                        print(f"  Excluded additional {elective_mask.sum()} electives found via column check")
                
                # NOW filter by credits (only non-elective courses)
                if 'Credits' not in all_half_sem.columns:
                    all_half_sem['Credits'] = 2
                all_half_sem['Credits'] = pd.to_numeric(all_half_sem['Credits'], errors='coerce').fillna(2)
                all_half_sem = all_half_sem[all_half_sem['Credits'] <= 2]
                all_half_sem = all_half_sem[all_half_sem['Department'].astype(str).isin(['CSE', 'CSE-A', 'CSE-B'])]
                all_half_sem = all_half_sem.drop_duplicates(subset=['Course Code']).reset_index(drop=True)
                
                # Final safety check: remove any remaining electives (should be none at this point)
                if not elective_courses.empty and 'Course Code' in elective_courses.columns:
                    elective_codes_set = set(elective_courses['Course Code'].dropna().astype(str))
                    all_half_sem = all_half_sem[~all_half_sem['Course Code'].astype(str).isin(elective_codes_set)].copy()
                
                half_sem_courses = all_half_sem

            # Sort courses by Course Code for consistent, deterministic splitting
            if 'Course Code' in half_sem_courses.columns:
                half_sem_courses = half_sem_courses.sort_values('Course Code').reset_index(drop=True)
            
            # Split half-sem courses equally: for n courses, split into n//2 and n - n//2
            n = len(half_sem_courses)
            if n == 0:
                pre_half = pd.DataFrame()
                post_half = pd.DataFrame()
            else:
                # For even numbers: split equally (e.g., 6 -> 3 and 3)
                # For odd numbers: give one extra to Pre-Mid (e.g., 5 -> 3 and 2)
                split_point = (n + 1) // 2  # This ensures: 6->3, 5->3, 4->2, 3->2
                pre_half = half_sem_courses.iloc[:split_point].copy()
                post_half = half_sem_courses.iloc[split_point:].copy()
            
            print(f"  Split {n} half-sem courses: {len(pre_half)} Pre-Mid, {len(post_half)} Post-Mid")

            # IMPORTANT: Electives are ALWAYS included in BOTH Pre-Mid and Post-Mid sessions
            # regardless of credit value. They are NOT split like half-semester courses.
            # Always include electives and full-sem courses in both sessions for CSE
            if not elective_courses.empty:
                print(f"  Adding {len(elective_courses)} electives to BOTH Pre-Mid and Post-Mid for {department}")
            # Add HSS to both sessions as well
            if not hss_courses.empty:
                print(f"  Adding {len(hss_courses)} HSS to BOTH Pre-Mid and Post-Mid for {department}")
            pre_mid_courses = pd.concat([elective_courses, hss_courses, full_sem_courses, pre_half], ignore_index=True)
            post_mid_courses = pd.concat([elective_courses, hss_courses, full_sem_courses, post_half], ignore_index=True)
        else:
            # For other departments (DSAI, ECE), deterministic split
            # Sort courses by Course Code for consistent, deterministic splitting
            if 'Course Code' in half_sem_courses.columns:
                half_sem_courses = half_sem_courses.sort_values('Course Code').reset_index(drop=True)
            
            n = len(half_sem_courses)
            if n == 0:
                pre_half = pd.DataFrame()
                post_half = pd.DataFrame()
            else:
                # For even numbers: split equally (e.g., 6 -> 3 and 3)
                # For odd numbers: give one extra to Pre-Mid (e.g., 5 -> 3 and 2)
                split_point = (n + 1) // 2  # This ensures: 6->3, 5->3, 4->2, 3->2
                pre_half = half_sem_courses.iloc[:split_point].copy()
                post_half = half_sem_courses.iloc[split_point:].copy()
            
            print(f"  Split {n} half-sem courses: {len(pre_half)} Pre-Mid, {len(post_half)} Post-Mid")
            
            # IMPORTANT: Electives are ALWAYS included in BOTH Pre-Mid and Post-Mid sessions
            # regardless of credit value. They are NOT split like half-semester courses.
            if not elective_courses.empty:
                print(f"  Adding {len(elective_courses)} electives to BOTH Pre-Mid and Post-Mid for {department}")
            if not hss_courses.empty:
                print(f"  Adding {len(hss_courses)} HSS to BOTH Pre-Mid and Post-Mid for {department}")
            pre_mid_courses = pd.concat([elective_courses, hss_courses, full_sem_courses, pre_half], ignore_index=True)
            post_mid_courses = pd.concat([elective_courses, hss_courses, full_sem_courses, post_half], ignore_index=True)

        # Deduplicate by Course Code if present
        def dedup_df(df):
            if df.empty:
                return df
            if 'Course Code' in df.columns:
                return df.drop_duplicates(subset=['Course Code']).reset_index(drop=True)
            return df.drop_duplicates().reset_index(drop=True)

        pre_mid_courses = dedup_df(pre_mid_courses)
        post_mid_courses = dedup_df(post_mid_courses)
        
        # Final verification: Ensure electives are present in BOTH sessions and have their flag set
        if not elective_courses.empty and 'Course Code' in elective_courses.columns:
            elective_col_final = elective_col if elective_col is not None else None
            if elective_col_final is None:
                # Try to find elective column in the merged courses
                for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
                    if colname in pre_mid_courses.columns:
                        elective_col_final = colname
                        break
            
            if elective_col_final is not None:
                # Ensure elective flag is set for all elective courses in both sessions
                elective_codes_set = set(elective_courses['Course Code'].dropna().astype(str))
                
                # Check and fix Pre-Mid courses
                if 'Course Code' in pre_mid_courses.columns:
                    pre_elective_mask = pre_mid_courses['Course Code'].astype(str).isin(elective_codes_set)
                    if pre_elective_mask.any():
                        pre_mid_courses.loc[pre_elective_mask, elective_col_final] = 'YES'
                        print(f"  Verified {pre_elective_mask.sum()} electives in Pre-Mid session for {department}")
                
                # Check and fix Post-Mid courses
                if 'Course Code' in post_mid_courses.columns:
                    post_elective_mask = post_mid_courses['Course Code'].astype(str).isin(elective_codes_set)
                    if post_elective_mask.any():
                        post_mid_courses.loc[post_elective_mask, elective_col_final] = 'YES'
                        print(f"  Verified {post_elective_mask.sum()} electives in Post-Mid session for {department}")
        
        # Store original course codes for validation (before 2-credit sharing modifies them)
        original_input_codes = set()
        if 'Course Code' in courses_df.columns:
            original_input_codes = set(courses_df['Course Code'].dropna().astype(str))

        # Apply shared 2-credit course coordination across CSE/DSAI/ECE
        pre_mid_courses, post_mid_courses = ExcelLoader._apply_two_credit_sharing(
            pre_mid_courses,
            post_mid_courses,
            department,
            all_sem_courses,
        )

        # Validate that all input courses are accounted for
        if original_input_codes:
            pre_codes = set(pre_mid_courses['Course Code'].dropna().astype(str)) if 'Course Code' in pre_mid_courses.columns and not pre_mid_courses.empty else set()
            post_codes = set(post_mid_courses['Course Code'].dropna().astype(str)) if 'Course Code' in post_mid_courses.columns and not post_mid_courses.empty else set()
            all_assigned = pre_codes | post_codes
            missing_courses = original_input_codes - all_assigned
            
            if missing_courses:
                print(f"  ERROR: {len(missing_courses)} courses not assigned to any session: {', '.join(sorted(missing_courses))}")
        
        print(f"  Session allocation for {department}:")
        print(f"     - Pre-Mid: {len(pre_mid_courses)} courses")
        print(f"     - Post-Mid: {len(post_mid_courses)} courses")
        
        # Calculate overlap (courses in both sessions)
        if 'Course Code' in pre_mid_courses.columns and 'Course Code' in post_mid_courses.columns:
            pre_codes = set(pre_mid_courses['Course Code'].dropna().astype(str))
            post_codes = set(post_mid_courses['Course Code'].dropna().astype(str))
            overlap = pre_codes & post_codes
            print(f"     - Overlap (full semester courses): {len(overlap)} courses")

        return pre_mid_courses, post_mid_courses

    @staticmethod
    def _apply_two_credit_sharing(pre_mid_df, post_mid_df, department, all_sem_courses):
        """Ensure shared 2-credit courses alternate sessions across CSE/DSAI/ECE.
        Rule: If a 2-credit shared course is in CSE Pre-Mid, it must be in DSAI/ECE Post-Mid (and vice versa)."""
        target_departments = {"CSE", "CSE-A", "CSE-B", "DSAI", "ECE"}
        if department not in target_departments:
            return pre_mid_df, post_mid_df

        if all_sem_courses is None or all_sem_courses.empty:
            return pre_mid_df, post_mid_df

        if 'Course Code' not in pre_mid_df.columns and 'Course Code' not in post_mid_df.columns:
            return pre_mid_df, post_mid_df

        # Determine semester id from available data
        semester_id = None
        if 'Semester' in all_sem_courses.columns:
            sem_values = all_sem_courses['Semester'].dropna().unique()
            if len(sem_values) == 1:
                semester_id = int(sem_values[0])

        if semester_id is None:
            return pre_mid_df, post_mid_df

        def _get_course_rows(df, code):
            if 'Course Code' not in df.columns:
                return df.iloc[0:0]
            return df[df['Course Code'] == code].copy()

        def _remove_course(df, code):
            if 'Course Code' not in df.columns:
                return df
            return df[df['Course Code'] != code].reset_index(drop=True)

        shared_map = ExcelLoader._two_credit_course_session_map.setdefault(semester_id, {})

        # Detect electives in the current session DataFrames to avoid moving them.
        # Electives must remain in BOTH sessions, regardless of credits.
        def _get_elective_codes(df):
            if df is None or df.empty:
                return set()
            elective_col = None
            for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
                if colname in df.columns:
                    elective_col = colname
                    break
            if elective_col is None or 'Course Code' not in df.columns:
                return set()
            mask = df[elective_col].astype(str).str.upper() == 'YES'
            return set(df.loc[mask, 'Course Code'].dropna().astype(str))

        pre_elective_codes = _get_elective_codes(pre_mid_df)
        post_elective_codes = _get_elective_codes(post_mid_df)
        elective_codes_union = pre_elective_codes | post_elective_codes

        # Helper to check if a course is 2-credit and shared across CSE, DSAI, ECE
        def _is_shared_two_credit(course_code):
            mask = all_sem_courses['Course Code'] == course_code
            subset = all_sem_courses[mask]
            if subset.empty:
                return False
            try:
                credit_vals = pd.to_numeric(subset['Credits'], errors='coerce').fillna(0)
            except Exception:
                credit_vals = pd.Series([0] * len(subset))
            if not (credit_vals <= 2).any():
                return False
            # Check if course appears in at least 2 of the target departments (CSE, DSAI, ECE)
            depts = set(subset['Department'].astype(str))
            # Normalize CSE variants
            normalized_depts = set()
            for d in depts:
                if d in {"CSE", "CSE-A", "CSE-B"}:
                    normalized_depts.add("CSE")
                elif d in target_departments:
                    normalized_depts.add(d)
            return len(normalized_depts) >= 2

        # Step 1: Record CSE session assignments (from CSE-A or CSE-B, whichever processes first)
        if department in {"CSE", "CSE-A", "CSE-B"}:
            # Record which session each 2-credit shared course is assigned to in CSE
            if 'Course Code' in pre_mid_df.columns:
                for code in pre_mid_df['Course Code'].dropna().unique():
                    if _is_shared_two_credit(code):
                        # Only record if not already set (CSE-A and CSE-B should have same assignments)
                        if code not in shared_map:
                            shared_map[code] = 'Pre'
            
            if 'Course Code' in post_mid_df.columns:
                for code in post_mid_df['Course Code'].dropna().unique():
                    if _is_shared_two_credit(code):
                        # Only record if not already set
                        if code not in shared_map:
                            shared_map[code] = 'Post'

        # Step 2: Apply opposite-session rule for DSAI/ECE
        if department in {"DSAI", "ECE"}:
            if 'Course Code' not in pre_mid_df.columns:
                return pre_mid_df, post_mid_df
                
            # Get all 2-credit shared courses that should be coordinated
            all_codes = set(pre_mid_df['Course Code'].dropna()) | set(post_mid_df['Course Code'].dropna())
            # Never move electives - they must remain in BOTH sessions
            all_codes = {code for code in all_codes if str(code) not in elective_codes_union}
            
            for code in all_codes:
                if not _is_shared_two_credit(code):
                    continue
                    
                cse_session = shared_map.get(code)
                if cse_session is None:
                    # CSE hasn't processed yet, skip for now
                    continue
                
                # Determine desired session for DSAI/ECE (opposite of CSE)
                desired_session = 'Post' if cse_session == 'Pre' else 'Pre'
                
                # Check current location
                in_pre = code in pre_mid_df['Course Code'].values if 'Course Code' in pre_mid_df.columns else False
                in_post = code in post_mid_df['Course Code'].values if 'Course Code' in post_mid_df.columns else False
                
                # Move course to correct session if needed
                if desired_session == 'Pre':
                    if in_post:
                        # Move from Post-Mid to Pre-Mid
                        row = _get_course_rows(post_mid_df, code)
                        if not row.empty:
                            post_mid_df = _remove_course(post_mid_df, code)
                            pre_mid_df = pd.concat([pre_mid_df, row], ignore_index=True)
                    # Ensure it's not in Post-Mid (remove if somehow still there)
                    if 'Course Code' in post_mid_df.columns and code in post_mid_df['Course Code'].values:
                        post_mid_df = _remove_course(post_mid_df, code)
                else:  # desired_session == 'Post'
                    if in_pre:
                        # Move from Pre-Mid to Post-Mid
                        row = _get_course_rows(pre_mid_df, code)
                        if not row.empty:
                            pre_mid_df = _remove_course(pre_mid_df, code)
                            post_mid_df = pd.concat([post_mid_df, row], ignore_index=True)
                    # Ensure it's not in Pre-Mid (remove if somehow still there)
                    if 'Course Code' in pre_mid_df.columns and code in pre_mid_df['Course Code'].values:
                        pre_mid_df = _remove_course(pre_mid_df, code)
                
                # Deduplicate
                if 'Course Code' in pre_mid_df.columns:
                    pre_mid_df = pre_mid_df.drop_duplicates(subset=['Course Code']).reset_index(drop=True)
                if 'Course Code' in post_mid_df.columns:
                    post_mid_df = post_mid_df.drop_duplicates(subset=['Course Code']).reset_index(drop=True)

        return pre_mid_df.reset_index(drop=True), post_mid_df.reset_index(drop=True)