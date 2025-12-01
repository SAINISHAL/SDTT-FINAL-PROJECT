"""Core scheduling logic for generating timetables from Excel data."""
import pandas as pd
import random
from config import DAYS, TEACHING_SLOTS, LECTURE_DURATION, TUTORIAL_DURATION, LAB_DURATION, MINOR_DURATION
from config import PRE_MID, POST_MID, MINOR_SUBJECT, MINOR_CLASSES_PER_WEEK, DEPARTMENTS
from config import MINOR_SLOTS, LUNCH_SLOTS
from excel_loader import ExcelLoader

class ScheduleGenerator:
    """Generates weekly class schedules for semesters and departments from Excel data."""
    
    def __init__(self, data_frames):
        """Initialize ScheduleGenerator with data frames."""
        self.dfs = data_frames
        # Track global slots per semester to avoid clashes between departments in same semester
        self.semester_global_slots = {}
        # Track room occupancy per (semester, day, slot)
        self.room_occupancy = {}
        # Track detailed room bookings per (sem_key, day, slot) for conflict validation
        self.room_bookings = {}
        # Load classrooms (room_name, capacity) and classify by type
        self.classrooms = []
        self.lab_rooms = []
        self.software_lab_rooms = []
        self.hardware_lab_rooms = []
        self.nonlab_rooms = []
        self.c004_room = None  # Special room for combined classes only
        self.shared_rooms = set()

        try:
            cls_df = self.dfs.get('classroom')
            if cls_df is not None and not cls_df.empty:
                name_col = None
                cap_col = None
                type_col = None
                for col in cls_df.columns:
                    cl = str(col).lower()
                    if name_col is None and any(k in cl for k in ['room', 'class', 'hall', 'name']):
                        name_col = col
                    if cap_col is None and any(k in cl for k in ['cap', 'seats', 'capacity']):
                        cap_col = col
                    if type_col is None and any(k in cl for k in ['type', 'category', 'room type']):
                        type_col = col
                if name_col is None:
                    name_col = cls_df.columns[0]
                    
                for _, row in cls_df.iterrows():
                    room_name = str(row.get(name_col, '')).strip()
                    try:
                        capacity = int(float(row.get(cap_col, 0))) if cap_col is not None else 0
                    except Exception:
                        capacity = 0
                    room_type = str(row.get(type_col, '')).strip().lower() if type_col is not None else ''
                    
                    if room_name:
                        # Special handling for C004 - combined class room only
                        room_name_upper = room_name.upper()
                        if room_name_upper == 'C004':
                            self.c004_room = (room_name, capacity)
                            self.shared_rooms.add(room_name_upper)
                            continue
                            
                        self.classrooms.append((room_name, capacity))
                        
                        if 'lab' in room_type:
                            self.lab_rooms.append((room_name, capacity))
                            if 'software' in room_type or 'soft' in room_type:
                                self.software_lab_rooms.append((room_name, capacity))
                            if 'hardware' in room_type or 'hard' in room_type:
                                self.hardware_lab_rooms.append((room_name, capacity))
                        else:
                            # Normal classrooms (excluding C004)
                            self.nonlab_rooms.append((room_name, capacity))
                
                # Sort all room lists by capacity
                self.classrooms.sort(key=lambda x: (x[1], x[0]))
                self.lab_rooms.sort(key=lambda x: (x[1], x[0]))
                self.software_lab_rooms.sort(key=lambda x: (x[1], x[0]))
                self.hardware_lab_rooms.sort(key=lambda x: (x[1], x[0]))
                self.nonlab_rooms.sort(key=lambda x: (x[1], x[0]))
                
                print(f"Room configuration loaded:")
                print(f"  - C004 (Combined): {self.c004_room if self.c004_room else 'Not found'}")
                print(f"  - Normal classrooms: {len(self.nonlab_rooms)}")
                print(f"  - Software labs: {len(self.software_lab_rooms)}")
                print(f"  - Hardware labs: {len(self.hardware_lab_rooms)}")
                
        except Exception as e:
            print(f"Error loading classroom data: {e}")
            self.classrooms = []
            self.lab_rooms = []
            self.software_lab_rooms = []
            self.hardware_lab_rooms = []
            self.nonlab_rooms = []
            self.c004_room = None
            
        # Store minor slots per semester
        self.semester_minor_slots = {}
        # Store elective slots per semester, keyed by (semester_id, elective_code)
        self.semester_elective_slots = {}
        # Store elective tutorial slots per semester, keyed by (semester_id, 'ALL_ELECTIVE_TUTORIALS') or (semester_id, elective_code, 'Tutorial')
        # This ensures all departments get elective tutorials at the same time slots for both Pre-Mid and Post-Mid
        self.semester_elective_tutorial_slots = {}
        # 240-seater combined-class capacity per semester: set of (day, slot)
        self.semester_combined_capacity = {}
        # Combined class assigned slots per course and component:
        # key=(semester_id, course_code, component['Lecture'|'Tutorial'|'Lab']) -> list[(day, slot)]
        self.semester_combined_course_slots = {}
        # Global combined course slots shared across semesters but per allowed pairing group:
        # key=('GLOBAL', group_key, course_code, component) -> list[(day, start_slot)]
        self.global_combined_course_slots = {}
        self.scheduled_slots = {}  # Track all scheduled slots by semester+department
        self.scheduled_courses = {}  # Track when each course is scheduled
        self.actual_allocations = {}  # Track actual allocated counts: key=(semester_id, dept, session, course_code), value={'lectures': X, 'tutorials': Y, 'labs': Z}
        self.assigned_rooms = {}  # Track room assignments: key=(semester_id, dept, session, course_code) -> room_name
        self.assigned_lab_rooms = {}  # Track lab room assignments: key=(semester_id, dept, session, course_code) -> room_name
        
    def _initialize_schedule(self):
        """Initialize an empty schedule with Days as rows and Time Slots as columns."""
        schedule = pd.DataFrame(index=DAYS, columns=TEACHING_SLOTS)
        
        # Initialize with 'Free'
        for day in DAYS:
            for slot in TEACHING_SLOTS:
                schedule.loc[day, slot] = 'Free'
        
        # Mark lunch break (now possibly multiple 30-min slots)
        for day in DAYS:
            for lunch_slot in LUNCH_SLOTS:
                if lunch_slot in schedule.columns:
                    schedule.loc[day, lunch_slot] = 'LUNCH BREAK'
        
        return schedule
    
    def _get_consecutive_slots(self, start_slot, duration):
        """Get consecutive time slots for a given duration."""
        try:
            start_index = TEACHING_SLOTS.index(start_slot)
            end_index = start_index + duration
            if end_index <= len(TEACHING_SLOTS):
                return TEACHING_SLOTS[start_index:end_index]
        except ValueError:
            pass
        return []
    
    def _ends_at_thirty(self, slots):
        """Check if a sequence of slots ends at :30."""
        if not slots:
            return False
        last_slot = slots[-1]
        # Extract end time from slot (format: 'HH:MM-HH:MM')
        try:
            end_time = last_slot.split('-')[1]  # Get the end time
            # Check if it ends at :30
            return end_time.endswith(':30')
        except (IndexError, AttributeError):
            return False
    
    def _get_preferred_start_slots(self, duration, regular_slots):
        """Get start slots that result in courses ending at :30.
        Returns (preferred_slots, remaining_slots)."""
        preferred = []
        remaining = []
        
        for start_slot in regular_slots:
            slots = self._get_consecutive_slots(start_slot, duration)
            # Check all slots are in regular_slots and not in excluded slots
            if len(slots) == duration and all(s in regular_slots for s in slots):
                if self._ends_at_thirty(slots):
                    preferred.append(start_slot)
                else:
                    remaining.append(start_slot)
        
        return preferred, remaining
    
    def _get_dept_from_global_key(self, dept_key):
        """Extract department label from a global slot key (e.g., 'CSE-A' from 'CSE-A_Pre-Mid')."""
        return dept_key.split('_')[0] if dept_key else ''

    def _departments_can_share_slots(self, dept_a, dept_b):
        """Return True if two departments are allowed to share the same time slots."""
        if not dept_a or not dept_b:
            return False

        share_groups = [
            {"CSE-A", "CSE-B"},
        ]

        for group in share_groups:
            if dept_a in group and dept_b in group:
                return True
        return False

    def _is_time_slot_available_global(self, day, slots, department, session, semester_id):
        """Enhanced slot availability check to prevent conflicts.
        Rules:
        - Same department + same session = conflict (same students can't be in two classes)
        - Same department + different session = OK (different students)
        - Different departments = OK (different students, can share slots)
        - CSE-A and CSE-B can share slots for the same courses."""
        semester_key = f"sem_{semester_id}"
        
        # Use the same tracking system as _mark_slots_busy_global
        if semester_key not in self.semester_global_slots:
            return True  # No slots booked yet for this semester
        
        # Check for conflicts
        for slot in slots:
            # Check semester-wide conflicts
            for dept_key, used_slots in self.semester_global_slots[semester_key].items():
                if (day, slot) in used_slots:
                    # Extract department from dept_key (format: "DEPT_SESSION")
                    dept_in_slot = dept_key.split('_')[0] if '_' in dept_key else dept_key
                    session_in_slot = dept_key.split('_')[1] if '_' in dept_key else ''
                    
                    # Allow CSE-A and CSE-B to share slots (they can have same courses at same time)
                    if self._departments_can_share_slots(department, dept_in_slot):
                        continue  # Allow sharing between CSE-A and CSE-B
                    
                    # Allow same department different sessions (different students)
                    if department == dept_in_slot and session != session_in_slot:
                        continue
                    
                    # Allow different departments (different students, can share slots)
                    if department != dept_in_slot:
                        continue
                    
                    # Block: same department + same session = conflict
                    # (This means department == dept_in_slot and session == session_in_slot)
                    return False
        return True

    def _mark_slots_busy_global(self, day, slots, department, session, semester_id):
        """Mark time slots as busy in global tracker."""
        key = f"{department}_{session}"
        semester_key = f"sem_{semester_id}"
        
        if semester_key not in self.semester_global_slots:
            self.semester_global_slots[semester_key] = {}
        
        if key not in self.semester_global_slots[semester_key]:
            self.semester_global_slots[semester_key][key] = set()
        
        for slot in slots:
            self.semester_global_slots[semester_key][key].add((day, slot))
            # prepare room occupancy tracker
            occ_key = (semester_key, day, slot)
            if occ_key not in self.room_occupancy:
                self.room_occupancy[occ_key] = set()
    
    def _is_time_slot_available_local(self, schedule, day, slots):
        """Check if time slots are available in local schedule."""
        for slot in slots:
            if schedule.loc[day, slot] != 'Free':
                return False
        return True
    
    def _mark_slots_busy_local(self, schedule, day, slots, course_code, class_type):
        """Mark time slots as busy in local schedule."""
        suffix = ''
        if class_type == 'Lab':
            suffix = ' (Lab)'
        elif class_type == 'Tutorial':
            suffix = ' (Tut)'
        elif class_type == 'Minor':
            suffix = ' (Minor)'
        
        for slot in slots:
            schedule.loc[day, slot] = f"{course_code}{suffix}"
    
    def _log_room_booking(self, semester_id, day, slot, room_name, department, course_code, session):
        """Record a room booking so conflicts can be detected later."""
        semester_key = f"sem_{semester_id}"
        slot_key = (day, slot)
        booking = {
            'room': room_name,
            'dept': department,
            'course': str(course_code).strip(),
            'session': session
        }
        if semester_key not in self.room_bookings:
            self.room_bookings[semester_key] = {}
        if slot_key not in self.room_bookings[semester_key]:
            self.room_bookings[semester_key][slot_key] = []
        self.room_bookings[semester_key][slot_key].append(booking)
    
    def _assign_room(self, day, slot, course_code, department, session, semester_id, is_lab=False, is_combined=False, required_capacity=0, slots=None):
        """Assign a room for a course at the specified slots with specific rules."""
        semester_key = f"sem_{semester_id}"
        slot_sequence = slots if slots else [slot]
        
        def _room_available(room_name):
            for slot_label in slot_sequence:
                occ_key = (semester_key, day, slot_label)
                if occ_key in self.room_occupancy and room_name in self.room_occupancy[occ_key]:
                    return False
            return True
        
        def _mark_room_usage(room_name, target_allocation):
            for slot_label in slot_sequence:
                occ_key = (semester_key, day, slot_label)
                if occ_key not in self.room_occupancy:
                    self.room_occupancy[occ_key] = set()
                self.room_occupancy[occ_key].add(room_name)
                if not room_name or room_name.upper() not in self.shared_rooms:
                    self._log_room_booking(semester_id, day, slot_label, room_name, department, course_code, session)
            target_allocation.append((day, slot, room_name))
        
        # RULE 1: Combined classes MUST use C004
        if is_combined:
            if self.c004_room:
                room_name, room_capacity = self.c004_room
                # Check capacity
                if required_capacity > 0 and room_capacity < required_capacity:
                    print(f"      WARNING: C004 capacity ({room_capacity}) insufficient for combined class {course_code} ({required_capacity} students)")
                    return None
                
                allocation_key = (semester_id, department, session, course_code)
                if allocation_key not in self.assigned_rooms:
                    self.assigned_rooms[allocation_key] = []
                
                _mark_room_usage(room_name, self.assigned_rooms[allocation_key])
                
                print(f"      Assigned C004 for combined class {course_code} at {day} {slot}")
                return room_name
            else:
                print(f"      ERROR: C004 not found for combined class {course_code}")
                return None
        
        # RULE 2: Labs - Department specific assignment
        if is_lab:
            # CSE and DSAI get software labs
            if department in ['CSE-A', 'CSE-B', 'CSE', 'DSAI']:
                available_rooms = self.software_lab_rooms.copy()
                lab_type = "software"
            # ECE gets hardware labs
            elif department in ['ECE']:
                available_rooms = self.hardware_lab_rooms.copy()
                lab_type = "hardware"
            else:
                # Default to all labs if department doesn't match
                available_rooms = self.lab_rooms.copy()
                lab_type = "any"
            
            # Filter out rooms already occupied at this time
            filtered_rooms = []
            for name, cap in available_rooms:
                if _room_available(name):
                    filtered_rooms.append((name, cap))
            available_rooms = filtered_rooms
            
            # Sort by capacity (prefer smallest fitting room)
            if required_capacity > 0:
                fitting_rooms = [(name, cap) for name, cap in available_rooms if cap >= required_capacity]
                if fitting_rooms:
                    fitting_rooms.sort(key=lambda x: x[1])  # Sort by capacity ascending
                    available_rooms = fitting_rooms
            
            # Select room (prefer smaller rooms for better utilization)
            if available_rooms:
                available_rooms.sort(key=lambda x: x[1])
                selected_room = available_rooms[0][0]
                
                # Mark room as occupied
                allocation_key = (semester_id, department, session, course_code)
                if allocation_key not in self.assigned_lab_rooms:
                    self.assigned_lab_rooms[allocation_key] = []
                
                _mark_room_usage(selected_room, self.assigned_lab_rooms[allocation_key])
                
                print(f"      Assigned {lab_type} lab {selected_room} for {course_code} at {day} {slot}")
                return selected_room
            else:
                print(f"      WARNING: No {lab_type} lab available for {course_code} at {day} {slot}")
                return None
        
        # RULE 3: Regular classes - Use normal classrooms (NOT C004)
        available_rooms = [(name, cap) for name, cap in self.nonlab_rooms.copy() if name.upper() != 'C004']
        available_rooms = [(name, cap) for name, cap in available_rooms if _room_available(name)]
        
        # Sort by capacity (prefer smallest fitting room)
        if required_capacity > 0:
            fitting_rooms = [(name, cap) for name, cap in available_rooms if cap >= required_capacity]
            if fitting_rooms:
                fitting_rooms.sort(key=lambda x: x[1])  # Sort by capacity ascending
                available_rooms = fitting_rooms
        
        # Select room (prefer smaller rooms for better utilization)
        if available_rooms:
            available_rooms.sort(key=lambda x: x[1])
            selected_room = available_rooms[0][0]
            
            allocation_key = (semester_id, department, session, course_code)
            if allocation_key not in self.assigned_rooms:
                self.assigned_rooms[allocation_key] = []
            _mark_room_usage(selected_room, self.assigned_rooms[allocation_key])
            
            return selected_room
        
        return None
    
    def _schedule_minor_classes(self, schedule, department, session, semester_id):
        """Schedule Minor subject classes ONLY in configured MINOR_SLOTS (e.g., 07:30-08:30 split).
        All departments/sections in a semester get the same minor slots."""
        # Skip minor scheduling entirely for semester 1 as per requirement
        if int(semester_id) == 1:
            return
        scheduled = 0
        attempts = 0
        max_attempts = 200
        
        # Compute valid minor start slots (so MINOR_DURATION consecutive slots are within MINOR_SLOTS)
        minor_starts = []
        for s in MINOR_SLOTS:
            seq = self._get_consecutive_slots(s, MINOR_DURATION)
            if len(seq) == MINOR_DURATION and all(x in MINOR_SLOTS for x in seq):
                minor_starts.append(s)
        
        if not minor_starts:
            # nothing to schedule if config mismatched
            return
        
        semester_key = f"sem_{semester_id}"
        # If already assigned for this semester, use the same slots
        if semester_key in self.semester_minor_slots:
            assigned = self.semester_minor_slots[semester_key]
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, MINOR_DURATION)
                self._mark_slots_busy_local(schedule, day, slots, MINOR_SUBJECT, 'Minor')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
            return

        assigned = []
        while scheduled < MINOR_CLASSES_PER_WEEK and attempts < max_attempts:
            attempts += 1
            
            day = random.choice(DAYS)
            start = random.choice(minor_starts)
            slots = self._get_consecutive_slots(start, MINOR_DURATION)
            
            if (len(slots) == MINOR_DURATION and
                self._is_time_slot_available_local(schedule, day, slots) and
                self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                
                self._mark_slots_busy_local(schedule, day, slots, MINOR_SUBJECT, 'Minor')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
                
                assigned.append((day, start))
                scheduled += 1
        # Save assigned slots for all departments in this semester
        if assigned:
            self.semester_minor_slots[semester_key] = assigned

    def _find_combined_slots(self, schedule, component, duration, department, session, semester_id, avoid_days=None):
        """Find available slots for combined classes across all departments in the same group.
        Returns list of (day, start_slot) tuples for combined scheduling."""
        combined_slots = []
        attempts = 0
        max_attempts = 500
        avoid_days = set(avoid_days or [])
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(duration, regular_slots)
        
        # Combine: preferred first, then remaining
        all_possible_starts = preferred_starts + remaining_starts
        
        # Determine which departments can share combined slots
        def _get_combined_group(dept_label):
            if dept_label in {'CSE-A', 'CSE-B'}:
                return 'CSE'
            if dept_label in {'DSAI', 'ECE'}:
                return 'DSAI_ECE'
            return None
        
        group_key = _get_combined_group(department)
        if not group_key:
            return combined_slots
        
        while attempts < max_attempts:
            attempts += 1
            day = random.choice([d for d in DAYS if d not in avoid_days])
            if not all_possible_starts:
                break
            start_slot = random.choice(all_possible_starts)
            slots = self._get_consecutive_slots(start_slot, duration)
            slots = [slot for slot in slots if slot in regular_slots]
            
            if len(slots) == duration:
                # Check if slots are available for ALL departments in the group
                all_available = True
                for dept in DEPARTMENTS:
                    if _get_combined_group(dept) == group_key:
                        # Check local availability for each department
                        dept_schedule = schedule  # We're working with current department's schedule
                        if not self._is_time_slot_available_local(dept_schedule, day, slots):
                            all_available = False
                            break
                        # Check global availability
                        if not self._is_time_slot_available_global(day, slots, dept, session, semester_id):
                            all_available = False
                            break
                
                if all_available:
                    combined_slots.append((day, start_slot))
                    break
        
        return combined_slots

    def _schedule_combined_class(self, schedule, course_code, component, duration, required_count, department, session, semester_id, avoid_days=None, required_capacity=0):
        """Schedule combined class for all departments in the same group with C004 room allocation."""
        scheduled_count = 0
        scheduled_slots = []
        assigned_rooms = []
        
        # Check if combined slots already exist for this course
        def _get_combined_group(dept_label):
            if dept_label in {'CSE-A', 'CSE-B'}:
                return 'CSE'
            if dept_label in {'DSAI', 'ECE'}:
                return 'DSAI_ECE'
            return None
        
        group_key = _get_combined_group(department)
        course_key = str(course_code).strip()
        is_lab = (component == 'Lab')
        
        # Check global combined slots first
        if group_key:
            global_key = ('GLOBAL', group_key, course_key, component)
            assigned = self.global_combined_course_slots.get(global_key, [])
            
            if assigned and len(assigned) >= required_count:
                # Use existing global slots
                for day, start in assigned[:required_count]:
                    slots = self._get_consecutive_slots(start, duration)
                    self._mark_slots_busy_local(schedule, day, slots, course_code, component)
                    self._mark_slots_busy_global(day, slots, department, session, semester_id)
                    scheduled_slots.extend([(day, slot) for slot in slots])
                    
                    # Assign C004 for this combined class
                    room = self._assign_room(day, start, course_code, department, session, semester_id, 
                                          is_lab=is_lab, is_combined=True, required_capacity=required_capacity, slots=slots)
                    if room:
                        assigned_rooms.append(room)
                    
                    scheduled_count += 1
                return scheduled_count, scheduled_slots, assigned_rooms
            
            # Check semester-specific combined slots
            sem_key = (semester_id, course_key, component)
            assigned = self.semester_combined_course_slots.get(sem_key, [])
            
            if assigned and len(assigned) >= required_count:
                # Use existing semester slots
                for day, start in assigned[:required_count]:
                    slots = self._get_consecutive_slots(start, duration)
                    self._mark_slots_busy_local(schedule, day, slots, course_code, component)
                    self._mark_slots_busy_global(day, slots, department, session, semester_id)
                    scheduled_slots.extend([(day, slot) for slot in slots])
                    
                    # Assign C004 for this combined class
                    room = self._assign_room(day, start, course_code, department, session, semester_id, 
                                          is_lab=is_lab, is_combined=True, required_capacity=required_capacity, slots=slots)
                    if room:
                        assigned_rooms.append(room)
                    
                    scheduled_count += 1
                return scheduled_count, scheduled_slots, assigned_rooms
        
        # Find new combined slots
        while scheduled_count < required_count:
            combined_slots = self._find_combined_slots(schedule, component, duration, department, session, semester_id, avoid_days)
            
            if not combined_slots:
                break
            
            day, start_slot = combined_slots[0]
            slots = self._get_consecutive_slots(start_slot, duration)
            
            # Assign C004 first (before marking slots globally)
            room = self._assign_room(day, start_slot, course_code, department, session, semester_id, 
                                  is_lab=is_lab, is_combined=True, required_capacity=required_capacity, slots=slots)
            
            if room:  # C004 must be available for combined classes
                # Mark slots for all departments in group
                for dept in DEPARTMENTS:
                    if _get_combined_group(dept) == group_key:
                        self._mark_slots_busy_global(day, slots, dept, session, semester_id)
                
                # Mark locally for current department
                self._mark_slots_busy_local(schedule, day, slots, course_code, component)
                scheduled_slots.extend([(day, slot) for slot in slots])
                scheduled_count += 1
                assigned_rooms.append(room)
                
                # Store combined slot
                if group_key:
                    global_key = ('GLOBAL', group_key, course_key, component)
                    if global_key not in self.global_combined_course_slots:
                        self.global_combined_course_slots[global_key] = []
                    self.global_combined_course_slots[global_key].append((day, start_slot))
                    
                    sem_key = (semester_id, course_key, component)
                    if sem_key not in self.semester_combined_course_slots:
                        self.semester_combined_course_slots[sem_key] = []
                    self.semester_combined_course_slots[sem_key].append((day, start_slot))
            else:
                # C004 not available, try different slot
                continue
        
        return scheduled_count, scheduled_slots, assigned_rooms

    def _schedule_lectures(self, schedule, course_code, lectures_per_week, department, session, semester_id, avoid_days=None, is_combined=False, is_elective=False, is_minor=False, required_capacity=0):
        """Schedule lecture sessions with room allocation, returns list of (day, slot) tuples.
        Prioritizes slots ending at :30, falls back to remaining slots if needed."""
        if lectures_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 2000  # Increased attempts for better allocation
        used_days = set()
        avoid_days = set(avoid_days or [])
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]

        # PRIORITY 1: Handle combined classes first
        if is_combined:
            scheduled_count, combined_slots, assigned_rooms = self._schedule_combined_class(
                schedule, course_code, 'Lecture', LECTURE_DURATION, 
                lectures_per_week, department, session, semester_id, avoid_days, required_capacity
            )
            if scheduled_count >= lectures_per_week:
                return combined_slots
            # If not all combined slots found, continue with regular scheduling for remaining
        
        # PRIORITY 2: Handle electives with common slots
        if is_elective:
            elective_slots = self._schedule_elective_classes(schedule, course_code, lectures_per_week, department, session, semester_id, avoid_days, required_capacity)
            if elective_slots:
                return elective_slots
        
        skip_room_assignment = is_elective or is_minor

        # PRIORITY 3: Regular lecture scheduling
        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(LECTURE_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations

        while len(scheduled_slots) < lectures_per_week * LECTURE_DURATION and attempts < max_attempts:
            attempts += 1
            
            # Try to use days where this course isn't already scheduled
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos)
            
            # Check all slots are available (both local and global)
            slots_available = True
            for slot in slots:
                if not (self._is_time_slot_available_local(schedule, day, [slot]) and
                       self._is_time_slot_available_global(day, [slot], department, session, semester_id)):
                    slots_available = False
                    break
            
            if slots_available:
                room_available = True
                if not skip_room_assignment:
                    room = self._assign_room(day, start_slot, course_code, department, session, semester_id, 
                                             is_lab=False, is_combined=False, required_capacity=required_capacity, slots=slots)
                    room_available = room is not None
                if room_available:
                    for slot in slots:
                        self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                        self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                        scheduled_slots.append((day, slot))
                    used_days.add(day)
                    avoid_days.add(day)
                    all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]

        scheduled_count = len(scheduled_slots) // LECTURE_DURATION
        if scheduled_count < lectures_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{lectures_per_week} lectures (attempts: {attempts})")
        return scheduled_slots

    def _schedule_elective_classes(self, schedule, course_code, elective_per_week, department, session, semester_id, avoid_days=None, required_capacity=0):
        """Schedule elective classes at the same slot/day for all departments/sections in a semester.
        ALL electives in a semester must use the same time slots across ALL departments (CSE, DSAI, ECE).
        Returns list of (day, slot) tuples."""
        if elective_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 1000  # Increased for better success rate
        semester_key = f"sem_{semester_id}"
        avoid_days = set(avoid_days or [])
        
        # Use a common key for ALL electives in a semester to ensure same time slots
        # This ensures CSE, DSAI, and ECE all have electives at the same time
        common_elective_key = (semester_key, 'ALL_ELECTIVES')
        elective_key = (semester_key, course_code)
        used_days = set()
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        # First, check if common elective slots have been assigned for this semester
        # If yes, ALL electives must use those same slots
        if common_elective_key in self.semester_elective_slots:
            assigned = self.semester_elective_slots[common_elective_key]
            print(f"      Using common elective slots for {course_code} (already assigned for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, LECTURE_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            # Also store under course-specific key for backward compatibility
            self.semester_elective_slots[elective_key] = assigned
            return scheduled_slots
        
        # If no common slots assigned yet, check if this specific course has slots assigned
        # (This handles legacy cases where course-specific slots were assigned first)
        if elective_key in self.semester_elective_slots:
            assigned = self.semester_elective_slots[elective_key]
            # Promote to common slots so all future electives use the same slots
            self.semester_elective_slots[common_elective_key] = assigned
            print(f"      Using existing elective slots for {course_code} (promoted to common slots for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, LECTURE_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            return scheduled_slots

        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(LECTURE_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations

        assigned = []
        scheduled = 0
        while scheduled < elective_per_week and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos) if available_combos else (None, None, None)
            if day is None:
                break
            
            if (self._is_time_slot_available_local(schedule, day, slots) and
                self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
                assigned.append((day, start_slot))
                used_days.add(day)
                avoid_days.add(day)
                scheduled += 1
                # Remove this combination from future consideration
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
        
        if assigned:
            # Store under common key so ALL electives in this semester use the same slots
            self.semester_elective_slots[common_elective_key] = assigned
            # Also store under course-specific key for backward compatibility
            self.semester_elective_slots[elective_key] = assigned
            print(f"      Assigned common elective slots for semester {semester_id}: {assigned}")
        
        scheduled_count = len(scheduled_slots) // LECTURE_DURATION
        if scheduled_count < elective_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{elective_per_week} elective classes (attempts: {attempts})")
        return scheduled_slots

    def _schedule_elective_tutorials(self, schedule, course_code, elective_tutorials_per_week, department, session, semester_id, avoid_days=None, required_capacity=0):
        """Schedule elective tutorials at the same slot/day for all departments/sections in a semester.
        ALL elective tutorials in a semester must use the same time slots across ALL departments (CSE-A, CSE-B, DSAI, ECE)
        for both Pre-Mid and Post-Mid sessions.
        Returns list of (day, slot) tuples."""
        if elective_tutorials_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 1000
        semester_key = f"sem_{semester_id}"
        avoid_days = set(avoid_days or [])
        
        # Use a common key for ALL elective tutorials in a semester to ensure same time slots
        # This ensures CSE-A, CSE-B, DSAI, and ECE all have elective tutorials at the same time
        # Works for both Pre-Mid and Post-Mid sessions
        common_elective_tutorial_key = (semester_key, 'ALL_ELECTIVE_TUTORIALS')
        elective_tutorial_key = (semester_key, course_code, 'Tutorial')
        used_days = set()
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        # First, check if common elective tutorial slots have been assigned for this semester
        # If yes, ALL elective tutorials must use those same slots (for both Pre-Mid and Post-Mid)
        if common_elective_tutorial_key in self.semester_elective_tutorial_slots:
            assigned = self.semester_elective_tutorial_slots[common_elective_tutorial_key]
            print(f"      Using common elective tutorial slots for {course_code} (already assigned for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, TUTORIAL_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Tutorial')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            # Also store under course-specific key for backward compatibility
            self.semester_elective_tutorial_slots[elective_tutorial_key] = assigned
            return scheduled_slots
        
        # If no common slots assigned yet, check if this specific course has tutorial slots assigned
        # (This handles legacy cases where course-specific slots were assigned first)
        if elective_tutorial_key in self.semester_elective_tutorial_slots:
            assigned = self.semester_elective_tutorial_slots[elective_tutorial_key]
            # Promote to common slots so all future elective tutorials use the same slots
            self.semester_elective_tutorial_slots[common_elective_tutorial_key] = assigned
            print(f"      Using existing elective tutorial slots for {course_code} (promoted to common slots for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, TUTORIAL_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Tutorial')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            return scheduled_slots

        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(TUTORIAL_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, TUTORIAL_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == TUTORIAL_DURATION:
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, TUTORIAL_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == TUTORIAL_DURATION:
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations

        assigned = []
        scheduled = 0
        while scheduled < elective_tutorials_per_week and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos) if available_combos else (None, None, None)
            if day is None:
                break
            
            if (self._is_time_slot_available_local(schedule, day, slots) and
                self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Tutorial')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
                assigned.append((day, start_slot))
                used_days.add(day)
                avoid_days.add(day)
                scheduled += 1
                # Remove this combination from future consideration
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
        
        if assigned:
            # Store under common key so ALL elective tutorials in this semester use the same slots
            # This works across all departments and both Pre-Mid and Post-Mid sessions
            self.semester_elective_tutorial_slots[common_elective_tutorial_key] = assigned
            # Also store under course-specific key for backward compatibility
            self.semester_elective_tutorial_slots[elective_tutorial_key] = assigned
            print(f"      Assigned common elective tutorial slots for semester {semester_id}: {assigned}")
        
        scheduled_count = len(scheduled_slots) // TUTORIAL_DURATION
        if scheduled_count < elective_tutorials_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{elective_tutorials_per_week} elective tutorials (attempts: {attempts})")
        return scheduled_slots

    def _schedule_tutorials(self, schedule, course_code, tutorials_per_week, department, session, semester_id, avoid_days=None, is_combined=False, is_elective=False, is_minor=False, required_capacity=0):
        """Schedule tutorial sessions with room allocation, returns list of (day, slot) tuples.
        Prioritizes slots ending at :30, falls back to remaining slots if needed."""
        if tutorials_per_week == 0:
            return []

        scheduled_slots = []
        attempts = 0
        max_attempts = 1000
        used_days = set()
        avoid_days = set(avoid_days or [])
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]

        # PRIORITY 1: Handle combined classes first
        if is_combined:
            scheduled_count, combined_slots, assigned_rooms = self._schedule_combined_class(
                schedule, course_code, 'Tutorial', TUTORIAL_DURATION, 
                tutorials_per_week, department, session, semester_id, avoid_days, required_capacity
            )
            if scheduled_count >= tutorials_per_week:
                return combined_slots
            # If not all combined slots found, continue with regular scheduling for remaining
        
        skip_room_assignment = is_elective or is_minor

        # PRIORITY 2: Handle electives with common slots
        if is_elective:
            elective_slots = self._schedule_elective_tutorials(schedule, course_code, tutorials_per_week, department, session, semester_id, avoid_days, required_capacity)
            if elective_slots:
                return elective_slots

        # PRIORITY 3: Regular tutorial scheduling
        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(TUTORIAL_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, TUTORIAL_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == TUTORIAL_DURATION:
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, TUTORIAL_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == TUTORIAL_DURATION:
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations

        while len(scheduled_slots) < tutorials_per_week * TUTORIAL_DURATION and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break

            day, start_slot, slots = random.choice(available_combos)
            
            room_available = True
            slots_available = True
            for slot in slots:
                if not (self._is_time_slot_available_local(schedule, day, [slot]) and
                        self._is_time_slot_available_global(day, [slot], department, session, semester_id)):
                    slots_available = False
                    break
            
            if slots_available:
                if not skip_room_assignment:
                    room = self._assign_room(day, start_slot, course_code, department, session, semester_id, 
                                              is_lab=False, is_combined=False, required_capacity=required_capacity, slots=slots)
                    room_available = room is not None
                if room_available:
                    for slot in slots:
                        self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Tutorial')
                        self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                        scheduled_slots.append((day, slot))
                    used_days.add(day)
                    avoid_days.add(day)
                    all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]

        scheduled_count = len(scheduled_slots) // TUTORIAL_DURATION
        if scheduled_count < tutorials_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{tutorials_per_week} tutorials (attempts: {attempts})")
        
        return scheduled_slots

    def _schedule_labs(self, schedule, course_code, labs_per_week, department, session, semester_id, avoid_days=None, is_combined=False, is_elective=False, is_minor=False, required_capacity=0):
        """Schedule lab sessions in regular time slots (multi-slot labs) with department-specific lab allocation.
        Prioritizes slots ending at :30, falls back to remaining slots if needed.
        Returns list of (day, slot) tuples."""
        if labs_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 1000  # Increased for better success rate
        used_days = set()
        avoid_days = set(avoid_days or [])
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        # PRIORITY 1: Handle combined classes first
        if is_combined:
            scheduled_count, combined_slots, assigned_rooms = self._schedule_combined_class(
                schedule, course_code, 'Lab', LAB_DURATION, 
                labs_per_week, department, session, semester_id, avoid_days, required_capacity
            )
            if scheduled_count >= labs_per_week:
                return combined_slots
            # If not all combined slots found, continue with regular scheduling for remaining

        skip_room_assignment = is_elective or is_minor

        # PRIORITY 2: Department-specific lab scheduling
        # Get preferred start slots (ending at :30) and remaining slots
        preferred_starts, remaining_starts = self._get_preferred_start_slots(LAB_DURATION, regular_slots)
        
        # Build list of all possible (day, start_slot) combinations, prioritizing preferred slots
        preferred_combinations = []
        remaining_combinations = []
        for day in DAYS:
            for start_slot in preferred_starts:
                slots = self._get_consecutive_slots(start_slot, LAB_DURATION)
                if len(slots) == LAB_DURATION and all(s in regular_slots for s in slots):
                    preferred_combinations.append((day, start_slot, slots))
            for start_slot in remaining_starts:
                slots = self._get_consecutive_slots(start_slot, LAB_DURATION)
                if len(slots) == LAB_DURATION and all(s in regular_slots for s in slots):
                    remaining_combinations.append((day, start_slot, slots))
        
        # Shuffle both lists for randomness
        random.shuffle(preferred_combinations)
        random.shuffle(remaining_combinations)
        
        # Combine: preferred first, then remaining
        all_combinations = preferred_combinations + remaining_combinations
        
        while len(scheduled_slots) < labs_per_week * LAB_DURATION and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos)
            
            room_available = True
            # Check if all slots are available
            all_available = True
            for slot in slots:
                if not (self._is_time_slot_available_local(schedule, day, [slot]) and
                       self._is_time_slot_available_global(day, [slot], department, session, semester_id)):
                    all_available = False
                    break
            
            if all_available:
                if not skip_room_assignment:
                    room = self._assign_room(day, start_slot, course_code, department, session, semester_id, 
                                             is_lab=True, is_combined=False, required_capacity=required_capacity, slots=slots)
                    room_available = room is not None
                if room_available:
                    for slot in slots:
                        self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lab')
                        self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                        scheduled_slots.append((day, slot))
                    used_days.add(day)
                    avoid_days.add(day)
                    all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]

        scheduled_count = len(scheduled_slots) // LAB_DURATION
        if scheduled_count < labs_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{labs_per_week} labs (attempts: {attempts})")
        return scheduled_slots

    def _schedule_course(self, schedule, course, department, session, semester_id):
        """Schedule all components of a course based on LTPSC with proper room allocation."""
        course_code = course['Course Code']
        lectures_per_week = course['Lectures_Per_Week']
        tutorials_per_week = course['Tutorials_Per_Week']
        labs_per_week = course['Labs_Per_Week']
        
        # Robust elective detection: check multiple column name variants + pattern overrides
        elective_flag = False
        for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
            if colname in course.index:
                elective_flag = str(course.get(colname, '')).upper() == 'YES'
                if elective_flag:
                    break
        # Pattern overrides: force ELEC as elective; force HSS as not elective
        course_code_str = str(course.get('Course Code', '')).upper()
        course_name_str = str(course.get('Course Name', '')).upper()
        if 'ELEC' in course_code_str or 'ELEC' in course_name_str:
            elective_flag = True
        if 'HSS' in course_code_str or 'HSS' in course_name_str:
            elective_flag = False
        # Determine if HSS
        is_hss = ('HSS' in course_code_str) or ('HSS' in course_name_str)
        is_minor_course = (MINOR_SUBJECT.upper() in course_code_str) or (MINOR_SUBJECT.upper() in course_name_str)
        
        # Read Combined Class column (handle multiple column name variants)
        combined_class_flag = False
        for colname in ['Combined Class', 'COMBINED CLASS', 'Combined Class ', 'COMBINED CLASS ']:
            if colname in course.index:
                combined_class_val = str(course.get(colname, '')).strip().upper()
                combined_class_flag = combined_class_val == 'YES'
                if combined_class_flag:
                    break
        
        # Get student count for room capacity
        registered_students = 0
        if 'Registered Students' in course.index:
            try:
                registered_students = int(float(course.get('Registered Students', 0)))
            except:
                registered_students = 0
        
        # IMPORTANT: Follow LTPSC strictly for ALL courses (including electives)
        # Do not override parsed weekly counts; use values from Excel (via parse_ltpsc)
        elective_status = " [ELECTIVE]" if elective_flag else ""
        combined_status = " [COMBINED]" if combined_class_flag else ""
        print(f"      Scheduling {course_code}{elective_status}{combined_status}: L={lectures_per_week}, T={tutorials_per_week}, P={labs_per_week}, Students={registered_students}")
        
        # Track used days for this course scoped to semester+department+session
        scoped_key = (semester_id, department, session, course_code)
        if scoped_key not in self.scheduled_courses:
            self.scheduled_courses[scoped_key] = set()

        success_counts = {'lectures': 0, 'tutorials': 0, 'labs': 0}
        scheduled_slots = []
        assigned_rooms = []
        assigned_lab_rooms = []

        # PRIORITY 1: Schedule combined classes first
        # This ensures combined classes get priority for available slots
        avoid_days = set()
        
        # Schedule lectures
        if lectures_per_week > 0:
            lecture_slots = self._schedule_lectures(
                schedule, course_code, lectures_per_week, department, session, semester_id,
                avoid_days, is_combined=combined_class_flag, is_elective=elective_flag, 
                is_minor=is_minor_course, required_capacity=registered_students
            )
            scheduled_slots.extend(lecture_slots)
            success_counts['lectures'] = len(lecture_slots) // LECTURE_DURATION
        
        # Schedule tutorials
        if tutorials_per_week > 0:
            tutorial_slots = self._schedule_tutorials(
                schedule, course_code, tutorials_per_week, department, session, semester_id,
                avoid_days, is_combined=combined_class_flag, is_elective=elective_flag,
                is_minor=is_minor_course, required_capacity=registered_students
            )
            scheduled_slots.extend(tutorial_slots)
            success_counts['tutorials'] = len(tutorial_slots) // TUTORIAL_DURATION
        
        # Schedule labs
        if labs_per_week > 0:
            lab_slots = self._schedule_labs(
                schedule, course_code, labs_per_week, department, session, semester_id,
                avoid_days, is_combined=combined_class_flag,
                is_elective=elective_flag, is_minor=is_minor_course,
                required_capacity=registered_students
            )
            scheduled_slots.extend(lab_slots)
            success_counts['labs'] = len(lab_slots) // LAB_DURATION

        # Get assigned rooms for this course
        allocation_key = (semester_id, department, session, course_code)
        if allocation_key in self.assigned_rooms:
            assigned_rooms = [room_info[2] for room_info in self.assigned_rooms[allocation_key]]
        if allocation_key in self.assigned_lab_rooms:
            assigned_lab_rooms = [room_info[2] for room_info in self.assigned_lab_rooms[allocation_key]]

        # Store actual allocation counts with room information
        self.actual_allocations[allocation_key] = {
            'lectures': success_counts['lectures'],
            'tutorials': success_counts['tutorials'],
            'labs': success_counts['labs'],
            'combined_class': combined_class_flag,
            'room': assigned_rooms[0] if assigned_rooms else '',
            'lab_room': assigned_lab_rooms[0] if assigned_lab_rooms else ''
        }

        # Store scheduled slots for conflict tracking
        if scoped_key not in self.scheduled_slots:
            self.scheduled_slots[scoped_key] = []
        self.scheduled_slots[scoped_key].extend(scheduled_slots)
        
        # Track days used by this course
        for day, slot in scheduled_slots:
            self.scheduled_courses[scoped_key].add(day)

        return success_counts

    def generate_department_schedule(self, semester_id, department, session):
        """Generate a complete weekly schedule for a department and session."""
        print(f"\nGenerating schedule for {department} {session} (Semester {semester_id})")
        
        # Initialize empty schedule
        schedule = self._initialize_schedule()
        
        # Get courses for this department and session
        sem_courses = ExcelLoader.get_semester_courses(self.dfs, semester_id)
        if sem_courses.empty:
            print(f"WARNING: No courses found for semester {semester_id}")
            return schedule
        
        # Parse LTPSC
        sem_courses = ExcelLoader.parse_ltpsc(sem_courses)
        if sem_courses.empty:
            print(f"WARNING: No valid courses after LTPSC parsing for semester {semester_id}")
            return schedule
        
        # Filter for department
        if 'Department' in sem_courses.columns:
            dept_mask = sem_courses['Department'].astype(str).str.contains(f"^{department}$", na=False, regex=True)
            dept_courses = sem_courses[dept_mask].copy()
        else:
            dept_courses = sem_courses.copy()
        
        if dept_courses.empty:
            print(f"WARNING: No courses found for {department} in semester {semester_id}")
            return schedule
        
        # Divide by session
        pre_mid_courses, post_mid_courses = ExcelLoader.divide_courses_by_session(dept_courses, department, all_sem_courses=sem_courses)
        
        # Select appropriate session
        if session == PRE_MID:
            session_courses = pre_mid_courses
        else:
            session_courses = post_mid_courses
        
        if session_courses.empty:
            print(f"WARNING: No courses assigned to {department} {session} session")
            return schedule
        
        # Schedule minor classes first (early morning)
        self._schedule_minor_classes(schedule, department, session, semester_id)
        
        # Schedule each course
        # Sort courses to ensure consistent scheduling order
        # Combined classes and electives get priority
        course_priority = []
        regular_courses = []
        
        for _, course in session_courses.iterrows():
            # Check if combined class
            combined_class_flag = False
            for colname in ['Combined Class', 'COMBINED CLASS', 'Combined Class ', 'COMBINED CLASS ']:
                if colname in course.index:
                    combined_class_val = str(course.get(colname, '')).strip().upper()
                    combined_class_flag = combined_class_val == 'YES'
                    break
            
            # Check if elective
            elective_flag = False
            for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
                if colname in course.index:
                    elective_flag = str(course.get(colname, '')).upper() == 'YES'
                    if elective_flag:
                        break
            # Pattern overrides
            course_code_str = str(course.get('Course Code', '')).upper()
            course_name_str = str(course.get('Course Name', '')).upper()
            if 'ELEC' in course_code_str or 'ELEC' in course_name_str:
                elective_flag = True
            if 'HSS' in course_code_str or 'HSS' in course_name_str:
                elective_flag = False
            
            if combined_class_flag or elective_flag:
                course_priority.append(course)
            else:
                regular_courses.append(course)
        
        # Schedule priority courses first (combined classes and electives)
        print(f"  Scheduling {len(course_priority)} priority courses (combined/electives)...")
        for course in course_priority:
            self._schedule_course(schedule, course, department, session, semester_id)
        
        # Schedule regular courses
        print(f"  Scheduling {len(regular_courses)} regular courses...")
        for course in regular_courses:
            self._schedule_course(schedule, course, department, session, semester_id)
        
        print(f"Schedule generated for {department} {session}")
        return schedule
    
    def get_actual_allocations(self, semester_id, department, session, course_code):
        """Get actual number of classes allocated for a course."""
        allocation_key = (semester_id, department, session, course_code)
        return self.actual_allocations.get(allocation_key, {
            'lectures': 0,
            'tutorials': 0,
            'labs': 0,
            'combined_class': False,
            'room': '',
            'lab_room': ''
        })
    
    def validate_room_conflicts(self):
        """Validate room allocation conflicts across all schedules."""
        conflicts = []
        
        for semester_key, semester_data in self.room_bookings.items():
            for (day, slot), bookings in semester_data.items():
                if len(bookings) > 1:
                    # Multiple courses in same room at same time
                    conflict = {
                        'semester': semester_key,
                        'day': day,
                        'slot': slot,
                        'room': bookings[0]['room'],
                        'entries': [(b['dept'], b['course'], b['session']) for b in bookings]
                    }
                    conflicts.append(conflict)
        
        return conflicts