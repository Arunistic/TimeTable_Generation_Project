from datetime import datetime, timedelta
import random
import pandas as pd
import openpyxl

# Constants
MAX_RETRIES = 100
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
TIME_SLOTS = [
    ('09:00', '10:00'),
    ('10:00', '10:30'),  # First half of session
    ('10:30', '11:00'),  # Fixed short break
    ('11:00', '12:00'),
    ('12:00', '13:00'),
    ('14:00', '15:00'),
    ('15:00', '16:00'),
    ('16:00', '17:00'),
    ('14:00', '15:30'),
    ('15:30', '17:00')
]

FIXED_BREAK = ('10:30', '11:00')

lunch_breaks = {
    'Monday': ('13:00', '14:00'),
    'Tuesday': ('13:00', '14:00'),
    'Wednesday': ('13:00', '14:00'),
    'Thursday': ('13:00', '14:00'),
    'Friday': ('13:00', '14:00')
}

timetable = {}
slot_usage = {}

def time_to_datetime(time_str):
    try:
        return datetime.strptime(time_str, '%H:%M')
    except ValueError:
        print(f"Error parsing time: {time_str}")
        return None

def is_lunch_conflict(start_time, end_time, day):
    lunch_start, lunch_end = lunch_breaks[day]
    lunch_start_time = time_to_datetime(lunch_start)
    lunch_end_time = time_to_datetime(lunch_end)
    return not (end_time <= lunch_start_time or start_time >= lunch_end_time)

def allocate_sessions(num_sessions, duration, session_type, branch_sem, course, name, faculty):
    count = 0
    retries = 0
    while count < num_sessions and retries < MAX_RETRIES:
        day = random.choice(DAYS)
        for i, (start, end) in enumerate(TIME_SLOTS):
            start_time = time_to_datetime(start)
            end_time = time_to_datetime(end)

            if start_time is None or end_time is None:
                continue

            if (start, end) == FIXED_BREAK:
                continue

            if is_lunch_conflict(start_time, end_time, day):
                continue

            if duration == 1.5 and (end_time - start_time) == timedelta(hours=1, minutes=30):
                if not slot_usage[branch_sem][(day, i)]:
                    slot_usage[branch_sem][(day, i)] = True
                    timetable[branch_sem].append({
                        'Course Code': course,
                        'Day': day,
                        'Time': f"{start} - {end}",
                        'Type': session_type
                    })
                    count += 1
                    break

            elif duration == 1.0 and (end_time - start_time) == timedelta(hours=1):
                if not slot_usage[branch_sem][(day, i)]:
                    slot_usage[branch_sem][(day, i)] = True
                    timetable[branch_sem].append({
                        'Course Code': course,
                        'Day': day,
                        'Time': f"{start} - {end}",
                        'Type': session_type
                    })
                    count += 1
                    break

        retries += 1

    if retries >= MAX_RETRIES:
        print(f"Warning: Failed to allocate {num_sessions} sessions for {course} after {MAX_RETRIES} retries.")

def parse_ltp(ltp):
    try:
        l, t, p, s, c = map(int, ltp.split('-'))
        return l, p
    except:
        return 0, 0

def generate_full_timetable(csv_path):
    df = pd.read_csv(csv_path)
    course_details = {}

    for _, row in df.iterrows():
        branch = row['Branch']
        sem = row['Semester']
        course = row['Course Code']
        name = row['Course Name']
        ltp = row['L-T-P-S-C']
        faculty = row['Faculty']

        branch_sem = f"{branch}_sem{sem}"

        if branch_sem not in timetable:
            timetable[branch_sem] = []
        if branch_sem not in slot_usage:
            slot_usage[branch_sem] = {(day, i): False for day in DAYS for i in range(len(TIME_SLOTS))}

        l_hours, p_hours = parse_ltp(ltp)

        for _ in range(l_hours):
            allocate_sessions(1, 1.0, 'Lecture', branch_sem, course, name, faculty)

        num_practical_sessions = round(p_hours / 2)
        for _ in range(num_practical_sessions):
            allocate_sessions(1, 1.5, 'Practical', branch_sem, course, name, faculty)

        # Store the course details per branch_sem
        if branch_sem not in course_details:
            course_details[branch_sem] = []
        
        course_details[branch_sem].append({
            'Course Code': course,
            'Course Name': name,
            'Faculty': faculty,
            'L-T-P-S-C': ltp
        })

    for branch_sem in timetable:
        for day in DAYS:
            timetable[branch_sem].append({
                'Course Code': '-',
                'Day': day,
                'Time': f"{FIXED_BREAK[0]} - {FIXED_BREAK[1]}",
                'Type': 'Break'
            })

    return timetable, course_details

def export_timetable_to_excel(timetable, course_details, filename="timetables.xlsx"):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Writing timetable and course details in the same sheet
        for class_sem, sessions in timetable.items():
            data = []
            for day in DAYS:
                row = {'Day': day}
                for time in TIME_SLOTS:
                    slot_str = f"{time[0]} - {time[1]}"
                    row[slot_str] = ''
                data.append(row)

            for s in sessions:
                day = s['Day']
                time = s['Time']
                course_code = s['Course Code']
                session_type = s['Type']

                cell = f"{course_code} ({session_type[0]})"

                for row in data:
                    if row['Day'] == day:
                        row[time] = cell

            df_timetable = pd.DataFrame(data)

            # Writing the timetable for the current class
            df_timetable.to_excel(writer, sheet_name=class_sem, index=False)

            # Adding course details below the timetable
            course_details_start_row = len(df_timetable) + 2  # Adding 2 to leave a row between the timetable and course details
            if class_sem in course_details:
                df_details = pd.DataFrame(course_details[class_sem])
                df_details.to_excel(writer, sheet_name=class_sem, startrow=course_details_start_row, index=False)

def print_timetable_terminal(timetable):
    for class_sem, sessions in timetable.items():
        print(f"\n{'-'*10} Timetable for {class_sem} {'-'*10}")
        grid = {slot: {day: '' for day in DAYS} for slot in [f"{s[0]} - {s[1]}" for s in TIME_SLOTS]}
        for s in sessions:
            grid[s['Time']][s['Day']] = f"{s['Course Code']} ({s['Type'][0]})"

        header = "{: <12}".format("Time") + "".join([f"{day: <12}" for day in DAYS])
        print(header)
        print("-" * len(header))
        for time, row in grid.items():
            line = "{: <12}".format(time)
            for day in DAYS:
                line += f"{row[day]: <12}"
            print(line)

if __name__ == "__main__":
    timetable = {}
    slot_usage = {}

    final_timetable, course_details = generate_full_timetable("cleaned_courses.csv")

    export_timetable_to_excel(final_timetable, course_details)
    print_timetable_terminal(final_timetable)
    print("\nExcel file 'timetables.xlsx' created with each class timetable in proper format (grid). Course details are listed below each class timetable.")
