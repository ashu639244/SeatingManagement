#Project
from datetime import datetime
start_time = datetime.now()

#Help
def proj_chat_tool():
	pass


###Code

import pandas as pd
from collections import defaultdict, deque
import sys
from datetime import datetime
import builtins  # Import builtins to use builtins.sum

# Read ip_1.csv to create the course dictionary
try:
    df1 = pd.read_csv('ip_1.csv', skiprows=1)
    df1.columns = ['rollno', 'register_sem', 'schedule_sem', 'course_code']
    course_dict = df1.groupby('course_code')['rollno'].apply(list).to_dict()
except FileNotFoundError:
    print("Error: 'ip_1.csv' not found.")
    sys.exit(1)
except Exception as e:
    print(f"Error reading 'ip_1.csv': {e}")
    sys.exit(1)

# Read ip_2.csv to extract course IDs
try:
    df2 = pd.read_csv('ip_2.csv', skiprows=1)
    df2.columns = ['date', 'day', 'morning_courses', 'evening_courses']
except FileNotFoundError:
    print("Error: 'ip_2.csv' not found.")
    sys.exit(1)
except Exception as e:
    print(f"Error reading 'ip_2.csv': {e}")
    sys.exit(1)

# Split and strip spaces from course IDs
df2['morning_courses'] = df2['morning_courses'].astype(str).str.split(';').apply(
    lambda x: [i.strip() for i in x] if isinstance(x, list) else [])
df2['evening_courses'] = df2['evening_courses'].astype(str).str.split(';').apply(
    lambda x: [i.strip() for i in x] if isinstance(x, list) else [])

# Read ip_3.csv to initialize room capacities
try:
    df3 = pd.read_csv('ip_3.csv', skiprows=1, header=None)
    df3.columns = ['room_no', 'capacity', 'extra_info']
except FileNotFoundError:
    print("Error: 'ip_3.csv' not found.")
    sys.exit(1)
except Exception as e:
    print(f"Error reading 'ip_3.csv': {e}")
    sys.exit(1)

# Take buffer and dense inputs with validation
try:
    buffer_value = int(input("Enter buffer value: "))
except ValueError:
    print("Error: Buffer value must be an integer.")
    sys.exit(1)

try:
    dense_input = input("Enter dense (1 for True, 0 for False): ")
    if dense_input not in ['0', '1']:
        raise ValueError
    dense = bool(int(dense_input))
except ValueError:
    print("Error: Dense must be 1 (True) or 0 (False).")
    sys.exit(1)

# Adjust room capacities based on buffer
room_dict = {row['room_no']: max(0, int(row['capacity']) - buffer_value) for _, row in df3.iterrows()}

# Exclude rooms with zero capacity if dense is True
if dense:
    room_dict = {room: capacity for room, capacity in room_dict.items() if capacity > 0}

# Sort rooms by capacity and handle "L" rooms at the end
l_rooms = {room: capacity for room, capacity in room_dict.items() if room.startswith('L')}
non_l_rooms = {room: capacity for room, capacity in room_dict.items() if not room.startswith('L')}

# Sort non-L rooms by descending capacity, then by room number
sorted_non_l_rooms = sorted(non_l_rooms.items(), key=lambda x: (-x[1], x[0]))
# Sort L rooms alphabetically
sorted_l_rooms = sorted(l_rooms.items(), key=lambda x: x[0])
sorted_room_dict = dict(sorted_non_l_rooms + sorted_l_rooms)

# Room allocations storage
room_allocations = []

# Excel data storage
excel_data = []

# Process each day
for _, row in df2.iterrows():
    date, day, morning_courses, evening_courses = row['date'], row['day'], row['morning_courses'], row['evening_courses']

    # Function to allocate courses
    def allocate_courses(courses, slot):
        # Initialize current rooms with courses and their assigned students
        current_rooms = {
            room: {
                'courses': defaultdict(list)  # Maps course_id to list of rollnos
            }
            for room in sorted_room_dict.keys()
        }
        for course_id in courses:
            if course_id not in course_dict:
                continue
            roll_numbers = deque(course_dict[course_id])
            for room, info in current_rooms.items():
                room_capacity = sorted_room_dict[room]
                max_per_course = room_capacity if dense else room_capacity // 2
                max_total = room_capacity
                # Calculate current total students in the room using builtins.sum
                current_total = builtins.sum(len(students) for students in info['courses'].values())
                available_space = max_total - current_total
                # Calculate how many can be assigned from this course
                current_course_count = len(info['courses'][course_id])
                assignable = min(
                    len(roll_numbers),
                    max_per_course - current_course_count,
                    available_space
                )
                for _ in range(assignable):
                    if not roll_numbers:
                        break
                    student = roll_numbers.popleft()
                    info['courses'][course_id].append(student)
                if not roll_numbers:
                    break  # All students for this course have been assigned

        # Convert allocations to a simpler structure and collect Excel data
        allocations = {}
        for room, info in current_rooms.items():
            for course_id, students in info['courses'].items():
                if students:  # Exclude courses with zero assigned students
                    allocations.setdefault(room, []).append({
                        'course_id': course_id,
                        'students': students
                    })
                    # Append to excel_data
                    excel_data.append({
                        'Date': date,
                        'Day': day,
                        'Course ID': course_id,
                        'Room Number': room,
                        'Count of Students Assigned': len(students),
                        'Roll Numbers Assigned': ', '.join(map(str, students))
                    })
        return allocations

    # Allocate Morning Courses
    morning_allocations = allocate_courses(morning_courses, 'Morning')

    # Allocate Evening Courses
    evening_allocations = allocate_courses(evening_courses, 'Evening')

    # Store the allocations
    room_allocations.append({
        'date': date,
        'day': day,
        'morning': morning_allocations,
        'evening': evening_allocations
    })

# Create a DataFrame from the excel_data
df_output = pd.DataFrame(excel_data, columns=[
    'Date',
    'Day',
    'Course ID',
    'Room Number',
    'Count of Students Assigned',
    'Roll Numbers Assigned'
])

# Define the output Excel file name with timestamp to avoid overwriting
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f'room_allocations_{timestamp}.xlsx'

# Write the DataFrame to an Excel file
try:
    df_output.to_excel(output_file, index=False)
    print(f"\nRoom allocations have been successfully written to '{output_file}'.")
except Exception as e:
    print(f"Error writing to Excel file: {e}")
    sys.exit(1)


proj_chat_tool()


#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))

