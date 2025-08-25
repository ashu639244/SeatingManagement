import os
import pandas as pd
from datetime import datetime

def load_data(ip1_file, ip2_file, ip3_file, ip4_file):
    try:
        ip1 = pd.read_csv(ip1_file, sep=',', header=1, on_bad_lines='skip')
        ip2 = pd.read_csv(ip2_file, sep=',', header=1, on_bad_lines='skip')
        ip3 = pd.read_csv(ip3_file, sep=',', header=0, on_bad_lines='skip')
        ip4 = pd.read_csv(ip4_file, sep=',', header=0, on_bad_lines='skip')

        # Strip whitespace from column names
        ip1.columns = ip1.columns.str.strip()
        ip2.columns = ip2.columns.str.strip()
        ip3.columns = ip3.columns.str.strip()
        ip4.columns = ip4.columns.str.strip()

        # Convert the 'Date' column in ip2 to datetime
        ip2['Date'] = pd.to_datetime(ip2['Date'], format='%d/%m/%Y', errors='coerce')

        return ip1, ip2, ip3, ip4
    except Exception as e:
        print(f"Error loading data: {e}")
        return None, None, None, None

def arrange_seating(ip1, ip2, ip3, date, session, original_room_capacities, buffer_size, arrangement_type):
    # Select courses for the specified date and session
    courses = ip2.loc[ip2['Date'] == date, session]

    # Check if any courses are found
    if courses.empty:
        print(f"No courses found for date: {date} and session: {session}.")
        return pd.DataFrame()  # Return an empty DataFrame

    courses = courses.values[0]
    course_list = [course.strip() for course in courses.split(';')] if courses != 'NO EXAM' else []

    # Prepare output DataFrame
    output = []
    room_capacities = original_room_capacities.copy()  # Reset room capacities for this session

    for course in course_list:
        # Strip spaces from course code for accurate matching
        course = course.strip()

        # Get students enrolled in the course
        enrolled_students = ip1[ip1['course_code'].str.strip() == course]['rollno'].tolist()
        num_students = len(enrolled_students)

        # Allocate students to rooms based on capacity
        remaining_students = num_students
        student_index = 0

        for index, room in ip3.iterrows():
            if remaining_students <= 0:
                break  # All students have been assigned

            room_no = room['Room No.']
            available_capacity = room_capacities[room_no]  # Get current available capacity

            # Calculate effective capacity based on arrangement type
            effective_capacity = available_capacity - buffer_size  # Apply buffer first

            # Ensure effective capacity is not negative
            if effective_capacity < 0:
                effective_capacity = 0

            # Adjust for sparse arrangement
            if arrangement_type == 'Sparse':
                effective_capacity = effective_capacity // 2  # Max half for one course

            # Check if room has capacity
            if effective_capacity > 0:
                # Determine how many students can fit in this room
                students_in_room = min(effective_capacity, remaining_students)

                # Prepare output with additional details
                assigned_students = enrolled_students[student_index:student_index + students_in_room]
                output.append({
                    'Date': date.strftime('%d-%m-%Y'),  # Format date to dd-mm-yyyy
                    'Session': session,
                    'Course Code': course,
                    'Room No.': room_no,
                    'Total Capacity': room['Exam Capacity'],
                    'Students Present': students_in_room,
                    'Available Capacity After Allocation': available_capacity - students_in_room,
                    'Roll Numbers': ', '.join(map(str, assigned_students))
                })

                # Update remaining students and index
                remaining_students -= students_in_room
                student_index += students_in_room

                # Update the remaining capacity of the room
                room_capacities[room_no] -= students_in_room  # Reduce the available capacity

    return pd.DataFrame(output)

def save_output(seating_arrangement, output_dir ):
    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    csv_file = 'seating_arrangement.csv'
    excel_file = 'seating_arrangement.xlsx'
    
    csv_path = os.path.join(output_dir, csv_file)
    excel_path = os.path.join(output_dir, excel_file)

    seating_arrangement.to_csv(csv_path, index=False)
    seating_arrangement.to_excel(excel_path, index=False)

    print(f"Files saved at:\nCSV: {csv_path}\nExcel: {excel_path}")

def create_attendance_sheets(seating_arrangement, ip1, ip4, output_dir):
    # Group by Date, Session, and Room No.
    grouped = seating_arrangement.groupby(['Date', 'Session', 'Room No.'])

    for (date, session, room_no), group in grouped:
        # Create a new Excel file for each combination of date, session, and room number
        file_name = f"{date}_{session}_{room_no}.xlsx"  # Date already formatted
        file_path = os.path.join(output_dir, file_name)

        with pd.ExcelWriter(file_path) as writer:
            for course_code, course_group in group.groupby('Course Code'):
                # Get enrolled students for the course
                enrolled_students = course_group['Roll Numbers'].values[0].split(', ')
                student_names = ip4[ip4['Roll'].isin(enrolled_students)]['Name'].tolist()
                enrolled_students_with_names = list(zip(enrolled_students, student_names))

                # Create a DataFrame for the attendance sheet
                attendance_df = pd.DataFrame(enrolled_students_with_names, columns=['Roll No', 'Name'])

                # Add a 'Signature' column initialized with empty strings
                attendance_df['Signature'] = [''] * len(attendance_df)

                # Add empty rows for signatures
                signature_rows = pd.DataFrame({
                    'Roll No': [''] * 5,
                    'Name': [' '] * 5,
                    'Signature': [''] * 5  # Add empty Signature column
                })
                attendance_df = pd.concat([attendance_df, signature_rows], ignore_index=True)

                # Write the attendance DataFrame to a new sheet in the Excel file
                attendance_df.to_excel(writer, sheet_name=course_code, index=False)
                
def main():
    ip1_file = r"C:\Classes IITP\Electives\Sem5_CS384 - Python Programming\Project\2201MM27_CS384_2024\proj1\ip_1.csv"
    ip2_file = r"C:\Classes IITP\Electives\Sem5_CS384 - Python Programming\Project\2201MM27_CS384_2024\proj1\ip_2.csv"
    ip3_file = r"C:\Classes IITP\Electives\Sem5_CS384 - Python Programming\Project\2201MM27_CS384_2024\proj1\ip_3.csv"
    ip4_file = r"C:\Classes IITP\Electives\Sem5_CS384 - Python Programming\Project\2201MM27_CS384_2024\proj1\ip_4.csv"

    output_dir_seating = r"C:\Classes IITP\Electives\Sem5_CS384 - Python Programming\Project\2201MM27_CS384_2024\proj1\output"
    output_dir_attendance = r'C:\Classes IITP\Electives\Sem5_CS384 - Python Programming\Project\2201MM27_CS384_2024\proj1\output\Attendance'

    ip1, ip2, ip3, ip4 = load_data(ip1_file, ip2_file, ip3_file, ip4_file)

    if ip1 is None or ip2 is None or ip3 is None or ip4 is None:
        print("Error loading data.")
        return

    original_room_capacities = {room['Room No.']: room['Exam Capacity'] for _, room in ip3.iterrows()}

    buffer_size = int(input("Enter buffer size (0 for no buffer): "))
    arrangement_type = input("Enter arrangement type (Dense or Sparse): ").title()
    if arrangement_type == 'Sparse' and buffer_size > 0:
        buffer_size = buffer_size - 1
    
    all_seating_arrangements = []

    unique_dates = ip2['Date'].unique()
    unique_sessions = ip2.columns[1:]

    for date in unique_dates:
        for session in unique_sessions:
            seating_arrangement = arrange_seating(ip1, ip2, ip3, date, session, original_room_capacities, buffer_size, arrangement_type)
            all_seating_arrangements.append(seating_arrangement)

    complete_seating_arrangement = pd.concat(all_seating_arrangements, ignore_index=True)

    save_output(complete_seating_arrangement, output_dir_seating)

    create_attendance_sheets(complete_seating_arrangement, ip1, ip4, output_dir_attendance)

    print(f"Seating arrangement and attendance sheets saved in directory: {output_dir_seating}, {output_dir_attendance}")

if __name__ == "__main__":
    main()

end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time))