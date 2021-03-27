import pandas as pd
from termcolor import cprint
import math


def set_breakout_rooms(data_file: str, workshop_column: str, ignore: list):

    """
    Sets the breakout rooms for attendees and writes the results
    to csv files.
    
    Parameters:
    data_file (str): the absolute path to the excel data file
    workshop_column (str): the column in the source excel file where the workshop data choices are located.
    ignore (list): the list of workshops/speakers who are in the data but are no longer participating/canceled.
  
    Returns:
    Nothing, but does write to csv files in project root directory called 'workshop_column + room_codes.csv' and 'workshop_column + room_assignments.csv'

    """

    df = pd.read_excel(data_file)
    # replace with your columns of interest
    df = df[['Name', 'School Email', workshop_column]]
    id_column = "ID"
    df[id_column] = df["Name"] + "," + df["School Email"]
    df[workshop_column] = df[workshop_column].apply(lambda x: str(x).split(", "))
    cprint("Data successfully read in from " + data_file, "cyan")

    # Determines the order in which we fill the breakout rooms
    rooms = {}
    for index, row in df.iterrows():
        i = len(row[workshop_column])
        for choice in row[workshop_column]:
            if choice in ignore:
                continue
            elif choice not in rooms:
                rooms[choice] = i
            else:
                rooms[choice] += i
            i -= 1

    # This is the final sorted order in which we assign breakout rooms
    room_list = list(rooms.keys())
    room_list.sort(key=lambda x: rooms[x])
    print("Order in which we assigned rooms: " + str(room_list))

    # Maximum participants per breakout room
    MAX = math.ceil(len(df) / len(room_list)) + 1

    # email addresses of people already assigned
    already_assigned = set()
    # maps speakers -> email addresses of attendees preassigned to their room
    room_assignments = {}

    # populates room_assignments by filling up one speaker breakout room at a time starting with 
    # attendees that ranked them first, then second, third, etc
    for speaker in room_list:
        room_assignments[speaker] = []
        # for loop w / a break condition once the room is full
        for index in range(len(room_list)):
            if len(room_assignments[speaker]) > MAX:
                break
            for i, row in df.iterrows():
                if len(room_assignments[speaker]) > MAX:
                    break
                if row[id_column] not in already_assigned and len(row[workshop_column]) > index and row[workshop_column][index] == speaker:
                            already_assigned.add(row[id_column])
                            room_assignments[speaker].append(row[id_column])

    # assigns attendees who did not enter preferences for workshop 1 so were not assigned earlier, based on workshops with most available space
    room_list.sort(key=lambda x: len(room_assignments[x]))
    for i, row in df.iterrows():
        if row[id_column] not in already_assigned:
            for speaker in room_list:
                if len(room_assignments[speaker]) < MAX:
                    room_assignments[speaker].append(row[id_column])
                    already_assigned.add(row[id_column])
                    break

    # Final output csv format compatable with zoom import from csv breakout room preassignment: room#, email
    output_string = "Pre-assign Room Name,Name,Email Address\n"
    # final output csv format mapping breakout room numbers to the speaker in the room
    speaker_to_breakout_room = "speaker,id\n"

    breakout_room_num = 0
    for room in room_assignments:
        breakout_room_num += 1
        speaker_to_breakout_room += room + "," + str(breakout_room_num) + "\n"
        for email in room_assignments[room]:
            output_string += 'room' + str(breakout_room_num) + "," + email + "\n"

    # writes to csv the breakout room number -> which speaker is assigned to it
    breakout_room_num_file = open(f"/Users/alexguo/Desktop/registration-startup/" + workshop_column + "_room_codes.csv", "w")
    breakout_room_num_file.write (speaker_to_breakout_room)
    breakout_room_num_file.close()

    # writes to csv the final attendee room preassignments
    output_file = open(f"/Users/alexguo/Desktop/registration-startup/" + workshop_column + "_assignments.csv", "w")
    output_file.write(output_string)
    output_file.close()
    cprint("Writing to the CSVs " + workshop_column + "_assignments.csv and " + workshop_column + "_room_codes.csv was successful!", "green")
    print()


# Calls above function using based on 2021 conference ignore lists and data

# Choices to ignore because speakers canceled/etc
workshop_1_ignore = ['Adam Alpert: 10 Lessons Learned on the Journey from B-Lab to Y Combinator', 'nan']
workshop_2_ignore = ['Adam Alpert: 10 Lessons Learned on the Journey from B-Lab to Y Combinator', "Cassandra Carothers: Breaking into VC – A Circuitous Path (spoiler alert: there's no perfect way)", 'Cheryl McCants: Workshop Title Coming Soon!','nan']
workshop_3_ignore = ['Lorine Pendleton: Workshop Title Coming Soon!', "Ben Simon: Investors Suck And You Don't Need Them", 'nan']

# Call to set_breakout_room to write to CSV
excel_data = "/Users/alexguo/Desktop/registration-startup/first_version_of_startup_registration.xlsx"
set_breakout_rooms(excel_data, "Workshop_1", workshop_1_ignore)
set_breakout_rooms(excel_data, "Workshop_2", workshop_2_ignore)
set_breakout_rooms(excel_data, "Workshop_3", workshop_3_ignore)
