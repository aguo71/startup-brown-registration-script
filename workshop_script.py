import pandas as pd
from termcolor import cprint
import math

def sets_breakout_rooms(workshop_column: str, ignore: list):

    """
    Sets the breakout rooms for attendees and writes the results
    to csv files.
    
    Parameters:
    workshop_column (str): the column in the source excel file where the workshop data choices are located.
    ignore (list): the list of workshops/speakers who are in the data but are no longer participating/canceled.
  
    Returns:
    Nothing, but does write to csv files in project root directory called 'workshop_column + room_codes.csv' and 'workshop_column + room_assignments.csv'

    """

    # replace with your filename
    df = pd.read_excel("/Users/alexguo/Desktop/registration-startup/first_version_of_startup_registration.xlsx")
    df = df[['School Email', workshop_column]]
    df[workshop_column] = df[workshop_column].apply(lambda x: str(x).split(", "))

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
    cprint(room_list, "cyan")

    # Maximum participants per breakout room
    MAX = math.ceil(len(df) / len(room_list)) + 1

    # speakers -> email addresses of people attending
    # set: email addresses of people already assigned
    already_assigned = set()
    room_assignments = {}

    for speaker in room_list:
        room_assignments[speaker] = []
        # for loop w / a break condition
        for index in range(len(room_list)):
            if len(room_assignments[speaker]) > MAX:
                break
            for i, row in df.iterrows():
                if len(room_assignments[speaker]) > MAX:
                    break
                if row['School Email'] not in already_assigned and len(row[workshop_column]) > index and row[workshop_column][index] == speaker:
                            already_assigned.add(row['School Email'])
                            room_assignments[speaker].append(row['School Email'])

    # assigns people who did not enter preferences for workshop 1 based on workshops with most available space
    room_list.sort(key=lambda x: len(room_assignments[x]))
    for i, row in df.iterrows():
        if row['School Email'] not in already_assigned:
            for speaker in room_list:
                if len(room_assignments[speaker]) < MAX:
                    room_assignments[speaker].append(row['School Email'])
                    already_assigned.add(row['School Email'])
                    break

    # Final output csv format: room#, email
    output_string = "Pre-assign Room Name,Email Address\n"
    speaker_to_breakout_room = "speaker,id\n"
    breakout_room_num = 0
    for room in room_assignments:
        breakout_room_num += 1
        speaker_to_breakout_room += room + "," + str(breakout_room_num) + "\n"
        for email in room_assignments[room]:
            output_string += 'room' + str(breakout_room_num) + "," + email + "\n"

    # writes to csv the zoom breakout room number each speaker is assigned to
    breakout_room_num_file = open(f"/Users/alexguo/Desktop/registration-startup/" + workshop_column + "_room_codes.csv", "w")
    breakout_room_num_file.write (speaker_to_breakout_room)
    breakout_room_num_file.close()

    # writes to csv the attendee preassignments
    output_file = open(f"/Users/alexguo/Desktop/registration-startup/" + workshop_column + "_assignments.csv", "w")
    output_file.write(output_string)
    output_file.close()
    cprint("Writing to the CSVs " + workshop_column + "_assignments.csv and " + workshop_column + "_room_codes.csv + " was successful!", "green")

# Choices to ignore because speakers canceled/etc
# workshop_1_ignore = ['Adam Alpert: 10 Lessons Learned on the Journey from B-Lab to Y Combinator', 'nan']

# sets_breakout_rooms("Workshop_1", workshop_1_ignore)

# workshop_2_ignore = ['Adam Alpert: 10 Lessons Learned on the Journey from B-Lab to Y Combinator', "Cassandra Carothers: Breaking into VC – A Circuitous Path (spoiler alert: there's no perfect way)", 'Cheryl McCants: Workshop Title Coming Soon!','nan']

# sets_breakout_rooms("Workshop_2", workshop_2_ignore)

workshop_3_ignore = ['Lorine Pendleton: Workshop Title Coming Soon!', "Ben Simon: Investors Suck And You Don't Need Them", 'nan']

sets_breakout_rooms("Workshop_3", workshop_3_ignore)