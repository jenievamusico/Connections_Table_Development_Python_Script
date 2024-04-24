import pandas as pd
import time
import math
import tkinter as tk
from tkinter import filedialog


# ------------------------------ FUNCTIONS ------------------------------
def get_filename_from_user(message): # ----------- Opens a file selection box for user to select input files
    
    root = tk.Tk()
    root.withdraw()
    filename = filedialog.askopenfilename(title=message)
    
    return filename


def connections_sheet_input_management(df_connectionsInput): # ----------- Format connections sheet input file
    
    # remove all rows after connections info
    for i in range(len(df_connectionsInput)):
        if  str(df_connectionsInput.loc[i, 'Set ID']) == "nan":
            endOfTimetable = i
            break

    df_connectionsInput.drop(df_connectionsInput.index[endOfTimetable:], inplace=True)

    return df_connectionsInput


def timetable_management(df_Timetable): # ----------- Format timetable dataframe headers, values etc

    # rename courseID header
    df_Timetable.rename(columns = {"// courseID ": "courseID"},  inplace = True) 

    # remove any spaces from column headers
    df_Timetable.columns = df_Timetable.columns.str.rstrip(' ')

    # remove first row of df (empty row)
    df_Timetable.drop(df_Timetable.index[0], inplace=True)
    df_Timetable.reset_index(drop=True, inplace=True)

    # remove all rows after timetable (connections table)
    for i in range(len(df_Timetable)):
        if str(df_Timetable.loc[i, 'courseID']) == '//' or str(df_Timetable.loc[i, 'courseID']) == "nan": # this can be changed to only check for 'nan' if the timetable definitely does not have a connections table at the bottom
            endOfTimetable = i
            break

    df_Timetable.drop(df_Timetable.index[endOfTimetable:], inplace=True)

    return df_Timetable


def createConnectionsTableOutput(df_connectionsOutput,df_Timetable,df_connectionsInput,df_connection_time_input): # ----------- Conduct all information gathering, comparing, and calculations, and then output to output dataframe

    # this section retrieves each courseID from each row in the connections sheet and saves it to a list 
    courseID = []
    for row_index in range(len(df_connectionsInput)):
        for col_index in range(8, len(df_connectionsInput.columns), 2):
            if str(df_connectionsInput.iloc[row_index, col_index+1]) != 'nan':
                courseID.append(df_connectionsInput.iloc[row_index, col_index])
            else:
                break

    # this section retrieves each connecting courseID from each row in the connections sheet and saves it to a list 
    connCourseID = []

    for row_index in range(len(df_connectionsInput)):
        for col_index in range(10, len(df_connectionsInput.columns), 2):
            if str(df_connectionsInput.iloc[row_index, col_index]) != 'nan':
                connCourseID.append(df_connectionsInput.iloc[row_index, col_index])
            else:
                break

    # this section creates two new lists of the reversed connections to be added to the end of each list later on
    courseID_reversed = []
    connCourseID_reversed = []

    for val in connCourseID:
        courseID_reversed.append(val)

    for val in courseID:
        connCourseID_reversed.append(val)


    # this section retrieves the station sign and station index of each courseID in the courseID list from the timetable and adds it to a list
    stationIndex = []
    stationSign = []

    for i in range(len(courseID)):
        section = df_Timetable[df_Timetable['courseID'] == courseID[i]]
        section.reset_index(drop=True, inplace=True)
        if section.empty:
            stationIndex.append('could not find courseID ' + courseID[i] + ' in timetable')
            stationSign.append('could not find courseID ' + courseID[i] + ' in timetable')
        else:
            stationIndex.append(section.loc[len(section)-1, 'stationIndex'])
            stationSign.append(section.loc[len(section)-1, 'stationSign'])

    # this section makes a list of station signs and station indixes for the reversed connections courseID list
    stationIndex_reversed = []
    stationSign_reversed = []

    for i in range(len(courseID_reversed)):
        section = df_Timetable[df_Timetable['courseID'] == courseID_reversed[i]]
        section.reset_index(drop=True, inplace=True)
        if section.empty:
            stationIndex_reversed.append('could not find courseID ' + courseID_reversed[i] + ' in timetable')
            stationSign_reversed.append('could not find courseID ' + courseID_reversed[i] + ' in timetable')
        else:
            stationIndex_reversed.append(section.loc[0, 'stationIndex'])
            stationSign_reversed.append(section.loc[0, 'stationSign'])
    

    # this section identifies the connChangeTime for each connection from the connections sheet based on if it is a turnback connection or through connection as well as its config listed in the connection sheet
    connChangeTime = []

    for row_index in range(len(df_connectionsInput)):
        for col_index in range(8, len(df_connectionsInput.columns), 2):
            if str(df_connectionsInput.iloc[row_index, col_index+1]) != 'nan':
                config = str(df_connectionsInput.loc[row_index, 'Config']) # stores the config value based on the row of the connections sheet
                current_courseID = df_connectionsInput.iloc[row_index, col_index]
                next_courseID = df_connectionsInput.iloc[row_index, col_index+2]
                current_section = df_Timetable[df_Timetable['courseID'] == current_courseID] # this is a smaller dataframe that consists of the rows of the timetable where the courseID is the same as the current courseID in the connection sheet that is being looked at
                current_section.reset_index(drop=True, inplace=True)
                next_section = df_Timetable[df_Timetable['courseID'] == next_courseID] # this is a smaller dataframe that consists of the rows of the timetable where the courseID is the same as the next courseID in the connection sheet that is being looked at
                next_section.reset_index(drop=True, inplace=True)
                if current_section.empty: # this is to make sure that the courseID from the connections sheet is listed in the timetable and if it isn't then the user knows what courseID is missing
                    connChangeTime.append('could not find courseID ' + current_courseID + ' in timetable')
                elif next_section.empty:
                    connChangeTime.append('could not find courseID ' + next_courseID + ' in timetable')
                else:
                    if current_courseID.find('VIA') < 0: # this checks that it is NOT a via trip connection
                        for row in range(df_connection_time_input[df_connection_time_input.columns[0]].count()): # this finds the origin and destination corridor and corridor position for each train based on the first and last stations of each train
                            if df_connection_time_input.iloc[row, 0] == current_section.loc[0,'stationSign']: # current train origin
                                current_train_origin_corridor = df_connection_time_input.iloc[row, 1]
                                current_train_origin_corridor_position = int(df_connection_time_input.iloc[row, 2])
                            if df_connection_time_input.iloc[row, 0] == current_section.loc[len(current_section)-1,'stationSign']: # current train destination
                                current_train_destination_corridor = df_connection_time_input.iloc[row, 1]
                                current_train_destination_corridor_position = int(df_connection_time_input.iloc[row, 2])
                            if df_connection_time_input.iloc[row, 0] == next_section.loc[0,'stationSign']: # next train origin
                                next_train_origin_corridor = df_connection_time_input.iloc[row, 1]
                                next_train_origin_corridor_position = int(df_connection_time_input.iloc[row, 2])
                            if df_connection_time_input.iloc[row, 0] == next_section.loc[len(next_section)-1,'stationSign']: # next train destination
                                next_train_destination_corridor = df_connection_time_input.iloc[row, 1]
                                next_train_destination_corridor_position = int(df_connection_time_input.iloc[row, 2])

                        # this finds the direction of the origin and destination corridor of each train in the connection
                        for row in range(df_connection_time_input[df_connection_time_input.columns[3]].count()): 
                            if df_connection_time_input.iloc[row, 3] == current_train_origin_corridor: # current train origin 
                                current_train_origin_corridor_direction = df_connection_time_input.iloc[row, 4]
                            if df_connection_time_input.iloc[row, 3] == current_train_destination_corridor: # current train destination
                                current_train_destination_corridor_direction = df_connection_time_input.iloc[row, 4]
                            if df_connection_time_input.iloc[row, 3] == next_train_origin_corridor: # next train origin
                                next_train_origin_corridor_direction = df_connection_time_input.iloc[row, 4]
                            if df_connection_time_input.iloc[row, 3] == next_train_destination_corridor: # next train destination
                                next_train_destination_corridor_direction = df_connection_time_input.iloc[row, 4]
                        

                        # if one of the trains starts or ends at union, set the corridor direction to be the non Union value, if both are Union, note it as an error
                        if current_train_origin_corridor_direction == "Union" and current_train_destination_corridor_direction != "Union":
                            current_train_corridor_direction = current_train_destination_corridor_direction
                        elif current_train_destination_corridor_direction == "Union" and current_train_origin_corridor_direction != "Union":
                            current_train_corridor_direction = current_train_origin_corridor_direction
                        elif current_train_origin_corridor_direction == "Union" and current_train_destination_corridor_direction == "Union":
                            print("CourseID: " + current_courseID + " starts and ends at Union Station, errors may occur")
                        else:
                            current_train_corridor_direction = current_train_destination_corridor_direction

                        if next_train_origin_corridor_direction == "Union" and next_train_destination_corridor_direction != "Union":
                            next_train_corridor_direction = next_train_destination_corridor_direction
                        elif next_train_destination_corridor_direction == "Union" and next_train_origin_corridor_direction != "Union":
                            next_train_corridor_direction = next_train_origin_corridor_direction
                        elif next_train_origin_corridor_direction == "Union" and next_train_destination_corridor_direction == "Union":
                            print("CourseID: " + next_courseID + " starts and ends at Union Station, errors may occur")
                        else:
                            next_train_corridor_direction = next_train_origin_corridor_direction


                        # this means the corridor positions will be negative and get smaller the further away from union
                        if current_train_corridor_direction == "east": 
                            if current_train_origin_corridor_position - current_train_destination_corridor_position < 0: # this determines if the current train is inbound or outbound based on the origin and destination corridor positions
                                current_train_direction = 'IB'
                            elif current_train_origin_corridor_position - current_train_destination_corridor_position > 0:
                                current_train_direction = 'OB'
                        elif current_train_corridor_direction == "west": # this means the corridor positions will be positive and get larger the further away from union
                            if current_train_origin_corridor_position - current_train_destination_corridor_position > 0: # this determines if the current train is inbound or outbound based on the origin and destination corridor positions
                                current_train_direction = 'IB'
                            elif current_train_origin_corridor_position - current_train_destination_corridor_position < 0:
                                current_train_direction = 'OB'

                        # this means the corridor positions will be negative and get smaller the further away from union
                        if next_train_corridor_direction == "east": 
                            if next_train_origin_corridor_position - next_train_destination_corridor_position < 0: # this determines if the next train is inbound or outbound based on the origin and destination corridor positions
                                next_train_direction = 'IB'
                            elif next_train_origin_corridor_position - next_train_destination_corridor_position > 0:
                                next_train_direction = 'OB'
                        elif next_train_corridor_direction == "west": # this means the corridor positions will be positive and get larger the further away from union
                            if next_train_origin_corridor_position - next_train_destination_corridor_position > 0: # this determines if the next train is inbound or outbound based on the origin and destination corridor positions
                                next_train_direction = 'IB'
                            elif next_train_origin_corridor_position - next_train_destination_corridor_position < 0:
                                next_train_direction = 'OB'

                        # identifying if it is a through or turnback connection and appending the specific connection change time 
                        if (current_train_corridor_direction == next_train_corridor_direction and current_train_direction == next_train_direction) or (current_train_corridor_direction != next_train_corridor_direction and current_train_direction != next_train_direction): # through connection
                            connChangeTime.append(str(df_connection_time_input.iloc[3, 7])) # this should be 00:05:00 in the input file
                        elif current_train_corridor_direction == next_train_corridor_direction and current_train_direction != next_train_direction: # turnback connection
                            config_found = False
                            for row in range(df_connection_time_input[df_connection_time_input.columns[6]].count()): 
                                if df_connection_time_input.iloc[row, 6] == config:
                                    connChangeTime.append(str(df_connection_time_input.iloc[row, 7]))
                                    config_found = True
                                    break
                            if config_found == False:
                                connChangeTime.append('unknown configuration ' + config)

                    else: # this section is for if it IS a VIA trip connection
                        for row in range(df_connection_time_input[df_connection_time_input.columns[8]].count()): # this gets the direction of the current and next train in the connection from the input file
                            if df_connection_time_input.iloc[row, 8] == current_courseID:
                                current_VIA_train_direction = df_connection_time_input.iloc[row, 9]
                            if df_connection_time_input.iloc[row, 8] == next_courseID:
                                next_VIA_train_direction = df_connection_time_input.iloc[row, 9]
                        
                        if df_connection_time_input.iloc[0, 10] == 'Initial': # Intercity Dwell Times (Initial Train Service Phase)
                            if (current_courseID == "VIA1" and next_courseID == "VIA-E1") or (current_courseID == "VIA-E1" and next_courseID == "VIA1"): # VIA Rail Train #1
                                connChangeTime.append(str(df_connection_time_input.iloc[3, 12])) # this should be 00:60:00 in the input file
                            elif (current_courseID == "VIA2" and next_courseID == "VIA-E2") or (current_courseID == "VIA-E2" and next_courseID == "VIA2"): # VIA Rail Train #2
                                connChangeTime.append(str(df_connection_time_input.iloc[4, 12])) # this should be 00:40:00 in the input file
                            elif (current_VIA_train_direction == "VIAE IB" or current_VIA_train_direction == "VIAE OB") and next_VIA_train_direction == "VIA OB": # Non-revenue Train Journey to scheduled outbound departure
                                connChangeTime.append(str(df_connection_time_input.iloc[0, 12])) # this should be 00:20:00 in the input file
                            elif current_VIA_train_direction == "VIA IB" and (next_VIA_train_direction == "VIAE IB" or next_VIA_train_direction == "VIAE OB"): # Schedule inbound arrival to Non-revenue Train Journey
                                connChangeTime.append(str(df_connection_time_input.iloc[1, 12])) # this should be 00:10:00 in the input file
                            elif current_VIA_train_direction == "VIA IB" and  next_VIA_train_direction == "VIA OB": # Scheduled inbound arrival to scheduled outbound departure
                                connChangeTime.append(str(df_connection_time_input.iloc[2, 12])) # this should be 00:40:00 in the input file
                        
                        elif df_connection_time_input.iloc[0, 10] == 'Early': # Intercity Dwell Times (Early Train Service Phase)
                            if (current_courseID == "VIA1" and next_courseID == "VIA-E1") or (current_courseID == "VIA-E1" and next_courseID == "VIA1"): # VIA Rail Train #1
                                connChangeTime.append(str(df_connection_time_input.iloc[3, 13])) # this should be 00:60:00 in the input file
                            elif (current_courseID == "VIA2" and next_courseID == "VIA-E2") or (current_courseID == "VIA-E2" and next_courseID == "VIA2"): # VIA Rail Train #2
                                connChangeTime.append(str(df_connection_time_input.iloc[4, 13])) # this should be 00:40:00 in the input file
                            elif (current_VIA_train_direction == "VIAE IB" or current_VIA_train_direction == "VIAE OB") and next_VIA_train_direction == "VIA OB": # Non-revenue Train Journey to scheduled outbound departure
                                connChangeTime.append(str(df_connection_time_input.iloc[0, 13])) # this should be 00:15:00 in the input file
                            elif current_VIA_train_direction == "VIA IB" and (next_VIA_train_direction == "VIAE IB" or next_VIA_train_direction == "VIAE OB"): # Schedule inbound arrival to Non-revenue Train Journey
                                connChangeTime.append(str(df_connection_time_input.iloc[1, 13])) # this should be 00:10:00 in the input file
                            elif current_VIA_train_direction == "VIA IB" and  next_VIA_train_direction == "VIA OB": # Scheduled inbound arrival to scheduled outbound departure
                                connChangeTime.append(str(df_connection_time_input.iloc[2, 13])) # this should be 00:26:00 in the input file
            else:
                break

    connChangeTime_reversed = []

    for row_index in range(len(df_connectionsInput)):
        for col_index in range(10, len(df_connectionsInput.columns), 2):
            if str(df_connectionsInput.iloc[row_index, col_index]) != 'nan':
                config = str(df_connectionsInput.loc[row_index, 'Config'])
                current_courseID = df_connectionsInput.iloc[row_index, col_index]
                last_courseID = df_connectionsInput.iloc[row_index, col_index-2]
                current_section = df_Timetable[df_Timetable['courseID'] == current_courseID]
                current_section.reset_index(drop=True, inplace=True)
                last_section = df_Timetable[df_Timetable['courseID'] == last_courseID]
                last_section.reset_index(drop=True, inplace=True)
                if current_section.empty: # this is to make sure that the courseID from the connections sheet is listed in the timetable and if it isn't then the user knows what courseID is missing
                    connChangeTime_reversed.append('could not find courseID ' + current_courseID + ' in timetable')
                elif last_section.empty:
                    connChangeTime_reversed.append('could not find courseID ' + last_courseID + ' in timetable')
                else:
                    if current_courseID.find('VIA') < 0: # this checks that it is NOT a via trip connection
                        for row in range(df_connection_time_input[df_connection_time_input.columns[0]].count()): # this finds the origin and destination corridor and corridor position for each train based on the first and last stations of each train
                            if df_connection_time_input.iloc[row, 0] == current_section.loc[0,'stationSign']: # current train origin
                                current_train_origin_corridor = df_connection_time_input.iloc[row, 1]
                                current_train_origin_corridor_position = int(df_connection_time_input.iloc[row, 2])
                            if df_connection_time_input.iloc[row, 0] == current_section.loc[len(current_section)-1,'stationSign']: # current train destination
                                current_train_destination_corridor = df_connection_time_input.iloc[row, 1]
                                current_train_destination_corridor_position = int(df_connection_time_input.iloc[row, 2])
                            if df_connection_time_input.iloc[row, 0] == last_section.loc[0,'stationSign']: # last train origin
                                last_train_origin_corridor = df_connection_time_input.iloc[row, 1]
                                last_train_origin_corridor_position = int(df_connection_time_input.iloc[row, 2])
                            if df_connection_time_input.iloc[row, 0] == last_section.loc[len(last_section)-1,'stationSign']: # last train destination
                                last_train_destination_corridor = df_connection_time_input.iloc[row, 1]
                                last_train_destination_corridor_position = int(df_connection_time_input.iloc[row, 2])

                        # this finds the direction of the origin and destination corridor of each train in the connection
                        for row in range(df_connection_time_input[df_connection_time_input.columns[3]].count()): 
                            if df_connection_time_input.iloc[row, 3] == current_train_origin_corridor: # current train origin 
                                current_train_origin_corridor_direction = df_connection_time_input.iloc[row, 4]
                            if df_connection_time_input.iloc[row, 3] == current_train_destination_corridor: # current train destination
                                current_train_destination_corridor_direction = df_connection_time_input.iloc[row, 4]
                            if df_connection_time_input.iloc[row, 3] == last_train_origin_corridor: # last train origin
                                last_train_origin_corridor_direction = df_connection_time_input.iloc[row, 4]
                            if df_connection_time_input.iloc[row, 3] == last_train_destination_corridor: # last train destination
                                last_train_destination_corridor_direction = df_connection_time_input.iloc[row, 4]
                        

                        # if one of the trains starts or ends at union, set the corridor direction to be the non Union value, if both are Union, note it as an error
                        if current_train_origin_corridor_direction == "Union" and current_train_destination_corridor_direction != "Union":
                            current_train_corridor_direction = current_train_destination_corridor_direction
                        elif current_train_destination_corridor_direction == "Union" and current_train_origin_corridor_direction != "Union":
                            current_train_corridor_direction = current_train_origin_corridor_direction
                        elif current_train_origin_corridor_direction == "Union" and current_train_destination_corridor_direction == "Union":
                            print("CourseID: " + current_courseID + " starts and ends at Union Station, errors may occur")
                        else:
                            current_train_corridor_direction = current_train_destination_corridor_direction

                        if last_train_origin_corridor_direction == "Union" and last_train_destination_corridor_direction != "Union":
                            last_train_corridor_direction = last_train_destination_corridor_direction
                        elif last_train_destination_corridor_direction == "Union" and last_train_origin_corridor_direction != "Union":
                            last_train_corridor_direction = last_train_origin_corridor_direction
                        elif last_train_origin_corridor_direction == "Union" and last_train_destination_corridor_direction == "Union":
                            print("CourseID: " + last_courseID + " starts and ends at Union Station, errors may occur")
                        else:
                            last_train_corridor_direction = last_train_origin_corridor_direction


                        # this means the corridor positions will be negative and get smaller the further away from union
                        if current_train_corridor_direction == "east": 
                            if current_train_origin_corridor_position - current_train_destination_corridor_position < 0: # this determines if the current train is inbound or outbound based on the origin and destination corridor positions
                                current_train_direction = 'IB'
                            elif current_train_origin_corridor_position - current_train_destination_corridor_position > 0:
                                current_train_direction = 'OB'
                        elif current_train_corridor_direction == "west": # this means the corridor positions will be positive and get larger the further away from union
                            if current_train_origin_corridor_position - current_train_destination_corridor_position > 0: # this determines if the current train is inbound or outbound based on the origin and destination corridor positions
                                current_train_direction = 'IB'
                            elif current_train_origin_corridor_position - current_train_destination_corridor_position < 0:
                                current_train_direction = 'OB'

                        # this means the corridor positions will be negative and get smaller the further away from union
                        if last_train_corridor_direction == "east": 
                            if last_train_origin_corridor_position - last_train_destination_corridor_position < 0: # this determines if the next train is inbound or outbound based on the origin and destination corridor positions
                                last_train_direction = 'IB'
                            elif last_train_origin_corridor_position - last_train_destination_corridor_position > 0:
                                last_train_direction = 'OB'
                        elif last_train_corridor_direction == "west": # this means the corridor positions will be positive and get larger the further away from union
                            if last_train_origin_corridor_position - last_train_destination_corridor_position > 0: # this determines if the last train is inbound or outbound based on the origin and destination corridor positions
                                last_train_direction = 'IB'
                            elif last_train_origin_corridor_position - last_train_destination_corridor_position < 0:
                                last_train_direction = 'OB'

                        # identifying if it is a through or turnback connection and appending the specific connection change time 
                        if (current_train_corridor_direction == last_train_corridor_direction and current_train_direction == last_train_direction) or (current_train_corridor_direction != last_train_corridor_direction and current_train_direction != last_train_direction): # through connection
                            connChangeTime_reversed.append(str(df_connection_time_input.iloc[3, 7])) # this should be 00:05:00 in the input file
                        elif current_train_corridor_direction == last_train_corridor_direction and current_train_direction != last_train_direction: # turnback connection
                            config_found = False
                            for row in range(df_connection_time_input[df_connection_time_input.columns[6]].count()): 
                                if df_connection_time_input.iloc[row, 6] == config:
                                    connChangeTime_reversed.append(str(df_connection_time_input.iloc[row, 7]))
                                    config_found = True
                                    break
                            if config_found == False:
                                connChangeTime_reversed.append('unknown configuration ' + config)

                    else: # this section is for if it IS a VIA trip connection
                        for row in range(df_connection_time_input[df_connection_time_input.columns[8]].count()): # this gets the direction of the current and last train in the connection from the input file
                            if df_connection_time_input.iloc[row, 8] == current_courseID:
                                current_VIA_train_direction = df_connection_time_input.iloc[row, 9]
                            if df_connection_time_input.iloc[row, 8] == last_courseID:
                                next_VIA_train_direction = df_connection_time_input.iloc[row, 9]
                        
                        if df_connection_time_input.iloc[0, 10] == 'Initial': # Intercity Dwell Times (Initial Train Service Phase)
                            if (current_courseID == "VIA1" and last_courseID == "VIA-E1") or (current_courseID == "VIA-E1" and last_courseID == "VIA1"): # VIA Rail Train #1
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[3, 12])) # this should be 00:60:00 in the input file
                            elif (current_courseID == "VIA2" and last_courseID == "VIA-E2") or (current_courseID == "VIA-E2" and last_courseID == "VIA2"): # VIA Rail Train #2
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[4, 12])) # this should be 00:40:00 in the input file
                            elif (current_VIA_train_direction == "VIAE IB" or current_VIA_train_direction == "VIAE OB") and next_VIA_train_direction == "VIA OB": # Non-revenue Train Journey to scheduled outbound departure
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[0, 12])) # this should be 00:20:00 in the input file
                            elif current_VIA_train_direction == "VIA IB" and (next_VIA_train_direction == "VIAE IB" or next_VIA_train_direction == "VIAE OB"): # Schedule inbound arrival to Non-revenue Train Journey
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[1, 12])) # this should be 00:10:00 in the input file
                            elif current_VIA_train_direction == "VIA IB" and  next_VIA_train_direction == "VIA OB": # Scheduled inbound arrival to scheduled outbound departure
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[2, 12])) # this should be 00:40:00 in the input file
                        
                        elif df_connection_time_input.iloc[0, 10] == 'Early': # Intercity Dwell Times (Early Train Service Phase)
                            if (current_courseID == "VIA1" and last_courseID == "VIA-E1") or (current_courseID == "VIA-E1" and last_courseID == "VIA1"): # VIA Rail Train #1
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[3, 13])) # this should be 00:60:00 in the input file
                            elif (current_courseID == "VIA2" and last_courseID == "VIA-E2") or (current_courseID == "VIA-E2" and last_courseID == "VIA2"): # VIA Rail Train #2
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[4, 13])) # this should be 00:40:00 in the input file
                            elif (current_VIA_train_direction == "VIAE IB" or current_VIA_train_direction == "VIAE OB") and next_VIA_train_direction == "VIA OB": # Non-revenue Train Journey to scheduled outbound departure
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[0, 13])) # this should be 00:15:00 in the input file
                            elif current_VIA_train_direction == "VIA IB" and (next_VIA_train_direction == "VIAE IB" or next_VIA_train_direction == "VIAE OB"): # Schedule inbound arrival to Non-revenue Train Journey
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[1, 13])) # this should be 00:10:00 in the input file
                            elif current_VIA_train_direction == "VIA IB" and  next_VIA_train_direction == "VIA OB": # Scheduled inbound arrival to scheduled outbound departure
                                connChangeTime_reversed.append(str(df_connection_time_input.iloc[2, 13])) # this should be 00:26:00 in the input file
            else:
                break


    # the connection type is just a 2 for regukar connections and a 0 for reversed connections
    connectionType = []

    for i in range(len(courseID)):
        connectionType.append('2')

    connectionType_reversed = []

    for i in range(len(courseID_reversed)):
        connectionType_reversed.append('0')


    # adding the reversed connections info to the end of each list
    courseID += courseID_reversed
    connCourseID += connCourseID_reversed
    stationIndex += stationIndex_reversed
    stationSign += stationSign_reversed
    connTimeType = [0]*len(courseID)
    connChangeTime += connChangeTime_reversed
    connMaxChangeTime = ['HH:MM:SS']*len(courseID)
    connectionType += connectionType_reversed
    

    # assigning each column of the output dataframe with each list of values
    df_connectionsOutput['// courseID '] = courseID
    df_connectionsOutput['connCourseID '] = connCourseID
    df_connectionsOutput['stationIndex '] = stationIndex
    df_connectionsOutput['connStationSign '] = stationSign
    df_connectionsOutput['connTimeType '] = connTimeType
    df_connectionsOutput['connChangeTime '] = connChangeTime
    df_connectionsOutput['connMaxChangeTime '] = connMaxChangeTime
    df_connectionsOutput['connectionType'] = connectionType

    df_connectionsOutput.sort_values(by = ['// courseID ', 'stationIndex '], inplace=True)
    df_connectionsOutput.reset_index(drop=True, inplace=True)

    return df_connectionsOutput


def main():

    # ask user to input connection sheet file
    connections_sheet_input = get_filename_from_user('Select the connections sheet you would like to use (must be csv format)')
    df_connectionsSheetInput = pd.read_csv(connections_sheet_input, header = 5, dtype = str)
    df_connectionsSheetInput = connections_sheet_input_management(df_connectionsSheetInput)

    # ask user to input timetable
    timetable_input = get_filename_from_user('Select the timetable you would like to use (must be csv format)')
    df_Timetable = pd.read_csv(timetable_input, header = 11, dtype = str)
    df_Timetable = timetable_management(df_Timetable)
    
    # ask user to input connection time input file
    connection_time_input = get_filename_from_user('Select the connection time input file you would like to use (must be xlsx format)')
    df_connectionTimeInput = pd.read_excel(connection_time_input, dtype = str)

    # create connections table output dataframe
    df_connectionsOutput = pd.DataFrame(columns=['// courseID ', 'connCourseID ', 'stationIndex ', 'connStationSign ', 'connTimeType ', 'connChangeTime ', 'connMaxChangeTime ', 'connectionType'])

    # run functions for ceating connections table
    df_connectionsOutput = createConnectionsTableOutput(df_connectionsOutput,df_Timetable,df_connectionsSheetInput,df_connectionTimeInput)
    
    # output connections table to excel workbook
    with pd.ExcelWriter('Connections Table.xlsx') as writer:
        df_connectionsOutput.to_excel(writer, startrow = 4, header = False, index = False)
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        column_list = df_connectionsOutput.columns
        for idx, val in enumerate(column_list):
            worksheet.write(2, idx, val)   
        worksheet.write(0, 0, "//")
        worksheet.write(1, 0, "// Connections:")
        worksheet.write(3, 0, "//")

    print("done")

    return

# ================ MAIN ================
if __name__ == "__main__":

    start_time = time.time() # store start time
    
    main()

    time_passed = time.time()-start_time
    minutes = math.trunc(time_passed/60)
    seconds = time_passed - (math.trunc(time_passed/60))*60
    print("Process finished ---", minutes, "minutes and", round(seconds, 2), "seconds ---") # print end time
