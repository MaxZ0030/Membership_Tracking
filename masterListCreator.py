# %%
import pandas as pd

import os
import glob

pd.set_option('display.max_colwidth', 0)
EIDfile = open("EID.txt", "r")
EID_Lookup = {}

# %%
def parse_csv():
    """
    Parse the csv and excel files in their respective directories and
    return as lists of DataFrames.
    """

    # Get the full file path
    events_path = os.path.join(os.getcwd(), "data/events")
    # Create a list of all the csv files in the directory
    events = glob.glob(os.path.join(events_path, '*.xlsx'))
    # Finally, parse all the excel files and place in a list
    attendanceData = [pd.read_excel(event) for event in events]
    #attendanceData.reverse()
    
    #Read in the events and members sheet

    members = pd.read_csv('data/IEEE Membership Database 2020-2021 - Form Responses 1.csv')
    
    eventsList = pd.read_excel('data/IEEE Events.xlsx')

    return attendanceData, members, eventsList


def check_attendance(name, meeting):
    """
    Checks the attendance of an EID for a given meeting.
    """
    # Iterate through list of EIDs
    for index, attendee in meeting.iterrows():
        
        attendee_name = str(attendee["First Name"].strip() + " " + attendee["Last Name"].strip())
        eid = EID_Lookup.get(attendee_name.lower(), "NA")
        #If the EID is not located in the database, add it
        if eid == 'NA':
            EID_Lookup[attendee_name.lower()] = attendee["What's your EID?"].lower()
            eid = EID_Lookup.get(attendee_name.lower(), "NA")
        #If the name and EID located in the database matches the one from the event, return true
        #attendee["What's your EID?"].lower() == EID_Lookup.get(name.lower(), "NA") and 
        if name.lower() == attendee_name.lower():
            return True
    return False    

def parse_events(event_attendance, memberlist):
    """
    Checks the attendance of an EID for a given meeting.
    """
    attendance = pd.DataFrame()
    

    attendance["Name:"] = memberlist["First Name"].str.strip() + " " + memberlist["Last Name"].str.strip()
    attendance["Total Spark Points:"] = 0
    #Go through all events in the folder and check if the person attended. If yes, add points
    for index, member in attendance.iterrows():
        for event in event_attendance:
            
            eventdata = list(event.columns)
            event_name = eventdata[-1]
            print(event_name)
            points_to_add = event.iat[0, len(eventdata)-1]
            attendance.at[index, event_name] = 0
            #If the person is in the database and at the meeting, sum up each person's spark points and add to their total
            if check_attendance(member["Name:"], event):
                    
                attendance.at[index, event_name] += points_to_add
                attendance.at[index, "Total Spark Points:"] += points_to_add

    
    
    return attendance

def parse_eventsAndMembers(eventsList, members):
    """
    Reformat the events and members excel sheets to remove unnecessary data
    """
    members = members.drop(columns = ['The FamilIEEE System places members into a small, tight-knit group of students ("family") who attend fun activities together and participate in a friendly competition between families. Would you be interested in learning more about the FamilIEEE System?'])
    members = members.sort_values(by=['Last Name'])

    eventsList = eventsList.drop(columns = ["Actions", 'Officer in Charge'])
    eventsList = eventsList[eventsList["Confirmed/Tentative/Cancelled"].astype(str) != 'Tentative']
    eventsList["Date"] = pd.to_datetime(eventsList["Date"]).dt.date
    eventsList = eventsList.loc[:, ~eventsList.columns.str.contains('^Unnamed')]
    return eventsList, members



# %%

# Get the list of known EIDs

for line in EIDfile.readlines():
    currline = str(line).strip().split(" : ")
    EID_Lookup[currline[0]] = currline[1]
EIDfile.close()

EVENTATTENDANCEDATA, MEMBERS, EVENTSLIST = parse_csv()



# Calculate spark points for each due-paying member 💰💰💰

EVENTSLIST, MEMBERS = parse_eventsAndMembers(EVENTSLIST, MEMBERS)
EVENTATTENDANCEDATA = parse_events(EVENTATTENDANCEDATA, MEMBERS)
# %%



MEMBERS = MEMBERS.drop(columns=['Timestamp'])

writer = pd.ExcelWriter('MasterSheet.xlsx', engine='xlsxwriter')
sheets = {'People': MEMBERS, 'Events': EVENTSLIST, 'Attendance and Spark Points':EVENTATTENDANCEDATA}

#Write the data to a new Master Sheet
for sheet_name, df in sheets.items():
    
        sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index = False)
        worksheet = writer.sheets[sheet_name]
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),
                len(str(series.name))
                )) + 1
            worksheet.set_column(idx, idx, max_len)

writer.save()



print("Finished")

#Save the EID List
EIDfilewriter = open("EID.txt", "w")
for key, value in EID_Lookup.items():
    EIDfilewriter.write(key + " : " + value)
EIDfilewriter.close()
