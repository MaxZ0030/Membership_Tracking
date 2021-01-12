# %%
import pandas as pd
import json
import os
import glob

pd.set_option('display.max_colwidth', 0)
#EIDfile = open("EID.txt", "r")
EID_Lookup = {}
student_list = {}
questionable_names = {}
lenient = True
valid_purchases = {}
invalid_purchases = {}



# %%
def parse_csv():
    """
    Parse the csv and excel files in their respective directories and
    return as lists of DataFrames.
    """

    # Get the full file path
    events_path = os.path.join(os.getcwd(), "data/eventst")
    # Create a list of all the csv files in the directory
    events = glob.glob(os.path.join(events_path, '*.xlsx'))
    # Finally, parse all the excel files and place in a list
    attendance_data = [pd.read_excel(event) for event in events]
    #attendanceData.reverse()
    
    #Read in the events and members sheet

    #pd.read_excel('data/IEEE Membership Database 2020-2021 - Form Responses 1.xlsx')
    member_list = pd.read_csv('data/IEEE Membership Database 2020-2021 - Form Responses 1.csv')
    
    events_list = pd.read_excel('data/IEEE Events.xlsx')

    shop_info = False#pd.read_excel('')

    shop_purchases = pd.read_excel('data/THIEEE SHOP (Responses).xlsx')

    bonus_points = False#pd.read_excel('')

    return attendance_data, member_list, events_list, shop_info, shop_purchases, bonus_points


def forgiveness(database_name,attendee_name,findex):
    """
    Returns true if the database name and the attendee name differ only by the findex amount

    """
    if(len(database_name) != len(attendee_name)):
        return False
    
    for i in range(0,len(database_name)):
        
        if(database_name[i] != attendee_name[i]):
            
            findex -= 1
            if findex < 0:
                return False
        

    return True



def check_attendance(database_name, meeting):
    """
    Return a boolean

    Checks the attendance of an EID for a given meeting.
    """
    # Iterate through list of EIDs
    for index, attendee in meeting.iterrows():
        
        attendee_name = str(attendee["First Name"].strip() + " " + attendee["Last Name"].strip()).lower()
        
        #If the name and EID located in the database matches the one from the event, return true
        #attendee["What's your EID?"].lower() == EID_Lookup.get(name.lower(), "NA") and
        if lenient:
            if forgiveness(database_name,attendee_name,1):
                return True
        elif database_name == attendee_name:
            return True
    return False    


def parse_bonus_points(attendance, member_list):
    return False

def validate_purchase(attendance, member_name, purchase_info):
    print(member_name)
    print(purchase_info)
    purchaseable = {"Cookies" : 100, "Chocolate" : 100, "Milk" : 100}
    member = student_list.get(member_name, "NA")
    if(member == "NA"):
        raise Exception
    points_needed = purchaseable[purhcase_info["What would you like to purchase?"]]


    return False


def parse_other(attendance, shop_purchases, bonus_points):
    """
    Return a pandas dataframe


    Goes through the shop and bonus points and updates the master sheet accordingly.
    """

    for aindex, member in attendance.iterrows():
        mname = member["Name:"]

        for pindex, purchaser in shop_purchases.iterrows():
                pname = str(purchaser["First Name"].strip() + " " + purchaser["Last Name"].strip()).lower()
                
                if lenient:
                    if forgiveness(mname.lower(),pname,1):
                        validate_purchase(mname, shop_purchases.loc[pindex].to_dict())                        
                elif mname.lower() == pname:
                    validate_purchase(mname, shop_purchases.loc[pindex].to_dict())                        
                




def parse_sheets(event_attendance, member_list):
    """
    Return a pandas dataframe

    Goes through all the events and tallies up the points of the students.
    """
    attendance = pd.DataFrame()

    #parse_bonus_points(attendance, member_list)


    attendance["Name:"] = member_list["First Name"].str.strip() + " " + member_list["Last Name"].str.strip()
    attendance["Total Spark Points:"] = 0
    print(attendance)
    print(member_list)
    #Go through all events in the folder and check if the person attended. If yes, add points
    for index, member in attendance.iterrows():
        #print(event_attendance)
        #print(member)
        name = str(member["Name:"].strip()).lower()
        valid = student_list.get(name, "NA")
        #Add a new student to the dictionary of student_list if not already there
        print(valid)
        if(valid == 'NA'):
            print("Added student " + name)
            student_list[name.lower()] = {'EID' : 0, 'Points' : 0}


        for event in event_attendance:
            
            eventdata = list(event.columns)
            event_name = eventdata[-1]
            #print(event_name)
            points_to_add = event.iat[0, len(eventdata)-1]
            attendance.at[index, event_name] = 0
            #If the person is in the database and at the meeting, sum up each person's spark points and add to their total


            if check_attendance(member["Name:"].lower(), event):
                    
                attendance.at[index, event_name] += points_to_add
                attendance.at[index, "Total Spark Points:"] += points_to_add

    
    #parse_shop()

    return attendance

def clean_events(events_list):
    """
    Return a pandas dataframe (Events Excel Sheet)

    Reformat the events excel sheet to remove unnecessary data.
    """
    
    events_list = events_list.drop(columns = ["Actions", 'Officer in Charge'])
    events_list = events_list[events_list["Confirmed/Tentative/Cancelled"].astype(str) != 'Tentative']
    events_list["Date"] = pd.to_datetime(events_list["Date"]).dt.date
    events_list = events_list.loc[:, ~events_list.columns.str.contains('^Unnamed')]
    return events_list

def clean_members(member_list):
    """
    Return a pandas dataframe (Member Excel Sheet)

    Reformat the events excel sheet to remove unnecessary data
    """

    member_list = member_list.drop(columns = ['The FamilIEEE System places members into a small, tight-knit group of students ("family") who attend fun activities together and participate in a friendly competition between families. Would you be interested in learning more about the FamilIEEE System?'])
    member_list = member_list.sort_values(by=['Last Name'])
    return member_list



def finished():
    print('Finished')




def createMasterSheet():
    
    # %%

    # Get the list of known EIDs

    #with open("EID.txt", 'r') as file:
    #    EID_Lookup = json.loads(file.read())  
    
    EVENTATTENDANCEDATA, MEMBERLIST, EVENTSLIST, SHOPINFO, SHOPPURCHASES, BONUSPOINTS = parse_csv()



    # Calculate spark points for each due-paying member 💰💰💰

    EVENTSLIST = clean_events(EVENTSLIST)
    MEMBERLIST = clean_members(MEMBERLIST)
    EVENTATTENDANCEDATA = parse_sheets(EVENTATTENDANCEDATA, MEMBERLIST)
    #print(student_list)
    EVENTATTENDANCEDATA = parse_other(EVENTATTENDANCEDATA, SHOPPURCHASES, BONUSPOINTS)
    

    MEMBERLIST = MEMBERLIST.drop(columns=['Timestamp'])

    writer = pd.ExcelWriter('MasterSheet.xlsx', engine='xlsxwriter')
    sheets = {'People': MEMBERLIST, 'Events': EVENTSLIST, 'Attendance and Spark Points':EVENTATTENDANCEDATA, 'Shop Info': SHOPINFO}

    #Write the data to a new Master Sheet
    for sheet_name, df in sheets.items():
            print(sheet_name)
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



    finished()

    #Save the EID List


    #with open("EID.txt", 'w') as file:
    #        file.write(json.dumps(EID_Lookup, indent=4))


if __name__ == "__main__":
    createMasterSheet()
    
