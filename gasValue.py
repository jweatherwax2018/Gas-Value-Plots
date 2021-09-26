import os
import matplotlib.pyplot as plt
#from gsheets import Sheets
from datetime import date
from apiclient import discovery
from google.oauth2 import service_account
import datetime

SHEET_ID = "14Pt05nmV8K6xQKCD2iFJ92FRAZXXVZcT2MhEJqIUyYQ"
WORKSHEETS = ["ArCO2","ArCO2MTS","N2", "CO2", "He"]
def get_data(service, range):

    request = service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=range)
    response = request.execute()
    
    #print(response)
    
    dates = []
    times = []
    pressures = []
    
    #Collecting Data from sheet
    for data in response['values'][1:]:
        if len(data) < 3:
            continue
        if len(data[1]) < 1:
            continue
        dates.append(datetime.datetime.strptime(data[0]+":"+data[1],"%Y-%m-%d:%H:%M"))
        times.append(data[1])
        pressures.append(int(data[2]))
    return (dates, pressures)
    
def parse_data(service):
    dates = []
    pressures = []
    
    #Assign Dates and Pressures for each gas cylinder
    for worksheet in WORKSHEETS:
        print(f"Downloading Data for {worksheet}...")
        d, p = get_data(service, range = f"{worksheet}!A:C")
        print(f"Finished Downloading Data for {worksheet}.")
        dates.append(d)
        pressures.append(p)
    return (dates, pressures)

def main():
    SCOPE = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

    print("Connecting...")
    # Authenticate
    creds = service_account.Credentials.from_service_account_file("C:\\Users\\Joseph\\Documents\\Research\\Gas Plots\\gasValues\\client_secret.json",scopes=SCOPE);
    service = discovery.build('sheets','v4', credentials=creds)
    
    #Plotting
    dates, pressures = parse_data(service)
    m = 0
    for i, date in enumerate(dates):
        plt.plot(date, pressures[i])
        m = max(m, max(pressures[i]))
    plt.ylim(0, m+500)
    plt.xlim(dates[0][0] + datetime.timedelta(days=-10),dates[0][-1] + datetime.timedelta(days=+30))
    plt.legend(WORKSHEETS,loc='upper right',bbox_to_anchor=(0,1.05,1.11,0))
    plt.xlabel("Date")
    plt.ylabel("Pressure (kPa)")
    plt.title('Gas Values', fontdict = {'fontsize' : 20})
    plt.show()

if __name__ == "__main__":
    main()
