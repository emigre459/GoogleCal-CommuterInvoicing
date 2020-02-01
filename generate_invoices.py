import datetime as dt
import pandas as pd
import logging
import argparse

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

# Cameron Station Commuters calendar
CALENDAR_ID = 'j03017amddesj4vcd6c7c0m1nc@group.calendar.google.com'

MORNING_COMMUTE_TIME = dt.time(hour = 8, minute = 0)
EVENING_COMMUTE_TIME = dt.time(hour = 17, minute = 15)

#Rate (in USD) per commuter per ride
GOING_RATE = 3

# NASA garage monthly parking permit cost in USD
PARKING_PERMIT_COST = 50

LOG_FORMAT = '%(levelname)s (%(name)s) - %(message)s'

logging.basicConfig(
    format=LOG_FORMAT,
    level=logging.ERROR,
    datefmt='%m/%d/%Y %H:%M:%S')

logger = logging.getLogger(__name__)

parser = argparse.ArgumentParser(description='Queries the Cameron Station Commuters Google \
calendar and extracts passenger trips for every day in a specified time period, \
then calculates the amount owed by each passenger.')

parser.add_argument('start_date', type=str, help='String of the format "05-05-2020" (May 5, 2020).\
Indicates the start date for reporting, inclusive.')

parser.add_argument('end_date', type=str, help='String of the format "05-05-2020" (May 5, 2020).\
Indicates the end date for reporting, inclusive.')

args = vars(parser.parse_args())



def check_google_api_credentials(token_filepath='./'):
    '''
    Checks to make sure you have the credentials needed to access the 
    commuter Google Calendar. If not, opens a browser tab to let you
    login to obtain the proper credentials.
    
    This only needs to be run whenever you need to create or overwrite
    your token.pickle file.
    
    
    Parameters
    ----------
    token_filepath: str. Relative or absolute filepath to use
        when looking for the credentials contained in token.pickle.
        Default is to look in the current working directory.
    
    Returns
    -------
    API credentials and a token.pickle file 
    may be created if one wasn't already present 
    in the working directory.
    '''
    
    creds = None
    
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(token_filepath + 'token.pickle'):
        with open(token_filepath + 'token.pickle', 'rb') as token:
            creds = pickle.load(token)
            
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'google_api_credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_filepath + 'token.pickle', 'wb') as token:
            pickle.dump(creds, token)
            
    return creds


def record_maker(event_summary):
    """
    Extract info about commuter event and return in a useful format
    
    
    Parameters
    ----------
    event_summary: icalendar component summary vText describing the title of the event. Expected
                    to be of the format "Passenger #X: Passenger_Name"
         
    Returns
    -------
    3-tuple of the format (Driver Name, Passenger Name, extra info)
    """
    
    #[Driver_Name Passenger Passenger_Num, Passenger_Name Extra_Info]
    passenger_info = pd.Series(event_summary.split(":")).str.strip()
    
    # Make DataFrame with rows that represent string portions before and after colon
    # and columns split on whitespace from pre- and post-colon substrings
    separated_text = passenger_info.str.split(" ", expand=True)
    
    passenger_name = separated_text.loc[1,0]
    extra_info = separated_text.loc[1,1:].str.cat(sep = " ")
    
    # Make extra_info an empty string if None
    if not extra_info: extra_info = ''

    # If no driver named, assume it's Becky
    if separated_text.loc[0,0] == "Passenger":
        driver_name = "Becky"

    else:
        driver_name = separated_text.loc[0,0]
        
    #Check to make sure it actually found a name to match to
    if len(passenger_info) > 1 and passenger_info[1].strip() != "":
        return driver_name.title().strip(), passenger_name.title().strip(), extra_info    
           
    else:
        return None
    
    
def get_events(start_date, end_date, creds):
    """
    Returns all commuter events between start_date and end_date, inclusive.
    
    
    Parameters
    ----------
    start_date: str of the format 05-05-2020 (May 5, 2020). 
        Indicates the start date for reporting
        
    end_date: str of the format 05-05-2020 (May 5, 2020). 
        Indicates the start date for reporting
        
    creds: Google OAuth2 credentials for API access
    
    
    Returns
    -------
    pandas DataFrame with columns 
    ['Date', 'Driver Name', 'Passenger Name', 'Extra Info'] 
    """
    
    service = build('calendar', 'v3', credentials=creds)
    print(f'Getting events between {start_date} and {end_date}, inclusive.')

    start_date = dt.datetime.strptime(start_date, '%m-%d-%Y')
    start = dt.datetime(year=start_date.year, 
                month=start_date.month, 
                day=start_date.day, 
                tzinfo=dt.timezone.utc).isoformat()[:-6] + 'Z'
    
    end_date = dt.datetime.strptime(end_date, '%m-%d-%Y')
    end = dt.datetime(year=end_date.year, 
                month=end_date.month, 
                day=end_date.day+1, 
                tzinfo=dt.timezone.utc).isoformat()[:-6] + 'Z'
    
    events_result = service.events().list(calendarId=CALENDAR_ID, 
                                          timeMin=start,
                                          timeMax=end,
                                          singleEvents=True,
                                          orderBy='startTime').execute()
    
    events = events_result.get('items', [])

    if not events:
        logging.error('No events found!')
    
    dates = []
    drivers = []
    passengers = []
    extra_info = []
    
    for event in events:
        event_details = record_maker(event['summary'])
        
        
        # Only append if not None
        if event_details:
            dates.append(dt.datetime.strptime(event['start']['dateTime'][:10], 
                                          '%Y-%m-%d').strftime('%m-%d-%Y'))
            drivers.append(event_details[0])
            passengers.append(event_details[1])
            extra_info.append(event_details[2])
        
    output = pd.DataFrame({'Date': dates,
                        'Driver Name': drivers,
                        'Passenger Name': passengers,
                        'Extra Info': extra_info})
    
    # Do some cleanup
    output.replace({
        'Kc': 'KC',
        'Pg': 'Paul'
    }, inplace=True)
        
    return output


def save_data(daily_record):
    '''
    Saves the day-by-day data we've pulled together and summarizes
    it in invoice form (AKA who-owes-whom format)
    
    
    Parameters
    ----------
    daily_record. Pandas DataFrame of the format output by get_events()
    
    
    Returns
    -------
    Nothing. Saves results as an Excel file in 
    /Commuting/<current_year>/<current date as YYYY-MM-DD>_CommuterInvoices.xlsx
    '''
    
    base_path = "/Users/emigre459/Dropbox/Shared Folder - Becky and Dave/Finances/Commuting"
    filepath = os.path.join(base_path, 
                            str(dt.date.today().year), 
                            dt.date.today().isoformat() + "_CommuterInvoices.xlsx")

    writer = pd.ExcelWriter(filepath)

    #Iterate through the unique passenger names in the dataframe
    for name in daily_record["Passenger Name"].unique():
        temp_df = daily_record.groupby("Passenger Name").get_group(name)\
        [['Date', 'Driver Name']]

        #Remove the Name field, as that will already be in the sheet name
        temp_df.to_excel(writer, sheet_name = name, index = False)
        
    #How much does each rider owe??
    invoice = daily_record.groupby(["Driver Name", "Passenger Name"]).count()['Extra Info'] * GOING_RATE

    # Becky gets $25 credit for rides with KC due to 
    # him owing us half of the monthly parking permit cost

    # Did Becky ride with KC?
    if invoice['KC'].index.str.contains("Becky").sum() > 0:
        invoice['KC', 'Becky'] -= int(PARKING_PERMIT_COST / 2)

    else:
        invoice['KC', 'Becky'] = -int(PARKING_PERMIT_COST / 2)

    invoice = invoice.reset_index().rename(columns = {
        'Extra Info': 'Amount Owed'
    })

    # Save results to 'Totals' tab of spreadsheet
    invoice.to_excel(writer, sheet_name = 'Totals', index = False)
    writer.save()


def main():
    
    # After this line executes, you may need to allow the app Quickstart
    # access to your Google Account via opened browser tab
    creds = check_google_api_credentials()

    full_record = get_events(start_date = args['start_date'],
                             end_date = args['end_date'],
                             creds=creds)

    save_data(full_record)
    

if __name__ == '__main__':
    main()