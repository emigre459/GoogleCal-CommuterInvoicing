# Commuter Invoicing from Google Calendar

## The General Concept
This project parses Google Calendar data exported as an iCalendar (ICS) file and produces data in a format that is useful for pasting into an invoice template.

## The Deets AKA Details

Specifically, using the pandas and icalendar Python packages, this project will pull data
from a file named `calendar_data.ics` in the code's working directory and generate a pandas Dataframe
that then is parsed for unique passenger names. It then outputs an Excel file with a sheet for each
passenger, named using that passenger's name. These data can then be pasted into the invoice template
of your choosing. The code also calculates how much each passenger owes for the time period specified.