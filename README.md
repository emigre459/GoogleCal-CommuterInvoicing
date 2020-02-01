# Commuter Invoicing from Google Calendar

This project parses Google Calendar data exported as an iCalendar (ICS) file and produces data in a format that is useful for pasting into an invoice template.

## Usage

The best approach is to:

1. Create a conda environment using the included `environment.yml` file
2. After activating your new conda environment, run `python generate_invoices.py start_date end_date`

## The Deets AKA Details

Specifically, using the Google API, this project will pull data
from my personal calendar *Cameron Station Commuters* and generate a pandas Dataframe
that then is parsed for unique passenger names. It then outputs an Excel file with a sheet for each
passenger, named using that passenger's name, as well as a summary tab reflecting how frequently each passenger rode with a given driver. These data can then be pasted into the invoice template
of your choosing. 