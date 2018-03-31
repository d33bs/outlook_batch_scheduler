"""
Outlook Batch Scheduler

Uses single o365 account with shared access to many o365 calendars to schedule
events in batch via a standardized spreadsheet format.

pre-reqs: Python 3.x, requests, o365, tzlocal, pytz
Last modified: March 2018
By: Dave Bunten

License: MIT - see license.txt
"""

import os
import sys
import logging
import argparse
import time
import datetime
import json
import requests
import urllib
import csv
import tzlocal
from pytz import timezone
import assets.o365.ExtendedSchedule as ExtendedSchedule
import assets.o365.ExtendedCalendar as ExtendedCalendar
import assets.o365.ExtendedEvent as ExtendedEvent

def run_batch_outlook_calendar_scheduling(run_path):
    """
    Main logic for performing batch outlook scheduling using
    spreadsheet provided from configuration.

    params:
        run_path: current run path for script for use in finding config file
    """

    config_file = open(run_path+"/config/config.json")
    config_data = json.load(config_file)
    
    csv_data = read_csv_data_from_filepath(config_data["schedule_csv_filepath"])

    auth = (config_data["username"],config_data["password"])

    schedule = ExtendedSchedule.ExtendedSchedule(auth)

    for event in csv_data:
        #collect calendar information
        try:
            result = schedule.getSharedCalendars(event["Calendar"])
        except Exception as e:
            logging.error('Failed.'+str(e))

        local_start_datetime = datetime.datetime.strptime(event["Start Date"] + "T" + event["Start Time"], "%m/%d/%YT%I:%M %p")
        local_end_datetime = datetime.datetime.strptime(event["Start Date"] + "T" + event["End Time"], "%m/%d/%YT%I:%M %p")
        recurrence_start_datetime = datetime.datetime.strptime(event["Start Date"], "%m/%d/%Y")
        recurrence_end_datetime = datetime.datetime.strptime(event["End Date"], "%m/%d/%Y")

        #"fix" the time based on DST standard for time localization
        if (tzlocal.get_localzone().localize(datetime.datetime.now()).dst() != datetime.timedelta(0) and
            tzlocal.get_localzone().localize(local_start_datetime).dst() == datetime.timedelta(0)):
            local_start_datetime = local_start_datetime + datetime.timedelta(hours=1)
            local_end_datetime = local_end_datetime + datetime.timedelta(hours=1)

        outlook_start_datetime = datetime.datetime.strftime(convert_datetime_local_to_utc(local_start_datetime), "%Y-%m-%dT%H:%M:%SZ")
        outlook_end_datetime = datetime.datetime.strftime(convert_datetime_local_to_utc(local_end_datetime), "%Y-%m-%dT%H:%M:%SZ")
        outlook_recurrence_start_date = recurrence_start_datetime.strftime("%Y-%m-%dT%H:%M:%SZ")
        outlook_recurrence_end_date = recurrence_end_datetime.strftime("%Y-%m-%dT%H:%M:%SZ")

        #event_type, event_interval, event_daysOfWeek, start_date, end_date):
        for cal in schedule.calendars:
            if cal.getOwner() == event["Calendar"]:      
                e = ExtendedEvent.ExtendedEvent(auth=auth, cal=cal)
                e.setSubject(event["Event Name"])
                e.setBody(event["Description"])
                e.setLocation(event["Location"])
                e.setStart(outlook_start_datetime, config_data["o365_timezone"])
                e.setEnd(outlook_end_datetime, config_data["o365_timezone"])
                e.setRecurrence("Weekly",1,outlook_days_of_week(event),outlook_recurrence_start_date, outlook_recurrence_end_date)
                new_e = e.create()
                continue

def read_csv_data_from_filepath(csv_file_path):
    """
    Reads csv from provided file path and returns csv data in dictionary

    params:
        run_path: current run path for script for use in finding config file

    returns:
        csv_data: dictionary of data read by the csv python lib
    """

    csv_data = []
    reader = csv.DictReader(open(csv_file_path, 'r'))
    for line in reader:
        csv_data.append(line)

    return csv_data

def convert_datetime_local_to_utc(datetime_local):
    """
    Translate local datetime to utc for use by internal Mediasite system
    
    params:
        datetime_local: local datetime object
    
    returns:
        converted utc datetime object
    """

    #find our current timezone
    local_tz = tzlocal.get_localzone()

    #find UTC times for datetimes due to Mediasite requirements
    UTC_OFFSET_TIMEDELTA = datetime.datetime.utcnow() - datetime.datetime.now() 

    return datetime_local + UTC_OFFSET_TIMEDELTA

def outlook_days_of_week(event_data):
    """
    Converts gathered event data to Outlook-API consumable weekday string

    params:
        event_data: dictionary containing event data specific to an outlook calendar occurrence

    returns:
        weekday_list: list containing days of the week for the calendar occurence in an 
            outlook-API friendly format.
    """

    weekday_list = []

    if event_data["Sun"] == "TRUE":
        weekday_list.append("Sunday")
    if event_data["Mon"] == "TRUE":
        weekday_list.append("Monday")
    if event_data["Tue"] == "TRUE":
        weekday_list.append("Tuesday")
    if event_data["Wed"] == "TRUE":
        weekday_list.append("Wednesday")
    if event_data["Thu"] == "TRUE":
        weekday_list.append("Thursday")
    if event_data["Fri"] == "TRUE":
        weekday_list.append("Friday")
    if event_data["Sat"] == "TRUE":
        weekday_list.append("Saturday")

    return weekday_list

if __name__ == "__main__":
    #gather our runpath for future use with various files
    run_path = os.path.dirname(os.path.realpath(__file__))

    #log file datetime
    current_datetime_string = '{dt.month}-{dt.day}-{dt.year}_{dt.hour}-{dt.minute}-{dt.second}'.format(dt = datetime.datetime.now())
    logfile_path = run_path+'/logs/outlook_batch_scheduler_'+current_datetime_string+'.log'

    #logger for log file
    logging_format = '%(asctime)s - %(levelname)s - %(message)s'
    logging_datefmt = '%m/%d/%Y - %I:%M:%S %p'
    logging.basicConfig(filename=logfile_path,
                        filemode='w',
                        format=logging_format,
                        datefmt=logging_datefmt,
                        level=logging.DEBUG
                        )

    #logger for console
    console = logging.StreamHandler()
    formatter = logging.Formatter(logging_format,
                                    datefmt=logging_datefmt)
    console.setFormatter(formatter)
    logging.getLogger().addHandler(console)

    #main program function
    run_batch_outlook_calendar_scheduling(run_path)