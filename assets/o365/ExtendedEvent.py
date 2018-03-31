"""
Outlook Batch Scheduler - ExtendedEvent

Extends o365 python lib for use with shared calendars

pre-reqs: Python 3.x, requests, o365, tzlocal, pytz
Last modified: March 2018
By: Dave Bunten

License: MIT - see license.txt
"""

import logging
import requests
import urllib
import time
import json
from O365 import Event

logging = logging.getLogger(__name__)

class ExtendedEvent(Event):

    #takes a calendar ID
    create_url = "https://outlook.office365.com/api/v1.0/users/{0}/calendars/{1}/events"

    def setRecurrence(self, event_type, event_interval, event_daysOfWeek, start_date, end_date):
        """
            "Recurrence": {
                "Pattern": {
                    "Type": "RelativeYearly",
                    "Interval": 1,
                    "Month": 6,
                    "DayOfMonth": 0,
                    "DaysOfWeek": ["Wednesday"],
                    "FirstDayOfWeek": "Sunday",
                    "Index": "Third"
                },
                "Range": {
                    "Type": "EndDate",
                    "StartDate": "2017-06-21",
                    "EndDate": "2020-07-08",
                    "RecurrenceTimeZone": "UTC",
                    "NumberOfOccurrences": 0
                }
            }
        """
        self.json["Recurrence"] = {}
        self.json["Recurrence"]["Pattern"] = {}
        self.json["Recurrence"]["Range"] = {}
        self.json["Recurrence"]["Pattern"]["Type"] = event_type
        self.json["Recurrence"]["Pattern"]["Interval"] = event_interval
        self.json["Recurrence"]["Pattern"]["DaysOfWeek"] = event_daysOfWeek
        self.json["Recurrence"]["Pattern"]["FirstDayOfWeek"] = "Sunday"
        self.json["Recurrence"]["Range"]["Type"] = "EndDate"
        self.json["Recurrence"]["Range"]["StartDate"] = start_date
        self.json["Recurrence"]["Range"]["EndDate"] = end_date

    def create(self,calendar=None):
        '''
        this method creates an event on the calender passed.

        IMPORTANT: It returns that event now created in the calendar, if you wish
        to make any changes to this event after you make it, use the returned value
        and not this particular event any further.
        
        calendar -- a calendar class onto which you want this event to be created. If this is left
        empty then the event's default calendar, specified at instancing, will be used. If no 
        default is specified, then the event cannot be created.

        Note: Extended to include the calendar owner information when making the post request
        '''
        if not self.auth:
            logging.debug('failed authentication check when creating event.')
            return False

        if calendar:
            calId = calendar.calendarId
            self.calendar = calendar
            logging.debug('sent to passed calendar.')
        elif self.calendar:
            calId = self.calendar.calendarId
            logging.debug('sent to default calendar.')
        else:
            logging.debug('no valid calendar to upload to.')
            return False

        headers = {'Content-type': 'application/json', 'Accept': 'application/json'}

        logging.debug('creating json for request.')
        data = json.dumps(self.json)

        response = None
        try:
            logging.debug('sending post request now')

            #modified to include the calendar owner in the post request
            response = requests.post(self.create_url.format(self.calendar.getOwner(),calId),data,headers=headers,auth=self.auth,verify=self.verify)
            logging.debug('sent post request.')
            if response.status_code > 399:
                logging.error("Invalid response code [{}], response text: \n{}".format(response.status_code, response.text))
                return False
        except Exception as e:
            if response:
                logging.debug('response to event creation: %s',str(response))
            else:
                logging.error('No response, something is very wrong with create: %s',str(e))
            return False

        logging.debug('response to event creation: %s',str(response))
        return Event(response.json(),self.auth,calendar)


    def setStart(self, start_datetime, start_timezone):
        '''
        sets event start time.
        
        Argument:
            val - this argument can be passed in three different ways. You can pass it in as a int
            or float, in which case the assumption is that it's seconds since Unix Epoch. You can
            pass it in as a struct_time. Or you can pass in a string. The string must be formated
            in the json style, which is %Y-%m-%dT%H:%M:%SZ. If you stray from that in your string
            you will break the library.


        Note: extended to include start timezone
        '''

        self.json["Start"] = {}

        self.json["StartTimeZone"] = start_timezone

        if isinstance(start_datetime,time.struct_time):
            self.json["Start"] = time.strftime(self.time_string,start_datetime)
        elif isinstance(start_datetime,int):
            self.json["Start"] = time.strftime(self.time_string,time.gmtime(start_datetime))
        elif isinstance(start_datetime,float):
            self.json["Start"] = time.strftime(self.time_string,time.gmtime(start_datetime))
        else:
            #this last one assumes you know how to format the time string. if it brakes, check
            #your time string!
            self.json["Start"] = start_datetime

    def setEnd(self, end_datetime, end_timezone):
        '''
        sets event end time.
        
        Argument:
            val - this argument can be passed in three different ways. You can pass it in as a int
            or float, in which case the assumption is that it's seconds since Unix Epoch. You can
            pass it in as a struct_time. Or you can pass in a string. The string must be formated
            in the json style, which is %Y-%m-%dT%H:%M:%SZ. If you stray from that in your string
            you will break the library.

        Note: extended to include end timezone
        '''

        self.json["End"] = {}

        self.json["EndTimeZone"] = end_timezone

        if isinstance(end_datetime,time.struct_time):
            self.json["End"] = time.strftime(self.time_string,end_datetime)
        elif isinstance(end_datetime,int):
            self.json["End"] = time.strftime(self.time_string,time.gmtime(end_datetime))
        elif isinstance(end_datetime,float):
            self.json["End"] = time.strftime(self.time_string,time.gmtime(end_datetime))
        else:
            #this last one assumes you know how to format the time string. if it brakes, check
            #your time string!
            self.json["End"] = end_datetime
