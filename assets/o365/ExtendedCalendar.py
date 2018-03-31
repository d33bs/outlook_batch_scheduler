"""
Outlook Batch Scheduler - ExtendedCalendar

Extends o365 python lib for use with shared calendars

pre-reqs: Python 3.x, requests, o365, tzlocal, pytz
Last modified: March 2018
By: Dave Bunten

License: MIT - see license.txt
"""

import logging
import requests
import urllib
from O365 import Calendar

logging = logging.getLogger(__name__)

class ExtendedCalendar(Calendar):
    def __init__(self, json=None, auth=None, verify=True, owner=""):
        '''
        Wraps all the information for managing calendars.

        Note: modified to include the owner information about a calendar
        '''
        self.json = json
        self.auth = auth
        self.events = []

        if json:
            logging.debug('translating calendar information into local variables.')
            self.calendarId = json['Id']
            self.name = json['Name']
            self.owner = owner

        self.verify = verify

    def getOwner(self):
        return self.owner
