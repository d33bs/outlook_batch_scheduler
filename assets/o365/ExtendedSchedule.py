"""
Outlook Batch Scheduler - ExtendedSchedule

Extends o365 python lib for use with shared calendars

pre-reqs: Python 3.x, requests, o365, tzlocal, pytz
Last modified: March 2018
By: Dave Bunten

License: MIT - see license.txt
"""

import logging
import requests
import urllib
from O365 import Schedule
import assets.o365.ExtendedCalendar as ExtendedCalendar

logging = logging.getLogger(__name__)

class ExtendedSchedule(Schedule):
    def shared_calendar_url(self, username):
        """
        Creates a shared calendar outlook API url for use with schedules

        params:
            username: full outlook username including domain specific to shared calendar

        returns: string for shared calendar outlook API url
        """

        return "https://outlook.office365.com/api/v1.0/users/"+urllib.parse.quote_plus(username)+"/calendars"

    def getSharedCalendars(self, username):
        '''Begin the process of downloading calendar metadata.'''
        logging.debug('fetching shared calendars under name: '+username)

        #change to include shared calendar url creation
        response = requests.get(self.shared_calendar_url(username),auth=self.auth,verify=self.verify)
        logging.info('Response from O365: %s', str(response))
        for calendar in response.json()['value']:
            try:
                duplicate = False
                logging.debug('Got a calendar with Name: {0} and Id: {1}'.format(calendar['Name'],calendar['Id']))
                for i,c in enumerate(self.calendars):
                    if c.json['Id'] == calendar['Id']:
                        c.json = calendar
                        c.name = calendar['Name']
                        c.calendarId = calendar['Id']
                        duplicate = True
                        log.debug('Calendar: {0} is a duplicate',calendar['Name'])
                        break

                if not duplicate:
                    #change to include Extended Calendar
                    self.calendars.append(ExtendedCalendar.ExtendedCalendar(json=calendar,auth=self.auth,owner=username))
                    logging.debug('appended calendar: %s',calendar['Name'])

                logging.debug('Finished with calendar {0} moving on.'.format(calendar['Name']))

            except Exception as e:
                logging.info('failed to append calendar: {0}'.format(str(e)))
        
        logging.debug('all calendars retrieved and put in to the list.')
        return True