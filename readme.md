# Outlook Batch Scheduler

Uses single o365 account with shared access to many o365 calendars to schedule events in batch via a standardized spreadsheet format.

This project specifically targets a single account with shared read/write access to many calendars under other account names. The user-provided csv file specifies these shared accounts by name to create the calendar events. For example:

* master_account@fakedomain.com
	* shared access to -> another_account1@fakedomain.com
	* shared access to -> another_account2@fakedomain.com
	* shared access to -> another_account3@fakedomain.com
	* ...etc...

## Prerequisites and Documentation

Before you get started, make sure to install or create the following prerequisites:

* Python 3.x: [https://www.python.org/downloads/](https://www.python.org/downloads/)
* Python Requests Library (non-native library used for HTTP requests): [http://docs.python-requests.org/en/master/](http://docs.python-requests.org/en/master/)
* python-O365: [https://github.com/Narcolapser/python-o365](https://github.com/Narcolapser/python-o365)
* pytz: [https://github.com/newvem/pytz](https://github.com/newvem/pytz)
* tzlocal: [https://github.com/regebro/tzlocal](https://github.com/regebro/tzlocal)
* A Microsoft O365 account with calendar access

## Special Note

This project makes use and extends the functionality of the Python-O365 library. The extended classes can be found under the assets/o365 folder within this repository. These modifications were made to address the use of shared calendars and recurring calendar events (functionality not originally provided within the library). Please note specific changes in each file: ExtendedCalendar.py, ExtendedSchedule.py, and ExtendedEvent.py .

## Usage

1. Ensure prerequisites outlined above are completed.
1. Fill in necessary information within config/sample_config.json and rename to project specifics
1. Remove the text "_sample" from config file
1. Run main.py with Python 3.x

## License

MIT - See license.txt

## Notice

The project is made possible by open source software. Please see the following listing for software used and respective licensing information:

* Python 3 - **PSF** [https://docs.python.org/3/license.html](https://docs.python.org/3/license.html)
* Requests - **Apache 2.0** [https://opensource.org/licenses/Apache-2.0](https://opensource.org/licenses/Apache-2.0)
* pytz - **MIT** [https://opensource.org/licenses/MIT](https://opensource.org/licenses/MIT)
* tzlocal - **MIT** [https://opensource.org/licenses/MIT](https://opensource.org/licenses/MIT)
* python-O365 - **Apache 2.0** [https://opensource.org/licenses/Apache-2.0](https://opensource.org/licenses/Apache-2.0)
