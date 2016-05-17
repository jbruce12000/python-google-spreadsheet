Google Spreadsheets API
========================
A simple Python wrapper for the Google Spreadsheeta API.
[![Build Status](https://secure.travis-ci.org/yoavaviram/python-google-spreadsheet.png?branch=master)](http://travis-ci.org/yoavaviram/python-google-spreadsheet)



Features
--------

* An object oriented interface for Worksheets
* Supports List Feed view of spreadsheet rows, represented as dictionaries
* Compatible with Google App Engine


Requirements
--------------
Before you get started, make sure you have:

* virtualenv venv
* source venv/bin/activate
* pip install -e git+git@github.com:jbruce12000/python-google-spreadsheet.git#egg=google_spreadsheet
* cd venv/
* pip install -r venv/src/google-spreadsheet/requirements.txt (that will install gdata, httplib2 and oauth2client)
* setup oauth2 for your app https://console.developers.google.com/apis/library
  * pick web app
  * set the redirect uri to http://localhost:8080/
  * grab the client_id and client_secret
* put the client_id and client_secret in a local json formatted file named ./client_secrets.json like this...

    {
    "web": {
        "client_id": "XXXXXXXXX",
        "client_secret": "XXXXXXXX",
        "redirect_uris": ["http://localhost"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://accounts.google.com/o/oauth2/token"
        }
    }

* the spec is here... https://developers.google.com/api-client-library/python/guide/aaa_client_secrets if you care
* the first time you run this, it will open a browser and have you verify it is ok for this app to access your google account.  after that, it'll be cached in a local file named ./creds.dat.

Usage
-----

List Spreadsheets and Worksheets:

    >>> from google_spreadsheet.api import SpreadsheetAPI
    >>> api = SpreadsheetAPI()
    >>> spreadsheets = api.list_spreadsheets()
    >>> spreadsheets
    [('MyFirstSpreadsheet', 'tkZQWzwHEjKTWFFCAgw'), ('MySecondSpreadsheet', 't5I-ZPGdXjTrjMefHcg'), 
    ('MyThirdSpreadsheet', 't0heCWhzCmm9Y-GTTM_Q')]
    >>> worksheets = api.list_worksheets(spreadsheets[0][1])
    >>> worksheets
    [('MyFirstWorksheet', 'od7'), ('MySecondWorksheet', 'od6'), ('MyThirdWorksheet', 'od4')]

Please note that in order to work with a Google Spreadsheet it must be accessible
to the user who's login credentials are provided. The `GOOGLE_SPREADSHEET_SOURCE`
argument is used by Google to identify your application and track API calls.

Working with a Worksheet:

    >>> from google_spreadsheet.api import SpreadsheetAPI
    >>> api = SpreadsheetAPI()
    >>> sheet = api.get_worksheet('tkZQWzwHEjKTWFFCAgw', 'od7')
    >>> rows = sheet.get_rows()
    >>> len(rows)
    18
    >>> row_to_update = rows[0]
    >>> row_to_update['name'] = 'New Name'
    >>> sheet.update_row(row_to_update)
    {'name': 'New Name'...}
    >>> row_to_insert = rows[0]
    >>> row_to_insert['name'] = 'Another Name'
    >>> row = sheet.insert_row(row_to_insert)
    {'name': 'Another Name'...}
    >>> sheet.delete_row(row)
    >>> sheet.delete_all_rows()

Advanced Queries:

    >>> from google_spreadsheet.api import SpreadsheetAPI
    >>> api = SpreadsheetAPI()
    >>> sheet = api.get_worksheet('tkZQWzwHEjKTWFFCAgw', 'od7')
    >>> rows = sheet.get_rows(query='name = "Joe" and height < 175')

Or filter in memory:

    >>> sheet = api.get_worksheet('tkZQWzwHEjKTWFFCAgw', 'od7')
    >>> filtered_rows = sheet.get_rows(
            filter_func=lambda row: row['status'] == "READY")

Sort:

    >>> sheet = api.get_worksheet('tkZQWzwHEjKTWFFCAgw', 'od7')
    >>> rows = sheet.get_rows(order_by='column:age', reverse='true')

Batch Update:

    This is experimental.  If you want to bulk update a spreadsheet and the
    first column is the key and the data has one header row, you can use this.
    If key exists that row is updated.  If not, the row is inserted.  Blank
    cells do NOT get updated.  This is much faster than delete / insert.

    >>> sheet = api.get_worksheet('tkZQWzwHEjKTWFFCAgw', 'od7')
    >>> sheet.batch_verify_key_content([['key1','a','1'],['key2','b','2']])

That's it.

For more information about these calls, please consult the [Google Spreadsheets
API Developer Guide](https://developers.google.com/google-apps/spreadsheets/).

Tests
------
To run the test suite please follow these steps:

* Make sure [Nose](http://readthedocs.org/docs/nose/en/latest/) is installed: (`pip install nose`)
* Create a local file named: `test_settings.py` with the following variables set to the relevant values: `GOOGLE_SPREADSHEET_USER`, `GOOGLE_SPREADSHEET_PASSWORD`, `GOOGLE_SPREADSHEET_SOURCE`, `GOOGLE_SPREADSHEET_KEY`, `GOOGLE_WORKSHEET_KEY`, `COLUMN_NAME`, `COLUMN_UNIQUE_VALUE`
* Run `nosetests`

License
-------

Copyright &copy; 2012 Yoav Aviram

See LICENSE for details.

