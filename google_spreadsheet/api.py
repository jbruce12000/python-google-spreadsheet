#!/usr/bin/python
#
# Copyright (C) 2012 Yoav Aviram.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
import gdata.spreadsheet.service
import gdata.service
import httplib2
from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client import tools

ID_FIELD = '__rowid__'


class WorksheetException(Exception):
    """Base class for spreadsheet exceptions.
    """
    pass


class SpreadsheetAPI(object):
    def __init__(self, client_secrets_file='./client_secrets.json', credentials_file='./creds.dat'):
        """Initialise a Spreadsheet API wrapper.

        :param client_secrets_file:
            A file containing xml secrets as defined here...
        :param client_credentials_file:
            A file used to cache credentials.
        """
        self.credentials_file = credentials_file
        self.client_secrets_file = client_secrets_file
        self.storage = Storage(self.credentials_file)
        self.credentials = self.storage.get()

        if self.credentials is None or self.credentials.invalid:
            self.flags = tools.argparser.parse_args(args=[])
            self.flow = flow = flow_from_clientsecrets(self.client_secrets_file,
                scope=["https://spreadsheets.google.com/feeds"])
            self.credentials = tools.run_flow(self.flow, self.storage,
                self.flags)

        if self.credentials.access_token_expired:
            self.credentials.refresh(httplib2.Http())

        self.client = gdata.spreadsheet.service.SpreadsheetsService(
            additional_headers={'Authorization': 'Bearer %s' %
            self.credentials.access_token})

    def _get_client(self):
        """Initialize a `gdata` client.

        :returns:
            A gdata client.
        """
        return self.client

    def list_spreadsheets(self):
        """List Spreadsheets.

        :return:
            A list with information about the spreadsheets available
        """
        sheets = self._get_client().GetSpreadsheetsFeed()
        return map(lambda e: (e.title.text, e.id.text.rsplit('/', 1)[1]),
            sheets.entry)

    def list_worksheets(self, spreadsheet_key):
        """List Spreadsheets.

        :return:
            A list with information about the spreadsheets available
        """
        wks = self._get_client().GetWorksheetsFeed(
            key=spreadsheet_key)
        return map(lambda e: (e.title.text, e.id.text.rsplit('/', 1)[1]),
            wks.entry)

    def get_worksheet(self, spreadsheet_key, worksheet_key):
        """Get Worksheet.

        :param spreadsheet_key:
            A string representing a google spreadsheet key.
        :param worksheet_key:
            A string representing a google worksheet key.
        """
        return Worksheet(self._get_client(), spreadsheet_key, worksheet_key)


class Worksheet(object):
    """Worksheet wrapper class.
    """
    def __init__(self, gd_client, spreadsheet_key, worksheet_key):
        """Initialise a client

        :param gd_client:
            A GDATA client.
        :param spreadsheet_key:
            A string representing a google spreadsheet key.
        :param worksheet_key:
            A string representing a google worksheet key.
        """
        self.gd_client = gd_client
        self.spreadsheet_key = spreadsheet_key
        self.worksheet_key = worksheet_key
        self.keys = {'key': spreadsheet_key, 'wksht_id': worksheet_key}
        self.entries = None
        self.query = None
        self.cells = self.gd_client.GetCellsFeed(self.spreadsheet_key,
            self.worksheet_key)
        self.batchRequest = gdata.spreadsheet.SpreadsheetsCellsFeed()
        self.header_row = self.set_header_row()

    def find_cell_by_contents(self, searchfor):
        """find the cell with the given contents"""
        for cell in enumerate(self.cells.entry):
            if cell[1].content.text == searchfor:
                return cell[1]
        return None

    def batch_verify_key_content(self, data=[]):
        """
        import a list of lists of data into a spreadsheet.  Check each row to
        see if it already exists.  If it does exist based on using the first
        column as the key, update any other columns that need it.  If it does
        not exist, insert it at the end of the spreadsheet.  This does not
        update cells that are empty.  This is intended to be used when you
        have data that changes and updates are necessary.

        :param data:
            list of lists where each inner list is a row
        """
        for row in data:
            cell = self.find_cell_by_contents(row[0])
            if cell and cell.cell.col == '1':
                for col in row:
                    cell.cell.inputValue = col
                    self.batchRequest.AddUpdate(cell)
                    cell = self.next_cell(cell)
                    if cell is None:
                        break
            else:
                self.insert_as_last(row)
        updated = self.gd_client.ExecuteBatch(self.batchRequest, self.cells.GetBatchLink().href)

    def next_cell(self, cell):
        """find the next cell in this row.  Empty cells do not exist to the
        google api for whatever reason.

        :param cell:
            gdata.spreadsheet.SpreadsheetsCell object
            used to get the current row
        :return:
            gdata.spreadsheet.SpreadsheetsCell object
            None if there is no matching cell
        """
        col = int(cell.cell.col)
        col = col + 1
        return self.find_cell(cell.cell.row, col)

    def find_cell(self, row=1, col=1):
        """find the cell with the given row and col"""
        row = str(row)
        col = str(col)
        for cell in enumerate(self.cells.entry):
            if cell[1].cell.row == row:
                if cell[1].cell.col == col:
                    return cell[1]
        return None

    def update_cell(self,row,col,val):
        # .77 sec / call or about 4675 cells/hour
        # this does create the cell though if it is blank or does not exist
        # which is cool
        return self.gd_client.UpdateCell(row,col,val,self.spreadsheet_key,self.worksheet_key)

    def batch(self, startxy=(2,1),endxy=(10,4), data=[]):
        """Batch Import a list of lists to a specific location.
        :param startxy:
            start row,column - integers
        :param column:
            end row,column - integers
        :param data:
            The data to batch import - list of lists

        data always starts at startxy
        data outside the given startxy / endxy range is ignored
        insufficient data is set as a blank
        """

        for r in range(startxy[0],endxy[0]+1):
            for c in range(startxy[1],endxy[1]+1):

                # get content from data passed in
                # is this cell within the data passed in? then grab it
                content = ''
                cell = self.find_cell(r, c)
                try:
                    content = str(data[r-startxy[0]][c-startxy[1]])
                except (IndexError):
                    if cell:
                        cell.cell.inputValue = ''
                        self.batchRequest.AddUpdate(cell)
                        # do not want to update_cell here
                        # it's already blank, leave it alone
                    continue
                if cell:
                    cell.cell.inputValue = content
                    self.batchRequest.AddUpdate(cell)
                else:
                    self.update_cell(r,c,content)


        updated = self.gd_client.ExecuteBatch(self.batchRequest, self.cells.GetBatchLink().href)

    def set_header_row(self):
        header = []
        c = 1
        while True:
            cell = self.find_cell(1, c)
            if cell:
                header.append(cell.content.text)
            else:
                break
            c = c + 1
        return header

    def insert_as_last(self, data):
        """
        Insert a row given a list.  Since I am taking a list,
        I must assume the first row is header row and use that
        to create a dict to pass to gdata.
        :param data:
            A list containing a row of data where each element is a column
        """

        # create a dict for insertion
        row_to_insert = {}
        for x in range(0, len(self.header_row)):
            row_to_insert[self.header_row[x]] = data[x]
        return self.insert_row(row_to_insert)

    def _row_to_dict(self, row):
        """Turn a row of values into a dictionary.
        :param row:
            A row element.
        :return:
            A dictionary with rows.
        """
        result = dict([(key, row.custom[key].text) for key in row.custom])
        result[ID_FIELD] = row.id.text.split('/')[-1]
        return result

    def _get_row_entries(self, query=None):
        """Get Row Entries.

        :return:
            A rows entry.
        """
        if not self.entries:
            self.entries = self.gd_client.GetListFeed(
                query=query, **self.keys).entry
        return self.entries

    def _get_row_entry_by_id(self, id):
        """Get Row Entry by ID

        First search in cache, then fetch.
        :param id:
            A string row ID.
        :return:
            A row entry.
        """
        entry = [entry for entry in self._get_row_entries()
                 if entry.id.text.split('/')[-1] == id]
        if not entry:
            entry = self.gd_client.GetListFeed(row_id=id, **self.keys).entry
            if not entry:
                raise WorksheetException("Row ID '{0}' not found.").format(id)
        return entry[0]

    def _flush_cache(self):
        """Flush Entries Cache."""
        self.entries = None

    def _make_query(self, query=None, order_by=None, reverse=None):
        """Make Query.

         A utility method to construct a query.

        :return:
            A :class:`~,gdata.spreadsheet.service.ListQuery` or None.
        """
        if query or order_by or reverse:
            q = gdata.spreadsheet.service.ListQuery()
            if query:
                q.sq = query
            if order_by:
                q.orderby = order_by
            if reverse:
                q.reverse = reverse
            return q
        else:
            return None

    def get_rows(self, query=None, order_by=None,
                 reverse=None, filter_func=None):
        """Get Rows

        :param query:
            A string structured query on the full text in the worksheet.
              [columnName][binaryOperator][value]
              Supported binaryOperators are:
              - (), for overriding order of operations
              - = or ==, for strict equality
              - <> or !=, for strict inequality
              - and or &&, for boolean and
              - or or ||, for boolean or.
        :param order_by:
            A string which specifies what column to use in ordering the
            entries in the feed. By position (the default): 'position' returns
            rows in the order in which they appear in the GUI. Row 1, then
            row 2, then row 3, and so on. By column:
            'column:columnName' sorts rows in ascending order based on the
            values in the column with the given columnName, where
            columnName is the value in the header row for that column.
        :param reverse:
            A string which specifies whether to sort in descending or ascending
            order.Reverses default sort order: 'true' results in a descending
            sort; 'false' (the default) results in an ascending sort.
        :param filter_func:
            A lambda function which applied to each row, Gets a row dict as
            argument and returns True or False. Used for filtering rows in
            memory (as opposed to query which filters on the service side).
        :return:
            A list of row dictionaries.
        """
        new_query = self._make_query(query, order_by, reverse)
        if self.query is not None and self.query != new_query:
            self._flush_cache()
        self.query = new_query
        rows = [self._row_to_dict(row)
            for row in self._get_row_entries(query=self.query)]
        if filter_func:
            rows = filter(filter_func, rows)
        return rows

    def update_row(self, row_data):
        """Update Row (By ID).

        Only the fields supplied will be updated.
        :param row_data:
            A dictionary containing row data. The row will be updated according
            to the value in the ID_FIELD.
        :return:
            The updated row.
        """
        try:
            id = row_data[ID_FIELD]
        except KeyError:
            raise WorksheetException("Row does not contain '{0}' field. "
                                "Please update by index.".format(ID_FIELD))
        entry = self._get_row_entry_by_id(id)
        new_row = self._row_to_dict(entry)
        new_row.update(row_data)
        entry = self.gd_client.UpdateRow(entry, new_row)
        if not isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            raise WorksheetException("Row update failed: '{0}'".format(entry))
        for i, e in enumerate(self.entries):
            if e.id.text == entry.id.text:
                self.entries[i] = entry
        return self._row_to_dict(entry)

    def update_row_by_index(self, index, row_data):
        """Update Row By Index

        :param index:
            An integer designating the index of a row to update (zero based).
            Index is relative to the returned result set, not to the original
            spreadseet.
        :param row_data:
            A dictionary containing row data.
        :return:
            The updated row.
        """
        entry = self._get_row_entries(self.query)[index]
        row = self._row_to_dict(entry)
        row.update(row_data)
        entry = self.gd_client.UpdateRow(entry, row)
        if not isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            raise WorksheetException("Row update failed: '{0}'".format(entry))
        self.entries[index] = entry
        return self._row_to_dict(entry)

    def insert_row(self, row_data):
        """Insert Row

        :param row_data:
            A dictionary containing row data.
        :return:
            A row dictionary for the inserted row.
        """
        entry = self.gd_client.InsertRow(row_data, **self.keys)
        if not isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            raise WorksheetException("Row insert failed: '{0}'".format(entry))
        if self.entries:
            self.entries.append(entry)
        return self._row_to_dict(entry)

    def delete_row(self, row):
        """Delete Row (By ID).

        Requires that the given row dictionary contains an ID_FIELD.
        :param row:
            A row dictionary to delete.
        """
        try:
            id = row[ID_FIELD]
        except KeyError:
            raise WorksheetException("Row does not contain '{0}' field. "
                                "Please delete by index.".format(ID_FIELD))
        entry = self._get_row_entry_by_id(id)
        self.gd_client.DeleteRow(entry)
        for i, e in enumerate(self.entries):
            if e.id.text == entry.id.text:
                del self.entries[i]

    def delete_row_by_index(self, index):
        """Delete Row By Index

        :param index:
            A row index. Index is relative to the returned result set, not to
            the original spreadsheet.
        """
        entry = self._get_row_entries(self.query)[index]
        self.gd_client.DeleteRow(entry)
        del self.entries[index]

    def delete_all_rows(self, header_rows=0):
        """Delete All Rows
        """
        entries = self._get_row_entries(self.query)
        if header_rows:
            stuff_to_delete = range(header_rows, len(entries))
            stuff_to_delete.reverse()
            for x in stuff_to_delete:
                self.delete_row_by_index(x)
        else:
            for entry in entries:
                self.gd_client.DeleteRow(entry)
        self._flush_cache()
