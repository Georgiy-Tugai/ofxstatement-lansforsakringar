import xlrd
from ofxstatement.parser import StatementParser
from ofxstatement.plugin import Plugin
from ofxstatement.statement import Statement, StatementLine, generate_transaction_id, recalculate_balance
from datetime import timedelta
from hashlib import sha1

import logging
log = logging.getLogger(__name__)

class LansforsakringarPlugin(Plugin):
    """Länsförsäkringar <https://www.lansforsakringar.se>"""

    def get_parser(self, filename):
        bank_id = self.settings.get('bank', 'ELLFSESS')
        account_id = self.settings.get('account')
        return LansforsakringarParser(filename, bank_id, account_id)


class LansforsakringarParser(StatementParser):
    statement = Statement(currency='SEK')

    def __init__(self, filename, bank_id, account_id):
        self.filename = filename
        self.statement.bank_id = bank_id
        self.statement.account_id = account_id
        self.sheet = None
        self.row_num = 0
        self.seen = {}

    def parse(self):
        with xlrd.open_workbook(self.filename) as book:
            self.sheet = book.sheet_by_index(0)
            return super().parse()

    def split_records(self):
        rows = self.sheet.get_rows()
        datestr = next(rows)  # statement date
        assert datestr[0].value.startswith("Kontoutdrag -")
        # end of statement date's day
        self.statement.end_date = self.parse_datetime(datestr[0].value[13:]) + timedelta(days=1)
        next(rows)  # headers
        return rows

    def parse_record(self, row):
        self.row_num += 1
        line = StatementLine()
        line.date = self.parse_datetime(row[0].value)
        line.date_user = self.parse_datetime(row[1].value)
        line.refnum = str(self.row_num)
        line.memo = row[2].value
        line.amount = row[3].value
        line.trntype = self.get_type(line)
        if self.statement.start_balance is None and self.row_num == 1:
            self.statement.start_balance = row[4].value - line.amount
            self.statement.start_date = line.date_user
        self.statement.end_balance = row[4].value
        line.id = self.generate_transaction_id(line)
        if line.id in self.seen:
            log.warn("Transaction with duplicate FITID generated:\n%s\n%s\n\n" % (line, self.seen[line.id]))
        else:
            self.seen[line.id] = line
        return line

    def generate_transaction_id(self, stmt_line):
        """Generate pseudo-unique id for given statement line.

        This function can be used in statement parsers when real transaction id is
        not available in source statement.

        Modified version of ofxstatement's function of the same name.
        Includes refnum (in our case, row number) into the hash; this is safe here as Kontoutdrag is only available after the reporting period is over i.e. it should never change.
        """
        h = sha1()
        h.update(stmt_line.date.strftime("%Y-%m-%d %H:%M:%S").encode("utf8"))
        h.update(stmt_line.refnum.encode("utf8"))
        h.update(stmt_line.memo.encode("utf8"))
        h.update(str(stmt_line.amount).encode("utf8"))
        return h.hexdigest()

    @staticmethod
    def get_type(line):
        if line.amount > 0:
            return 'CREDIT'
        elif line.amount < 0:
            return 'DEBIT'
        else:
            return 'OTHER'
