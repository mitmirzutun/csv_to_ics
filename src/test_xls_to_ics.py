import unittest
from converter import xls_to_ics


# check if correct columns in intermediate csv?
class TestFunction(unittest.TestCase):
    def test_only_completed_bookings(self):
        xls_to_ics('test_xlstoics.xlsx', 'test_xlstoics.ics')
        with open('test_xlstoics.ics') as ics_output:
            assert ics_output.read().count("BEGIN:VEVENT") == 5
