import unittest
import unittest.mock
from converter import xls_to_ics

# check if correct columns in intermediate csv?
# from https://gist.github.com/pitrk/9db7dd9eb1c1e24f2e53edb116d09880 
# check if failed and cancelld bookings are filtered out 
class TestFunction(unittest.TestCase):
    mock_file_content = """
bookingCode,persons,price.groupInternal,status,bookingStartAtDate,bookingStartAtTime,bookingEndAtDate,bookingEndAtTime,customerNote
K-088A2D7A,1,FALSE,completedAndCancellationCompleted,04.01.2025,12:30,04.01.2025,20:29,
K-088A7330,30,FALSE,failed,04.01.2025,12:30,04.01.2025,20:59,
K-484D8FA8,30,FALSE,completed,04.01.2025,12:30,04.01.2025,20:29,
K-481C659E,1,TRUE,completed,04.01.2025,22:30,04.01.2025,23:59,Schlüsseltest	
K-F803ADD6,10,FALSE,completed,07.01.2025,17:30,07.01.2025,20:29,Grete Community Music	
K-4845457D,25,FALSE,completedAndCancellationCompleted,17.01.2025,15:30,17.01.2025,23:29,30 Geburtstag:) 	
    """
    def test_only_completed_bookings(self):
        xls_to_ics('test_xlstoics.xlsx', 'test_xlstoics.ics')
        with open('test_xlstoics.ics') as ics_output:
            assert ics_output.read().count("BEGIN:VEVENT") == 5
            

# K-C1B168E5	25	TRUE	completed	08.01.2025	18:00	08.01.2025	20:59	Bewohner*innentreffen	
# K-F8134239	5	TRUE	completed	09.01.2025	18:30	09.01.2025	19:59	Yogastunde	
# K-4863DBED	5	TRUE	completed	10.01.2025	18:00	10.01.2025	20:59	Der Westen ist bunt!	
# K-485494F2	35	FALSE	completed	12.01.2025	12:00	12.01.2025	18:59	Kindergeburtstag	
# K-F89C0F4D	1	TRUE	completed	13.01.2025	00:00	13.01.2025	23:59		
# K-F803D82A	10	FALSE	completed	14.01.2025	17:30	14.01.2025	20:29	Grete Community Music	
# K-48504616	10	TRUE	completed	15.01.2025	18:00	15.01.2025	21:59		
# K-F8137359	5	TRUE	completed	16.01.2025	18:30	16.01.2025	19:59	Yogastunde	
# K-48A37760	25	FALSE	completed	17.01.2025	14:00	17.01.2025	23:59	30 Geburtstag 	
# K-C11A8742	30	TRUE	completedAndCancellationCompleted	18.01.2025	12:00	18.01.2025	21:59	Platzhalter für FEIERWERK	
# K-B5B7B1EA	50	FALSE	completed	18.01.2025	12:00	18.01.2025	23:59	Bitte als Verwendungszweck in der Rechnung angeben: Feierwerk Boom - Coffee & Cake - 48511	
# K-81A3EE07	20	FALSE	completedAndCancellationCompleted	19.01.2025	17:00	19.01.2025	22:59		
# K-089EF8A0	20	TRUE	completed	20.01.2025	19:15	20.01.2025	20:44	Zumba	



