import typing
import datetime
import zoneinfo
import csv
import pathlib
import uuid
import icalendar


class Config(typing.TypedDict):
    HEADER_ROWS_TO_SKIP: int
    CSV_NAME: int
    CSV_START_DATE: int | tuple[int, int]
    CSV_END_DATE: int | tuple[int, int]
    CSV_DESCRIPTION: int
    CSV_LOCATION: int | list[int]
    UID: typing.Optional[int]
    EMAIL: typing.Optional[int]
    CSV_DELIMITER: str
    TIME_FORMAT: str
    INPUT_TZ: typing.Optional[str]
    OUTPUT_TZ: str


class ConfigOverrides(Config, total=False):
    pass


DEFAULT_CONFIG: Config = {
    'HEADER_ROWS_TO_SKIP':  0,

    # The variables below refer to the column indexes in the CSV
    'CSV_NAME': 0,
    'CSV_START_DATE': 1,
    'CSV_END_DATE': 1,
    'CSV_DESCRIPTION': 2,
    'CSV_LOCATION': 3,
    'UID': None,
    'EMAIL': None,

    # Delimiter used in CSV file
    'CSV_DELIMITER': ',',

    # Time Format used in CSV file
    'TIME_FORMAT': '%Y-%m-%dT%H%M',
    'INPUT_TZ':None,
    "OUTPUT_TZ": "UTC",

}
OV_DEFAULT_CONFIG: Config = {
    'HEADER_ROWS_TO_SKIP':  2,

    # The variables below refer to the column indexes in the CSV
    'CSV_NAME': 1,
    'CSV_START_DATE': (0, 2),
    'CSV_END_DATE': (0, 3),
    'CSV_DESCRIPTION': 10,
    'CSV_LOCATION': [4, 5],
    'UID': 12,
    'EMAIL': None,

    # Delimiter used in CSV file
    'CSV_DELIMITER': ',',

    # Time Format used in CSV file
    'TIME_FORMAT': '%d.%m.%YT%H:%M',
    'INPUT_TZ': 'CEST',
    "OUTPUT_TZ": "UTC",
}


class Convert():
    def __init__(self) -> None:
        self.csv_data: typing.List[typing.List[typing.Any]] = []
        self.cal: typing.Optional[icalendar.Calendar] = None

    def _generate_configs_from_default(
        self,
        overrides: typing.Optional[ConfigOverrides] = None,
    ) -> Config:
        """ Generate configs by inheriting from defaults """
        config = DEFAULT_CONFIG.copy()
        non_optional_overrides: ConfigOverrides = {}  # type: ignore
        if overrides:
            non_optional_overrides = overrides
        for k, v in non_optional_overrides.items():
            config[k] = v  # type: ignore
        return config

    def read_ical(self, ical_file_location: typing.Union[str, pathlib.Path]) \
            -> icalendar.Calendar:
        """ Read the ical file """
        with open(ical_file_location, 'r', encoding='utf-8') as ical_file:
            data = ical_file.read()
        self.cal = icalendar.Calendar.from_ical(data)
        return self.cal

    def read_csv(
        self,
        csv_location: typing.Union[str, pathlib.Path],
        csv_configs: typing.Optional[Config] = None,
    ) -> typing.List[typing.List[typing.Any]]:
        """ Read the csv file """
        csv_configs = self._generate_configs_from_default(csv_configs)
        with open(csv_location, 'r', encoding='utf-8') as csv_file:
            csv_reader = csv.reader(csv_file,
                                    delimiter=csv_configs['CSV_DELIMITER'])
            self.csv_data = list(csv_reader)
        self.csv_data = self.csv_data[csv_configs['HEADER_ROWS_TO_SKIP']:]
        return self.csv_data

    def make_ical(
        self,
        csv_configs: typing.Optional[Config] = None,
    ) -> icalendar.Calendar:
        """ Make iCal entries """
        csv_configs = self._generate_configs_from_default(csv_configs)
        output_timezone = zoneinfo.ZoneInfo(csv_configs["OUTPUT_TZ"])
        input_timezone = csv_configs["INPUT_TZ"]
        self.cal = icalendar.Calendar()
        for row in self.csv_data:
            event = icalendar.Event()
            name = row[csv_configs['CSV_NAME']]
            start_date_str: str
            if isinstance(csv_configs['CSV_START_DATE'], int):
                start_date_str = row[csv_configs['CSV_START_DATE']]
            else:
                idx = csv_configs['CSV_START_DATE']
                start_date_str = row[idx[0]]
                if row[idx[1]] == '':
                    start_date_str += "T00:00"
                else:
                    start_date_str += "T" + row[idx[1]]
            if input_timezone is None:
                start_date = datetime.datetime.strptime(start_date_str, csv_configs["TIME_FORMAT"])
            else:
                start_date_str += f" {input_timezone}"
                start_date = datetime.datetime.strptime(start_date_str, csv_configs["TIME_FORMAT"]+" %Z")
            start_date = start_date.astimezone(output_timezone)
            end_date_str: str
            if isinstance(csv_configs['CSV_END_DATE'], int):
                end_date_str = row[csv_configs['CSV_END_DATE']]
            else:
                idx = csv_configs['CSV_END_DATE']
                end_date_str = row[idx[0]]
                if row[idx[1]] == '':
                    end_date_str += "T23:59"
                else:
                    end_date_str += "T" + row[idx[1]]
            if input_timezone is None:
                end_date = datetime.datetime.strptime(end_date_str, csv_configs["TIME_FORMAT"])
            else:
                end_date_str += f" {input_timezone}"
                end_date = datetime.datetime.strptime(end_date_str, csv_configs["TIME_FORMAT"]+" %Z")
            end_date = end_date.astimezone(output_timezone)
            location_idx: list[int]
            if isinstance(csv_configs['CSV_LOCATION'], int):
                location_idx = [csv_configs['CSV_LOCATION']]
            else:
                location_idx = csv_configs['CSV_LOCATION']
            location = ", ".join(row[val] for val in location_idx)
            event.add('summary', name)
            event.add('dtstart', start_date)
            event.add('dtend', end_date)
            event.add('description', row[csv_configs['CSV_DESCRIPTION']])
            event.add('location', location)
            if csv_configs['UID'] is None:
                event_uuid = uuid.uuid5(uuid.NAMESPACE_DNS,
                                        f"{name}-{start_date}").hex
                event.add('uid', event_uuid)
            elif row[csv_configs["UID"]] == '':
                event_uuid = uuid.uuid5(uuid.NAMESPACE_DNS,
                                        f"{name}-{start_date}").hex
                event.add('uid', event_uuid)
                row[csv_configs["UID"]] = event_uuid

            else:
                event.add('uid', row[csv_configs['UID']])
            if csv_configs['EMAIL'] is not None:
                event.add('email', row[csv_configs['EMAIL']])
            event.add('dtstamp', datetime.datetime.now())
            self.cal.add_component(event)
        return self.cal

    def make_csv(self) -> None:
        """ Make CSV """
        if self.cal is None:
            return
        for event in self.cal.subcomponents:
            if event.name != 'VEVENT':
                continue
            dtstart = ''
            if event.get('DTSTART'):
                dtstart = event.get('DTSTART').dt
            dtend = ''
            if event.get('DTEND'):
                dtend = event.get('DTEND').dt
            email = ''
            if event.get("EMAIL"):
                email = event.get("EMAIL")
            row = [
                event.get('SUMMARY'),
                dtstart,
                dtend,
                event.get('DESCRIPTION'),
                event.get('LOCATION'),
                event.get('UID'),
                email,
            ]
            row = [str(x) for x in row]
            self.csv_data.append(row)

    def save_ical(self, ical_location: str) -> None:
        """ Save the calendar instance to a file """
        if self.cal is None:
            return None
        data = self.cal.to_ical()
        with open(ical_location, 'wb') as ical_file:
            ical_file.write(data)

    def save_csv(
        self,
        csv_location: str,
        csv_delimiter: str = ',',
    ) -> None:
        """ Save the csv to a file """
        with open(csv_location, 'w', encoding='utf-8') as csv_handle:
            writer = csv.writer(csv_handle, delimiter=csv_delimiter,
                                quoting=csv.QUOTE_MINIMAL)
            writer.writerows(self.header())
            for row in self.csv_data:
                writer.writerow([r.strip() for r in row])

    def get_title(self) -> typing.Optional[str]:
        if self.cal is None:
            return None
        for event in self.cal.subcomponents:
            if event.name != 'VEVENT':
                continue
            return event.get("SUMMARY")
        return None

    def get_start_date(self) -> typing.Optional[str]:
        if self.cal is None:
            return None
        for event in self.cal.subcomponents:
            if event.name != 'VEVENT':
                continue
            return event.get("DTSTART").dt
        return None

    def header(self) -> list[list[typing.Any]]:
        return []


class OVConvert(Convert):
    def __init__(self) -> None:
        self.csv_data: typing.List[typing.List[typing.Any]] = []
        self.cal: typing.Optional[icalendar.Calendar] = None

    def _generate_configs_from_default(
        self,
        overrides: typing.Optional[ConfigOverrides] = None,
    ) -> Config:
        """ Generate configs by inheriting from defaults """
        config = OV_DEFAULT_CONFIG.copy()
        non_optional_overrides: ConfigOverrides = {}  # type: ignore
        if overrides:
            non_optional_overrides = overrides
        for k, v in non_optional_overrides.items():
            config[k] = v  # type: ignore
        return config

    def header(self) -> list[list[typing.Any]]:
        return [['DATE', 'SUMMARY', 'START', 'END', 'LOCATION_SHORT',
                 'LOCATION_ADDRESS', 'LOCATION_TEXT', 'AUDIENCE', 'STATUS',
                 'Kürzel', 'DESCRIPTION', 'REMARKS', 'UID', '', '', '', '', '',
                 '', '', '', '', ''],
                ['date', 'Text', 'time', 'time', 'Text', 'address', 'Text',
                 'One of: Vorstand, OV, OVplus, other',
                 ('Zero to many of: Planned, booked, homepage-published,'
                  ' email-invited, done, cancelled, location-needed'),
                 'Text', 'text', 'Text', 'UID capital letters, generated']]
