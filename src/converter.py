import typing
import datetime
import zoneinfo
import csv
import pathlib
import os.path
import uuid
import icalendar
from slugify import slugify


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

HTWK_DEFAULT_CONFIG: Config = {
    'HEADER_ROWS_TO_SKIP':  2,

    # The variables below refer to the column indexes in the CSV
    'CSV_NAME': 1,
    'CSV_START_DATE': (0, 2),
    'CSV_END_DATE': (0, 3),
    'CSV_DESCRIPTION': 5,
    'CSV_LOCATION': [4],
    'UID': None,
    'EMAIL': None,

    # Delimiter used in CSV file
    'CSV_DELIMITER': ';',

    # Time Format used in CSV file
    'TIME_FORMAT': '%d.%m.%YT%H:%M',
    'INPUT_TZ': 'CEST',
    "OUTPUT_TZ": "UTC",
}


def load_config(base_config: Config, config_overrides: typing.Optional[ConfigOverrides]) -> Config:
    new_config=base_config.copy()
    if config_overrides is None:
        return new_config
    for key,value in config_overrides.items():
        new_config[key]=value
    return new_config


class Converter():
    def __init__(self, config: typing.Optional[ConfigOverrides] = None) -> None:
        self.__config = load_config(DEFAULT_CONFIG, config)
        self.__csv_data: typing.List[typing.List[typing.Any]] = []
        self.__cal: typing.Optional[icalendar.Calendar] = None

    def change_config(self, config_overrides: ConfigOverrides):
        self.__config = load_config(self.__config,config_overrides)

    def read_ical(self, ical_file_location: typing.Union[str, pathlib.Path]) \
            -> icalendar.Calendar:
        """ Read the ical file """
        with open(ical_file_location, 'r', encoding='utf-8') as ical_file:
            data = ical_file.read()
        self.__cal = icalendar.Calendar.from_ical(data)  # type: ignore
        return self.__cal  # type: ignore

    def read_csv(
        self,
        csv_location: typing.Union[str, pathlib.Path],
    ) -> typing.List[typing.List[typing.Any]]:
        """ Read the csv file """
        csv_configs = self.__config
        with open(csv_location, 'r', encoding='utf-8') as csv_file:
            csv_reader = csv.reader(csv_file,
                                    delimiter=csv_configs['CSV_DELIMITER'])
            self.__csv_data = list(csv_reader)
        self.__csv_data = self.__csv_data[csv_configs['HEADER_ROWS_TO_SKIP']:]
        return self.__csv_data

    def make_ical(
        self,
    ) -> icalendar.Calendar:
        """ Make iCal entries """
        csv_configs = self.__config
        output_timezone = zoneinfo.ZoneInfo(csv_configs["OUTPUT_TZ"])
        input_timezone = csv_configs["INPUT_TZ"]
        self.__cal = icalendar.Calendar()
        for row in self.__csv_data:
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
            print(csv_configs)
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
            self.__cal.add_component(event)
        return self.__cal

    def make_csv(self) -> None:
        """ Make CSV """
        if self.__cal is None:
            return
        for event in self.__cal.subcomponents:
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
            self.__csv_data.append(row)

    def save_ical(self, ical_location: str) -> None:
        """ Save the calendar instance to a file """
        if self.__cal is None:
            return None
        if not os.path.exists(ical_location):
            os.mkdir(ical_location)
        if os.path.isfile(ical_location) or ical_location.endswith(".ics"):
            data = self.__cal.to_ical()
            with open(ical_location, 'wb') as ical_file:
                ical_file.write(data)
            return None
        for event in self.__cal.subcomponents:
            if event.name !="VEVENT":
                continue
            calendar = icalendar.Calendar()
            title = event.get("SUMMARY")
            start = event.get("DTSTART").dt.strftime("%Y%m%d_%H%M")
            calendar.add_component(event)
            data = calendar.to_ical()
            filename = slugify(f"{title}_{start}.ics")
            with open(os.path.join(ical_location,filename),"wb") as ical_file:
                ical_file.write(data)
        with open(os.path.join(ical_location,"bundled.ics"),"wb") as ical_file:
            ical_file.write(self.__cal.to_ical())

    def save_csv(self, csv_location: str) -> None:
        """ Save the csv to a file """
        csv_delimiter = self.__config['CSV_DELIMITER']
        with open(csv_location, 'w', encoding='utf-8') as csv_handle:
            writer = csv.writer(csv_handle, delimiter=csv_delimiter,
                                quoting=csv.QUOTE_MINIMAL)
            writer.writerows(self.header())
            for row in self.__csv_data:
                writer.writerow([r.strip() for r in row])

    def get_title(self) -> typing.Optional[str]:
        if self.__cal is None:
            return None
        for event in self.__cal.subcomponents:
            if event.name != 'VEVENT':
                continue
            return event.get("SUMMARY")
        return None

    def get_start_date(self) -> typing.Optional[str]:
        if self.__cal is None:
            return None
        for event in self.__cal.subcomponents:
            if event.name != 'VEVENT':
                continue
            return event.get("DTSTART").dt
        return None

    def header(self) -> list[list[typing.Any]]:
        return []
