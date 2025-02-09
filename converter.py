import csv_ical
convert = csv_ical.Convert()


def ics_to_csv(ics_file_name, csv_file_name):
    convert.read_ical(ics_file_name)
    convert.make_csv()
    convert.save_csv(csv_file_name)


def xls_to_ics(xlsx_file_name, ics_file_name):
    import openpyxl
    import csv
    import datetime
    xlsx_file = openpyxl.load_workbook(xlsx_file_name)
    worksheet = xlsx_file.active
    with open(ics_file_name + ".tmp.csv", "w") as csv_file:
        writer = csv.writer(csv_file)
        titles = list(val.value for val in next(worksheet.rows))
        XLSX_BOOKING_ID: int = titles.index("bookingCode")
        XLSX_INTERNAL: int = titles.index("price.groupInternal")
        XLSX_STATUS: int = titles.index("status")
        XLSX_PERSONS: int = titles.index("persons")
        XLSX_RESOURCE: int = titles.index("resourceTitle")
        XLSX_START_DATE: int = titles.index("bookingStartAtDate")
        XLSX_START_TIME: int = titles.index("bookingStartAtTime")
        XLSX_END_DATE: int = titles.index("bookingEndAtDate")
        XLSX_END_TIME: int = titles.index("bookingEndAtTime")
        XLSX_DESCRIPTION: int = titles.index("customerNote")
        for row in range(2, worksheet.max_row + 1):
            print(worksheet.cell(column=XLSX_BOOKING_ID, row=row).value)
            input(worksheet.cell(column=XLSX_STATUS,row=row).value)
            input(XLSX_STATUS)
            if worksheet.cell(column=XLSX_STATUS, row=row).value != "completed":
                continue
            input(worksheet.cell(column=XLSX_RESOURCE,row=row).value)
            if worksheet.cell(column=XLSX_RESOURCE,row=row).value!="Pavillon Gemeinschaftsraum - Kooperative Großstadt ":
                continue
            booking_id = worksheet.cell(column=XLSX_BOOKING_ID, row=row).value
            if worksheet.cell(column=XLSX_INTERNAL, row=row).value=="TRUE":
                booking_id += " (INTERNAL)"
            persons = worksheet.cell(column=XLSX_PERSONS, row=row).value
            start_date = worksheet.cell(column=XLSX_START_DATE, row=row).value + " " + worksheet.cell(column=XLSX_START_TIME, row=row).value
            end_date = worksheet.cell(column=XLSX_END_DATE, row=row).value + " " + worksheet.cell(column=XLSX_END_TIME, row=row).value
            desc = worksheet.cell(column=XLSX_DESCRIPTION, row=row).value
            writer.writerow([booking_id, start_date, end_date, desc, "Freihampton", persons])
            input([booking_id, start_date, end_date, desc, "Freihampton", persons])
    config = {
        'CSV_NAME': 0,
        'CSV_START_DATE': 1,
        'CSV_END_DATE': 2,
        'CSV_DESCRIPTION': 3,
        'CSV_LOCATION': 4,
        'CSV_PERSONS': 5,
        'CSV_DELIMITER': ","
    }
    convert.read_csv(ics_file_name + ".tmp.csv", config)
    for i in range(len(convert.csv_data)):
        convert.csv_data[i][config["CSV_START_DATE"]] = datetime.datetime.strptime(convert.csv_data[i][config["CSV_START_DATE"]], "%Y-%m-%d %H:%M")
        convert.csv_data[i][config["CSV_END_DATE"]] = datetime.datetime.strptime(convert.csv_data[i][config["CSV_END_DATE"]], "%Y-%m-%d %H:%M")
    convert.make_ical(config)
    convert.save_ical(ics_file_name)
