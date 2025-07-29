import converter
import csv
import uuid

def main():
    import os
    file_name = input("File name: ")
    save_file_name: str
    if file_name.endswith(".ics"):
        initialfile = os.path.basename(file_name)
        initialfile = ".".join(initialfile.split(".")[:-1]) + ".csv"
        tmp = input(f".csv file name (default {initialfile}):")
        save_file_name = initialfile if tmp == "" else tmp
        converter.ics_to_csv(file_name, save_file_name)
    if file_name.endswith(".xlsx"):
        initialfile = os.path.basename(file_name)
        initialfile = ".".join(initialfile.split(".")[:-1]) + ".ics"
        tmp = input(f".ics file name (default {initialfile}):")
        save_file_name = initialfile if tmp == "" else tmp
        converter.xls_to_ics(file_name, save_file_name)
    if file_name.endswith(".csv"):
        initialfile = os.path.basename(file_name)
        with open(file_name, "r",) as fd:
            reader = csv.reader(fd)
            header = next(reader)
            for row in reader:
                pass
                id = uuid.uuid5(uuid.NAMESPACE_DNS, row[0]+row[2]+row[1])
                with open(str(id)+".csv", "w") as out:
                    writer = csv.writer(out)
                    writer.writerow(header)
                    writer.writerow(row)
                calender = converter.OVBuilder()
                calender.from_path(str(id)+".csv").to_ical("../data_out/"+str(id)+".ics")


if __name__ == "__main__":
    main()
