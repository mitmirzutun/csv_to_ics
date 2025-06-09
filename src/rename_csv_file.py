
import csv
import datetime
import os
# rename csv file with first field 18.2.2003, last field "bank" to 20030218_bank.csv
def main():
    project_root_dir = os.path.dirname(os.path.dirname(__file__))
    data_in_dir = os.path.join(project_root_dir, "data_in")
    file_name = input("File name, in {} directory: ".format(data_in_dir))
    with open(os.path.join(data_in_dir, file_name), newline="") as csvfile:
        lines = list(csv.reader(csvfile, delimiter=";", lineterminator="\n"))
    # check: only two lines
    if len(lines) != 2:
        print("expecting csv with 2 lines")
        return
    eventdate = datetime.datetime.strptime(lines[1][0], "%d.%m.%Y")
    eventdate = eventdate.strftime("%Y%m%d")
    eventshortname = lines[1][8]
    new_filename = "{}_{}.csv".format(eventdate, eventshortname)
    os.popen("cp {} {}".format(os.path.join(data_in_dir, file_name), os.path.join(data_in_dir,new_filename) ))

if __name__ == "__main__":
    main()