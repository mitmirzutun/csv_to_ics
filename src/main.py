import converter

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
        outputdir = initialfile.split(".")[:-1] 
        initialfile = os.path.join(os.pardir, "data_out",
                                   ".".join(initialfile.split(".")[:-1])
                                   + ".ics")
        tmp = input(f".ics file name (default {initialfile}):")
        save_file_name = initialfile if tmp == "" else tmp
        calender = converter.OVBuilder().from_path(file_name)
        calender.to_ical(save_file_name)

if __name__ == "__main__":
    main()
