import converter
TKINTER_SUCESSFUL: bool
try:
    import tkinter.filedialog
    TKINTER_SUCESSFUL = True
except ModuleNotFoundError:
    TKINTER_SUCESSFUL = False


def main():
    import os
    if TKINTER_SUCESSFUL:
        file_name = tkinter.filedialog.askopenfilename(title="Open file",
                                                       filetypes=[("CSV File", ".csv"),
                                                                  ("Carlender File", ".ics"),
                                                                  ("Exel 2010 File", ".xlsx")])
    else:
        file_name = input("File name: ")
    if file_name.endswith(".ics"):
        initialfile = os.path.basename(file_name)
        initialfile = ".".join(initialfile.split(".")[:-1]) + ".csv"
        save_file_name: str
        if TKINTER_SUCESSFUL:
            save_file_name = tkinter.filedialog.asksaveasfilename(defaultextension="csv",
                                                                  filetypes=[("CSV Files", ".csv")],
                                                                  title="Save file as",
                                                                  initialfile=initialfile)
        else:
            tmp = input(f".csv file name (default {initialfile}):")
            save_file_name = initialfile if tmp == "" else tmp
        converter.ics_to_csv(file_name, save_file_name)
    if file_name.endswith(".xlsx"):
        initialfile = os.path.basename(file_name)
        initialfile = ".".join(initialfile.split(".")[:-1]) + ".ics"
        save_file_name: str
        if TKINTER_SUCESSFUL:
            save_file_name = tkinter.filedialog.asksaveasfilename(defaultextension="ics",
                                                                  filetypes=[("Carlender File", ".ics")],
                                                                  title="Save file as",
                                                                  initialfile=initialfile)
        else:
            tmp = input(f".ics file name (default {initialfile}):")
            save_file_name = initialfile if tmp == "" else tmp
        converter.xls_to_ics(file_name, save_file_name)


if __name__ == "__main__":
    main()
