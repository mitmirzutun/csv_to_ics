import converter


def main():
    import tkinter.filedialog
    import os
    if False:
        file_name = tkinter.filedialog.askopenfilename(title="Open file",
                                                       filetypes=[("CSV File", ".csv"),
                                                                  ("Carlender File", ".ics"),
                                                                  ("Exel 2010 File", ".xlsx")])
    else:
        file_name="data/Klink Buchungen 2025.xlsx"
    if file_name.endswith(".ics"):
        save_file_name = os.path.basename(file_name)
        save_file_name = ".".join(save_file_name.split(".")[:-1])+".csv"
        save_file_name = tkinter.filedialog.asksaveasfilename(defaultextension="csv",
                                                              filetypes=[("CSV Files", ".csv")],
                                                              title="Save file as",
                                                              initialfile=save_file_name)
        converter.ics_to_csv(file_name, save_file_name)
    if file_name.endswith(".xlsx"):
        initialfile = os.path.basename(file_name)
        initialfile = ".".join(initialfile.split(".")[:-1])+".ics"
        if False:
            save_file_name = tkinter.filedialog.asksaveasfilename(defaultextension="ics",
                                                                  filetypes=[("Carlender File",
                                                                              ".ics")],
                                                                  title="Save file as",
                                                                  initialfile=initialfile)
        else:
            save_file_name="Klink Buchungen 2025.ics"
        converter.xls_to_ics(file_name, save_file_name)


if __name__ == "__main__":
    main()
