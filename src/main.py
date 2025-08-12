import converter


def main():
    conv=converter.OVConvert()
    conv.read_csv("data_in/ov_termine.csv")
    conv.make_ical()
    conv.save_ical("data_out/ov_termine.ics")
    conv.save_csv("data_out/ov_termine.csv")


if __name__ == "__main__":
    main()
