def main():
    import converter
    converter = (converter.CSVBuilder().number_of_headers(1).title_col(0)
                 .location_col(1).start_col(2).end_col(3))
    calender = converter.from_path("test.csv")
    calender.to_csv("test1.csv")
    calender.to_ical("test1.ics")


if __name__ == "__main__":
    main()
