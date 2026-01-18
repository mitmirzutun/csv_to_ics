import converter
import argparse


def main():
    conv = converter.Converter(converter.HTWK_DEFAULT_CONFIG)
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input")
    parser.add_argument("-tz", "--time-zone")
    args = parser.parse_args()
    print("This program will generate a file per event if given a folder")
    input_file: str
    if args.input is None:
        input_file = input("Path to input file: ")
    else:
        input_file = args.input
    output_dir: str
    output_dir = input("Path to output (folder or file, DEFAULT=data_out/)")
    if output_dir == "":
        output_dir = "data_out/"
    timezone: str
    if args.time_zone is None:
        timezone = input("Timezone (DEFAULT=CEST): ")
        if timezone == "":
            timezone = "CEST"
    else:
        timezone = args.time_zone
    conv.change_config({"INPUT_TZ":timezone})
    conv.read_csv(input_file)
    conv.make_ical()
    conv.save_ical(output_dir)


if __name__ == "__main__":
    main()
