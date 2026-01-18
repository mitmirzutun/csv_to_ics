import converter
import argparse


def main():
    conv = converter.OVConvert()
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input")
    parser.add_argument("-o", "--output")
    args = parser.parse_args()
    print(args.input, args.output)
    input_file: str
    if args.input is None:
        input_file = input("Path to input file: ")
    else:
        input_file = args.input
    timezone_string = input("Timezone (default=CEST): ")
    if timezone_string == "":
        timezone_string = "CEST"
    conv.read_csv(input_file, csv_configs=converter.HTWK_DEFAULT_CONFIG)
    conv.make_ical(csv_configs=converter.HTWK_DEFAULT_CONFIG)
    output_file: str
    if args.output is None:
        title = conv.get_title().replace(" ", "_")
        start_date = conv.get_start_date().strftime("%Y-%m-%dT%H%M")
        default = f"data_out/{start_date}_{title}.ics"
        output_file = input(f"Path to output file (default={default}): ")
        if output_file == "":
            output_file = default
    else:
        output_file = args.output
    conv.save_ical(output_file)
    conv.save_csv(input_file)


if __name__ == "__main__":
    main()
