import os
import argparse
from helpers import read_input_file, write_to_excel


def main():
    parser = argparse.ArgumentParser(description='Process Excel file')
    parser.add_argument(
        '--file', help='Name of the filename that you want to process')
    parser.add_argument(
        '--sheet', help='Name of the sheetname that you want to process')
    parser.add_argument(
        '--outputdir', help='The output directory you want to write to')
    parser.add_argument(
        '--outputfile', help='The output directory you want to write to')
    args = parser.parse_args()
    output_path = os.path.join(args.outputdir, args.outputfile)
    dataframe = read_input_file(args.file, args.sheet, to_df=True)
    write_to_excel(output_path, dataframe)


if __name__ == '__main__':
    main()
