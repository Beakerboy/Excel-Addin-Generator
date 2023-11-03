import argparse
import os
import sys
import excel_addin_generator.main as gen


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input",
                        help="The input bin or xlam file.")
    parser.add_argument("output",
                        help="The output path and file name.")
    args = parser.parse_args()
    base_path = os.path.dirname(__file__)
    if len(args) > 2:
        # check the extension on sys.argv[1] to determine which function to call
        if args.input.endswith('.xlam'):
            gen.createFromZip(args.input, base_path + '/data', args.output)
        elif args.input.endswith('.bin'):
            gen.create_from_bin(args.input, base_path + '/data', args.output)
        else:
            raise Exception(args.input, " is not a valid file format.")


main()
