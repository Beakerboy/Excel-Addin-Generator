import os
import sys
import excel_addin_generator.main as gen


def main(args: list) -> None:

    if len(args) > 2:
        # check the extension on sys.argv[1] to determine which function to call
        input_file = args[1]
        output_file_name = args[2]
        if input_file.endswith('.xlam'):
            gen.createFromZip(input_file, args[0] + '/data', output_file_name)
        elif input_file.endswith('.bin'):
            gen.create_from_bin(input_file, args[0] + '/data', output_file_name)
        else:
            raise Exception(input_file, " is not a valid file format.")

args = []
args[0] = os.path.dirname(sys.argv[0])
args[1] = sys.argv[1]
args[2] = sys.argv[2]
main(args)
