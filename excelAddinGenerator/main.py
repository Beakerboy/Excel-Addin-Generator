import shutil, sys, os, zipfile

def main():
    if len(sys.argv) > 2:
        # check the extension on sys.argv[1] to determine which function to call
        input_file = sys.argv[1]
        output_file_name = sys.argv[2]
        if input_file.endswith('.xlam'):
            createFromZip(input_file, output_file_name)
        elif input_file.endswith('.bin'):
            createFromBin(input_file, os.path.dirname(sys.argv[0]) + '/../src/data', output_file_name)
        else:
            raise Exception(input_file, " is not a valid file format.")

def createFromBin(input_file, wrapper_dir, output_file_name):
    """Create a zip file containing the provided bin"""
    # check that inputfile has at least 8 bytes
    if os.path.getsize(input_file) < 8:
        raise Exception('Bin file too small')
    # file must start with 'd0 cf 11 e0 a1 b1 1a e1'
    fileSig = open(input_file, "rb").read(8).hex()
    if fileSig != 'd0cf11e0a1b11ae1':
        raise Exception('File signature {} is not as expected.', format(fileSig))
    shutil.move(input_file, wrapper_dir + "/xl/vbaProject.bin")
    shutil.make_archive(output_file, 'zip', wrapper_dir)
    shutil.move(output_file + ".zip", output_file_name)

def createFromZip(input_file, output_file_name):
    """Create a zip file containing the bin file within the provided zip file"""
    # check that input is a zip file
    if zipfile.is_zipfile(input_file):
        # check that the zip archive contains /xl/vbaProject.bin
        with zipfile.ZipFile(input_file, 'r') as zip:
          binFile = zip.read('xl/vbaProject.bin')
          createFromBin(binFile, os.path.dirname(sys.argv[0]) + '/../src/data', output_file_name)
    else:
        raise Exception(input_file, " is not a valid file format.")

if __name__ == "__main__":
    main()
