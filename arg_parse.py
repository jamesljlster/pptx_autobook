import argparse as argp


def get_arg():
    parser = argp.ArgumentParser()

    parser.add_argument('--pptx-src', type=str, help='PPTX source file path')
    parser.add_argument('--docx-out', type=str,
                        help='DOCX output file path. Overwrite if file is existed')

    args = parser.parse_args()
    return args
