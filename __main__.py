import sys
import os


if __package__ == 'pptx_autobook':
    from pptx_autobook.pptx_proc import pptx_get_outline, pptx_slide_export
    from pptx_autobook.arg_parse import get_arg
else:
    from pptx_proc import pptx_get_outline, pptx_slide_export
    from arg_parse import get_arg

if __name__ == '__main__':

    # Get program options
    args = get_arg()
    print(args)

    print('Hello World')
