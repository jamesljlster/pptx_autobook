import uuid

from win32com.client import Dispatch, constants
from os.path import realpath


def pptx_slide_export(pptPath, imgDir, format):

    # Open PowerPoint application
    ppt = Dispatch('PowerPoint.Application')
    ppt.Visible = 1

    # Open presentation session
    pptSess = ppt.Presentations.Open(realpath(pptPath))

    # Export slides as images
    pptSess.Export(realpath(imgDir), format)

    # Close application
    ppt.Quit()


# Test
if __name__ == '__main__':

    pptx_slide_export('./pptx_file/test.pptx',
                      './pptx_file/' + str(uuid.uuid4()), 'PNG')
