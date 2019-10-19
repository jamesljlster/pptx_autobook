from win32com.client import Dispatch, constants
from os.path import realpath
from lxml import etree


def pptx_slide_get_id_index_map(pptXml):

    sldIdMap = {}

    # Get slide list node
    sldRoot = None
    ret = pptXml.xpath('/p:presentation/p:sldIdLst',
                       namespaces=pptXml.nsmap)
    if len(ret) != 1:
        raise RuntimeError('Broken PPTX file?!')
    else:
        sldRoot = ret[0]

    # Get slide id list
    sldIdLst = sldRoot.findall('.//p:sldId', namespaces=pptXml.nsmap)
    for i in range(len(sldIdLst)):
        sldIdMap[sldIdLst[i].get('id')] = i + 1

    return sldIdMap


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

    import uuid

    from zipfile import ZipFile

    fPath = './pptx_file/test.pptx'

    zFile = ZipFile(fPath)
    pptXml = etree.fromstring(zFile.read('ppt/presentation.xml'))
    print('Slide ID Map:', pptx_slide_get_id_index_map(pptXml))

    pptx_slide_export(fPath, './pptx_file/' + str(uuid.uuid4()), 'PNG')
