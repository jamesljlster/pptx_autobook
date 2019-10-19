from win32com.client import Dispatch, constants
from os.path import realpath
from lxml import etree


def pptx_slide_get_id_list(sldRoot):

    ret = []

    # Get slide id list
    sldIdLst = sldRoot.findall('.//{*}sldId')
    for i in range(len(sldIdLst)):
        ret.append(sldIdLst[i].get('id'))

    return ret


def pptx_slide_get_section_info(sectNode):

    # Get section name and child slide id list
    return {
        'name': sectNode.get('name'),
        'list': pptx_slide_get_id_list(sectNode)
    }


def pptx_slide_get_list(pptXml):

    sldList = []

    # Find section node
    ret = pptXml.findall('.//{*}section')
    if len(ret) > 0:
        for child in ret:
            sldList.append(pptx_slide_get_section_info(child))

    else:
        sldList.append({
            'name': None,
            'list': pptx_slide_get_id_list(pptXml)
        })

    return sldList


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
    ret = pptx_slide_get_id_list(sldRoot)
    for i in range(len(ret)):
        sldIdMap[ret[i]] = i + 1

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
    sldIdMap = pptx_slide_get_id_index_map(pptXml)
    print('Slide ID Map:', sldIdMap)
    print()

    slideList = pptx_slide_get_list(pptXml)
    for sld in slideList:
        print('===', sld['name'], '===')
        for sldId in sld['list']:
            print('ID:', sldId, 'Index:', sldIdMap[sldId])
        print()

    pptx_slide_export(fPath, './pptx_file/' + str(uuid.uuid4()), 'PNG')
