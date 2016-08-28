#!/usr/bin/env python3

import zipfile
import xml.etree.ElementTree as ET

ns = {'visio': 'http://schemas.microsoft.com/office/visio/2012/main'}


class Shape:
    """Contains a single shape object"""
    
    def __init__(self, xml_shape):
        """Initialise the shape into python object

        :param xml_shape: the shape from xml.etree.ElementTree
        """
        self.id = xml_shape.attrib['ID']
        self.type = xml_shape.attrib['Type']
        self.name = xml_shape.attrib.get('Name', '')

        self.x = None
        self.y = None

        # Very ugly parsing for the moment
        for item in xml_shape:
            if 'Cell' in item.tag:
                if item.attrib['N'] == 'PinX':
                    self.x = item.attrib['V']
                elif item.attrib['N'] == 'PinY':
                    self.y = item.attrib['V']

class Connect:
    """Contains a single connect object"""

    def __init__(self, xml_connect):
        """Initialise the connect into python object

        :param xml_connect: the connect from xml.etree.ElementTree
        """
        self.from_cell = xml_connect.attrib['FromCell']
        self.from_part = xml_connect.attrib['FromPart']
        self.from_sheet = xml_connect.attrib['FromSheet']
        self.to_cell = xml_connect.attrib['ToCell']
        self.to_part = xml_connect.attrib['ToPart']
        self.to_sheet = xml_connect.attrib['ToSheet']

class Page:
    """Holds a single Visio page"""

    def __init__(self, xml_page):
        """Initialises a page from xml data

        :param xml_page: the page from xml.etree.ElementTree
        """
        tree = ET.parse(xml_page)
        root = tree.getroot()

        self.shapes = []
        self.connects = []

        for shapes in root.findall('visio:Shapes', ns):
            for shape in shapes:
                self.shapes.append(Shape(shape))

        for connects in root.findall('visio:Connects', ns):
            for connect in connects:
                self.connects.append(Connect(connect))


class Diagram:

    def __init__(self, filename=None):
        self.unzip(filename)
        ns = {'visio': 'http://schemas.microsoft.com/office/visio/2012/main'}

        self.pages = []
        self.pages.append(Page(self.full_directory + '/visio/pages/page1.xml'))

    def unzip(self, filename, targetdir='.'):
        self.full_directory = '{}/{}'.format(targetdir, filename.rsplit('.', 1)[0])
        print('Extracting {} to directory {}'.format(filename, self.full_directory))
        with zipfile.ZipFile(filename, "r") as zip_ref:
           zip_ref.extractall(self.full_directory)


def main():
   filename = 'network_diag_detailed.vsdx'
   print('Loading diagram {}'.format(filename))
   diag = Diagram(filename)

   for page in diag.pages:
       for shape in page.shapes:
           print('ID: {}, type: {}, name: {}, x, y: {}x{}'.format(shape.id,
                                                                  shape.type,
                                                                  shape.name,
                                                                  shape.x,
                                                                  shape.y))


if __name__ == '__main__':
    main()
