#!/usr/bin/env python3

import zipfile
import xml.etree.ElementTree as ET
import re
import shutil
import os
from tmphacks import (VisioRelationships, DocumentRelationships,
                      WindowsProperties, DocumentProperties, ContentTypes)

ns = {'visio': 'http://schemas.microsoft.com/office/visio/2012/main'}


def from_camel(name):
    """Converts CamelCase to snake_case format

    :param name: A CamelCase string to convert
    """
    s1 = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', name)
    return re.sub('([a-z0-9])([A-Z])', r'\1_\2', s1).lower()


def to_camel(name):
    """Converts snake_case to CamelCase

    :param name: A snake_case string to convert
    """
    return name.title().replace("_", "")


class Shape:
    """Contains a single shape object"""

    def __init__(self):
        pass

    @staticmethod
    def from_xml(xml_shape):
        """Initialise the shape into python object

        :param xml_shape: the shape from xml.etree.ElementTree
        """

        # TODO: Ugly parsing at the moment
        #       we should probably check for certain fields etc

        # Parse the main shape attributes

        cls = Shape()
        for key, value in xml_shape.items():
            setattr(cls, from_camel(key), value)

        # Parse all shape data
        for item in xml_shape:
            if 'Cell' in item.tag:
                setattr(cls,
                        from_camel(item.attrib['N']),
                        item.attrib['V'])
            elif 'Section' in item.tag:
                pass
            elif 'Shapes' in item.tag:
                pass

        return cls


class Connect:
    """Contains a single connect object"""

    def __init__(self):
        pass

    @staticmethod
    def from_xml(xml_connect):
        """Initialise the connect into python object

        :param xml_connect: the connect from xml.etree.ElementTree
        """
        cls = Connect()
        # Parse the main shape attributes
        for key, value in xml_connect.items():
            # TODO decide if we want to convert the value from_camel
            setattr(cls, from_camel(key), value)

        return cls


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
                self.shapes.append(Shape.from_xml(shape))

        for connects in root.findall('visio:Connects', ns):
            for connect in connects:
                self.connects.append(Connect.from_xml(connect))


class Document:
    """Class holding a visio (*.vsdx) document

    The *.vsdx format is a container based on the Open Packaging Conventions
    (OPC) introduced by Microsoft. Its a zip file that contains various files
    together building a visio document.

    This is the tree view of an empty *.vsdx file

    .
    |- _rels/
    |  |- .rels
    |
    |- docProps/
    |  |- app.xml
    |  |- core.xml
    |  |- custom.xml
    |  |- thumbnail.emf
    |
    |- visio/
    |  |- _rels
    |  |  |- document.xml.rels
    |  |
    |  |- pages
    |  |  |- _rels
    |  |  |  |- pages.xml.rels
    |  |  |
    |  |  |- page1.xml
    |  |  |- pages.xml
    |  |
    |  |- document.xml
    |  |- windows.xml
    |
    |- [Content_Types].xml
    """

    def __init__(self, pages=None):
        self.pages = pages

        # apps.xml data
        self.company = ''
        self.manager = ''

        # core.xml data
        self.creator = ''
        self.language = 'en-US'
        self.keywords = []
        self.title = ''
        self.subject = ''
        self.description = ''

        # Relationships
        self.visio_rels = VisioRelationships()
        self.document_rels = DocumentRelationships()

        # Document properties
        self.windows_properties = WindowsProperties()
        self.document_properties = DocumentProperties()

        # Content_Types
        self.content_types = ContentTypes()

        # custom.xml data
        self.is_metric = True  # Using the metric system

    def to_file(self, filename, tmp_folder='./tmp_folder'):
        """Writes visio diagram to file

        :param filename: The filename to write to
        """
        if filename.endswith('.vsdx'):
            filename.strip('.vsdx')

        # Create the directory structure
        if not os.path.exists(tmp_folder):
            os.makedirs(tmp_folder)
            os.makedirs(tmp_folder + '/_rels')
            os.makedirs(tmp_folder + '/docProps')
            os.makedirs(tmp_folder + '/visio')
            os.makedirs(tmp_folder + '/visio/_rels')
            os.makedirs(tmp_folder + '/visio/pages')
            os.makedirs(tmp_folder + '/visio/pages/_rels')
            os.makedirs(tmp_folder + '/visio/masters')
            os.makedirs(tmp_folder + '/visio/masters/_rels')
            os.makedirs(tmp_folder + '/visio/theme')

        else:
            raise IOError('Error writing to file {}. '
                          'tmp folder {} already exists'
                          .format(filename, tmp_folder))

        # Add the neccesary files
        # TODO:
        #     write pages.xml/page1.xml/.. files
        #     write docOps files app.xml, core.xml, custom.xml, thumbnail.emf
        #     write _rels files
        #     write visio document.xml and window.xml files
        #     write [Content-Types].xml file
        #     Convert all data from snake to CamelCase again
        #     Do i have to set the RecalcDocument in custom.xml to True?

        # Create [content_Types].xml
        with open(tmp_folder + '/[Content_Types].xml', 'w') as f:
            f.write(self.content_types.get_xml())

        # Create _rels files
        with open(tmp_folder + '/visio/_rels/document.xml.rels', 'w') as f:
            f.write(self.visio_rels.get_xml())

        with open(tmp_folder + '/_rels/.rels', 'w') as f:
            f.write(self.document_rels.get_xml())

        # Create visio document and window properties
        with open(tmp_folder + '/visio/windows.xml', 'w') as f:
            f.write(self.windows_properties.get_xml())

        with open(tmp_folder + '/visio/document.xml', 'w') as f:
            f.write(self.document_properties.get_xml())

        # Zip the tmp_folder contents and rename file to vsdx
        shutil.make_archive(filename, 'zip', tmp_folder)
        shutil.move(filename + '.zip', filename + '.vsdx')

        # Remove the temporary folder
        shutil.rmtree(tmp_folder)

    @classmethod
    def from_file(cls, filename):
        # unzip vsdx file to temp. folder
        directory = './{}'.format(filename.rsplit('.', 1)[0])
        with zipfile.ZipFile(filename, "r") as zip_ref:
            zip_ref.extractall(directory)

        # Read pages
        pages_rels = get_pages_rels('{}/visio/pages/_rels/pages.xml.rels'
                                    .format(directory))

        pages = [Page('{}/visio/pages/{}'.format(directory, page['Target']))
                 for page in pages_rels]

        # Remove extracted folder again
        shutil.rmtree(directory)
        return cls(pages=pages)


def get_pages_rels(xml_file):
    """Returns all page relations from pages.xml.rels"""
    tree = ET.parse(xml_file)
    root = tree.getroot()
    return [page.attrib for page in root]


def main():
    filename = 'SimpleDrawing.vsdx'
    print('Loading diagram {}'.format(filename))
    diag = Document.from_file(filename)

    print('Writing to file')
    diag.to_file('mytmpfile')


if __name__ == '__main__':
    main()
