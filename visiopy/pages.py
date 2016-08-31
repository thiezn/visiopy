# -*- coding: utf-8 -*-

"""
visiopy.pages

This module handles Visio pages

:copyright: (c) 2016 by Mathijs Mortimer.
"""

import xml.etree.ElementTree as ET
from relationships import Relationship


class PageCollection:
    """Holds a collection of :class:`Page` classes including their properties
    and relationships"""

    def __init__(self, rels=None, pages=None):
        """Initialise pages

        :param rels: Instance of :class:`Relationship` from pages.xml.rels
        :param pages: List of :class:`Page` classes
        """
        self.rels = rels  # holds relationships to pages
        self.pages = pages  # holds instances of :class:`Page` class

    @classmethod
    def from_xml(cls, dir):
        """Generate PageCollection from files

        :param dir: The directory of the extracted visio package
        """

        rel_dir = '{}/visio/pages/_rels/'.format(dir)
        page_dir = '{}/visio/pages/'.format(dir)

        rels = Relationship.from_xml(rel_dir + 'pages.xml.rels')
        pages = [Page(page_dir + page_filename)
                 for _, (page_filename, _) in rels.rels.items()]

        return cls(rels, pages)

    def to_xml(self):
        """Generate XML data for pages

        :return: strings containing (pages.xml, pages.xml.rels)
        """
        root = ET.Element('Pages', {'xmlns': 'http://schemas.microsoft.com/office/visio/2012/main',
                                    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                                    'xml:space': 'preserve'})

        for page in self.pages:
            xml_page = ET.SubElement(root,
                                     'Page',
                                     {'ID': page.id,
                                      'NameU': page.name,
                                      'IsCustomNameU': '1',
                                      'Name': page.name,
                                      'IsCustomName': '1',
                                      'ViewScale': '0.82',
                                      'ViewCenterX': '4.1275082550165',
                                      'ViewCenterY': '8.5852171704343'})
            xml_pagesheet = ET.SubElement(xml_page,
                                          'PageSheet',
                                          {'LineStyle': '0',
                                           'FillStyle': '0',
                                           'TextStyle': '0'})
            # TODO Add all data within pagesheet
            ET.SubElement(xml_page, 'Rel', {'r:id': page.rel_id})

        pages_xml = ET.tostring(root, encoding='unicode')
        pages_xml_rels = self.rels.to_xml()

        return pages_xml, pages_xml_rels


class Page:
    """Holds a single Visio page"""

    ns = {'visio': 'http://schemas.microsoft.com/office/visio/2012/main'}

    def __init__(self, shapes=None, connects=None, rels=None):
        """Initialises a page

        :param shapes: List of :class:`Shape` classes
        :param connects: List of :class:`Connect` classes
        :param rels: Instance of :class:`Relationship` from page?.xml.rels
        """
        self.name = 'HAVETOPARSETHISCANBEEMPTY'
        self.id = 'HAVETOPARSETHISSHOULDBE0'
        self.rel_id = 'HAVETOPARSESHOULDBErId0'
        self.shapes = shapes
        self.connects = connects
        self.rels = rels

    def to_xml(self, xml_file):
        """Writes Page object to page?.xml file"""
        # TODO: the to_xml functions all around shouldn't write to file
        # but just spit out the xml as a string. Then I can have this
        # function for each object like Shape, Connect, etc. This will
        # allow me to inject these objects into the Page.to_xml()
        # function to create the full picture
        # pheww, this is going to be a lot of work until i figured
        # all the types out!
        pass

    @classmethod
    def from_xml(cls, xml_file):
        """Create a Page object from an existing xml_file"""

        shapes = []
        connects = []

        # Parse main page?.xml file
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for items in root.findall('visio:Shapes', cls.ns):
            for item in items:
                shapes.append(Shape.from_xml(item))

        for items in root.findall('visio:Connects', cls.ns):
            for item in items:
                connects.append(Connect.from_xml(item))

        # Parse _rels/page?.xml.rels file
        dir, filename = xml_file.rsplit('/', 1)
        rels = Relationship.from_xml('{}/_rels/{}.rels'
                                     .format(dir, filename))

        return cls(shapes=shapes, connects=connects, rels=rels)


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
            setattr(cls, key, value)

        # Parse all shape data
        for item in xml_shape:
            if 'Cell' in item.tag:
                setattr(cls,
                        item.attrib['N'],
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
            setattr(cls, key, value)

        return cls
