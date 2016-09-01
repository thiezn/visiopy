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

    def __init__(self, content_types, rels=None, pages=None):
        """Initialise pages

        :param content_types: Instance of :class:`ContentType`
        :param rels: Instance of :class:`Relationship` from pages.xml.rels
        :param pages: List of :class:`Page` classes
        """
        self.content_types = content_types
        self.rels = rels
        self.pages = pages

    def add_page(self, name):
        """Add a page to the collection"""

        filename = 0
        rel_id = 0
        id = 0

        for page in self.pages:
            if int(page.rel_id.lstrip('rId')) >= rel_id:
                rel_id = int(page.rel_id.strip('rId'))
            if int(page.filename.lstrip('page').strip('.xml')) >= filename:
                filename = int(page.filename.lstrip('page').strip('.xml'))
            if int(page.id) >= id:
                id = int(page.id)

        filename += 1
        rel_id += 1
        id += 1

        filename = 'page{}.xml'.format(filename)
        rel_id = 'rId{}'.format(rel_id)
        id = str(id)

        self.rels.add(rel_id, filename, 'http://schemas.microsoft.com/visio/2010/relationships/page')
        self.content_types.add('/visio/pages/{}'.format(filename), 'application/vnd.ms-visio.page+xml')
        self.pages.append(Page(filename,
                               name,
                               id,
                               rel_id))

    @classmethod
    def from_xml(cls, dir, content_types):
        """Generate PageCollection from files

        :param dir: The directory of the extracted visio package
        :param content_types: Instance of :class:`ContentType`
        """

        rel_dir = '{}/visio/pages/_rels/'.format(dir)
        page_dir = '{}/visio/pages/'.format(dir)
        pages = []

        rels = Relationship.from_xml(rel_dir + 'pages.xml.rels')

        # Parse pages.xml for page info
        tree = ET.parse(page_dir + 'pages.xml')
        root = tree.getroot()

        for child in root:
            name = child.attrib.get('Name', '')
            id = child.attrib['ID']
            # TODO Parse these namespaces properly
            rel_id = child[1].attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            pages.append(Page.from_xml(page_dir + rels.rels[rel_id][0],
                                       name, id, rel_id))

        return cls(content_types, rels, pages)

    def to_xml(self):
        """Generate XML data for pages

        :return: XML string for (pages.xml, pages.xml.rels)
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

            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'PageWidth', 'V': '8.26771653543307'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'PageHeight', 'V': '11.69291338582677'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'ShdwOffsetX', 'V': '0.1181102362204724'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'ShdwOffsetY', 'V': '-0.1181102362204724'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'PageScale', 'V': '0.03937007874015748', 'U': 'MM'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'DrawingScale', 'V': '0.03937007874015748', 'U': 'MM'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'DrawingSizeType', 'V': '0'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'DrawingScaleType', 'V': '0'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'InhibitSnap', 'V': '0'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'PageLockReplace', 'V': '0', 'U': 'BOOL'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'PageLockDuplicate', 'V': '0', 'U': 'BOOL'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'UIVisibility', 'V': '0'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'ShdwType', 'V': '0'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'ShdwObliqueAngle', 'V': '0'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'ShdwScaleFactor', 'V': '1'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'DrawingResizeType', 'V': '1'})
            ET.SubElement(xml_pagesheet, 'Cell', {'N': 'PageShapeSplit', 'V': '1'})

            ET.SubElement(xml_page, 'Rel', {'r:id': page.rel_id})

        pages_xml = ET.tostring(root, encoding='unicode')
        pages_xml_rels = self.rels.to_xml()

        return pages_xml, pages_xml_rels


class Page:
    """Holds a single Visio page"""

    ns = {'visio': 'http://schemas.microsoft.com/office/visio/2012/main'}

    def __init__(self, filename, name=None, id=None, rel_id=None, shapes=None, connects=None):
        """Initialises a page

        :param shapes: List of :class:`Shape` classes
        :param connects: List of :class:`Connect` classes
        """
        self.filename = filename
        self.name = name
        self.id = id
        self.rel_id = rel_id
        self.shapes = shapes
        self.connects = connects

    def to_xml(self):
        """Writes Page object to page?.xml file"""
        # TODO: the to_xml functions all around shouldn't write to file
        # but just spit out the xml as a string. Then I can have this
        # function for each object like Shape, Connect, etc. This will
        # allow me to inject these objects into the Page.to_xml()
        # function to create the full picture
        # pheww, this is going to be a lot of work until i figured
        # all the types out!

        root = ET.Element('PageContents', {'xmlns': 'http://schemas.microsoft.com/office/visio/2012/main',
                                           'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                                           'xml:space': 'preserve'})
        if self.shapes:
            xml_shapes = ET.SubElement(root, 'Shapes')
            for shape in self.shapes:
                xml_shapes.append(ET.fromstring(shape.to_xml()))

        if self.connects:
            xml_connects = ET.SubElement(root, 'Connects')
            for connect in self.connects:
                xml_connects.append(ET.fromstring(connect.to_xml()))

        return ET.tostring(root, encoding='unicode')

    @classmethod
    def from_xml(cls, xml_file, name, id, rel_id):
        """Create a Page object from an existing xml_file"""

        if '/' in xml_file:
            filename = xml_file.rsplit('/', 1)[1]
        else:
            filename = xml_file

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

        return cls(filename, name, id, rel_id, shapes, connects)


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

    def to_xml(self):
        # TODO stop faking it baby!
        return "<Shape ID='1' Type='Shape' LineStyle='3' FillStyle='3' TextStyle='3' />"


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

    def to_xml(self):
        pass
