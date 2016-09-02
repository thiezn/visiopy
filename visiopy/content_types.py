# -*- coding: utf-8 -*-

"""
visiopy.content_types

This module handles the definitions in [Content_Types.xml]

:copyright: (c) 2016 by Mathijs Mortimer.
"""

import xml.etree.ElementTree as ET


class ContentTypes:
    """Holds a collection of :class:`Page` classes including their properties
    and relationships"""

    def __init__(self, **kwargs):
        """Initialise ContentTypes

        :param defaults: Default content types
        :param overrides: Override content types
        """
        self.defaults = kwargs.get('defaults', {})
        self.overrides = kwargs.get('overrides', {})

        # Add the default visio file content types
        # TODO: Have to decide if we want to be allowed to overwrite these by keyword args?
        self.add_default('emf', 'image/x-emf')
        self.add_default('rels', 'application/vnd.openxmlformats-package.relationships+xml')
        self.add_default('xml', 'application/xml')
        self.add('/visio/document.xml', 'application/vnd.ms-visio.drawing.main+xml')
        self.add('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml')
        self.add('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')
        self.add('/docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml')
        self.add('/visio/pages/pages.xml', 'application/vnd.ms-visio.pages+xml')
        self.add('/visio/windows.xml', 'application/vnd.ms-visio.windows+xml')

    def add(self, part_name, content_type):
        """Add a item to the Types

        :param part_name: Name of the part. e.g. '/visio/pages/page1.xml'
        :param content_type:  e.g. 'application/vnd.ms-visio.page+xml'
        """
        self.overrides[part_name] = content_type

    def add_default(self, extension, content_type):
        """Add a default item to Types

        :param extension: The filename extension. e.g. xml or rels
        :param content_type: The ContentType. e.g. 'image/x-emf'
        """
        self.defaults[extension] = content_type

    def rm(self, part_name):
        """Remove a part from the overrides"""
        del self.overrides[part_name]

    def rm_default(self, extension):
        """Remove a part from the defaults"""
        del self.defaults[extension]

    def to_xml(self):
        """Generate XML data for [Content_Types].xml

        :return: XML string
        """
        root = ET.Element('Types', {'xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types'})

        for key, value in self.defaults.items():
            ET.SubElement(root, 'Default', {'Extension': key, 'ContentType': value})

        for key, value in self.overrides.items():
            ET.SubElement(root, 'Override', {'PartName': key, 'ContentType': value})

        return ET.tostring(root, encoding='unicode')

    @classmethod
    def from_xml(cls, xml_file):
        """Generate ContentTypes from xml

        :param xml_file: the location of the [Content_Types].xml file
        """
        tree = ET.parse(xml_file)
        root = tree.getroot()

        defaults = {}
        overrides = {}

        for item in root:
            if item.tag == '{http://schemas.openxmlformats.org/package/2006/content-types}Default':
                defaults[item.attrib['Extension']] = item.attrib['ContentType']
            if item.tag == '{http://schemas.openxmlformats.org/package/2006/content-types}Override':
                overrides[item.attrib['PartName']] = item.attrib['ContentType']

        return cls(defaults=defaults, overrides=overrides)
