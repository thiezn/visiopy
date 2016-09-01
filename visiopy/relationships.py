# -*- coding: utf-8 -*-

"""
visiopy.relationships

This module implements the Visio package relationships

:copyright: (c) 2016 by Mathijs Mortimer.
"""

import xml.etree.ElementTree as ET


class Relationship:

    # TODO define __getitem__ classes so you can just iterate through this object
    # to get the relationships

    doc_schema = "http://schemas.openxmlformats.org/package/2006/relationships"

    def __init__(self):
        """Initialise the relationship"""
        self.rels = {}

    def add(self, rel_id, target, type):
        """Add a relationship to the document

        :param rel_id: A relationship id for example: rId1
        :param target: The page to refer for example: page1.xml
        :param type: The schema of the target for example:
                     http://schemas.microsoft.com/visio/2010/relationships/page
        """
        if rel_id not in self.rels:
            self.rels[rel_id] = (target, type)
        else:
            raise ValueError('rel_id {} already exists'.format(rel_id))

    def rm(self, rel_id):
        """Remove relationship from pages relationships"""
        del self.rels[rel_id]

    def to_xml(self):
        """Generate XML from current relationships"""

        root = ET.Element('Relationships', {'xmlns': self.doc_schema})

        for Id, (Target, Type) in self.rels.items():
            ET.SubElement(root, 'Relationship', {'Id': Id, 'Target': Target, 'Type': Type})

        return ET.tostring(root, encoding='unicode')

    @staticmethod
    def from_xml(xml_file):
        """Generate class from XML document

        :param xml_file: The xml file to parse
        """
        cls = Relationship()
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for page in root:
            cls.add(page.attrib['Id'], page.attrib['Target'], page.attrib['Type'])
        return cls
