# -*- coding: utf-8 -*-

"""
visiopy.docprops

This module implements the Visio document properties files in docProps

:copyright: (c) 2016 by Mathijs Mortimer.
"""

import xml.etree.ElementTree as ET


class DocProps:
    def __init__(self):
        """Initialise the relationship"""
        pass

    def to_app_xml(self):
        """Generate app.xml from :class:`DocProps`

        :return: app.xml string
        """

        root = ET.Element('Properties', {'xmlns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
                                         'xmlns:vt': "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"})

        ET.SubElement(root, 'Template')

        application = ET.SubElement(root, 'Application')
        application.text = 'Microsoft Visio'

        scale_crop = ET.SubElement(root, 'ScaleCrop')
        scale_crop.text = 'false'

        heading_pairs = ET.SubElement(root, 'HeadingPairs')
        vt_vector1 = ET.SubElement(heading_pairs, 'vt:vector', {'size': '2', 'baseType': 'variant'})
        vt_variant1 = ET.SubElement(vt_vector1, 'vt:variant')
        item1 = ET.SubElement(vt_variant1, 'vt:lpstr')
        item1.text = 'Pages'
        vt_variant2 = ET.SubElement(vt_vector1, 'vt:variant')
        item2 = ET.SubElement(vt_variant2, 'vt:i4')
        item2.text = '2'

        titles_of_parts = ET.SubElement(root, 'TitlesOfParts')
        vt_vector2 = ET.SubElement(titles_of_parts, 'vt:vector', {'size': '2', 'baseType': 'lpstr'})
        item3 = ET.SubElement(vt_vector2, 'vt:lpstr')
        item3.text = 'MyFirstPage'
        item4 = ET.SubElement(vt_vector2, 'vt:lpstr')
        item4.text = 'MySecondPage'

        ET.SubElement(root, 'Manager')
        ET.SubElement(root, 'Company')
        links_up_to_date = ET.SubElement(root, 'LinksUpToDate')
        links_up_to_date.text = 'false'
        shared_doc = ET.SubElement(root, 'SharedDoc')
        shared_doc.text = 'false'
        ET.SubElement(root, 'HyperlinkBase')
        hyperlinks_changed = ET.SubElement(root, 'HyperlinksChanged')
        hyperlinks_changed.text = 'false'
        app_version = ET.SubElement(root, 'AppVersion')
        app_version.text = '15.0000'

        return ET.tostring(root, encoding='unicode')

    def to_core_xml(self):
        root = ET.Element('cp:coreProperties', {'xmlns:cp': "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", 
                                                'xmlns:dc': "http://purl.org/dc/elements/1.1/",
                                                'xmlns:dcterms': "http://purl.org/dc/terms/",
                                                'xmlns:dcmitype': "http://purl.org/dc/dcmitype/",
                                                'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance"})
        ET.SubElement(root, 'dc:title')
        ET.SubElement(root, 'dc:subject')
        ET.SubElement(root, 'dc:creator')
        ET.SubElement(root, 'cp:keywords')
        ET.SubElement(root, 'dc:description')
        ET.SubElement(root, 'cp:lastModifiedBy')
        ET.SubElement(root, 'cp:lastPrinted')
        created = ET.SubElement(root, 'dcterms:created', {'xsi:type': 'dcterms:W3CDTF'})
        created.text = '2016-08-29T06:42:32Z'
        modified = ET.SubElement(root, 'dcterms:modified', {'xsi:type': 'dcterms:W3CDTF'})
        modified.text = '2016-08-31T19:23:29Z'
        ET.SubElement(root, 'cp:category')
        language = ET.SubElement(root, 'dc:language')
        language.text = 'en-US'

        return ET.tostring(root, encoding='unicode')

    def to_custom_xml(self):
        # There's some documentation on this here: https://msdn.microsoft.com/en-us/library/windows/desktop/aa380374(v=vs.85).aspx
        root = ET.Element('Properties', {'xmlns': "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
                                         'xmlns:vt': "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"})

        ET.SubElement(root, 'Property')

        property1 = ET.SubElement(root, 'property', {'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
                                                     'pid': '2',
                                                     'name': "_VPID_ALTERNATENAMES"})
        value1 = ET.SubElement(property1, 'vt:lpwstr')

        property2 = ET.SubElement(root, 'property', {'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
                                                     'pid': '3',
                                                     'name': 'BuildNumberCreated'})
        value2 = ET.SubElement(property1, 'vt:i4')
        value2.text = '1006637809'

        property3 = ET.SubElement(root, 'property', {'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
                                                     'pid': '4',
                                                     'name': 'BuildNumberEdited'})
        value3 = ET.SubElement(property1, 'vt:i4')
        value3.text = '1006637809'

        property4 = ET.SubElement(root, 'property', {'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
                                                     'pid': '5',
                                                     'name': 'IsMetric'})
        value4 = ET.SubElement(property1, 'vt:boot')
        value4.text = 'true'

        property5 = ET.SubElement(root, 'property', {'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
                                                     'pid': '4',
                                                     'name': 'TimeEdited'})
        value5 = ET.SubElement(property1, 'vt:filetime')
        value5.text = '2016-08-31T19:23:09Z'

        return ET.tostring(root, encoding='unicode')

    def to_xml(self):
        """Return app.xml, core.xml and custom.xml strings"""
        return self.to_app_xml(), self.to_core_xml(), self.to_custom_xml()

    @classmethod
    def from_xml(cls, xml_file):
        """Generate class from existing document propery files

        :param xml_file: The xml file to parse
        """
        tree = ET.parse(xml_file)
        root = tree.getroot()
        # TODO: We should parse these files and fill the __init__().
        # for now only statically written the to_xml files
        return cls()
