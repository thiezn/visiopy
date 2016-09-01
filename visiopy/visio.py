#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import zipfile
import shutil
import os
from relationships import Relationship
from content_types import ContentTypes
from pages import PageCollection
from hacks import (WindowsProperties, DocumentProperties)
from docprops import DocProps


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

    def __init__(self,
                 page_collection=None,
                 package_rels=None,
                 document_rels=None,
                 content_types=None):

        self.page_collection = page_collection

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
        self.package_rels = package_rels
        self.document_rels = document_rels

        # Document properties
        self.doc_props = DocProps()
        self.windows_properties = WindowsProperties()
        self.document_properties = DocumentProperties()

        # Content_Types
        self.content_types = content_types

        # custom.xml data
        self.is_metric = True  # Using the metric system

    def to_file(self, filename, tmp_folder='./tmp'):
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
            # TODO the masters folder seems to be generated when you
            # create a connection. Lets tackle this when i've got
            # the basics worked out.
            # os.makedirs(tmp_folder + '/visio/masters')
            # os.makedirs(tmp_folder + '/visio/masters/_rels')
            # TODO the theme folder lets you put themes on your diagram?
            # not present when theres no theme customisation
            # os.makedirs(tmp_folder + '/visio/theme')
        else:
            raise IOError('Error writing to file {}. '
                          'tmp folder {} already exists'
                          .format(filename, tmp_folder))

        xml_decl = '<?xml version="1.0" encoding="utf-8" ?>'
        xml_decl_standalone = '<?xml version="1.0" encoding="utf-8" standalone="yes" ?>'

        # Create [content_Types].xml
        with open(tmp_folder + '/[Content_Types].xml', 'w') as f:
            f.write(xml_decl_standalone)
            f.write(self.content_types.to_xml())

        # Create docProps files
        app_xml, core_xml, custom_xml = self.doc_props.to_xml()

        with open('{}/docProps/app.xml'.format(tmp_folder), 'w') as f:
            f.write(xml_decl_standalone)
            f.write(app_xml)

        with open('{}/docProps/core.xml'.format(tmp_folder), 'w') as f:
            f.write(xml_decl_standalone)
            f.write(core_xml)

        with open('{}/docProps/custom.xml'.format(tmp_folder), 'w') as f:
            f.write(xml_decl_standalone)
            f.write(custom_xml)

        shutil.copy('thumbnail.emf', '{}/docProps/'.format(tmp_folder))

        # Write pages.xml and pages.xml.rels
        pages_xml, pages_xml_rels = self.page_collection.to_xml()

        with open('{}/visio/pages/_rels/pages.xml.rels'.format(tmp_folder), 'w') as f:
            f.write(xml_decl_standalone)
            f.write(pages_xml_rels)

        with open('{}/visio/pages/pages.xml'.format(tmp_folder), 'w') as f:
            f.write(xml_decl)
            f.write(pages_xml)

        # Write page?.xml and page?.xml.rels
        # TODO, page?.xml.rels not generated yet
        for page in self.page_collection.pages:
            with open('{}/visio/pages/{}'.format(tmp_folder, page.filename), 'w') as f:
                f.write(xml_decl)
                f.write(page.to_xml())

        # Create _rels files
        xml_declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

        with open('{}/_rels/.rels'.format(tmp_folder), 'w') as f:
            f.write(xml_decl_standalone)
            f.write(self.package_rels.to_xml())

        with open('{}/visio/_rels/document.xml.rels'.format(tmp_folder), 'w') as f:
            f.write(xml_decl_standalone)
            f.write(self.document_rels.to_xml())

        # Create visio document and window properties
        with open(tmp_folder + '/visio/windows.xml', 'w') as f:
            f.write(self.windows_properties.to_xml())

        with open(tmp_folder + '/visio/document.xml', 'w') as f:
            f.write(self.document_properties.to_xml())

        # Zip the tmp_folder contents and rename file to vsdx
        shutil.make_archive(filename, 'zip', tmp_folder)
        shutil.move(filename + '.zip', filename + '.vsdx')

        # Remove the temporary folder
        print('normally we remove the tmp folder but leave it for debuggin now')
        # shutil.rmtree(tmp_folder)

    @classmethod
    def from_file(cls, filename):
        # unzip vsdx file to temp. folder
        directory = './{}'.format(filename.rsplit('.', 1)[0])
        with zipfile.ZipFile(filename, "r") as zip_ref:
            zip_ref.extractall(directory)

        # Read [Content_Types].xml
        content_types = ContentTypes.from_xml(directory + '/[Content_Types].xml')

        # Read relationships
        package_rels = Relationship.from_xml('{}/_rels/.rels'.format(directory))
        document_rels = Relationship.from_xml('{}/visio/_rels/document.xml.rels'.format(directory))

        # Read pages and relationships
        page_collection = PageCollection.from_xml(directory, content_types)

        # Remove extracted folder again
        shutil.rmtree(directory)
        return cls(page_collection=page_collection,
                   package_rels=package_rels,
                   document_rels=document_rels,
                   content_types=content_types)

    def add_page(self, name):
        """Add a page to the document"""
        self.page_collection.add_page(name)


def main():
    filename = 'SimpleDrawingMultiplePages.vsdx'
    new_filename = 'mytmpfile'
    print('Loading diagram {}'.format(filename))
    diag = Document.from_file(filename)

    print('Adding empty page')
    diag.add_page('MyNewPage')

    print('Writing to file {}'.format(new_filename))
    diag.to_file(new_filename)


if __name__ == '__main__':
    main()
