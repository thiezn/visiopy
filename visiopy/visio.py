#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import zipfile
import shutil
import os
from relationships import Relationship
from pages import Page, PageCollection
from hacks import (VisioRelationships, WindowsProperties,
                   DocumentProperties, ContentTypes)


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
                 document_rels=None):

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
        self.windows_properties = WindowsProperties()
        self.document_properties = DocumentProperties()

        # Content_Types
        self.content_types = ContentTypes()

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
            os.makedirs(tmp_folder + '/visio/masters')
            os.makedirs(tmp_folder + '/visio/masters/_rels')
            # TODO the theme folder lets you put themes on your diagram?
            # not present when theres no theme customisation
            os.makedirs(tmp_folder + '/visio/theme')

        else:
            raise IOError('Error writing to file {}. '
                          'tmp folder {} already exists'
                          .format(filename, tmp_folder))

        # Add the neccesary files
        # TODO:
        #     write pages.xml/page1.xml/.. files
        #     write docOps files app.xml, core.xml, custom.xml, thumbnail.emf
        #     DONE: write _rels files
        #     write visio document.xml and window.xml files
        #     write [Content-Types].xml file
        #     Convert all data from snake to CamelCase again
        #     Do i have to set the RecalcDocument in custom.xml to True?

        xml_decl = '<?xml version="1.0" encoding="utf-8" ?>'
        xml_decl_standalone = '<?xml version="1.0" encoding="utf-8" standalone="yes" ?>'

        # Create [content_Types].xml
        with open(tmp_folder + '/[Content_Types].xml', 'w') as f:
            f.write(self.content_types.to_xml())

        # Write pages.xml and pages.xml.rels
        # TODO: need to extract all the page?.xml and related page?.xml.rels
        pages_xml, pages_xml_rels, page_xml_list = self.page_collection.to_xml()

        with open('{}/visio/pages/_rels/pages.xml.rels'.format(tmp_folder), 'w') as f:
            f.write(xml_decl_standalone)
            f.write(pages_xml_rels)

        with open('{}/visio/pages/pages.xml'.format(tmp_folder), 'w') as f:
            f.write(xml_decl)
            f.write(pages_xml)

        for page_xml in page_xml_list:
            with open('{}/visio/pages/needtofigureoutfilename.page?.xml'.format(tmp_folder), 'w') as f:
                f.write(xml_decl)
                f.write(page_xml)

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

        # Read relationships
        package_rels = Relationship.from_xml('{}/_rels/.rels'.format(directory))
        document_rels = Relationship.from_xml('{}/visio/_rels/document.xml.rels'.format(directory))

        # Read pages and relationships
        page_collection = PageCollection.from_xml(directory)

        # Remove extracted folder again
        shutil.rmtree(directory)
        return cls(page_collection=page_collection,
                   package_rels=package_rels,
                   document_rels=document_rels)


def main():
    filename = 'SimpleDrawingMultiplePages.vsdx'
    print('Loading diagram {}'.format(filename))
    diag = Document.from_file(filename)

    print('Writing to file')
    diag.to_file('mytmpfile')


if __name__ == '__main__':
    main()
