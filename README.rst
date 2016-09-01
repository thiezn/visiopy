visiopy
=======

Handling Microsoft Visio 2013 vsdx files from Python

**WARNING, DO NOT USE. Very much a work in progress**

A nice short description of the format found here:
http://www.digitalpreservation.gov/formats/fdd/fdd000021.shtml

The Visio VSDX Graphics File Format, developed and documented to a substantial degree by Microsoft, carries data that represents diagrams that employ vector graphics and supplementary information related to the creation, modification, and review of a collection of related diagrams. Typical domains of usage include: flowcharts; database schema models; organization charts; plans for building wiring, furniture layout, etc. A VSDX document represents a collection of Drawing Pages, Masters, Shapes, Images, Comments, Data Connections, and recalculation information relating to dynamic data connections. Starting with Visio 2013, this format is the default format for the Visio products for creating diagrams. The VSDX file can also be rendered as a Web Drawing through SharePoint (with Visio Services enabled) starting with SharePoint 2013 and also through an Office 365 plan that includes Visio Services.

The VSDX format is related to the OOXML family of formats, in that a VSDX file is a ZIP-based package that conforms to the Open Packaging Conventions as specified in ISO/IEC29500-2:2011 (see OPC/OOXML_2012), the further packaging restrictions for OOXML documents as specified in clause 9 of ISO/IEC29500-1:2011, and the MS-VSDX documentation.

The conceptual structure of a minimal Visio Drawing includes:
• VisioDocument - a collection of resources including a sequence of drawing pages
• Masters - a collection of page layouts shared by several pages, typically as framework or background to elements specific to each page
• Pages - graphical elements and layout for individual pages
• Shapes - individual graphical elements and their properties. Shapes may be incorporated into pages or masters. See Notes below for more about properties for shapes.
• Data connections - specifying how data can be retrieved from external sources to affect various aspects of a drawing, inlcuding its visual appearance.

The package structure for a VSDX file may have many parts, including:
• document.xml -- part that holds properties of the drawing needed for rendering or editing, such as style sheets, colors, and fonts used in the drawing.
• masters/masters.xml -- list of master page layouts
• masters/master1.xml, masters/master2.xml, ... -- individual master page layouts
• pages/pages.xml -- list of drawing pages 
• pages/page1.xml, pages/page2.xml, ... layouts and graphical elements for individual pages
• data/connections.xml -- specifications for any connections to remote data
• data/recordsets.xml -- each record set holds data retrieved most recently via a particular connection. The data is arranged in rows and columns to allow a row of field values to be associated with a particular shape instance.
• comments.xml -- text for comments on the drawing

The same markup and schema can be used to produce stencils and templates for use with Visio or other graphic applications that support VSDX. A Visio drawing file uses the extension .vsdx. Stencils, using the extension .vssx, are collections of shapes and icons considered useful for a particular type of diagram, such as flowcharts or room plans. Stencils can be loaded into the Visio interface for convenient use. Templates, using the extension .vstx, are drawings intended for repeated use as the basis for a common type of diagram, for example, a map of counties in a state or the bracket for a sports tournament. There is a market for third-party stencils and templates; some are sold and others are made freely available. These three diagram types are not permitted to incorporate macros. Three separate extensions are used for macro-enabled drawings: .vsdm, .vssm, and .vstm.


