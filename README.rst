visiopy
=======

Handling Microsoft Visio 2013 vsdx files from Python

**WARNING, DO NOT USE. Very much a work in progress**


Here's how I would LIKE to be able to generate diagrams

.. code:: python

    >>> import visiopy

    >>> my_diagram = visiopy('MyFirstVisio.vsdx', author='Mathijs Mortimer', use_metric=True)
    >>> page_rel_id = diag.add_page('MyFirstPage')

    >>> rect1 = my_diagram.add_rect(page_rel_id, pin_x=2.0, pin_y=5.0, width=2.0, height=2.0)
    >>> rect2 = my_diagram.add_rect(page_rel_id, pin_x=6.0, pin_y=5.0, width=2.0, height=2.0)
    >>> connect1 = my_diagram.add_connect(page_rel_id, rect1, rect2)

    >>> my_diagram.to_file('MyFirstVisio.vsdx')


A nice short description of the format found here:
http://www.digitalpreservation.gov/formats/fdd/fdd000021.shtml
