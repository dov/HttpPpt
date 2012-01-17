# Description

A http server and COM client for remote controlling powerpoint through python.

This script runs an http server that speaks to a power point presentation,
allowing remote controlling the presentation by going backwards and
forwards in the presentation.

The server currently listens to port 8001 and returns a simple html
gui that allows remote controlling the presentation in a browser.

Alternatively an application may be written that use the following
commands directly:

* http://server:8001/nextpage
* http://server:8001/prevpage
* http://server:8001/gotopage=42

# Installation

This programs requires the following tools to be installed externally:

* python from http://python.org
* Win32 Extensions for Python from http://sourceforge.net/projects/pywin32/

Once these tools have been installed HttpPpt may be run by double clicking on HttpPpt.py .



# License

This script is released under the AGPL licence version 3.0

# Author

Dov Grobgeld <dov.grobgeld@gmail.com>
Tuesday 2012-01-17 00:05 
 
