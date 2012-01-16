"""
This script runs an http server that speaks to a power point presentation,
allowing remote controlling the presentation by going backwards and
forwards in the presentation.

The server currently listens to port 8001 and returns a simple html
gui that allows remote controlling the presentation in a browser.

Alternatively an application may be written that use the following
commands directly:

  http://server:8001/nextpage
  http://server:8001/prevpage
  http://server:8001/gotopage=42

Dov Grobgeld
This script is released under the AGPL license.
"""

import SimpleHTTPServer
import SocketServer
import win32com.client
import PPTSlideshow
import datetime
import socket

PORT = 8001

# My n900 generates two requests at a time, this is solved
# by filtering requests by at least 1s. This is reasonable
# for a presentation.
last_slidetime = datetime.datetime.now()
slide_command_inactive = 1.0

class TestHandler(SimpleHTTPServer.SimpleHTTPRequestHandler):
  """The test example handler."""
  # Recognize the requests and serve different commands depending
  # on the page requested.
  def do_GET(self):
    global last_slidetime, slide_command_inactive
    CurrentSlide = "?"
    time_since_last_slide_change = (datetime.datetime.now() - last_slidetime).total_seconds()

    if time_since_last_slide_change > slide_command_inactive:
      if self.path=="/nextpage":
        CurrentSlide = PPTSlideshow.GotoSlide(Relative=1)
        last_slidetime = datetime.datetime.now()
      elif self.path=="/prevpage":
        CurrentSlide = PPTSlideshow.GotoSlide(Relative=-1)
        last_slidetime = datetime.datetime.now()
      elif "/gotopage=" in self.path:
        new_page = int(self.path.split("=")[1])
        CurrentSlide = PPTSlideshow.GotoSlide(Absolute=new_page)
        last_slidetime = datetime.datetime.now()
    
    # Always serve the interactive page
    message = """
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Buttons</title>
<SCRIPT LANGUAGE="JavaScript">
function testResults (form) {{
    var TestVar = form.inputbox.value;
    window.location = "/gotopage="+TestVar
}}
</SCRIPT>
</head>
<style href>a {{text-decoration: none}} </style>
<body>
<h1 align="center">Remote control</h1>
<center>
<a href="/prevpage">
<button
  style="width:40%;height:150px;background-color:#FFe0e0">
  <b>&lt;&lt;&lt;Previous</b></button></a>
<a href="/nextpage">
<button
  style="width:40%;height:150px;background-color:#e0FFe0">
  <b>Next &gt;&gt;&gt;</b></button></a>
<FORM NAME="myform" ACTION="" METHOD="GET">
<table>
<td>Page:<td><input type="text" name="inputbox" value="{0}">
<td><INPUT
       style="width:100px;height:50px;background-color:#e0e0FF"
       TYPE="button"
       NAME="button"
       Value="Goto"
       onClick="testResults(this.form)"
       > 
</table>
</form>
</center>
</body>
</html>
""".format(CurrentSlide)
    self.send_response(200)
    self.send_header("Content-type", "text/html")
    self.send_header("Content-Length", len(message))
    self.end_headers()

    self.wfile.write(message)


Handler = SimpleHTTPServer.SimpleHTTPRequestHandler

httpd = SocketServer.TCPServer(("", PORT), TestHandler)

print "Connect to http://%s:%d"%(socket.gethostbyname(socket.gethostname()),PORT)
httpd.serve_forever()
