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
import webbrowser
from string import Template

PORT = 8001

# My n900 generates two requests at a time, this is solved
# by filtering requests by at least 1s. This is reasonable
# for a presentation.
last_slidetime = datetime.datetime.now()
slide_command_inactive = 1.0

class HttpPptHandler(SimpleHTTPServer.SimpleHTTPRequestHandler):
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
      elif "gotopage=" in self.path:
        new_page = int(self.path.split("gotopage=")[1])
        CurrentSlide = PPTSlideshow.GotoSlide(Absolute=new_page)
        last_slidetime = datetime.datetime.now()
    
    # Always serve the interactive page
    message = (Template(open("index-template.html").read())
               .substitute(current_page=CurrentSlide))
                       
    self.send_response(200)
    self.send_header("Content-type", "text/html")
    self.send_header("Content-Length", len(message))
    self.end_headers()

    self.wfile.write(message)


Handler = SimpleHTTPServer.SimpleHTTPRequestHandler

httpd = SocketServer.TCPServer(("", PORT), HttpPptHandler)
host = socket.gethostbyname(socket.gethostname())


fh=open("usage.html","w")
fh.write(Template(open("usage-template.html").read())
         .substitute(host=host,
                     port=PORT))
fh.close()
webbrowser.open("usage.html")
                            
print "Connect to http://%s:%d"%(socket.gethostbyname(socket.gethostname()),PORT)
httpd.serve_forever()
