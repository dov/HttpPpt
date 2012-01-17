"""Control a powerpoint presentation"""

import win32com.client
import time

def GotoSlide(Relative=None, Absolute=None):
  # Goto a slide in the current presentation. Does nothing
  # if no slideshow is active.
  app = win32com.client.Dispatch("PowerPoint.Application")
  try:
    # Goto next slide in the current view
    if not Absolute is None:
      slideIndex = Absolute
    else:
      slideIndex = app.SlideShowWindows(1).View.CurrentShowPosition + Relative
    app.SlideShowWindows(1).View.GotoSlide(slideIndex)
    return slideIndex
  except:
    pass
  return "?"
  
def Test():
  SlidesCount=app.ActivePresentation.Slides.Count
  
  for i in range(20):
    try:
      # Goto next slide in the current view
      slideIndex = app.ActiveWindow.View.Slide.SlideIndex 
      app.ActiveWindow.View.GotoSlide(slideIndex+1)
    except:
      try:
        # Goto next slide in the slideshow
        slideIndex = app.SlideShowWindows(1).View.CurrentShowPosition 
        app.SlideShowWindows(1).View.GotoSlide(slideIndex+1)
      except:
        pass
    
    time.sleep(2)


