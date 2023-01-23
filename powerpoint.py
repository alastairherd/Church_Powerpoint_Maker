from slide_making import *
import pandas as pd
from pptx import Presentation

# print(f"Words: {author}\nTune: {tune}\nComposer: {composer}\n©: {copyr}\nCCLI: 522221")


template_file = 'Presentation1.pptx'

# Open the template file
# Create new class instance (prs)
prs = Presentation(template_file)


## Currently trying to sort out adding catechism pages
for i in range (3,15):
    try:
        print(i)
        slide_writer(i, prs)
    except:
        print(f"Error on slide {i}")



# Save the PowerPoint file
prs.save('example_from_template.pptx')