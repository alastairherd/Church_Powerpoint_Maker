from slide_making import *
from pptx import Presentation

template_file = 'template.pptx'

# Open the template file
# Create new class instance (prs)
prs = Presentation(template_file)


## Currently trying to sort out adding catechism pages
for i in range (1,23):
    try:
        print(i)
        slide_writer(i, prs)
    except:
        print(f"Error on slide {i}")



# Save the PowerPoint file
prs.save('example_from_template.pptx')