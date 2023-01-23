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

'''
Sub html_fixer()
    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        Dim shp As Shape
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Text = Replace(shp.TextFrame.TextRange.Text, "_x000D_", "")
                End If
            End If
        Next shp
    Next sld
End Sub

'''