from functions import *
import pandas as pd
from pptx import Presentation
import json


# print(f"Words: {author}\nTune: {tune}\nComposer: {composer}\n©: {copyr}\nCCLI: 522221")


def create_json_var(name):
    with open(f"{name}.json", "r") as f:
        value = json.load(f)
        globals()[name + "_json"] = value

create_json_var("components")
create_json_var("psalms")
create_json_var("wsc")



df = pd.read_csv('https://docs.google.com/spreadsheets/d/e/2PACX-1vSIEtCzAWNJVZK7T1OEo1oGhnTib2bgFfdYRfFON1gpbG7LkHTFRJKSioV087Ys1oBZci80cRIRm0_u/pub?gid=0&single=true&output=csv')

cycle = df["Example Column"]


try:
    call_to_worship = get_esv_text(cycle[1])
    song1 = song_details(cycle[2], cycle[3])
    ## (psalm, meter,lst)
    psalm = psalm_getter(cycle[4])
    first_reading = cycle[6]
    catechism_reading = catechism_finder(cycle[7])
    song2 = song_details(cycle[8], cycle[9])
    second_reading = cycle[10]
    song3 = song_details(cycle[11], cycle[12])
    song4 = song_details(cycle[13], cycle[14])
except:
    pass


# If value is a list, then use second value as slide title
slide_dict = {
    2: "Notices",
    3: ["call_to_worship","Call to worship"],
    4: "song1",
    5: ["confession","Confession"],
    6: ["assurance", "Assurance of Pardon"],
    7: ["lords_prayer","The Lord's Prayer"],
    8: "psalm",
    9: ["first_reading","First Reading"],
    10: "catechism_reading",
    11: "song2",
    12: ["second_reading","Second Reading"],
    13: ["apostles_creed","Apostles' Creed (180 AD)"],
    14: ["Prayers","Prayers of Intercession"],
    15: "song3",
    16: "song4",
    17: ["the_grace","The Grace"],
    18: "Goodbye"
}

# Make a dictionary of the keys where slide_dict value starts with "song" or "psalm", some values are lists, so need to check for that
song_list = []
component_list = []
reading_list = []
catechism_list = []
for elem in slide_dict:
    if type(slide_dict[elem]) != list:
        if slide_dict[elem].startswith("song") or slide_dict[elem].startswith("psalm"):
            song_list.append(elem)
        elif slide_dict[elem].startswith("catechism"):
            catechism_list.append(elem)
    elif type(slide_dict[elem]) == list:
        if "reading" in slide_dict[elem][0]:
            reading_list.append(elem)
        elif slide_dict[elem][0] not in ['call_to_worship']:
            component_list.append(elem)



template_file = 'Presentation1.pptx'

# Open the template file
# Create new class instance (prs)
prs = Presentation(template_file)


## Currently trying to sort out adding catechism pages
for i in range (10,11):
    try:
        print(i)
        slide_writer(i)
    except:
        print(f"Error on slide {i}")



# Save the PowerPoint file
prs.save('example_from_template.pptx')