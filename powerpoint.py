from functions import *
import pandas as pd
from pptx import Presentation


# print(f"Words: {author}\nTune: {tune}\nComposer: {composer}\n©: {copyr}\nCCLI: 522221")


df = pd.read_csv('https://docs.google.com/spreadsheets/d/e/2PACX-1vSIEtCzAWNJVZK7T1OEo1oGhnTib2bgFfdYRfFON1gpbG7LkHTFRJKSioV087Ys1oBZci80cRIRm0_u/pub?gid=0&single=true&output=csv')

cycle = df["Example Column"]


## Function to instantiate class and extract details
def song_details(url1,url2):
    scraper = HymnScraper(url1, url2)
    scraper.hymn_lyrics()
    scraper.tune_details()
    return scraper.get_lyrics()

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

print(song1)

# Create a new presentation
prs = Presentation()

## Test to see if this works