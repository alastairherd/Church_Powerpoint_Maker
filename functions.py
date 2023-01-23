from variables import *
import re
import re
import json
import pandas as pd
import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import os


## Set correct working directory
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
  
## Extracts the relevant details from the components JSON
def component_assigner(val, flag = None):
    if flag != 'communion':
        res = list(filter(lambda item: item["Component"] == val, components_json))
    else:
        initial = list(filter(lambda item: item["Component"] == "communion", components_json))
        res = list(filter(lambda item: item["Component"] == val, initial[0]["Content"]))
        
    speaker = res[0]["Content"]["speaker"]
    return_list = []
    for i in range(1,5):
        try:
            return_list.append(res[0]["Content"][str(i)])
        except:
            pass
    return speaker, return_list
'''
Returns the speaker, which most likely goes on every slide
And returns the text for each slide, in a list, with the content for each slide in each element
'''

## Function to extract the appropriate catechism question
def catechism_finder(val):
    real_val = int(val) - 1
    question = wsc_json['Data'][real_val]["Question"]
    answer = wsc_json['Data'][real_val]["Answer"]
    return (question,answer, val)


## Function to convert American language text to British
def replace_text(text, rep_text=anglo_list):
    for pattern, replacement in rep_text:
        pattern = re.compile(pattern)
        text = pattern.sub(replacement, text)
    return text

## Function to convert digits in verse to superscript
def super_verse(text):
    pattern = re.compile(r'\[([^\[\]]*)\]\s')
    matches = pattern.findall(text)
    for match in matches:
        superscript = ''.join(['⁰¹²³⁴⁵⁶⁷⁸⁹'[int(x)] if x.isdigit() else x for x in match])
        text = text.replace('[' + match + '] ', superscript)
    return text
    

## ESV API to return text from passage text
def get_esv_text(passage):
  params = {
    'q': passage,
    'include-headings': False,
    'include-footnotes': False,
    'include-verse-numbers': True,
    'include-short-copyright': False,
    'include-passage-references': False
  }
  
  headers = {
    'Authorization': 'Token %s' % API_KEY
  }
  
  response = requests.get(API_URL, params=params, headers=headers)
  
  passages = response.json()['passages']

  result = passages[0].strip()
  result = super_verse(result)
  result = replace_text(result)
  
  if passages:
    return (result,passage)
  else:
    return 'Error: Passage not found'

## Hymn Scraper Class
class HymnScraper:
    def __init__(self, url1 = None, url2 = None):
        self.url1 = url1
        self.url2 = url2
        self.title = None
        self.author = None
        self.copyr = None
        self.n_verses = None
        self.composer = None
        self.tune = None
        self.meter = None
        
    def hymn_lyrics(self):
        ## Word Extraction
        response = requests.get(self.url1)
        soup = BeautifulSoup(response.content, 'html.parser')
        try:
            self.title = soup.find("span", class_="hy_infoLabel", string="Title:").find_next().text
        except:
            None
        try:
            self.author = soup.find("span", class_="hy_infoLabel", string="Author:").find_next().text
        except:
            try:
                self.author = soup.find("span", class_="hy_infoLabel", string="Author (attributed to):").find_next().text
                self.author = self.author + " (atrb)"
            except:
                self.author = None
        try:
            self.copyr = soup.find("span", class_="hy_infoLabel", string="Copyright:").find_next().text
        except:
            self.copyr = None

        try:
            body_text = soup.select_one('div#at_fulltext.authority_section div div.authority_columns')
            paragraphs = body_text.find_all('p')
            verses = []
            for p in paragraphs:
                verses.append(p.text)

            self.n_verses = []
            for item in verses:
                if item[0].isdigit():
                    items = item.split(" ")
                    item = " ".join(items[1:])
                    if item.rstrip().endswith(" [Refrain]"):
                        item = item.split(" [Refrain]")[0]
                        self.n_verses.append(item)
                        self.n_verses.append(refrain)
                    else:
                        self.n_verses.append(item)
                if item.split(":")[0] == 'Refrain':
                    refrain = item.split(":")[1] + "\n"
                    self.n_verses.append(refrain)
        except:
            self.n_verses = None

    def tune_details(self):
        ## Tune Extraction
        response = requests.get(self.url2)
        soup = BeautifulSoup(response.content, 'html.parser')
        try:
            self.composer = soup.find("span", class_="hy_infoLabel", string="Composer:").find_next().text
        except:
            try:
                self.composer = soup.find("span", class_="hy_infoLabel", string="Composer (attributed to):").find_next().text
                self.composer = self.composer + " (atrb)"
            except:
                self.composer = "Unknown"
        try:
            self.tune = soup.find("span", class_="hy_infoLabel", string="Title:").find_next().text
        except:
            self.tune = None
        try:
            self.meter = soup.find("span", class_="hy_infoLabel", string="Meter:").find_next().text
        except:
            self.meter = None


    def get_lyrics(self):
        return (self.n_verses,self.title, self.author, self.composer, self.tune, self.copyr)
    
    def get_tune(self):
        return (self.composer, self.tune, self.meter)

## Hymn Scraper Functions
## Function to instantiate class and extract details
def song_details(url1,url2):
    scraper = HymnScraper(url1, url2)
    scraper.hymn_lyrics()
    scraper.tune_details()
    return scraper.get_lyrics()

def tune_details(url2):
    scraper = HymnScraper(None, url2)
    scraper.tune_details()
    return scraper.get_tune()

# function to convert to superscript
def get_super(x):
    normal = "0123456789"
    super_s = "⁰¹²³⁴⁵⁶⁷⁸⁹"
    res = x.maketrans(''.join(normal), ''.join(super_s))
    return x.translate(res)

def psalm_getter(psalm_name, psalms_json=psalms_json):
    
    match = re.search(r"^(Psalm )?(\d{1,3})(:(([1-9]\d{0,2})-([1-9]\d{0,2})))?(\s\(([a-zA-Z])\))?(\s\((\d{0,2})\))?$", psalm_name)

    if match:
        if match.group(2):
            psalm = match.group(2)
        if match.group(7):
            version = match.group(8)
        else:
            version = "a"
        if match.group(9):
            section = match.group(10)
        else:
            section = None
        if match.group(4):
            verses = match.group(4)
        elif section != None:
            verses = "1-300"
        else:
            verses = "1-30"

    if section == None:
        body = [d for d in psalms_json if d["Psalm"] == str(psalm) and d["Content"]["Version"] == version][0]['Content']['Body']
        meter = [d for d in psalms_json if d["Psalm"] == str(psalm) and d["Content"]["Version"] == version][0]['Content']['Meter']
    else:
        body = [d for d in psalms_json if d["Psalm"] == str(psalm) and d["Content"]["Section"] == str(section)][0]['Content']['Body']
        meter = [d for d in psalms_json if d["Psalm"] == str(psalm) and d["Content"]["Section"] == str(section)][0]['Content']['Meter']

    lst = []
    rng = [int(n) for n in verses.split("-")]
    for ints in range(rng[0],rng[1]):
        try:
            lst.append(body[str(ints)])
        except:
            pass

    lst = list(map(get_super,lst))

    return (psalm_name, meter,lst)

'''
# Pass in the string of the Psalm in the below format
##string = "Psalm 19:2-10 (a)"
# Swapping a either for section (number) or version (letter)
psalm, meter, lst = psalm_getter(string, psalms_json)
print(psalm,meter,lst)
'''
## Creates all the variables that we want to use in the template
## Using the dataframe we feed in the cycle number
df = pd.read_csv(online_csv)

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