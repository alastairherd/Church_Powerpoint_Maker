#############################################################
## Creates the American to British ESV Dictionary
list1 = [
         "(\W)([Ss])avior(s)?(\W)",
         "(\W)neighbor(s|ing)?(\W)",
         "(\W)favor(able|ite|s|ed|itism)?(\W)",
         "(\W)Favor(\W)",
         "(\W)labor(ed|s)?(\W)",
         "(\W)(vap|vig)or(\W)",
         "(\W)clamor(\W)",
         "(\W)([Ss])plendor(\W)",
         "(\W)color(s|ed)?(\W)",
         "(\W)([Hh])onor(s|able|ing|ed)?(\W)",
         "(\W)dishonor(s|able|ing|ed)?(\W)",
         "(\W)travel(ed|er|ers|ing)(\W)",
         "(\W)marvel(ous|ously|ed|ing)(\W)",
         "(\W)([Cc])ounsel(or|ors|ed)(\W)",
         "(\W)plow(s|ed|ers|ing|man|men|share|shares)?(\W)",
         "(\W)judgment(s)?(\W)",
         "(\W)(recogn|[Rr]eal|[Oo]rgan|[Ss]ymbol|bapt|critic|apolog|sympath)iz(e|ed|es|ing)?(\W)",
         "(\W)(un)?authorized(\W)",
         "(\W)(centi)?meters(\W)",
         "(\W)liter(s)?(\W)",
         "(\W)scepter(s)?(\W)",
         "(\W)worship(ed|er|ers|ing)(\W)",
         "(\W)quarrel(ed|ing)(\W)",
         "(\W)benefited(\W)",
         "(\W)signaled(\W)",
         "(\W)paralyzed(\W)",
         "(\W)fulfill(s|ment)?(\W)",
         "(\W)skillful(ly)?(\W)",
         "(\W)jewelry(\W)",
         "(\W)(De|de|of)fense(s|less)?(\W)",
         "(\W)([Ss])ulfur?(\W)"
    ]

list2 = [
        r'\1\2aviour\3\4',
        r'\1neighbour\2\3',
        r'\1favour\2\3',
        r'\1Favour\2',
        r'\1labour\2\3',
        r'\1\2our\3',
        r'\1clamour\2',
        r'\1\2plendour\3',
        r'\1colour\2\3',
        r'\1\2onour\3\4',
        r'\1dishonour\2\3',
        r'\1travell\2\3',
        r'\1marvell\2\3',
        r'\1\2ounsell\3\4',
        r'\1plough\2\3',
        r'\1judgement\2\3',
        r'\1\2is\3\4',
        r'\1\2authorised\3',
        r'\1\2metres\3',
        r'\1litre\2\3',
        r'\1sceptre\2\3',
        r'\1worshipp\2\3',
        r'\1quarrell\2\3',
        r'\1benefitted\2',
        r'\1signalled\2',
        r'\1paralysed\2',
        r'\1fulfil\2\3',
        r'\1skilful\2\3',
        r'\1jewellery\2',
        r'\1\2fence\3\4',
        r'\1\2ulphur\3'
    ]
#############################################################

## Generates list to use in below function
anglo_list = list(zip(list1, list2))

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