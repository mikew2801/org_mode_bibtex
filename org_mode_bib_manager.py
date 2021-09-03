"""
I use Bibtex, and I also use Emacs Org-Mode to write basically everything. I also like to use LaTeX, but many publishers in the humanities require submission of bibliographies in word format.

The following script takes as its input a bibtex file and a CSV file containing a series of styles for given bibliographic systems. It then outputs an ORG-Mode file with some simple formatting. I can then export this from org-mode to libre office, and save it as a word-file. At present I've only inputted the conventions for the MLA styleguide, but it can easily be expanded.
"""

from tabulate import tabulate
import pandas as pd
import bibtexparser
from bibtexparser.bwriter import BibTexWriter
import clipboard
import pyperclip
import os, fnmatch
import re

def sort_bibtex_db(bibtex_df):
    return bibtex_df.sort_values(["author", "year"], ascending = True)

def format_author_name(name, abbreviate = True):
    if len(name) < 1:
        return -1
    new_name = ""
    new_name += name.split()[-1]
    if len(name.split()) > 1:
        new_name += ", "
        if abbreviate == False:
            new_name += name.split()[0] + " "
            for item in name.split()[1:len(name.split())-1]:
                new_name += item[0] + ". "
        else:
            for item in name.split()[0:-1]:
                new_name += item[0] + ". "
    if new_name[-1] == " ":
        new_name = new_name[0:-1]
    return new_name[0:-1]

def update_author_name(bibtex_df):
    bibtex_df["author"].fillna("[NO AUTHOR]", inplace = True)
    bibtex_df["author"] = bibtex_df["author"].apply(format_author_name)
    return bibtex_df

def load_style(filename, style):
    """This loads the various formats for a particular style from the configuration file as a dictionary."""
    try:
        styles_df = pd.read_csv(filename)
    except FileNotFoundError:
        print("Error! File not found!")
        exit()

    style_row = styles_df[styles_df["Style_Name"] == style]

    if len(style_row) > 1:
        print("Error in file! Duplicate entries for style " + style + ". Please check formats.csv")
    else:
        style_dict = {}
        for column in style_row.columns:
            if not pd.isnull(style_row[column][0]):
                style_dict[column] = style_row[column][0].lower()
        return style_dict

def print_single_entry(df, style_dict):
    dict_keys = df.to_dict().keys()
    dict = df.to_dict()[list(dict_keys)[0]]
    
    if "ENTRYTYPE" in dict:
        if dict["ENTRYTYPE"] in style_dict:
            format_string = style_dict[dict["ENTRYTYPE"]]
        else:
            # In case there is no style for this entry type in the formats file
            # Perhaps return a dictionary entry
            return {"missing":dict["ENTRYTYPE"]}
    else:
        return None

    fields = re.findall('\<.*?\>', format_string)
    fields = [field.replace("<", "").replace(">", "") for field in fields]

    format_string = re.sub('\<.*?\>', "%s", format_string).replace("<", "").replace(">", "")

    values = []
    for field in fields:
        if not pd.isnull(dict[field]):
            values.append(dict[field].rstrip().replace("\n", ""))
        else:
            error_msg = "[%s missing]" % field
            values.append(error_msg.upper())
    
    return format_string % tuple(values)
    
def load_bib_file(filename):
    with open(filename) as file:    
        bibtex_db = bibtexparser.load(file)
    return pd.DataFrame(bibtex_db.entries)

def print_entire_db(filename, bib_df, style):
    style = load_style(filename, style)
    missing_styles = []
    for i in range(0, len(bib_df.index)):
        row = pd.DataFrame(bib_df.iloc[i, :])
        printed = print_single_entry(row, style)
        # Add a check here in case it is a missing style
        if printed is not None:
            if type(printed) is dict:
                missing_styles.append(printed["missing"])
            else:
                print(printed)     
                print("\n")                
    print("The following styles were present in the database but could not be found in the style document: ")
    [print(missing) for missing in set(missing_styles)]

to_open = "/home/michael/Dropbox/Python/1_Bib_Manager/Final/formats.csv"
mla = load_style(to_open, "MLA")
df = load_bib_file("main.bib")

with open("main.bib") as file:
    bibtex_db = bibtexparser.load(file)

authored = update_author_name(pd.DataFrame(bibtex_db.entries))
sorted_df = sort_bibtex_db(authored)
print_entire_db("formats.csv", sorted_df, "MLA")

sorted_df.drop(df.columns.difference(["author", "year", "title"]), 1, inplace=True)
