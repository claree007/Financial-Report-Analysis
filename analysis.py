import csv
import pandas as pd
import re
import xlsxwriter
import html
import urllib.request
from nltk.tokenize import word_tokenize, sent_tokenize


def pos_neg_dict():
    master_file = open("LoughranMcDonald_MasterDictionary_2016.csv")
    master_csv = csv.reader(master_file, delimiter=',')
    # skip first row i.e. header
    next(master_csv)
    for row in master_csv:
        # row 0 is name, 7 is if negative, 8 is if positive
        # if negative word
        if not int(row[7])==0:
            neg_list.append(row[0].lower())
        # if positive word
        elif not int(row[8])==0:
            pos_list.append(row[0].lower())
    master_file.close()

def polar_score(positive_score, negative_score):
    score = (positive_score - negative_score)/((positive_score + negative_score) + 0.000001)
    return score

def avg_sent_len(total_words, total_sentences):
    return total_words/total_sentences

def complex_word(word):
    complex_letters = 0
    if not re.match(r'es|ed$',word):
        for letter in word:
            if letter in ['a','e','i','o','u']:
                complex_letters+= 1
    return complex_letters

def complex_word_percent(total_complex, total_words):
    return total_complex/total_words 

def fog_ind(average_sentence_length, percentage_of_complex_words):
    index = 0.4 * (average_sentence_length + percentage_of_complex_words)
    return index

def pos_word_prop(positive_score, total_words):
    return positive_score/total_words

def neg_word_prop(negative_score, total_words):
    return negative_score/total_words

def uncertain_word_prop(uncertain_score, total_words):
    return uncertain_score/total_words

def constrain_word_prop(constrain_score, total_words):
    return constrain_score/total_words


pos_list = []
neg_list = []
stop_words = []
uncertain_list = []
constrain_list = []
c = html.unescape('&#149;')
puntuations = [',','.','!','\'','"',':',';','-','(',')','%',c]
excel_row = 1
fixed_link = "https://www.sec.gov/Archives/"

variables = [
    "CIK",
    "CONAME",
    "FYRMO",
    "FDATE",
    "FORM",
    "SECFNAME",
    "mda_positive_score",
    "mda_negative_score",
    "mda_polarity_score",
    "mda_average_sentence_length",
    "mda_percentage_of_complex_words",
    "mda_fog_index",
    "mda_complex_word_count",
    "mda_word_count",
    "mda_uncertainty_score",
    "mda_constraining_score",
    "mda_positive_word_proportion",
    "mda_negative_word_proportion",
    "mda_uncertainty_word_proportion",
    "mda_constraining_word_proportion",
    "qqdmr_positive_score",
    "qqdmr_negative_score",
    "qqdmr_polarity_score",
    "qqdmr_average_sentence_length",
    "qqdmr_percentage_of_complex_words",
    "qqdmr_fog_index",
    "qqdmr_complex_word_count",
    "qqdmr_word_count",
    "qqdmr_uncertainty_score",
    "qqdmr_constraining_score",
    "qqdmr_positive_word_proportion",
    "qqdmr_negative_word_proportion",
    "qqdmr_uncertainty_word_proportion",
    "qqdmr_constraining_word_proportion",
    "rf_positive_score",
    "rf_negative_score",
    "rf_polarity_score",
    "rf_average_sentence_length",
    "rf_percentage_of_complex_words",
    "rf_fog_index",
    "rf_complex_word_count",
    "rf_word_count",
    "rf_uncertainty_score",
    "rf_constraining_score",
    "rf_positive_word_proportion",
    "rf_negative_word_proportion",
    "rf_uncertainty_word_proportion",
    "rf_constraining_word_proportion",
    "constraining_words_whole_report"
    ]

# creating positive and negative list from master dict
pos_neg_dict()

# populating stop words list
with open("StopWords_Generic.txt") as stop_file:
    word = stop_file.readline()
    stop_words.append(word.lower())

# populating uncertainity list
with open("uncertainty_dictionary.xlsx",'rb') as uncertain_file:
    un_df = pd.read_excel(uncertain_file)
    for word in un_df["Word"]:
        uncertain_list.append(word.lower())

# populating constraining list
with open("constraining_dictionary.xlsx",'rb') as constrain_file:
    co_df = pd.read_excel(constrain_file)
    for word in co_df["Word"]:
        constrain_list.append(word.lower())

# opening input file
links_file = open("cik_list.xlsx",'rb')
link_df = pd.read_excel(links_file, sheetname="cik_list_ajay")

# opening output file
output_file = xlsxwriter.Workbook('output.xlsx')
worksheet = output_file.add_worksheet()

# writing header to output file
bold = output_file.add_format({'bold': True})
for x in range(len(variables)):
    worksheet.write(0, x, variables[x], bold)

for i in range(len(link_df)):
    # reading variables from excel
    cik,coname,fyrmo,fdate,form,secfname = [str(v) for v in link_df.iloc[i][:]]
    secfname = "".join([fixed_link, secfname])
    # getting document from the link
    req = urllib.request.urlopen(secfname)
    rr = req.read().decode("utf-8")
    # cleaning
    rr = html.unescape(rr)
    rr = re.sub(r'<.+?>', '', rr)
    f = re.sub(r'\s*\n+\s*', '\n', rr)
    # searching the report for the terms
    constraining_words_whole_report = 0
    tokenized_report = word_tokenize(rr)
    for word in tokenized_report:
        if word in constrain_list:
            constraining_words_whole_report+= 1

    selection = ['', '', '']

    # Management's Discussion and Analysis
    pattern1 = re.compile(r'\n+(?P<pat>I[Tt][Ee][Mm])\s+[0-9IVX]\.?\s+((?i:Management.s\n?\s*Discussion\n?\s*and\n?\s*Analysis)(\n|.)*?)\n((Risk\n?\s*Factors)|(?P=pat))')
    pattern2 = re.compile(r'\n(?P<pat>I[Tt][Ee][Mm])\s+[0-9IVX]\.?\n((?i:Management.s\n?\s*Discussion\n?\s*and\n?\s*Analysis)(\n|.)*?)\n(Risk\n?\s*Factors)|(?P=pat)')
    search1 = re.findall(pattern1, rr)
    if search1:
        if len(search1[0][1])<150:
            if len(search1)>1 and len(search1[1][1])>150:
                selection[0] = search1[1][1]
        else:
            selection[0] = search1[0][1]
    else:
        search1 = re.findall(pattern1, f)
    if search1:
        if len(search1[0][1])<150:
            if len(search1)>1 and len(search1[1][1])>150:
                selection[0] = search1[1][1]
        else:
            selection[0] = search1[0][1]
    else:
        search1 = re.findall(pattern2, f)
    if search1:
        if len(search1[0][1])<150:
            if len(search1)>1 and len(search1[1][1])>150:
                selection[0] = search1[1][1]
        else:
            selection[0] = search1[0][1]

    # Quantitative and Qualitative Disclosures about Market Risk
    pattern1 = re.compile(r'\n+(?P<pat>I[Tt][Ee][Mm])\s+[0-9IVX]\.?\s+((?i:Quantitative\n?\s*and\n?\s*Qualitative\n?\s*Disclosures?\n?\s*about\n?\s*Market\n?\s*Risk)(\n|.)*?)\n(?P=pat)')
    pattern2 = re.compile(r'\n(?P<pat>I[Tt][Ee][Mm])\s+[0-9IVXA-Z]+\.?\n((?i:Quantitative\n?\s*and\n?\s*Qualitative\n?\s*Disclosures?\n?\s*about\n?\s*Market\n?\s*Risk)(\n|.)*?)\n(?P=pat)')
    search1 = re.findall(pattern1, rr)
    if search1:
        if len(search1[0][1])<150:
            if len(search1)>1 and len(search1[1][1])>150:
                selection[1] = search1[1][1]
        else:
            selection[1] = search1[0][1]
    else:
        search1 = re.findall(pattern1, f)
    if search1:
        if len(search1[0][1])<150:
            if len(search1)>1 and len(search1[1][1])>150:
                selection[1] = search1[1][1]
        else:
            selection[1] = search1[0][1]
    else:
        search1 = re.findall(pattern2, f)
    if search1:
        if len(search1[0][1])<150:
            if len(search1)>1 and len(search1[1][1])>150:
                selection[1] = search1[1][1]
        else:
            selection[1] = search1[0][1]

    # Risk Factors
    pattern1 = re.compile(r'\n(Risk\n?\s*Factors(\n|.)*?)\nI[Tt][Ee][Mm]', re.I)
    pattern2 = re.compile(r'\n(?P<pat>I[Tt][Ee][Mm])\s+[0-9IVXA-Z]+\.?\s+(Risk\n?\s*Factors(\n|.)*?)\n(?P=pat)', re.I)
    search1 = re.findall(pattern2, f)
    if search1:
        if len(search1[0][1])<150:
            if len(search1)>1 and len(search1[1][1])>150:
                selection[2] = search1[1][1]
        else:
            selection[2] = search1[0][1]
    else:
        search1 = re.findall(pattern1, rr)
    if search1:
        if len(search1[0][0])<150:
            if len(search1)>1 and len(search1[1][0])>150:
                selection[2] = search1[1][0]
        else:
            selection[2] = search1[0][0]
    else:
        search1 = re.findall(pattern1, f)
    if search1:
        if len(search1[0][0])<150:
            if len(search1)>1 and len(search1[1][0])>150:
                selection[2] = search1[1][0]
        else:
            selection[2] = search1[0][0]

    # procesing
    output = [cik, coname, fyrmo, fdate[:10], form, secfname]
    for term in range(3):
        if selection[term] == '':
            for j in range(14):
                output.append('')
            continue
        words = word_tokenize(selection[term])
        sent_pattern = re.compile(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s')
        sentences = sent_pattern.split(selection[term])
        sentence_count = len(sentences)
        word_count = 0        
        complex_word_count = 0
        positive_score = 0
        negative_score = 0
        uncertainity_score = 0
        constraining_score = 0
        for word in words:
            if words in stop_words or word in puntuations:
                continue
            word_count+= 1
            word = word.lower()
            # positive
            if word in pos_list:
                positive_score+= 1
            # negative
            elif word in neg_list:
                negative_score-= 1
            # uncertain
            if word in uncertain_list:
                uncertainity_score+= 1
            # constraining
            if word in constrain_list:
                constraining_score+= 1
            # complex
            if complex_word(word) > 2:
                complex_word_count+= 1
        # calculations
        negative_score*= -1
        polarity_score = polar_score(positive_score, negative_score)
        average_sentence_length = avg_sent_len(word_count, sentence_count)
        percentage_of_complex_words = complex_word_percent(complex_word_count, word_count)
        fog_index = fog_ind(average_sentence_length, percentage_of_complex_words)
        positive_word_proportion = pos_word_prop(positive_score, word_count)
        negative_word_proportion = neg_word_prop(negative_score, word_count)
        uncertainity_word_proportion = uncertain_word_prop(uncertainity_score, word_count)
        constraining_word_proportion = constrain_word_prop(constraining_score, word_count)

        # populating list with output for writing to file
        output.append(positive_score)
        output.append(negative_score)
        output.append(polarity_score)
        output.append(average_sentence_length)
        output.append(percentage_of_complex_words)
        output.append(fog_index)
        output.append(complex_word_count)
        output.append(word_count)
        output.append(uncertainity_score)
        output.append(constraining_score)
        output.append(positive_word_proportion)
        output.append(negative_word_proportion)
        output.append(uncertainity_word_proportion)
        output.append(constraining_word_proportion)

    output.append(constraining_words_whole_report)

    # writing output to excel file
    for x in range(len(output)):
        worksheet.write(excel_row, x, output[x])
    excel_row+= 1

output_file.close()
links_file.close()