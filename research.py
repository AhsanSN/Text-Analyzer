import textstat
textstat.set_lang("en_US")

import xlsxwriter

txt = "my name is ahsan \n asd ahsan. hey!"

def unique(list1): 
    # intilize a null list 
    unique_list = [] 
      
    # traverse for all elements 
    for x in list1: 
        # check if exists in unique_list or not 
        if x not in unique_list: 
            unique_list.append(x) 
    # print list 
    return unique_list

def charNoSpaces(txt):
    n=0
    for i in txt:
        if(i!= ' '):
            n+=1
    return n

def wordsLongestSentence(txt):
    txt = txt.split(".")
    maax = 0
    for i in txt:
        if(len(i)>maax):
            maax=len(i)
    return(maax)

def wordsShortestSentence(txt):
    txt = txt.split(".")
    maax = 100000
    for i in txt:
        if(len(i)<maax):
            maax=len(i)
    return(maax)

def Average(lst): 
    return sum(lst) / len(lst)

def avgSentenceChars(txt):
    txt = txt.split(".")
    txt_n = []
    for i in (txt):
        txt_n.append(len(i))        
    avg = Average(txt_n)
    return(avg)

def avgSentenceWords(txt):
    txt = txt.split(".")
    txt_n = []
    for i in (txt):
        txt_n.append(len(i.split()))        
    avg = Average(txt_n)
    return(avg)

def avgWordLength(txt):
    txt = txt.split()
    txt_n = []
    for i in (txt):
        txt_n.append(len(i))        
    avg = Average(txt_n)
    return(avg)
    
def syllables(txt):
    n = 0
    n = n + txt.count("a")
    n = n + txt.count("e")
    n = n + txt.count("i")
    n = n + txt.count("o")
    n = n + txt.count("u")
    return n

def words_publisher(txt):
    txt = txt.replace('\n', '')
    return len(txt)/6

nChars = len(txt)
nWords = len(txt.split())
nUniqueWords = len(unique(txt.split()))
ncharNoSpaces = charNoSpaces(txt)
nSentences = (txt.count("."))
wordsLongestSentence = wordsLongestSentence(txt)
wordsShortestSentence = wordsShortestSentence(txt)
avgSentenceWords = avgSentenceWords(txt)
avgSentenceChars = avgSentenceChars(txt)
avgWordLength = avgWordLength(txt)
nParagraphs = txt.count("\n")
syllables = syllables(txt)
words_publisher = words_publisher(txt)
readingLevel = textstat.dale_chall_readability_score_v2(txt)
readingTime = textstat.reading_time(txt)

print("charLength", nChars)
print("nWords", nWords)
print("nUniqueWords", nUniqueWords)
print("ncharNoSpaces", ncharNoSpaces)
print("nSentences", nSentences)
print("wordsLongestSentence", wordsLongestSentence)
print("wordsShortestSentence", wordsShortestSentence)
print("avgSentenceWords", avgSentenceWords)
print("avgSentenceChars", avgSentenceChars)
print("avgWordLength", avgWordLength)
print("nParagraphs", nParagraphs)
print("syllables", syllables)
print("words_publisher", words_publisher)
print("readingLevel", readingLevel)
print("readingTime", readingTime)

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Stats.xlsx')
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.
expenses = (
    ['charLength', nChars],
    ['nWords',   nWords],
    ['nUniqueWords',  nUniqueWords],
    ['ncharNoSpaces',    ncharNoSpaces],
    ['nSentences',    nSentences],
    ['wordsLongestSentence',    wordsLongestSentence],
    ['wordsShortestSentence',    wordsShortestSentence],
    ['avgSentenceWords',    avgSentenceWords],
    ['avgSentenceChars',    avgSentenceChars],
    ['avgWordLength',    avgWordLength],
    ['nParagraphs',    nParagraphs],
    ['syllables',    syllables],
    ['words_publisher',    words_publisher],
    ['readingLevel',    readingLevel],
    ['readingTime',    readingTime]   
)

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for item, cost in (expenses):
    worksheet.write(row, col,     item)
    worksheet.write(row, col + 1, cost)
    row += 1

workbook.close()














