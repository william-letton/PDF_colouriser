#Package install required.
#pip install PyMuPDF
#pip install pywin32

##################################################

#Import packages
import fitz
import os
import random
import csv
import time
import win32com.client

##################################################

startTime=time.time()

##################################################

##Define accepted edge characters for words.
foreChar=[" ",
    "-",
    "(",
    ]

endChar=[" ",
    ".",
    ",",
    "-",
    ")",
    ":",
    ";"
    ]

##################################################

##Define functions for this program.

##Takes two tuples of length 4 as inputs, compares the co-ordinates, and outputs
##TRUE if the co-ordinates overap, FALSE otherwise.
def DetectOverlap(rect1,rect2):
    ##First see if the top left point of rect1 lies inside rect2.
    if rect1[0]>=rect2[0] and rect1[0]<=rect2[2] and rect1[1]>=rect2[1] and rect1[1]<=rect2[3]:
        return(True)
    ##Then see if the bottom right point of rect1 lies inside rect2
    if rect1[2]>=rect2[0] and rect1[2]<=rect2[2] and rect1[3]>=rect2[1] and rect1[3]<=rect2[3]:
        return(True)
    ##First see if the top left point of rect2 lies inside rect1.
    if rect2[0]>=rect1[0] and rect2[0]<=rect1[2] and rect2[1]>=rect1[1] and rect2[1]<=rect1[3]:
        return(True)
    ##Then see if the bottom right point of rect2 lies inside rect1
    if rect2[2]>=rect1[0] and rect2[2]<=rect1[2] and rect2[3]>=rect1[1] and rect2[3]<=rect1[3]:
        return(True)
    return(False)

##This function takes two tuples of length 4, compares the co-oridnates, and returns
##a single tuple of length 4 that represents the common area. If the rectangles do not
##share common area it will return the rectangle that links the closest corners!
def CommonArea(rect1,rect2):
    xcoords=list()
    ycoords=list()
    ##make the list of x co-ordinates.
    xcoords.append(rect1[0])
    xcoords.append(rect1[2])
    xcoords.append(rect2[0])
    xcoords.append(rect2[2])
    ##make the list of y co-ordinates.
    ycoords.append(rect1[1])
    ycoords.append(rect1[3])
    ycoords.append(rect2[1])
    ycoords.append(rect2[3])
    ##order both lists.
    xcoords.sort()
    ycoords.sort()
    ##extract the middle two values from each list.
    theOutput=(xcoords[1],ycoords[1],xcoords[2],ycoords[2])
    return(theOutput)

##This function takes a list of co-ordinate pairs, finds overlaps, and resolves the
##overlapping area.
def ResolveOverlapWithinList(rectList):
    startList=rectList
    ##Set overlapFound to true so that the resolver runs at least once.
    overlapFound=True
    ## run the resolver. any areas that do not overlap with any other are discarded.
    ## any that overlap are resolved into one common area.
    ##each of those overlapping areas is added to the list.
    newList=list()
    for rect1 in startList:
        for rect2 in startList:
            if rect1!=rect2:
                if DetectOverlap(rect1,rect2):
                    rect1_resolved=CommonArea(rect1,rect2)
                    newList.append(rect1_resolved)
    ##Remove any rectangles with 0 area
    newList2=list()
    for area in newList:
        if area[0]!=area[2] and area[1]!=area[3]:
            newList2.append(area)
    ##Take the set of newList
    startList=list(set(newList2))

    return(startList)

##################################################
##Search for and convert .docx files.
##detect if there are any word documents in the folder.
convertWordDocs="N"
current_directory= os.getcwd()
arr = os.listdir()
print("Searching for .docx files in the folder...")
wordDocList=list()
for fle in arr:
    if fle[len(fle)-4:]=="docx":
        wordDocList.append(fle)
if len(wordDocList)>0:
    print("MSword documents detected in folder:")
    print(wordDocList)
    convertWordDocs=input("Would you like to create .pdf versions of these documents and colourise them? (y/n)")
else:
    print("No MSword documents detected in folder.")
if convertWordDocs=="y" or convertWordDocs=="Y":
    print("File conversion requires closing all active MSword documents.")
    input("Please close all open MSword documents before continuing and then hit ENTER.")
    print("Creating .pdf versions of .docx documents...")
    wdFormatPDF = 17
    word = win32com.client.Dispatch('Word.Application')
    for i in range(0,len(wordDocList)):
        out_file=wordDocList[i][:len(wordDocList[i])-5]+"_converted.pdf"
        #print("out_file: ",out_file)
        doc=word.Documents.Open(os.path.abspath(wordDocList[i]))
        #input("press enter to continue")
        doc.SaveAs(os.path.join(current_directory,out_file),FileFormat=wdFormatPDF)
        doc.Close()
    word.Quit()
    print(".pdf documents created.")

##################################################
#Define desired words list
print("Reading searchWords file...")
with open("searchWords.csv") as csvfile:
    readCSV=csv.reader(csvfile,delimiter=',')
    RGBcodes=list()
    stringList=list()
    wordList=list()
    stringAndWordList=list()
    for row in readCSV:
        try:
            RGBcodes.append((float(row[1]),float(row[2]),float(row[3])))
        except:
            RGBcodes.append((row[1],row[2],row[3]))
        stringList.append(row[4])
        wordList.append(row[5])
        ##The point of string and wordList is to keep track of which colours
        ## are used for either strings or whole words.
        stringAndWordList.append(row[4]+row[5])
    print("strangAndWordList: ",stringAndWordList)
    ##Remove first row data.
RGBcodes=RGBcodes[1:]
stringList=stringList[1:]
wordList=wordList[1:]
stringAndWordList=stringAndWordList[1:]
#print("RGBcodes: ",RGBcodes)
#print("stringList: ",stringList)


##Remove blank searchTerms and associated colours.
while len(stringAndWordList[-1])==0:
    del stringAndWordList[-1]
    del stringList[-1]
    del wordList[-1]
    del RGBcodes[-1]
#print("RGBcodes: ",RGBcodes)
#print("stringList: ",stringList)


##split searchstrings and searchWords by commmas
stringList_split=list()
for group in stringList:
    stringList_split.append(group.split(","))
wordList_split=list()
for group in wordList:
    wordList_split.append(group.split(","))

print("stringList_split: ",stringList_split)
print("WordList_split: ",wordList_split)
print("RGBcodes: ",RGBcodes)

##################################################

#create list of pdf filenames.
arr = os.listdir()
#print(arr)
print("Searching for PDF files in the folder...")
pdfList=list()
for fle in arr:
    if fle[len(fle)-4:]==".pdf":
        pdfList.append(fle)
print(pdfList)

##Create new directory for results.
print("Creating folder for results...")
current_directory= os.getcwd()
#print("current_directory: ",current_directory)
final_directory = os.path.join(current_directory,r'Colourised_PDF_files')
#print("final_directory: ",final_directory)
if not os.path.exists(final_directory):
    os.makedirs(final_directory)


##run highlighting on each pdf
print("Highlighting words...")
for pdf in pdfList:
    print("     ",pdf)
    doc = fitz.open(pdf)
    #print("Document length: ",len(doc))
    #################
    #Find and highlight desired words on all pages of document.
    for i in range(0,len(doc)):
        ##load each page in turn to be searched.
        print("          Page ",(i+1)," of ",len(doc))
        page=doc.loadPage(i)
        ##For each colour-grouped list of words.
        for pos in range(0,len(stringList_split)):
            ##For each word within the colour-grouped list.
            for word in wordList_split[pos]:
                ##only proceed if the word is not an empty string.
                if word != '':
                    #print("word: ",word)
                    wordAreas=list()
                    ##create an 'areas' list containing the co-ordinates for found words.
                    for char1 in foreChar:
                        areas1=page.searchFor(char1+word,hit_max=1000)
                        #print("length of areas1: ",len(areas1))
                        #print("areas: ",areas)
                        for area in areas1:
                            wordAreas.append(area)
                    for char2 in endChar:
                        areas2=page.searchFor(word+char2,hit_max=1000)
                        #print("length of areas2: ",len(areas2))
                        #print("areas: ",areas)
                        for area in areas2:
                            wordAreas.append(area)
                            #print("masterAreas for ",word,": ",masterAreas)
                    wordAreas2=ResolveOverlapWithinList(wordAreas)
                    #print("length of wordAreas2: ",len(wordAreas2))
                    annot=page.addHighlightAnnot(wordAreas2)
                    try:
                        annot.setColors(stroke=RGBcodes[pos])
                        #print("colours set")
                        annot.update()
                    except:
                        continue
                    ##and now for each string within the colour-grouped list.
            for strng in stringList_split[pos]:
                ##only proceed if the word is not an empty string.
                if strng != '':
                    #print("string: ",strng)
                    stringAreas=list()
                    areas3=page.searchFor(strng,hit_max=1000)
                    #print("length of areas3: ",len(areas3))
                    for area in areas3:
                        stringAreas.append(area)
                    annot=page.addHighlightAnnot(stringAreas)
                    try:
                        annot.setColors(stroke=RGBcodes[pos])
                        #print("colours set")
                        annot.update()
                    except:
                        continue
    ################
    #Save result
    docFile=os.path.join(final_directory,(str(pdf)))
    doc.save(docFile)
    #doc.save("HL_"+pdf)




##Prompt to exit program.
print("Colouring complete.")

##################################################

##Report on program runtime.
print("Program took",int(time.time()-startTime),"seconds to run.")
input('Press ENTER to exit')
