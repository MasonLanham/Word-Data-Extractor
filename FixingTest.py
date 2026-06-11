from VariableUtility import *
from ProcedureParser import *
from Parser import *
from docx2python import docx2python

vu = VariableUtility()
vu.setInputFilePath("C:\\Users\\Marti\\Desktop\\WDS\\Code\\4-28\\Tests\\UserGuideTest.docx")
vu.setOutputFileLocation("C:\\Users\\Marti\\Desktop\\WDS\\Code\\4-28\\Output")
vu.setOutputFileName("Output.xml")
ProcedureParser = ProcedureParser()
Parser = Parser(vu.getInputFilePath())
prePattern = ""
postPattern = ")\t"
subPrePattern = "\t"
subPostPattern = ")\t"
ProcedureParser.generateTokens1(subPrePattern, subPostPattern, 0)
x = input("Any Procedures?")
if x == "y" or x == "Y" or x == "Yes" or x == "yes":
    x = input("What type? 1 = Numeric, 2 = Lowercase, 3 = Uppercase, 4 = Lowercase Roman Numerals, 5 = Uppercase Roman Numerals")
    if x.isdigit():
        ProcedureParser.generateTokens0(prePattern, postPattern, int(x) - 1)
x = input("Any Subprocedures?")
if x == "y" or x == "Y" or x == "Yes" or x == "yes":
    x = input("What type? 1 = Numeric, 2 = Lowercase, 3 = Uppercase, 4 = Lowercase Roman Numerals, 5 = Uppercase Roman Numerals")
    if x.isdigit():
        ProcedureParser.generateTokens1(subPrePattern, subPostPattern, int(x) - 1)
if x.isdigit():
    ProcedureParser.generateTokens1(subPrePattern, subPostPattern, int(x) - 1)
document = docx2python(vu.getInputFilePath())
docHeader = document.header
docFooter = document.footer
docTables = document.body
docTables = ProcedureParser.removeHeaderOrFooterParagraphs(docTables, docHeader)
docTables = ProcedureParser.removeHeaderOrFooterParagraphs(docTables, docFooter)
docTables = ProcedureParser.removeTabParagraphs(docTables)
docTables = ProcedureParser.removeEmptyParagraphs(docTables)    
structure = ProcedureParser.identifyDocumentStructure(docTables)
ProcedureStringList = ProcedureParser.identifyProcedureStrings(docTables)
tableList = Parser.tablesParse()
paragraphList = Parser.paragraphsParse()
graphicsList = Parser.graphicsParse(vu.getInputFilePath(), vu.getOutputFileLocation())
WordDocList = tableList + paragraphList + graphicsList
WordDocList = ProcedureParser.orderByDocumentStructure(paragraphList, graphicsList, tableList, structure)
print(ProcedureParser.Resets0)
print(ProcedureParser.Resets1)
print(ProcedureStringList)
print(docTables)
ProcedureParser.identifyProcedure(WordDocList, ProcedureStringList[0], ProcedureStringList[1])
xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\n<WordDoc\nxmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\nxsi:noNamespaceSchemaLocation=\"DEO3.xsd\">\n"
for element in WordDocList:
    xml += element.XMLReturn(1)
    xml += "\n"
xml +="</WordDoc>"
print(xml) #Optional
outputFile = open(vu.getOutputFileLocation() + "\\" + vu.getOutputFileName(), "wt")
outputFile.write(xml)
outputFile.close()
print("Parsing Completed")
