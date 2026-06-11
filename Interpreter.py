import os  #To check if directory or files exist
from Help import*
from VariableUtility import*
from Parser import *
from docx2python import docx2python
from docx2python.iterators import enum_at_depth
from ProcedureParser import *

class Interpreter():
    def __init__(self):
        self.Parser = None
        self.ProcedureParser = None
        self.prePattern = ""
        self.postPattern = ")\t"
        self.subPrePattern = "\t"
        self.subPostPattern = ")\t"

    def InterpretCommand(self):
        command = ""
        userexit = False
      

        vu=VariableUtility()
        
        
        

        while userexit != True:  
            correctDirectory = False
            correctFile = False
            correctName = False
           
            print("\nData Extractor Menu - Enter a command \noutname, outloc, input, parse, help, exit") 
            command= input(">>>")


            if command.lower() == "exit":#exit command
                userexit = True

            elif command.lower() == "outname":#outname command
                while correctName != True:    
                    outfilename = input("Enter a valid name for the output or enter q to go back:")
                    if(vu.setOutputFileName(outfilename)):
                       print("Filename Set to " + outfilename)
                       correctName=True
                    else:
                        print("Name Contains Illegal Characters")
                

                    
                        

                    if outfilename.lower() == "q":
                        continue
      
            elif command.lower() == "outloc":#outloc command
                
                while correctDirectory != True:  
                    outloc = input("Enter a directory location for the output file or enter q to go back:")
                    vu.setOutputFileLocation(outloc)
                    isdir= vu.OutputCheckValidity() 

                    if outloc.lower() == "q":
                        break

                    if isdir== True:
                        print("Output location successfully set")
                        correctDirectory = True

                    else:
                        print("Directory Does not exist")

            elif command.lower() == "input":#input command
                while correctFile != True:
                    inputfile = input("Enter filepath to the .docx input file or enter q to go back:")
                    vu.setInputFilePath(inputfile)
                    isFile = vu.InputCheckValidity()

                    if inputfile.lower() == "q":
                        break

                    if  isFile== True: 
                        print("File ready enter parse command to begin extraction")
                        vu.setInputFilePath(inputfile)
                        self.Parser = Parser(vu.getInputFilePath())
                        self.ProcedureParser = ProcedureParser()
                        correctFile = True
                        
                    elif isFile == False:
                        print("File does not exist")
                    
        

            elif command.lower() == "help": #help command 
                print("You are now in the help menu enter help again for more information...")
                hp= Help()
                hp.HelpDisplay()
                
                
            elif command.lower() == "print":
                print(vu.getInputFilePath())
                
    

            elif command.lower() == "parse":#parse command
                if vu.getInputFilePath() != "" and vu.getOutputFileLocation() != "" and vu.getOutputFileName != "":
                    self.ProcedureParser = ProcedureParser()
                    x = input("Any Procedures?")
                    if x == "y" or x == "Y" or x == "Yes" or x == "yes":
                        x = input("What type? 1 = Numeric, 2 = Lowercase, 3 = Uppercase, 4 = Lowercase Roman Numerals, 5 = Uppercase Roman Numerals")
                        if x.isdigit():
                            self.ProcedureParser.generateTokens0(self.prePattern, self.postPattern, int(x) - 1)
                        x = input("Any Subprocedures?")
                    if x == "y" or x == "Y" or x == "Yes" or x == "yes": 
                        x = input("What type? 1 = Numeric, 2 = Lowercase, 3 = Uppercase, 4 = Lowercase Roman Numerals, 5 = Uppercase Roman Numerals")
                        if x.isdigit():
                            self.ProcedureParser.generateTokens1(self.subPrePattern, self.subPostPattern, int(x) - 1)
                    
                    document = docx2python(vu.getInputFilePath())
                    docHeader = document.header
                    docFooter = document.footer
                    docTables = document.body
                    docTables = self.ProcedureParser.removeHeaderOrFooterParagraphs(docTables, docHeader)
                    docTables = self.ProcedureParser.removeHeaderOrFooterParagraphs(docTables, docFooter)
                    docTables = self.ProcedureParser.removeTabParagraphs(docTables)
                    docTables = self.ProcedureParser.removeEmptyParagraphs(docTables)    
                    structure = self.ProcedureParser.identifyDocumentStructure(docTables)
                    ProcedureStringList = self.ProcedureParser.identifyProcedureStrings(docTables)
                    tableList = self.Parser.tablesParse()
                    paragraphList = self.Parser.paragraphsParse()
                    graphicsList = self.Parser.graphicsParse(vu.getInputFilePath(), vu.getOutputFileLocation())
                    WordDocList = tableList + paragraphList + graphicsList
                    WordDocList = self.ProcedureParser.orderByDocumentStructure(paragraphList, graphicsList, tableList, structure)
                    self.ProcedureParser.identifyProcedure(WordDocList, ProcedureStringList[0], ProcedureStringList[1])

                    xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\n<WordDoc\nxmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\nxsi:noNamespaceSchemaLocation=\"DEO3.xsd\">\n"
                    for element in WordDocList:
                        xml += element.XMLReturn(1)
                        xml += "\n"
                    xml +="</WordDoc>"
                    #print(xml) #Optional
                    outputFile = open(vu.getOutputFileLocation() + "\\" + vu.getOutputFileName(), "wt")
                    outputFile.write(xml)
                    outputFile.close()
                    print("Parsing Completed")
                else:
                    print("Not all file information entered yet!")
                    
            elif command.lower() == "view":
                if vu.getInputFilePath() != "":
                    document = docx2python(vu.getInputFilePath())
                    docTables = document.body
                    print(docTables)
                else:
                    print("Input File Path not set yet!")

            elif command.lower() == "patterns":
                print("Note: Use the Tab key for any occurrences of \"\\t\" in the pattern.")
                x = input("What is the procedure pre-pattern?")
                self.prePattern = x
                x = input("What is the procedure post-pattern?")
                self.postPattern = x
                x = input("What is the sub-procedure pre-pattern?")
                self.subPrePattern = x
                x = input("What is the sub-procedure post-pattern?")
                self.subPostPattern = x
                
            else:   #invalid commands
                print("Invalid command\n")
            
    
