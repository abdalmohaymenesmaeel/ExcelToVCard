import pandas as pd
import codecs

class Card:
    def __init__(self,firstname,lastname,program,place,phone):
        self.firstname=firstname
        self.lastname=lastname
        self.program=program
        self.place=place
        self.phone=phone

class ImportCardsFromExcel:
    def __init__(self,excelFile,sheetName):
        self.excelFile=excelFile
        self.sheetName=sheetName
        self.lstOfCards=[]
    def generatCards(self):
        df = pd.read_excel(open(self.excelFile, 'rb'), sheet_name=self.sheetName)  
        for i in df.index:
            self.lstOfCards.append(Card(df['firstname'][i],df['lastname'][i],df['programm'][i],df['place'][i],df['phone'][i]))
        return self.lstOfCards
 
       
class VCardConvertor:
    def __init__(self,fileCard):
        self.fileCard=fileCard
        self.card=""        
    def writeToFile(self):
        conn=codecs.open(self.fileCard,"w", "utf-8")
        conn.write(self.card)
        conn.close()        
    def buildCard(self,listOfCards):     
        for i in listOfCards:
              self.card+="BEGIN:VCARD\nVERSION:2.1\n"    
              self.card+="N:%s;%s;;;\n"%(i.lastname,i.firstname)
              self.card+="FN:\n"
              self.card+="NOTE:%s\n"%(i.program)
              self.card+="TEL;TYPE=CELL:%s\n"%(i.phone)
              self.card+="X-MS-IMADDRESS;CHARSET=utf-8:%s\n"%(i.place)
              self.card+="END:VCARD\n" 

excelData=ImportCardsFromExcel("testContacts.xlsx","Sheet1")
excelData.generatCards()
VC=VCardConvertor("55555.vcf")
VC.buildCard(excelData.lstOfCards)
VC.writeToFile()
