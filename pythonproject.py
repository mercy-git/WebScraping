# Level_3

# Implemented using:
    # Excel
    # Files
    # Database
    # Natural Language Tool Kit  
    # Regular Expression
    # Exceptions 
    # Class & methods
    # Dictionary

import requests
from bs4 import BeautifulSoup
import io
import xlsxwriter
import sqlite3
from datetime import datetime
import re
import nltk
from nltk.corpus import stopwords,wordnet
import sys
import traceback

class Website:
    "Web search engine"
    # retrive all the stop words
    stopWords = set(stopwords.words('english'))
    stopWords = stopWords.union({'&','-'})
    
    def __init__(self,wname):
        self.websiteURL = wname
        self.pageText = ''
        self.soupText = ''
        self.inputWord = ''
        self.wordsDict = {}
           
    def writeSoupToFile(self):
        # Returns requests.models.Response Object
        page = requests.get(self.websiteURL.format(4980))
        # Raise an HTTPError if the HTTP Request returned an unsuccessful status code
        page.raise_for_status()
        # Returns BeautifulSoup Object
        soup = BeautifulSoup(page.text, 'html.parser')
        
        # Remove the java script coding: for element.Tag object in element.ResultSet Object
        for script in soup(["script","style","title"]):
            script.extract()
        # Get only the text from the HTML Content: Returns a string
        soupText = soup.get_text()
        
        # break into lines and remove leading and trailing space on each
        lines = (line.strip() for line in soupText.splitlines())
        # break multi-headlines into line each
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        # drop blank lines
        self.soupFilteredText = '\n'.join(chunk for chunk in chunks if chunk)
        
        # write the contents of the prettified soup, the complete html source code to text file
        with io.open("PrettySoup.txt", "w", encoding = "utf-8") as file:
            file.write(soup.prettify())
            
        # write the contents of the extracted HTML text to text file 
        with io.open("SoupText.txt", "w", encoding = "utf-8") as file:
            file.write(self.soupFilteredText)
        print("\nThe complete HTML code and the extracted HTML Text written successfully to text files PrettySoup.txt and SoupText.txt respectively")
        
    def setWordsDict(self):
        # Make the words list - connvert the soup text to lower case and split by space
        wordsList = self.soupFilteredText.lower().split()
        # Sort the words list
        wordsList.sort()
            
        for word in wordsList:
            # remove if any of the special characters .,:;?! is present at the end ; re.sub(pattern,repl,string,count=0,flags=0)
            word = re.sub("[.,:;?!'\")]$", '', word)
            # remove if any special characters '( present at the begining of the word
            word = re.sub("^['(]",'',word)
            
            # to write new words
            if word not in self.wordsDict:
                # write new words if it is not a stop word and if it is a valid engligh word
                if word not in self.stopWords and wordnet.synsets(word):
                    self.wordsDict[word] = 1
            # to update the count of the existing words
            else:
                self.wordsDict[word] += 1
        
    def printInpWordCount(self):
            self.inputWord = input("\nEnter the word to search: ")
            # re.findall(pattern,string,flags)
            findList = re.findall(self.inputWord,self.soupFilteredText,re.I)
            # find the number of times the input word present in the webpage
            self.inpWordCount = len(findList)
            # update the wordsDict with the Count value
            if self.wordsDict.get(self.inputWord.lower()):
                self.wordsDict[self.inputWord.lower()] = self.inpWordCount
                
            # print the number of times the input word present in the webpage    
            if self.inpWordCount > 1:
                print(f'\nThe word "{self.inputWord}" is present {self.inpWordCount} times.')
            elif self.inpWordCount == 1:
                print(f'\nThe word "{self.inputWord}" is present 1 time.')
            else:
                print(f'\nThe word "{self.inputWord}" is not present in the website.')

    def printSearchHistory(self):
        # establishes a connection with sqlite3 database
        connection = sqlite3.connect("WordsDB.db")

        createTable = '''CREATE TABLE IF NOT EXISTS search_history
                        (search_id integer PRIMARY KEY, website_url text, search_word text,
                        word_count integer, last_search_date int, search_count integer)'''
        # creates the table if it doesn't exist
        connection.execute(createTable)

        # max of search_id is saved in cursor_max
        cursor_max = connection.execute("SELECT MAX(search_id) FROM search_history")
        # calculate the SearchID to be inserted
        for row in cursor_max:
            # incremented by 1 if max of search_id is available
            if row[0] != None:
                insertSearchID = row[0] + 1
            # assigned as 1 if no row exists
            else:
                insertSearchID = 1
                
        # to check if the word and website already exists in search history
        existsSQL = "SELECT search_id,website_url,search_word,last_search_date FROM search_history WHERE website_url = ? AND search_word = ?"
        cursor_exists = connection.execute(existsSQL,(self.websiteURL,self.inputWord.lower()))
        
        rowExists = 0
        for row in cursor_exists:
            rowExists = 1
            # update if already exists in history and the word is present
            if self.inpWordCount != 0:
                updateSQL = "UPDATE search_history SET word_count = ?, last_search_date = ?, search_count = search_count + 1 WHERE search_id = ?"
                connection.execute(updateSQL,(self.inpWordCount,datetime.now().strftime("%B %d, %Y %I:%M%p"),row[0]))
            # delete if row exists and the word is not present now
            elif self.inpWordCount == 0:
                connection.execute("DELETE FROM search_history WHERE search_id = ? and search_word = ?", (row[0],self.inputWord.lower()))
                
        # insert if not exists already in history and if the word is present in the webpage
        if rowExists ==0 and self.inpWordCount != 0:
            insertSQL = "INSERT INTO search_history VALUES (?,?,?,?,?,?)"        
            connection.execute(insertSQL,(insertSearchID,self.websiteURL,self.inputWord.lower(),self.inpWordCount,datetime.now().strftime("%B %d, %Y %I:%M%p"),1))
            
        # make the changes to the database permanent
        connection.commit()

        # retrieve the latest 5 rows                       
        cursor = connection.execute("SELECT * FROM search_history ORDER BY last_search_date DESC LIMIT 5")

        print(f'\n{"-"*30}Recent Search History{"-"*30}')
        print("\nWebsite URL \t\t Search Word \t Word Count \t\t Last Search Date \t Search Count")
        # display the rows
        for row in cursor:
            print(f'{row[1]} \t {row[2].ljust(6)} \t {row[3]} \t\t {row[4]} \t {row[5]}')

        # if user says yes, delete the search history    
        if input("\nDo you want to delete the recent search history? (Y/N) ") in ('Y','y'):
            connection.execute("DELETE FROM search_history")
            connection.commit()
            print("Deleted successfully!")
            
    def writeWordsToExcel(self):
        # Returns a Workbook object
        workbook = xlsxwriter.Workbook('D:\python\project\WebsiteWords.xlsx')
        # Returns a Format object
        bold = workbook.add_format({'bold': True})
        
        #Returns a worksheet object
        worksheet = workbook.add_worksheet('Words')
        worksheetStop = workbook.add_worksheet('StopWords')

        # increase the column width    
        worksheet.set_column('A:A',20)
        worksheetStop.set_column('A:A',13)
            
        # write(row,column,token,[format]) or write(cell notation,token,[format])
        worksheet.write('A1','Words',bold)
        # make the column headings bold
        worksheet.write('B1','Count',bold)
        worksheetStop.write('A1','Stop Words',bold)
        
        row,col = 1,0
        # write the words and the count to the spreadsheet
        for word,count in self.wordsDict.items():
            worksheet.write_string(row,col,word)
            worksheet.write_number(row,col+1,count)
            row += 1
            
        row,col = 1,0
        # write the stop words to the stop words worksheet
        for word in self.stopWords:
            worksheetStop.write_string(row,col,word)
            row += 1
                
        workbook.close()
        print("\nWebsite words and Stop words written successfully to excel WebsiteWords.xlsx") 
        
web = Website(input("Enter the Website URL: "))

try:
    web.writeSoupToFile()
    web.setWordsDict()
    web.printInpWordCount()
    web.printSearchHistory()
    web.writeWordsToExcel()
    
# handle the exceptions raised if the user provides an invalid url
except requests.exceptions.InvalidSchema:
    print("Invalid Website Schema")
except requests.exceptions.MissingSchema:
    print("Missing Website schema")
except requests.exceptions.InvalidURL:
    print("Invalid Website URL")
except requests.exceptions.ConnectionError:
    print("Connection Failed")
except requests.exceptions.HTTPError as e:
    # re.match(pattern,string) returns a Match object
    print("HTTP Error: ",re.match(".+(?=for url:)", str(e)).group())
# handle the exception raised if the excel file is open
except PermissionError:
    print("Permission error: Close the excel file and try again")
except Exception as e:
    print("\nException:")
    # Returns a tuple
    exc_info = sys.exc_info()
    # * unpacks a tuple into positional arguments
    traceback.print_exception(*exc_info)
