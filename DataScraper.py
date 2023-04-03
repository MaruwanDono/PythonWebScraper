#Dec 2020
from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import os
import time

#Commands to remotely control Firefox browser
FirefoxDriverPath = os.getcwd()+'/geckodriver_folder/geckodriver_linux'
driver = webdriver.Firefox(executable_path=r'{}'.format(FirefoxDriverPath))

#Function to sign in to example.com
def SignIn(email,password):
    driver.get("https://example.com/account/login")
    driver.find_element_by_name("email").send_keys(email)
    driver.find_element_by_name("password").send_keys(password)
    #driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/main/div[2]/form/button").click()
    driver.find_element_by_xpath("/html/body/div[2]/div[1]/main/div/div/div/form/button").click()
    time.sleep(3)


#Function to extract title
def ExtractTitle(Data):
    Title = ""
    counter = 0
    while(counter<10):
        if [-1,-1] != [Data[counter].getText().find("Company Participants"),Data[counter].getText().find("Executives")]:
            break
        else:
            Title += Data[counter].getText()+"\n"
            counter += 1
    return Title


#Function to extract company participants
def ExtractCompanyParticipants(Data):
    CompanyParticipants = []
    Key = False
    for element in Data:
        if(element.find("strong")):
            if [-1,-1] != [element.getText().find("Company Participants"),element.getText().find("Executives")]:
                Key = True
            elif CompanyParticipants:
                Key = False
                break
        if Key:
            CompanyParticipants.append(element.getText())
    return CompanyParticipants[1:]


#Function to extract Conference Call Participants
def ExtractConfCallParticipants(Data):
    ConfCallParticipants = []
    Key = False
    for element in Data:
        if(element.find("strong")):
            if [-1,-1] != [element.getText().find("Conference Call Participants"),element.getText().find("Analysts")]:
                Key = True
            elif ConfCallParticipants:
                Key = False
                break
        if Key:
            ConfCallParticipants.append(element.getText())
    return ConfCallParticipants[1:]


#function to extract Intros
def ExtractOpeningWords(Data,CompanyParticipants):
    OpeningWords = {}
    names = [name.split(" - ")[0].split(" – ")[0] for name in CompanyParticipants]
    names.extend(CompanyParticipants)
    Key = False
    name = ""
    counter = 0
    for element in Data:
        elementText = element.getText()
        if(element.find("strong")):
            if elementText in names:
                Key = True
                name = elementText.split(" - ")[0].split(" – ")[0]+"_{}".format(counter)
                counter += 1
            elif elementText.lower().find("Question-and-Answer Session".lower()) != -1:
                Key = False
                break
            else:
                Key = False
        if Key and name.split("_")[0] != elementText.split(" - ")[0].split(" – ")[0]:
            #Multiple Intros for 1 person
            if name in OpeningWords:
                OpeningWords[name][0] += elementText + "\n"
            else:
                OpeningWords.update({name:[elementText]})


    return OpeningWords


#Function to extract Dialogue
def ExtractCompParticipantsDialogue(Data,CompanyParticipants):
    ParticipantsTurns, ParticipantsText = [], []
    names = [name.split(" - ")[0].split(" – ")[0] for name in CompanyParticipants]
    names.extend(CompanyParticipants)
    FirstKey = False
    SecondKey = False
    name = ""
    Text = ""
    counter = 1
    for element in Data:
        elementText = element.getText()
        if (not FirstKey) and (elementText.lower().find("Question-and-Answer Session".lower())!= -1):
            FirstKey = True
        elif FirstKey:
            #Elements in the Q and A section
            #Condition to identify a name from a reply
            #Bug from adsk-q4-2017, point written in bold, and others
            if (element.find("strong")):
                if len(element.find("strong").getText())>1:
                    if elementText in names:
                        SecondKey = True
                        name = elementText
                        ParticipantsTurns.append("QA"+str(counter)+"_"+name.split(" - ")[0].split(" – ")[0])
                        counter += 1
                    else:
                        SecondKey = False

            if SecondKey and name != elementText:
                Text += elementText + "\n"
            elif name == elementText and Text:
                ParticipantsText.append(Text)
                Text = ""
    #last text to be appended manually
    ParticipantsText.append(Text)
    return ParticipantsTurns, ParticipantsText


#Function that uses the above functions to extract data from a given link
def ExtractDataFromLink(link):
        print("Extracting data from "+link)
        driver.get(link)
        #Get the source code of the page from the given link
        content = driver.page_source
        soup = BeautifulSoup(content,'html.parser')
        PageData = soup.find(id="a-body")
        #Organizing Page code source into a list called Data
        NumP = len(PageData.findAll("div", attrs={"class":"p_count"})) + 1
        Data = []
        for i in range(NumP):
            Data.extend(PageData.findAll("p", attrs={"class":"p p"+str(i+1)}))

        #Data Analysis using the above functions
        title = ExtractTitle(Data)
        CompanyParticipants = ExtractCompanyParticipants(Data)
        ConfCallParticipants = ExtractConfCallParticipants(Data)
        OpeningWords = ExtractOpeningWords(Data,CompanyParticipants)
        ParticipantsTurns, ParticipantsText = ExtractCompParticipantsDialogue(Data,CompanyParticipants)

        #### Generating a DataFrame with pandas, a matrix-like structure
        #### DataFrame will be later printed into a file
        CsvDictionary = {title:[""]}
        CsvDictionary.update({"Company Participants": CompanyParticipants})
        CsvDictionary.update({"Conference Call Participants":ConfCallParticipants})
        CsvDictionary.update(OpeningWords)
        for i in range(len(ParticipantsTurns)):
            CsvDictionary.update({ParticipantsTurns[i]:[ParticipantsText[i]]})
        df = pd.DataFrame.from_dict(CsvDictionary, orient="index").T
        return df, title

#Function to extract links from txt file
def getLinks(path):
    LinkDict = {}
    Links = []
    name = ""
    with open(path,'r') as linksFile:
        Lines = linksFile.read().splitlines()
        for Line in Lines:
            if(Line):
                if "http" in Line:
                    Links.append(Line)
                elif Links:
                    LinkDict.update({name:Links[:]})
                    Links.clear()
                    name = Line
                else:
                    name = Line
                    LinkDict.update({name:Links})




        #Dictionary updated manually for last Links
        LinkDict.update({name:Links})
    return LinkDict


#####################
#### Main Program ###
#####################


if __name__=="__main__":
    #Variables for testing
    #LinkA = "https://example.com/article-A"
    #LinkB = "https://example.com/article-B"
    #LinkC = "https://example.com/article-C"
    #LinkD = "https://example.com/article-D"
    #LinkE = "https://example.com/article-E"
    #Links = [LinkD,LinkA,LinkE]


    LinksFileName = "AllLinks.txt"
    email = "test.example@gmail.com"
    password = "TestPassword"
    OutFolder = "Excel_Files/ExcelFiles"


    ###Execution
    #Extracting Links
    LinkDict = getLinks(LinksFileName)
    #Signing in to example.com
    SignIn(email,password)

    #For single link testing, remove the quotation marks
    """
    dfC, titleC = ExtractDataFromLink(LinkC)
    dfA, titleA = ExtractDataFromLink(LinkA)
    dfC.to_csv("Csv_Files/testFileC.csv", index=False, encoding='utf-8')
    dfA.to_csv("Csv_Files/testFileA.csv", index=False, encoding='utf-8')
    """

    #Iterate over All links and extract data
    try:
        for name, Links in LinkDict.items():
            with pd.ExcelWriter(OutFolder+"/{}.xlsx".format(name)) as writer:
                print("\n"+name)
                for counter, Link in enumerate(Links):
                    try:
                        df, title = ExtractDataFromLink(Link)
                    except Exception as e:
                        print("Data extraction Error: "+str(e))
                        df = pd.DataFrame(columns=[''])
                        title = "Error"
                    #Process title
                    if(title == "Error"):
                        title = "Error {}".format(counter)
                    elif(title.find(")") != -1):
                        title = title.split("Results")[0].split("Earnings")[0]
                        title = title[title.find(")")+2:]
                    else:
                        title = "sheet {}".format(counter)

                    print("Writing sheet: "+title)
                    #All sheet names must be different within one excel file
                    #If not the sheet with the same name will be overwritten
                    df.to_excel(writer,sheet_name=title)

    except Exception as e:
        print("ExcelWriter error: "+str(e))
