# Author: Jordan La Croix
#
# Queries open movie database and uploads contents
# to a google drive spreadsheet under the username/password

import sys
import urllib2
import json
import re
import time
import gdata.spreadsheet.service
import gdata.docs.service

sheetkey = "1q8JUMZ0a4VywZZmg4bh4YFpoQyNBNVb6Xo5Ta4Ikxj8"
worksheetkey = "default"
days = ["Monday" , "Tuesday" , "Wednesday" , "Thursday" , "Friday" , "Saturday" , "Sunday"]
months = [0 , "January" , "February" , "March" , "April" , "May" , "June" , "July" , "August" ,
          "September" , "October" , "November" , "December"]

def main():
    username = sys.argv[1]
    password = sys.argv[2]
    inputfile = sys.argv[3]

    # Makes a list containing movie ids
    movieids = []
    with open(inputfile , "r") as inputfile:
        for line in inputfile:
            movieids.append(line[:len(line) - 1])
    inputfile.close()

    # Logs into google spreadsheet account
    client = gdata.spreadsheet.service.SpreadsheetsService()
    client.ClientLogin(username, password)
    
    #client.UpdateCell(1 , 1 , "TEST" , sheetkey , "default")

    row = 1
    for movie in movieids:
        query = "http://www.omdbapi.com/?i=" + movie + "&r=json&tomatoes=true"
        result = urllib2.urlopen(query)
        jsonData = json.loads( result.read() )
        createrow(client , row , jsonData)
        row += 1


def createrow(client , row , jsonData):
    # Title
    client.UpdateCell(row , 1 , jsonData['Title'] , sheetkey , "default")

    # Genre
    genres = jsonData['Genre'].split(", ")
    client.UpdateCell(row , 2 , genres[0] , sheetkey , "default")

    # Runtime
    runtime = jsonData['Runtime'].split(" ")
    client.UpdateCell(row , 3 , runtime[0] , sheetkey , "default")

    # Rating
    client.UpdateCell(row , 4 , jsonData['Rated'] , sheetkey , "default")

    # Production Company
    client.UpdateCell(row , 5 , jsonData['Production'] , sheetkey , "default")

    # Box Office Release 10 Dec 2013
    boxrelease = time.strptime(jsonData['Released'] , "%d %b %Y")
    client.UpdateCell(row , 6 , days[boxrelease.tm_wday] , sheetkey , "default")
    client.UpdateCell(row , 7 , months[boxrelease.tm_mon] , sheetkey , "default")
    client.UpdateCell(row , 8 , str(boxrelease.tm_year) , sheetkey , "default")

    # DVD Release
    dvdrelease = time.strptime(jsonData['DVD'] , "%d %b %Y")
    client.UpdateCell(row , 9 , days[dvdrelease.tm_wday] , sheetkey , "default")
    client.UpdateCell(row , 10 , months[dvdrelease.tm_mon] , sheetkey , "default")
    client.UpdateCell(row , 11 , str(dvdrelease.tm_year) , sheetkey , "default")

    # Box Office Revenue
    revenue = jsonData['BoxOffice']
    revenuemils =  revenue[1:re.search("[.]" , revenue).pos]
    client.UpdateCell(row , 12 , revenuemils , sheetkey , "default")

    # Director
    client.UpdateCell(row , 16 , jsonData['Director'] , sheetkey , "default")

    # Tomato Score
    client.UpdateCell(row , 17 , jsonData['tomatoMeter'] , sheetkey , "default")

    # IMDB Rating
    client.UpdateCell(row , 18 , jsonData['imdbRating'] , sheetkey , "default")

    # Metascore
    client.UpdateCell(row , 19 , jsonData['Metascore'] , sheetkey , "default")

main()
