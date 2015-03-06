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

genre_priority = {'Action' : 12 , 'Sci-Fi' : 8 , 'Thriller' : 7 , 'Horror' : 9 , 'Mystery' : 5 ,
                  'Adventure' : 11 , 'Crime' : 6 , 'Drama' : 13 , 'Animation' : 14 , 'Comedy' : 1 ,
                  'Fantasy' : 10 , 'History' : 8 , 'Biography' : 7 , 'Romance' : 4 , 'Music' : 3 ,
                  'War' : 6 , 'NULL' : 100}

def main():
    username = sys.argv[1]
    password = sys.argv[2]
    inputfile = sys.argv[3]
    sheetkey = sys.argv[4]

    # Makes a list containing movie ids
    movieids = []
    with open(inputfile , "r") as inputfile:
        for line in inputfile:
            movieids.append(line[:len(line) - 1])
    inputfile.close()

    # Logs into google spreadsheet account
    client = gdata.spreadsheet.service.SpreadsheetsService()
    client.ClientLogin(username, password)
    

    row = 1
    for movie in movieids:
        query = "http://www.omdbapi.com/?i=" + movie + "&r=json&tomatoes=true"
        try:
            result = urllib2.urlopen(query)
            jsonData = json.loads( result.read() )
            create_entry(movie , client , row , sheetkey , jsonData)
        except Exception as ex:
            print ex
            pass
        row += 1

# Returns a primary genre from the genre list based on priority
def get_genre(genres):
    primary_genre = 'NULL'
    for genre in genres:
        if genre_priority[genre] < genre_priority[primary_genre]:
            primary_genre = genre
    return primary_genre

def create_entry(movieID , client , row , sheetkey ,jsonData):
    # Primary Key/IMDB ID
    client.UpdateCell(row , 1 , movieID , sheetkey , "default")

    # Title
    if jsonData['Title'] is not None:
        client.UpdateCell(row , 2 , jsonData['Title'] , sheetkey , "default")

    # Genre
    if jsonData['Genre'] is not None:
        genre = get_genre(jsonData['Genre'].split(", "))
        client.UpdateCell(row , 3 , genre , sheetkey , "default")

    # Runtime
    if jsonData['Runtime'] is not None:
        tmp = jsonData['Runtime'].split(" ")
        client.UpdateCell(row , 4 , tmp[0] , sheetkey , "default")

    # Rating
    if jsonData['Rated'] is not None:
        client.UpdateCell(row , 5 , jsonData['Rated'] , sheetkey , "default")

    # Production Company
    if jsonData['Production'] is not None:
        client.UpdateCell(row , 6 , jsonData['Production'] , sheetkey , "default")


    # Box Office Release 10 Dec 2013
    if jsonData['Released'] is not None:
        date = time.strptime(jsonData['Released'] , "%d %b %Y")
        tmp = str(date.tm_year) + "-" + str(date.tm_mon) + "-" +str(date.tm_mday)
        client.UpdateCell(row , 7 , tmp , sheetkey , "default")

    # DVD Release
    if jsonData['DVD'] is not None:
        date = time.strptime(jsonData['DVD'] , "%d %b %Y")
        tmp = str(date.tm_year) + "-" + str(date.tm_mon) + "-" +str(date.tm_mday)
        client.UpdateCell(row , 8 , tmp , sheetkey , "default")

    # Box Office Revenue
    if jsonData['BoxOffice'] is not None:
        revenue = jsonData['BoxOffice']
        client.UpdateCell(row , 9 , revenue[1 : len(revenue) - 1] , sheetkey , "default")

    # Director
    if jsonData['Director'] is not None:
        client.UpdateCell(row , 10 , jsonData['Director'] , sheetkey , "default")

    # Tomato Score
    if jsonData['tomatoMeter'] is not None:
        client.UpdateCell(row , 11 , jsonData['tomatoMeter'] , sheetkey , "default")

    # IMDB Rating
    if jsonData['imdbRating'] is not None:
        client.UpdateCell(row , 12 , jsonData['imdbRating'] , sheetkey , "default")

    # Metascore
    if jsonData['Metascore'] is not None:
        client.UpdateCell(row , 13 , jsonData['Metascore'] , sheetkey , "default")

main()
