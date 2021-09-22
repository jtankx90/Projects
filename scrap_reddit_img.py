# -*- coding: utf-8 -*-
"""
Created on Tue Sep 21 19:12:55 2021

@author: nehneh
"""
import praw
import urllib
import os
import configparser

config = configparser.ConfigParser()
config.read(r'C:\Users\jtank\Desktop\config.ini')

reddit = praw.Reddit(
    client_id=config.get('credentials','client_id'),
    client_secret=config.get('credentials','client_secret'),
    password=config.get('credentials','password'),
    user_agent=config.get('credentials','username'),
    username=config.get('credentials','username'),
)


#subreddit name
subreddit = reddit.subreddit("dankmemes")

current_path = r'C:\Users\jtank\Desktop\cs50 scripts\cs50\final project\webshop'
count = 1
item_count = 1
myPath = os.path.join(current_path,'static','img')
for submission in subreddit.top("all"):
    url = str(submission.url)
    if url.endswith("jpg") or url.endswith("jpeg") or url.endswith("png"):
       fullfilename = os.path.join(myPath, f"item{item_count}_{count}.jpg")
       print(fullfilename)
       urllib.request.urlretrieve(url, fullfilename)
       count += 1
       #3 image for each item
       if count == 4:
            item_count+=1
            count =1
    #10 items
    if item_count ==11:
        break