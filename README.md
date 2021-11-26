# Scholarship Reporting

This is a project I worked on as the Executive Intern for the Office of Advancement at St. John Fisher College. 
I was given a few different spreadsheets that I had to pull information from. 
They had different formats so I used this to combine them into one sheet. 

This is definitely a quick and dirty solution.
It is messy and could be more efficient.
I was also teaching myself pandas while I was doing it, so I wouldn't say it necessarily followed pandas best practices. 
But it did the job, and that is the most important thing. 
In rereadign the code, I definitely see a lot that I would do different. 

## Statement on Privacy of Data
These spreadsheets contained a lot of personal information on both scholarship recipients and donors. 
None of these spreadsheets are on this GitHub repo. 
I made sure to go through the code and remove any and all personal information. 
Subsequently, a random individual perusing the internet will not be able to run this code. 

## This repo is here to: 
* Add work to my code portfolio.
* Show future people who have to do this work what I did. 
* Help me remember it for next time. 


## Aim of Project
The aim of the project is to send info packets to all donors who contribute to scholarships.
These info packets contain an update on the financials of the scholarhip 
(endowment value, growth in value YOY, amount of money recieved this year, amount of money awarded this year).

## Process of Project
I was first given a spreadsheet that had all of the donors who recieved info last year, 
as well as a list of people who were supposed to recieve letters this year. 
I compared the two, and removed the duplicates. Then I wrote merge.py

### merge.py
This file takes the new sheet from the database, and adds in the names that are missing from that 
sheet but on the sheet from the previous year. The output is a new primary sheet that has all of 
the names that need to recieve updates on their scholarships this year. 
The output of this file became the primary file that we worked off of/on/around for the rest of the project. 

### group_recips.py
This file takes a sheet that contained every recipient of a scholarship. Each row was a different individual that contained their scholarship name. But there could be multiple entries for recipient if they recieved multiple scholarships. So in database terms the primary key was a primary composite key of the values scholarship name and recipient. I needed this grouped by scholarship. This program groups the recipeints by scholarship. 

This was used multiple times, so I added a GUI file selector so that I could call it and keep chaning the sheet I was pulling from. 

### merge_recips.py
This was used to add the names of each scholarship's recipients into the main file that we were working out of. This main file was housed on Google Drive so that multiple people could be working on it. I downloaded it as an XLSX, ran the Python code to add the recipients, then copy and pasted just that column back up into the Google Sheet. The advantage to keeping the list of recipients in one column was that they were quick and easy to move around. Then I used the text to columns on Google Sheets to break it up. 

### group_donors.py
This is a quick and script that groups the scholarships by donors. This was used to find who needed to recieve multiple letters, so that we could combine them into one packet. 



## A Note on Racism
A while back, I read a 
[very interesting New York Times Article](https://www.nytimes.com/2021/04/13/technology/racist-computer-engineering-terms-ietf.html)
about the use of 'master' in computing. 
Because of this I try to avoid this word, and instead use HGTV's alternate word which is primary. 
One of the .xlsx files I was given did have master in the name, which I left as is for data integrity. 
However once I was using it in code, I referred to this sheet as primary. 
