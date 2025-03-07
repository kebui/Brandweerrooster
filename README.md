# Brandweerrooster
A script to fetch your availability schedule from brandweerrooster and convert it into an ics file which you can import / subscribe to in your calendar app.

I have created this PowerShell script to be able to see in my calendar app when I am on stand-by (duty) in Brandweerrooster. This way I can also easily share my duty schedule with the rest of the family. When you run the script, it will fetch your availability from Brandweerrooster for the next 30 days and turn it into in .ics (calendar) file. This file can be imported into your favorite calendar app like Google Agenda Microsoft Outlook, Apple agenda etc. If you host the .ics file on a webserver you can also subscribe to it to avoid recurring manual imports.

# How to get started
You will need your Brandweerrooster login credentials (username and password)
You will also need your membership ID. You can find this in the URL when you login to Brandweerrooster and go to  the exception schedule.
The URL will look something like this: https://www.brandweerrooster.nl/memberships/12345/schedule/exception_schedule?locale=nl. The number in the URL is your membership ID.

Once you have this information you can edit the script and enter in the variables.
- First set a working directory. This is where the script and other files will be stored. In my case I use "C:\Script". You can change this if you like.
- Enter your username (usually emailaddress) you use to login to Brandweerrooster.
- Fill in your membership ID

In PowerShell change directory to the working directory you chose, e.g. "cd C:\Brandweerrooster\PowerShell" and run these two commands:
- $securePassword = Read-Host "Enter password" -AsSecureString
- $securePassword | ConvertFrom-SecureString | Set-Content ".\password.txt"

This will prompt with a window that will ask for your Brandweerrooster password. It will store your password encrypted.

# How to run the script
The script can be run manually in PowerShell or as a scheduled task. It will output a file calendar called brandweercalendar.ics inside the working directory you specified. In my case the file gets copied over to a webserver so I can add this to my calendar app as a 'subscribed calendar'.
For troubleshooting purposes the script will output a log file that you can review to troubleshoot any potential errors.







