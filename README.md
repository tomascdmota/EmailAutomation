# Email automation
Simple python script that checks the companies email for receipts or important PDFs and automatically saves and downloads them, for accounting purposes.
 It checks for the date, if a folder with that date isnt yet created on the desktop, it creates one then adds the files there. By the end of the month it should send an email to our accounting firm with every odf
 #TODO label files (name and date) for an easier understanding


# To run as a service

 1- Change email and pw to the correct ones
 2- Run pyinstaller --onefile main.py
 3- nssm install EmailAutomation (for starting the service)