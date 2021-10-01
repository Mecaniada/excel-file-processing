# excel-file-processing
This program is made for a better understanding on how openpyxl library works. It uses os, re, pyperclip, openpyxl and pandas. 
In addition to these libraries I've used a library called "alive_progress" wich it creates a progress/loading bar while the program executes a command.
I've used alive_progress in "Make me a sheet" command because it might take a little bit to create a new xlsx file with the filtered data. In "Make me a sheet" if you want to append some
data you have to copy a text ( I've used this pdf: https://www.blackbaud.com/files/customreports/PhoneDirectory_web.pdf ) and enter "Append data". If you want you can modify the regex
for your needs. For example the regex for email is: 

email_regex = re.compile(r'''
# some)>+thjing@gmail.com
[a-zA-z0-9_.+]+       # name part
@                     # @symbol
[a-zA-z0-9_.+]+       # domain name part

''', re.VERBOSE)

extract_mail_wd = email_regex.findall(data)
