import win32com.client
import os
import codecs

signature_name = 'new'
recipients = ('phofmannh@gmail.com','phofmannh@gmail.com','phofmannh@gmail.com')

#load signature
sig_files_path = 'AppData\Roaming\Microsoft\Signatures\\' + signature_name + '_fichiers\\'
sig_html_path = 'AppData\Roaming\Microsoft\Signatures\\' + signature_name + '.htm'

signature_path = os.path.join((os.environ['USERPROFILE']), sig_files_path) # Finds the path to Outlook signature files with signature name "Work"
html_doc = os.path.join((os.environ['USERPROFILE']),sig_html_path)     #Specifies the name of the HTML version of the stored signature
html_doc = html_doc.replace('\\\\', '\\')

html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore') #Opens HTML file and converts to UTF-8, ignoring errors
signature_code = html_file.read()               #Writes contents of HTML signature file to a string

signature_code = signature_code.replace((signature_name + '_fichiers/'), signature_path)      #Replaces local directory with full directory path
html_file.close()

ol=win32com.client.Dispatch("outlook.application")
olmailitem=0x0 #size of the new email

for recipient in recipients:
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= 'CESI: Testing Mail'
    newmail.To='phofmannh@gmail.com'
    #newmail.CC='xyz@gmail.com'

    newmail.BodyFormat = 3 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
    newmail.HTMLBody = "Test message" + signature_code
    # attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)
    # To display the mail before sending it
    newmail.Display()