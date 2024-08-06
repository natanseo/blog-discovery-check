import pandas as pd
import requests
import time
import re
from urllib.parse import urlparse
from email.message import EmailMessage
from email import encoders
import ssl
import smtplib
import pandas as pd
import openpyxl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

#make sure to add new URLs in urls list with "/" at the beginning and end.

urls = ["/discover/", "/discover/blog/", "/discover/blog/sustainability/save-energy-by-washing-clothes-at-30-degrees/",
    "/discover/blog/sustainability/10-ways-to-extend-the-life-of-your-clothes/", "/discover/blog/sustainability/guide-to-embracing-sustainable-living/",
    "/discover/blog/sustainability/easy-energy-saving-tips-for-home-appliances/", "/discover/blog/laundry/maximize-small-laundry-space/",
    "/discover/blog/laundry/front-load-vs-top-load-washing-machine/", "/discover/blog/laundry/how-to-organize-your-wardrobe/",
    "/discover/blog/laundry/how-to-wash-linen-silk-and-delicate-material/", "/discover/blog/cleaning/simple-tips-to-get-rid-of-refrigerator-smell/",
    "/discover/blog/cleaning/house-cleaning-checklist/", "/discover/blog/kitchen/the-complete-guide-to-steam-cooking/",
    "/discover/blog/cleaning/how-to-remove-stains-from-clothes/"]

def one_lang():  #only one language domains.
    #make sure to add new values both dictionaries below but urls with no trailing slash at the end.
    one_lang_domains = {"AEG UK":"https://www.aeg.co.uk", "AEG DE":"https://www.aeg.de", "ELX PL":"https://www.electrolux.pl",
            "AEG ES" : "https://www.aeg.com.es", "AEG NL": "https://www.aeg.nl", "ELX HU": "https://www.electrolux.hu", "ELX SE": "https://www.electrolux.se",
            "ELX IT": "https://www.electrolux.it", "ELX CZ": "https://www.electrolux.cz", "ELX FR" : "https://www.electrolux.fr", "ELX FI" : "https://www.electrolux.fi",
            "ELX ES": "https://www.electrolux.es", "ELX NO": "https://www.electrolux.no", "ELX DK": "https://www.electrolux.dk", "ELX SK": "https://www.electrolux.sk",
            "ELX HR": "https://www.electrolux.hr", "ELX TR": "https://www.electrolux.com.tr", "AEG IT":"https://www.aeg.it", "AEG FR":"https://www.aeg.fr", 
            "AEG AT":"https://www.aeg.at", "AEG FI":"https://www.aeg.fi", "AEG DK":"https://www.aeg.dk", "AEG NO":"https://www.aeg.no", "AEG SK":"https://www.aeg.sk",
            "AEG RO":"https://www.aeg.ro", "AEG PT":"https://www.aeg.com.pt", "AEG BG":"https://www.aeg.bg", "AEG LU":"https://www.aeg.lu" }
    
    one_lang_dic = {'AEG UK': {}, 'AEG DE': {}, "ELX PL": {}, "AEG ES": {}, "AEG NL": {}, "ELX HU": {}, "ELX SE": {}, "ELX IT": {}, "ELX CZ": {}, "ELX FR": {}, "ELX FI": {},
                    "ELX ES": {}, "ELX NO": {}, "ELX DK": {}, "ELX SK": {}, "ELX HR": {}, "ELX TR": {}, "AEG IT": {}, "AEG FR": {}, "AEG AT": {}, "AEG FI": {}, "AEG DK": {}, 
                    "AEG NO": {}, "AEG SK": {}, "AEG RO": {}, "AEG PT": {}, "AEG BG": {}, "AEG LU": {}}

    s = requests.Session()
    s.headers.update({'User-Agent': 'dontlockmeout'})

    for a,b in one_lang_domains.items():
        for c in urls:
            url = b+c   #merging domain name with full URL string from given list of URLs.
            
            scode = s.get(url, verify=False, allow_redirects=False).status_code    #checking header status for full URL
            scode = "No" if scode != 200 else "Y"

            cleaned_urls = f"/{url.split('://', maxsplit=1)[1].split('/', maxsplit=1)[1]}"
            one_lang_dic[a][cleaned_urls] = scode  #updating the master with header status.

    one_lang_df = pd.DataFrame.from_dict(one_lang_dic)
    one_lang_df.reset_index(inplace=True)
    one_lang_df = one_lang_df.rename(columns={'index': 'URLs'})

    return one_lang_df
def multi_lang():  #multiple language domains.

    #make sure to add new values both dictionaries below but urls with no trailing slash at the end.
    multi_lang_domains = {"AEG NL/BE": "https://www.aeg.be/nl-be", "AEG FR/BE": "https://www.aeg.be/fr-be",
            "ELX FR/CH": "https://www.electrolux.ch/fr-ch", "ELX IT/CH": "https://www.electrolux.ch/it-ch",
            "ELX DE/CH": "https://www.electrolux.ch/de-ch"}
    multi_lang_dic = {"AEG NL/BE": {}, "AEG FR/BE": {}, "ELX FR/CH": {}, "ELX IT/CH": {}, "ELX DE/CH": {}}

    s = requests.Session()
    s.headers.update({'User-Agent': 'dontLockMeOut'})

    for a,b in multi_lang_domains.items():
        for c in urls:
            url = b+c   #merging domain name with full URL string from given list of URLs.
            
            scode = s.get(url, verify=False, allow_redirects=False).status_code    #checking header status for full URL
            scode = "No" if scode != 200 else "Y"

            cleaned_urls = re.sub(r'^https?://(?:www\.)?(?:[^/]+/){2}', '', url)
            cleaned_urls = "/" + cleaned_urls
            multi_lang_dic[a][cleaned_urls] = scode  #updating the master with header status.

    multi_lang_df = pd.DataFrame.from_dict(multi_lang_dic)
    multi_lang_df.reset_index(inplace=True)
    multi_lang_df = multi_lang_df.rename(columns={'index': 'URLs'})
    return multi_lang_df

single_language = one_lang()
multiple_language = multi_lang()
df_merged = pd.merge(single_language, multiple_language, on="URLs") #merging both dataframes into one

#below is the code to send email with attachment

df_merged.to_excel("discovery_blogs.xlsx", index=False)  # Set index=False to avoid index row in excel

smtp_port = 587                 # Standard secure SMTP port
smtp_server = "smtp.gmail.com"  # Google SMTP Server
password = 'ckqg ojiv cvfb svhf'

email_from = "nabeelikea@gmail.com"
email_list = ["nabeel.tanveer@electrolux.com", "elyssha.pantig@electrolux.com", "marina.de-souza-nilsson@electrolux.com",
              "anders.lofgren@electrolux.com", "asadej.kitnitchiva@electrolux.com"]
subject = "Discovery Blogs (Publish_Status)"

# source: https://github.com/The-Intrigued-Engineer/python_emails/blob/main/with_attachments.py
def send_emails(email_list):
    for person in email_list:
        # the body of the email
        body = f"""

        This is an automated email generated by a script that runs a check across a list of markets and domains to see which markets
        have published the below URLs and which ones have not.
        
        {''.join(['- ' + url + '\n' for url in urls])}

        The results are attached where 'Y' indicates published, 'No' indicates Not Published.
        
        PS: Check with Nabeel if you wish to include/exclude more markets/domains.

        """

        # MIME object to define parts of the email
        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = person
        msg['Subject'] = subject

        # Attach the body of the message
        msg.attach(MIMEText(body, 'plain'))

        # Define the file to attach
        filename = "discovery_blogs.xlsx"

        # Open the file in python as a binary
        attachment= open(filename, 'rb')  # r for read and b for binary

        # Encode as base 64
        attachment_package = MIMEBase('application', 'octet-stream')
        attachment_package.set_payload((attachment).read())
        encoders.encode_base64(attachment_package)
        attachment_package.add_header('Content-Disposition', "attachment; filename= " + filename)
        msg.attach(attachment_package)

        # Cast as string
        text = msg.as_string()

        # Connect with the server
        print("Connecting to server...")
        TIE_server = smtplib.SMTP(smtp_server, smtp_port)
        TIE_server.starttls()
        TIE_server.login(email_from, password)
        print("Succesfully connected to server")
        print()


        # Send emails to "person" as list is iterated
        print(f"Sending email to: {person}...")
        TIE_server.sendmail(email_from, person, text)
        print(f"Email sent to: {person}")
        print()

    # Close the port
    TIE_server.quit()

# Run the function
send_emails(email_list)


#filtering the table by market/brands

# "AEG UK", "AEG DE", "AEG UK", "AEG ES", "AEG NL", "ELX HU", "ELX SE", "ELX IT", "ELX CZ", "ELX FR", "ELX FI", "ELX ES",
# "ELX NO",	"ELX DK", "ELX SK",	"ELX HR", "ELX TR", "AEG IT", "AEG FR", "AEG AT", "AEG FI",	"AEG DK", "AEG NO",	"AEG SK",
# "AEG RO",	"AEG PT", "AEG BG",	"AEG LU", "AEG NL/BE", "AEG FR/BE",	"ELX FR/CH", "ELX IT/CH", "ELX DE/CH",

# use the above site names to filter.

# filtered_df = df_merged[["URLs", "ELX TR", "AEG IT", "AEG FR", "AEG AT", "AEG FI", "AEG ES", "AEG NL", "ELX HU", "ELX SE", "ELX IT", "ELX CZ"]]
# filtered_df.to_clipboard()

# df_merged.to_clipboard()


