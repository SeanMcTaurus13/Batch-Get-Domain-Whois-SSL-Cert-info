'''
Created on Aug 18, 2017

@author: Sean McCombs
This Program will accept a list of Domains via a file called 'domaintestlist.txt' (one domain per line). 
The file will need to be in the same directory of the program. When executed it will batch get whois data on the domains.
It will create an output file in the same directory called 'domain_results.xlsx'.
'''

import asyncio
import async_timeout
import whois
from _overlapped import NULL
import xlsxwriter
import re

async def getwhois(domain = ''):
    '''getwhois will return whois info for domain. '''
    with async_timeout.timeout(30):
        try:
            url = domain
            response = whois.whois(url)
            return response
        except whois.parser.PywhoisError:
            return NULL
                  
async def main2():
    domainsList = []
    resultsList = []

    print("Start Domain detail processing...")
    
    for dom in urls:
        print("now working on", dom)
        domainsList.append(dom)
        domaindetails = await getwhois(dom)
        if domaindetails == 0:
            resultsList.append("no match for \"%s\"" %dom)
        else:
            resultsList.append(domaindetails)
    print('Domain List:',domainsList)
    print('Whois RAW list:',resultsList)
    print("done.\nPlease check \"domain_results.xlsx\" in root directory.")

    global domains
    domains = domainsList
    global results
    results = resultsList

filename = 'domaintestlist.txt' 
urls = []

print('Batch Get Whois Domain info program.')
try: 
    f = open(filename,'r') 
    print("File \"", filename, "\" opened successfully.")
except IOError: 
    print("Couldn't open %s" %  filename)

while 1:
        line = f.readline()
        if not line:
            print("File details:")
            break
        line = re.sub('\n', '', line)
        urls.append(line)
print(urls)
            
#url = 'marketo.com'
#urls = ['marketo.com','google.com','blahblahnoidonotexist.com']
domains = results = []

loop = asyncio.get_event_loop()
loop.run_until_complete(main2())

workbook = xlsxwriter.Workbook('domain_results.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1','Domains:')
worksheet.write('B1','Whois Results:')
count = 2
results_count = 0
resultstring = ''
for dom in domains:
    worksheet.write('A'+ str(count), dom)
    resultstring = str(results[results_count])
    worksheet.write('B'+ str(count), resultstring)
    #print(results[results_count])
    count = count + 1
    results_count = results_count + 1
    
workbook.close()