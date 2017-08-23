'''
Created on Aug 19, 2017

@author: Sean McCombs
This Program will accept a list of Domains via a file called 'domaintestlist.txt' (one domain per line). 
The file will need to be in the same directory of the program. When executed it will batch get SSL Cert Info, key and expiration.
It will create an output file in the same directory called 'domain_SSL_results.xlsx'.
'''

import ssl
import OpenSSL
import re
import xlsxwriter

filename = 'domaintestlist.txt' 
hosts = certs = certdetails = []

print('Batch Get SSL cert info.')
try: 
    f = open(filename,'r') 
    print("File \"", filename, "\" opened successfully.")
except IOError: 
    print("Couldn't open %s" %  filename)

workbook = xlsxwriter.Workbook('domain_SSL_results.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1','Domains:')
worksheet.write('B1','CERT:')
worksheet.write('C1','CERT Details:')
worksheet.write('D1','Cert Expiration date:')
worksheet.write('E1','Friendly Expiration date:')

A_count = 2
results_count = 0
resultstring = ''

while 1:
        line = f.readline()
        if not line:
            print("File details:")
            break
        line = re.sub('\n', '', line)
        hosts.append(line)
        worksheet.write('A'+ str(A_count), line)
        A_count = A_count + 1
print(hosts)



#print (c.getpeercert())'''
#hosts = ['marketo.com','smmblackbox.com','microsoft.com','google.com','yahoo.com']
host = 'marketo.com'
cert = certinfo = errorStr = ''
B_count = C_count = D_count = E_count = 2

for doms in hosts:
    print("now working on", doms)
    try:
        cert =  (ssl.get_server_certificate((doms, 443)))
        print(cert)
        #certs.append(cert)
        worksheet.write('B'+ str(B_count), str(cert))
        B_count = B_count + 1
    except ConnectionRefusedError:
        print("\n---------------------------\n",doms,"\nConnectionRefusedError: [WinError 10061] No connection could be made because the target machine actively refused it")
        errorStr = "ConnectionRefusedError: [WinError 10061] No connection could be made because the target machine actively refused it"
        #certs.append(errorStr)
        worksheet.write('B'+ str(B_count), str(errorStr))
        B_count = B_count + 1
        errorStr = ''

    except TimeoutError:
        print("\n---------------------------\n",doms,"\nTimeoutError: [WinError 10060] A connection attempt failed because the connected party did not properly respond after a period of time, or established connection failed because connected host has failed to respond")
        errorStr = "TimeoutError: [WinError 10060] A connection attempt failed because the connected party did not properly respond after a period of time, or established connection failed because connected host has failed to respond"
        #certs.append(errorStr)
        worksheet.write('B'+ str(B_count), str(errorStr))
        B_count = B_count + 1
        errorStr = ''
        
    except:
        print("\n---------------------------\n",doms,"\nGeneral Error, \"socket.gaierror: [Errno 11004] getaddrinfo failed\" (get address info function exception), or Domain does not resolve to an IP address.")
        errorStr = "General Error, \"socket.gaierror: [Errno 11004] getaddrinfo failed\" (get address info function exception), or Domain does not resolve to an IP address."
        #certs.append(errorStr)
        worksheet.write('B'+ str(B_count), str(errorStr))
        B_count = B_count + 1
        errorStr = ''
        
    # OpenSSL
    if cert != '':
        x509 = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, cert)
        certinfo = x509.get_subject().get_components()
        print(certinfo)
        #certdetails.append(x509.get_subject().get_components())
        worksheet.write('C'+ str(C_count), str(certinfo))
        
        certinfo = x509.get_subject().get_components()
        certExpDate = x509.get_notAfter()
        #print(certinfo)
        print("Cert Expiration date:\n",certExpDate)
        worksheet.write('D'+ str(D_count), str(certExpDate))
        datestring = year = month = day = hour = minute = second = friendlycertExpDate = ''
        datestring = str(certExpDate)
        year = (datestring[2:6])
        month = (datestring[6:8])
        day = (datestring[8:10])
        hour = (datestring[10:12])
        minute = (datestring[12:14])
        second = (datestring[14:17])
        friendlycertExpDate = ("%s-%s-%s %s:%s:%s"%(year,month,day,hour,minute,second))
        print('Friendly Expiration date:\n',friendlycertExpDate)
        worksheet.write('E'+ str(E_count), str(friendlycertExpDate))
        
        C_count = C_count + 1
        D_count = D_count + 1
        E_count = E_count + 1
        cert = certinfo = ''
        
    else:
        print('No cert.\n---------------------------')
        errorStr = 'No cert.'
        #certdetails.append('No cert.\n---------------------------')
        worksheet.write('C'+ str(C_count), str(errorStr))
        C_count = C_count + 1
        D_count = D_count + 1
        E_count = E_count + 1
        cert = errorStr = ''
            
workbook.close()
print("\nBatch work completed.")
