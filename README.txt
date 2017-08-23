-----------------------------------------------------------------------------------------------------------------------------
Program Name: Batch_Get_Whois_Info.py
Python Verison: Python 3.5.1
Created on Aug 18, 2017
@author: Sean McCombs
Program module dependencies: asyncio, async_timeout, whois, overlapped, xlsxwriter, re
Some modules may require you to install them via the pip installer. For example "whois", see 'pip_install_python-whois.png' .
and below resource:
https://stackoverflow.com/questions/4750806/how-do-i-install-pip-on-windows

This Program will accept a list of Domains via a file called 'domaintestlist.txt' (one domain per line). 
The file will need to be in the same directory of the program. When executed it will batch get whois data on the domains.
It will create an output file in the same directory called 'domain_results.xlsx'.
-----------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------------------------------
Program Name: Batch_Get_SSL_Cert_Info.py
Python Verison: Python 3.5.1
Created on Aug 19, 2017
@author: Sean McCombs
Program module dependencies: ssl, OpenSSL, re, xlsxwriter
Some modules may require you to install them via the pip installer. For example "whois", see 'pip_install_python-whois.png' .
and below resource:
https://stackoverflow.com/questions/4750806/how-do-i-install-pip-on-windows

This Program will accept a list of Domains via a file called 'domaintestlist.txt' (one domain per line). 
The file will need to be in the same directory of the program. When executed it will batch get SSL Cert Info, key and expiration.
It will create an output file in the same directory called 'domain_SSL_results.xlsx'.
-----------------------------------------------------------------------------------------------------------------------------

