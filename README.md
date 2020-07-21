#doccoment
Calculate how many time passed between the first and the last review comment in Word documents.  

Put *.docx files into "resources" directory and run the script.  
Output example:
```bash
Checking: resources/from_office365.docx
Created on 2020-07-06 09:54:59
Comments are from:
2020-07-06T11:55:25
2020-07-06T11:55:32
2020-07-06T11:55:44
2020-07-06T14:34:33
Oldest: 2020-07-06 11:55:25
Newest: 2020-07-06 14:34:33
Time passed between first and last comment: 2:39:08
Time passed creation and last comment: 4:39:34
Checking: resources/from_libreoffice.docx
Created on 2020-07-03 15:58:08
Comments are from:
2020-07-03T15:59:36Z
2020-07-03T16:00:02Z
2020-07-06T14:11:14Z
Oldest: 2020-07-03 15:59:36
Newest: 2020-07-06 14:11:14
Time passed between first and last comment: 2 days, 22:11:38
Time passed creation and last comment: 2 days, 22:13:06
```