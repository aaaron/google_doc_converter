# google_doc_converter
Google apps script to retroactively convert already uploaded Word, Powerpoint and Excel files to the equivalent Google Docs formats

To run:
1. Go to http://script.google.com
2. Paste contents into blank Code.gs
3. Run
4. You will likely be prompted to turn on Advanced Drive service

Notes:
* Doesn't recurse any folders with "archive" in the title
* Doesn't convert any files containing a "~" in title (eg, recovered docs, which often can't be converted)
* If a file can't be converted successfully, it will rename it with ~ at end
* Google Apps only lets scripts run for ~6min before terminating them so you may need to write an Automator (MacOS) to replay script every 7 minutes or so


