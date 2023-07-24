# Studio-Workflow-Automation-and-Database-Integration-

Developed a script to handle multiple workflows and users for a studio, using Python and argparse
for argument handling.
Integrated the script with AutoDesk Flame files and parsed file names for database insertion.
Utilized MongoDB for data storage and retrieval, running the script multiple times to populate the
database.
Performed various database queries to answer specific questions related to user and machine data.
Exports are basic CSV files to XLS files with timecode and thumbnail preview that can uploaded to Frame.IO and/or Shotgrid

instructions:
1. Install argparse using pip: pip install argparse
2. Install pymongo using pip: pip install pymongo
3. Install xlsxwriter using pip: pip install xlsxwriter
4. Install MongoDB https://www.mongodb.com/
5. Install FFmpeg to capture thumbnail timecode (not need unless you use --process -videoname-)

Demo: 
https://youtu.be/apQBuAq_LKs
