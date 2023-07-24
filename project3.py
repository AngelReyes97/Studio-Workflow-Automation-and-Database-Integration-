# Install argparse using pip: pip install argparse
import argparse
# sys module is part of the Python standard library, no installation needed
import sys
#csv module is part of the Python standard library, no installation needed
import csv
#Install pymongo using pip: pip install pymongo
import pymongo
#re module is part of the Python standard library, no installation needed
import re
#getpass module is part of the Python standard library, no installation needed
import getpass
#datetime module is part of the Python standard library, no installation needed
from datetime import datetime
#subprocess module is part of the Python standard library, no installation needed
import subprocess
#Install xlsxwriter using pip: pip install xlsxwriter
import xlsxwriter
#os module is part of the Python standard library, no installation needed
import os


myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["Frames_to_Fix"]
collection = mydb["Files"]
collection2 = mydb["Frames"]

fields = []
line2 = []
line3 = ["Show location", "frames to fix"]
xytech_folders = []
files_folder = []
frames_to_fix = []
frames_within_video = []

parser = argparse.ArgumentParser(description="Process frame files")

parser.add_argument("--files", dest = "filename", type = str, nargs = "+")
parser.add_argument("-xytech", dest ="xytech_file", type = str)
parser.add_argument("--verbose", dest = "verbose", action = "store_true", help='Enable verbose console output')
parser.add_argument("--output", dest = "output", choices = ['CSV', 'DB', 'XLS'], help='Output format (CSV or DB or XLS)')
parser.add_argument("--process", dest="video_file", help='Process a video file')

args = parser.parse_args()

input_files = args.filename
xytech_file = args.xytech_file
verbose = args.verbose
output = args.output
video_file = args.video_file

if input_files is None:
    print("No BL/Flames files selected.")
    sys.exit(2)

if xytech_file is None:
    print("No Xytech files selected.")
    sys.exit(2)

if xytech_file:
    try:
        with open(xytech_file, "r") as tech_file:
            for line in tech_file:
                if line.startswith("Producer:"): # in the file search for the line that begins with producer
                    fields.append(line.strip().split(':')[0]) # get the field which is only the first word split on semi colon and get the first word
                    line2.append(line.strip().split('Producer: ')[1]) # get the name

                if line.startswith("Operator:"):
                    fields.append(line.strip().split(':')[0]) # get the field which is only the first word split on semi colon and get the first word
                    line2.append(line.strip().split('Operator: ')[1]) # get the name
        
                if line.startswith("Job:"):
                    fields.append(line.strip().split(':')[0]) # get the field which is only the first word split on semi colon and get the first word
                    line2.append(line.strip().split('Job: ')[1]) #get the discription

                if line.startswith("Notes:"):
                    fields.append(line.strip().split(':')[0]) #get the field which is only the first word
                    line = next(tech_file) #since the note is not on the same line go to the next line
                    line2.append(line.strip()) #get the note
        
                if line.startswith("/"):
                    xytech_folders.append(line)

    except FileNotFoundError:
        print(f"Error : file '{xytech_file}' does not exists.")

if input_files:
    if output == 'DB':
        if collection2.count_documents({}) > 0:
            # Delete existing data from the collection
            collection2.delete_many({})
    for file in input_files:
        try:
            with open(file, "r") as current_file:
                for line in current_file:
                    if re.search(r"/images1/Avatar\b", line):
                        line_parse = line.split(" ")
                        current_folder = line_parse.pop(0)
                        sub_folder = current_folder.replace("/images1/Avatar", "")
                        files_folder.append(sub_folder)
                        
                    elif re.search(r"/net/flame-archive\b", line):
                        line_parse = line.split(" ")
                        current_folder = line_parse.pop(1)
                        sub_folder = current_folder.replace("Avatar", "")
                        flame_folder_waste = line_parse.pop(0)
                        files_folder.append(sub_folder)
                        

                    new_location = ""

                    for xytech_line in xytech_folders:
                        if sub_folder in xytech_line:
                            new_location = xytech_line.strip()
                    
                    first = ""
                    pointer = ""
                    last = ""
                    parts = file.split('_')
                    submitted_user = getpass.getuser()
                    machine = parts[0]
                    user = parts[1]
                    date_str = parts[2].split('.')[0] # Remove file extension from date
                    date = datetime.strptime(date_str, '%Y%m%d').strftime('%m/%d/%Y')  # Convert date to desired format
                    submitted_date = datetime.now().strftime('%m/%d/%Y %H:%M:%S')

                    for numeral in line_parse:
                        if not numeral.strip().isnumeric(): # check if data is numeric and if not numeric continue
                            continue

                        if first == "":
                            first = int(numeral) # assign that first number to first convert it to an int
                            pointer = first #pointer is also assigned to first number
                            continue
                        if int(numeral) == (pointer+1): #if both numbers are the same not out of range
                            pointer = int(numeral) # update the pointer
                            continue
                        else:
                            #Range ends or no sucession, output
                            last = pointer
                            if first == last:
                                frame = "%s %s" % (new_location, first)
                                frames_to_fix.append(frame)
                                if output == 'DB':
                                    frames_DB = {
                                                'User': user,
                                                'Date': date,
                                                'Location': new_location,
                                                'Frame(s)': "%d" % (first)
                                                }
                                    collection2.insert_one(frames_DB)

                                
                            else:
                                frame = "%s %s-%s" % (new_location,first,last)
                                frames_to_fix.append(frame)
                                if output == 'DB':
                                    frames_DB = {
                                                'User': user,
                                                'Date': date,
                                                'Location': new_location,
                                                'Frame(s)': "%d-%d" % (first,last)
                                                }
                                    collection2.insert_one(frames_DB)

                            first= int(numeral)
                            pointer=first
                            last=""
                #Working with last number each line
                    last = pointer
                    if first != "":
                        if first == last:
                            frame = "%s %s" % (new_location, first)
                            frames_to_fix.append(frame)
                            if output == 'DB':
                                frames_DB = {
                                            'User': user,
                                            'Date': date,
                                            'Location': new_location,
                                            'Frame(s)': "%d" % (first)
                                            }
                                collection2.insert_one(frames_DB)

                        else:
                            frame = "%s %s-%s" % (new_location,first,last)
                            frames_to_fix.append(frame)
                            if output == 'DB':
                                sequence = "%d-%d" % (first,last)
                                frames_DB = {
                                            'User': user,
                                            'Date': date,
                                            'Location': new_location,
                                            'Frame(s)': "%d-%d" % (first,last)
                                            }
                                collection2.insert_one(frames_DB)

            if output == 'DB':
                print('Outputting to DB')
                #Collection 1 <User that ran script> <Machine> <Name of User on file> <Date of file> <submitted date>
                files_DB = {
                        'User that ran script': submitted_user,
                        'Machine': machine,
                        'User': user,
                        'Date': date,
                        'Submitted Date': submitted_date
                        }
                collection.insert_one(files_DB)
                


        except FileNotFoundError:
            print(f"Error : file '{xytech_file}' does not exists.")

if verbose:
    print('Verbose mode enabled')
    for lines in frames_to_fix:
        print(lines)

if output == 'CSV':
    print('Outputting to CSV')

    with open('frames_to_fix.csv', 'w', newline = '') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter= "/")
        csvwriter2 = csv.writer(csvfile, delimiter = "\n")
        csvwriter.writerow(fields)
        csvwriter.writerow(line2)
        csvwriter.writerow(line3)
        csvwriter2.writerow(frames_to_fix)


def frame_to_timecode(frame_number, fps): #function to conver to timecode
    # Calculate hours, minutes, seconds, and frames
    hours = frame_number // (fps * 3600)
    minutes = (frame_number % (fps * 3600)) // (fps * 60)
    seconds = (frame_number % (fps * 60)) // fps
    frames = frame_number % fps

    # Format timecode as string
    timecode = '{:02d}:{:02d}:{:02d}:{:02d}'.format(hours, minutes, seconds, frames)

    return timecode

def middle_range_timcode(start,end, fps = 60): #fucntion to get middle range then convert to timecode
    middle_frame = (start + end) // 2
    hours = middle_frame // (fps * 3600)
    minutes = (middle_frame % (fps * 3600)) // (fps * 60)
    seconds = (middle_frame % (fps * 60)) // fps
    frames = middle_frame % fps
    timecode = '{:02d}:{:02d}:{:02d}.{:02d}'.format(hours, minutes, seconds, frames)
    return timecode
    

if video_file:
    #video fps
    fps = 60 

    #Get the video file duration in seconds with FFmpeg command using subprocess
    video_duration = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of', 'default=noprint_wrappers=1:nokey=1', video_file]
    result = subprocess.run(video_duration, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    res = result.stdout.strip() #gets reuslt 
    duration = float(res) #convert it to float

    #convert to give the range of video, results in FPS of 5977
    frames_per_sec = int(duration * fps)


   # Retrieve all documents from the collection
    get_DB_frames= collection2.find()

    # Process the retrieved documents
    for frame in get_DB_frames:
        location = frame['Location']    #gets the location
        frame_range = frame['Frame(s)'] #gets the frame
        if '-' in frame_range:  #if its not a signle frame 
            start, end = map(int, frame_range.split('-')) #split it the range and assign first int to start and last int to end
            if start <= frames_per_sec or end <= frames_per_sec: #checks if range falls in the video file range
                start_timecode = frame_to_timecode(start, fps) #converts start to timecode
                end_timecode = frame_to_timecode(end, fps)  #converts end to timecode
                timecode = "%s / %s" % (start_timecode,end_timecode) #apply as a timecode range
                thumb_nail_timecode = middle_range_timcode(start,end) #get the middle range timecode for thumbnail
                #append the location, orignal frame ranges, the converted time code of ranges, and the thumbnail timecode
                frames_within_video.append((location, frame_range, timecode, thumb_nail_timecode))
        else:   # else if its a signle frame we dont include it and do nothing
            start = int(frame_range)
            end = start

    if output == 'XLS':
        # Create a new Excel file
        workbook = xlsxwriter.Workbook('project3.xlsx')
        worksheet = workbook.add_worksheet('Frames')

        # Write the headers
        headers = ['Location', 'Frame Range', 'Timecode', 'Thumbnail']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        # Set the dimensions for the thumbnail images
        thumb_width = 96
        thumb_height = 74

        # Write the frames within the video range
        for row, frame in enumerate(frames_within_video, start=1):
            location, frame_range, timecode, thumb_nail_timecode = frame
            worksheet.write(row, 0, location)
            worksheet.write(row, 1, frame_range)
            worksheet.write(row, 2, timecode)

            # Generate thumbnail
            thumbnail_folder = "thumbnails"  # Provide the folder path to save the thumbnails
            os.makedirs(thumbnail_folder, exist_ok=True)
            thumbnail_path = os.path.join(thumbnail_folder, f"thumbnail_{row}.png")

            thumb = [
            "ffmpeg", "-ss", thumb_nail_timecode, "-i", video_file, "-vframes", "1",
            "-vf", f"scale={thumb_width}:{thumb_height}:force_original_aspect_ratio=decrease",
            "-f", "image2", "-y", thumbnail_path
            ]
            subprocess.run(thumb)

            # Add thumbnail to XLS file
            worksheet.insert_image(row, 3, thumbnail_path, {'x_scale': 1, 'y_scale': 1})

            # Adjust row height to fit the thumbnail
            worksheet.set_row(row, thumb_height)


        # Adjust column widths based on data length
        for col, header in enumerate(headers):
            max_length = max([len(str(header))] + [len(str(frame[col])) if col < len(frame) else 0 for frame in frames_within_video])
            worksheet.set_column(col, col, max_length + 1)  # Add some extra padding

        # Close the workbook
        workbook.close()