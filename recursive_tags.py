import stagger
import tkinter, tkinter.filedialog as filedialog
from os import chdir, startfile, listdir, rename, getcwd, remove
import os.path
import csv
from time import sleep
from datetime import datetime
# =IFERROR(TRIM(LEFT(E2,FIND("(",E2)-1)),TRIM(E2))&".mp3"
'''
Get rid of .mp3.mp3's
Get rid of empty folders
Get rid of double DIcs albums and collapse into one folder
Consolidate Greatest Hits to one folder higher track counts
'''

mp3_tags_raw = '''
            id
            file_key : DONT MODIFY
            new_filename
            track_total
            track
            title
            album
            album_artist
            artist
            date
            disc
            disc_total
            genre
            comment
            action
            filepath
            '''
mp3_tag_list = [x.strip() for x in mp3_tags_raw.split('\n') if x.strip() != '']
special_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
indentation = '\t'
issues = []
track_count = 1

#root = tkinter.Tk()
#root.withdraw()
#workpath = filedialog.askdirectory(parent=root, initialdir="/", title='Please enter the artist directory for MP3 Tagging:')
#workpath = diropenbox('Choose mp3 folder for bulk tag editing:')
#=SUBSTITUTE(SUBSTITUTE(TRIM(RIGHT("00"&E2,2)&" "&F2)&".mp3","?",""),"*","-")
workpath = "C:\\Users\\erchampe\\Onedrive\\Music"
#workpath = 'D:\\Music'
#workpath = r'C:\Users\erchampe\OneDrive\Audiobooks'
#workpath = input('Please enter the artist directory for MP3 Tagging:\n')

def explore_dirs(directory, layer, log, quiet = True):

    chdir(directory)
    current_dir = os.path.split(getcwd())[1]
    if not quiet:
        print(indentation*layer+current_dir)
    subfolders = sorted([x for x in listdir('.') if os.path.isdir(x)])
    if len(subfolders) >= 1:
        for subfolder in subfolders:
            explore_dirs(os.path.join(getcwd(), subfolder), layer+1, log, quiet)
    tracks = sorted([x for x in listdir('.') if x[-4:].lower() == '.mp3'])
    if len(tracks) >= 1:
        for track in tracks:
            if not quiet:
                print(indentation*(layer+1)+'[.]', track)
            add_tracks(track, layer, log)
    chdir('..')

def add_tracks(track, layer, log):
    global track_count
    record = [str(track_count), track, '']
    track_count += 1
    try:
        mp3_tag = stagger.read_tag(track)
        for tag in mp3_tag_list[3:-2]:
            record.append(str(getattr(mp3_tag, tag)))
        comment_col = 13
        record[comment_col] = record[comment_col].replace('\n', '').replace('\r', '')
        if len(record[comment_col]) > 99:
            record[comment_col] = record[comment_col][:99]
        record.append('')   # action placeholder
        record.append(getcwd().strip())
        try:
            log.writerow(commas_out(record))
        except:
            pass
    except stagger.errors.NoTagError:
        #print(indentation*(layer+1), track, 'Tag cannot be read')
        for count in range(11):
            record.append('')
        record.append('Tag cannot be read;')
        record.append(getcwd().strip())
        log.writerow(commas_out(record))


def commas_out(string_list):
    for i, element in enumerate(string_list):
        if ',' in element:
            string_list[i] = element.replace(',', '{c}')
    return string_list

def commas_in(string_list):
    for i, element in enumerate(string_list):
        if '{c}' in element:
            string_list[i] = element.replace('{c}', ',')
    return string_list


# Begin Main #
track_count = 1
chdir(workpath)
response = ''
while response != 'w' and response != 'r' and response != 'rw':
    response = input('You can (r)ead, (w)rite, or (rw) MP3 tags.\nPlease Enter Mode:')

if 'r' in response:
    if os.path.exists('Tags.csv'):
        os.rename('Tags.csv', 'Tags{}.csv'.format(datetime.now().strftime('_%Y%m%d_%H%M')))
    with open('Tags.csv', 'w', newline='') as csv_file:
        log = csv.writer(csv_file, delimiter=',', quoting=csv.QUOTE_MINIMAL, quotechar='"') # quotechar='|', , escapechar='\\')
        log.writerow(mp3_tag_list)  # Write the header
        explore_dirs(workpath, 0, log)
        chdir(workpath)
        startfile(os.path.join(workpath, 'Tags.csv'))

if 'r' in response and 'w' in response:
    input('Once Tag.csv has been saved and closed, press Enter to write the metadata.\n')

if 'w' in response:
    with open('Tags.csv', 'r', newline='') as csv_file:
        irecords = csv.reader(csv_file, delimiter=',', quoting=csv.QUOTE_MINIMAL, quotechar='"')  # quotechar='|', , escapechar='\\')
        records = [commas_in(x) for x in irecords]
    os.rename('Tags.csv', 'Tags{}.csv'.format(datetime.now().strftime('_%Y%m%d_%H%M')))
    with open('Tags.csv', 'w', newline='') as csv_file:
        log = csv.writer(csv_file, delimiter=',', quoting=csv.QUOTE_MINIMAL, quotechar='"')
        cursor = iter(records)
        log.writerow(next(cursor))  # Write the header back
        print('    Writing all mp3 metadata from Tag.csv to mp3s')
        for record in cursor:
            filename = record[1]
            new_filename = record[2].strip()
            mp3_tag_dict = dict(zip(mp3_tag_list[3:-2], record[3:-2]))
            #print(mp3_tag_dict)
            track_path = record[-1].strip()
            chdir(track_path)
            track_action = record[-2].strip()    # For possible avoidance when writing
            if not 'Tag cannot be read;' in track_action:
                # File Tagging Steps
                record[-2] = ''
                if not os.path.exists(os.path.join(track_path, filename)):
                    #print('File <{}> cannot be found in {}. It may have been renamed or removed.'.format(filename, track_path))
                    record[-2] = 'Not Found; ' + record[-2]
                    log.writerow(commas_out(record))
                    continue
                if record[9] and len(record[9]) != 4:
                    record[-2] = 'Invalid Year; ' + record[-2]
                    log.writerow(commas_out(record))
                    continue
                mp3_tag = stagger.read_tag(filename)
                for tag in mp3_tag_list[3:-2]:
                    #print('Set', [tag], 'as', [mp3_tag_dict[tag].strip()])      # For Verbosity
                    setattr(mp3_tag, tag, mp3_tag_dict[tag].strip())
                try:
                    mp3_tag.write()
                except PermissionError:
                    record[-2] = 'Permission Issue; ' + record[-2]
                    log.writerow(commas_out(record))
                    continue
                except TypeError as e:
                    record[-2] = 'Tag Write Error {}; '.format(e) + record[-2]
                    log.writerow(commas_out(record))
                    continue
                record[-2] = 'Tagged; ' + record[-2]
            # File Rename Steps
            if os.path.exists(os.path.join(track_path, new_filename)) and new_filename:
                record[-2] = 'Name Conflict; ' + record[-2]
                log.writerow(commas_out(record))
                #print('File cannot be renamed. {} already exists in {}.'.format(new_filename, track_path))
                continue
            if not new_filename.endswith('.mp3') and new_filename:
                record[-2] = 'Invalid Extension; ' + record[-2]
                log.writerow(commas_out(record))
                #print('File cannot be renamed. {} already exists in {}.'.format(new_filename, track_path))
                continue
            if sum([(x in new_filename) for x in special_chars]) > 0:
                for special_char in special_chars:
                    if special_char in new_filename:
                        print('Found "{}" in "{}"'.format(special_char, new_filename))
                        break
                record[-2] = 'Filename Special Char; ' + record[-2]
                log.writerow(commas_out(record))
                #print('File cannot be renamed. {} contains invalid character [{}].'.format(new_filename, track_path))
                continue
            if new_filename != '' and new_filename != filename:
                #print('Changing {} to {}'.format(filename, new_filename))
                rename(filename, new_filename)
                record[-2] = 'Renamed; '+record[-2]
            log.writerow(commas_out(record))
    startfile(os.path.join(workpath, 'Tags.csv'))
'''
Useful Code Snippet

            #if '#' in track:
                #issues.append('Bar in filename will cause issue,' + getcwd() + track)
                #old_track_filename = track
                #track = track.replace(',', ';')
                #input('Rename {} with \n\t {}?'.format(old_track_filename, track))
                #os.rename(old_track_filename, track)
'''
