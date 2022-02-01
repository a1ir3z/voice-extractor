# import things here
import os
from zipfile import ZipFile
from pydub import AudioSegment

# Make folders
if not 'voice' in os.listdir():
    os.makedirs('voice')
if not 'power' in os.listdir():
    os.makedirs('power')

# get a PPTX and copy it into a folder
# convert copied pptx into a zip file

POWERLIST =[]

for file in os.listdir():
	if file.endswith('.ppt'):
		file=file.replace('.ppt','')
		os.rename(file+'.ppt',file+'.zip')
		POWERLIST.append(file+'.zip')
	if file.endswith('.pptx'):
		file=file.replace('.pptx','')
		os.rename(file+'.pptx',file+'.zip')
		POWERLIST.append(file+'.zip')
for file in os.listdir():
	if file.endswith('.zip'):
		POWERLIST.append(file)



# extract zip file and take the voices out

for power in POWERLIST:
	with ZipFile(power,'r') as zipObj:
		for content in zipObj.namelist():
			if content.endswith('.wav') or content.endswith('.wma') or content.endswith('.m4a') or content.endswith('.mp3') or content.endswith('.MP3') or content.endswith('.WAV') or content.endswith('.wav') 	:
				zipObj.extract(content,f'{power}-voice')


segments_path= os.path.join(f'{power}-voice','ppt','media')
segments = []
sort(segments_path)
for segment in os.listdir(segments_path) :
	



#convert all voices to mp3 and merge them
