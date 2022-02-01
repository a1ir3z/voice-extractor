import os
from zipfile import ZipFile
from pydub import AudioSegment
import shutil

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


#change directory to voice directory
source_dir = os.path.join(f'{power}-voice','ppt','media')
target_dir = 'voice'
    
file_names = os.listdir(source_dir)
    
for file_name in file_names:
    shutil.move(os.path.join(source_dir, file_name), target_dir)

os.removedirs(source_dir)


#combine 2 voices for test 
# i have to write a loop to combine all and export them
voices = os.listdir('voice')
voice1=  AudioSegment.from_wav('voice\media1.wav')
voice2=  AudioSegment.from_wav('voice\media2.wav')
combined = voice1 + voice2
combined.export ('joined.wav', format="wav")

#convert the file and rename it to the powerpoint name
