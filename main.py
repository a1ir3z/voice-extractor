# import things here
import os



# Make folders
if not 'voice' in os.listdir():
    os.makedirs('voice')
if not 'power' in os.listdir():
    os.makedirs('power')

# get a PPTX and copy it into a folder
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



# convert copied pptx into a zip file





# extract zip file and take the voices out





#convert all voices to mp3 and merge them