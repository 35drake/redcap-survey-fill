# Make a list of the settings from the settings.txt file
settings_file = open("files/settings.txt","r")
settings = []
for line in settings_file:
	if line[-1] == "\n": # the lines (but not necessarily the final line) will end in an unneeded newline
		line = line[0:-1]
	settings = settings + [line]

# Take only the data items and not the field names
settings[0] = settings[0][17:] # "Spreadsheet_path:" is a 17 character-long phrase
settings[1] = settings[1][9:] # This is for Username
settings[2] = settings[2][9:] # This is for Password

# Leave out the space at the start of the item if one exists
for count in range(0,len(settings)):
	if settings[count][0] == " ":
		settings[count] = settings[count][1:]


print(settings)

