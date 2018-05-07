import csv
rows = {}

# Open csv, extract all rows to rows object
with open('Schedule.csv') as csvfile:
	reader = csv.DictReader(csvfile)	
	csvLength = 1
	for row in reader:
		rows[str(csvLength)] = dict(row)
		csvLength += 1
csvfile.close()



for x in range(1,csvLength):
	a = str(x)

	if (rows[a]["Day"] == "1"):
		event = "Day: {}. Class: {}. StartTime: {}. EndTime: {}".format(rows[a]["Day"],
			rows[a]["Class"],rows[a]["Start Time"],rows[a]["End Time"])
		print(event)
	#print(rows[str(x)]["Class"])

