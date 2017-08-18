# Developed by Winston Zhou for use by Moneythink CMU
# August 2017

# How to use:
# 0. Read documentation.
# 1. Instantiate a class: a = iCal()
# 2. Read and perform analysis on all schedules: a.read()
# 3. Write available times to CSV | XLSX: a.writeToCSV() | a.writeToXLSX()
# 4. Match mentors to sites: a.matchFromCSV()
# 5. Find what sites a mentor can go to: a.findSites("firstName lastName")

from math import floor, ceil
import csv
from os.path import dirname, abspath, split
from os import sep, listdir
from time import sleep
import xlsxwriter # Comment out this line if xlsxwriter is not installed

iCalPath = dirname(abspath(__file__)) + sep + "iCals" + sep
rootPath = dirname(abspath(__file__)) + sep 

class iCal(object):
    def __init__(self):
        self.totalIntervals = 22
        # Twenty-two (22) is the number of half-hour intervals between
        # 8:00 AM (08:00) and 7:00 PM (19:00)
        self.extensionReduce = len(".ics") # Remove extension to get the name

        self.year = "2017" # Must be correct for analysis of non-SIO calendars
        self.startDatesReg = ["20170821", # 1-week back
                              "20170822",
                              "20170823",
                              "20170824",
                              "20170825",
                              "20170826",
                              "20170827",
                              "20170828", # First day of classes
                              "20170829",
                              "20170831",
                              "20170901",
                              "20170902",
                              "20170903",
                              "20170904",
                              "20170905",
                              "20170906"] # 1-week forward
        self.startDatesMini = ["20171016", # 1-week back
                              "20171017",
                              "20171018",
                              "20171019",
                              "20171020",
                              "20171021",
                              "20171022",
                              "20171023", # First day of Mini-2/Mini-4 ONLY!!
                              "20170824",
                              "20171025",
                              "20171026",
                              "20171027",
                              "20171028",
                              "20171029",
                              "20171030",
                              "20171031"] # 1-week forward
        self.endDatesReg = ["20171201", # 1-week back
                             "20171202",
                             "20171203",
                             "20171204",
                             "20171205",
                             "20171206",
                             "20171207",
                             "20171208", # Most likely end date for Mini-1/Mini-3
                             "20171209",
                             "20171211",
                             "20171212",
                             "20171213",
                             "20171214",
                             "20171215",
                             "20171216",
                             "20171217"]
        self.endDatesMini = ["20171009", # 1-week back
                             "20171010",
                             "20171011",
                             "20171012",
                             "20171013",
                             "20171014",
                             "20171015",
                             "20171016", # Most likely end date
                             "20171017",
                             "20171018",
                             "20171019",
                             "20171020",
                             "20171021",
                             "20171022",
                             "20171023",
                             "20171024"] # 1-week forward
        self.alarmDetected = False # True if BEGIN:VALARM is found
                                   # While TRUE, ignore SUMMARY until END:VALARM
        self.deliver = False # For a NON-SIO schedule, if this variable is TRUE,
                             # then the start and end times are stored
                             # on the next time END:VEVENT is reached

        self.header = None
        self.masterSchedule = []

        # Common ways people write weekdays
        self.mondaySet = set(["Monday", "MONDAY", "MON", "mon", "MO", "mo",
                              "monday", "Mo", "Mon", "Mondays", "mondays",
                              "MONDAYS"])
        self.tuesdaySet = set(["Tuesday", "TUESDAY", "TUE", "tue", "TU", "tu",
                               "tuesday", "tues", "TUES", "Tues", "Tue",
                               "Tuesdays", "tuesdays", "TUESDAYS"])
        self.wednesdaySet = set(["Wednesday", "WEDNESDAY", "WED", "wed", "we",
                                 "WE", "wednesday", "We", "Wednesdays",
                                 "WEDNESDAYS", "wednesdays"])
        self.thursdaySet = set(["Thursday", "THURSDAY", "thursday", "THU", "thu",
                                "TH", "th", "Thurs", "THURS", "thurs", "Thu",
                                "Th", "thursdays", "THURSDAYS", "Thursdays"])
        self.fridaySet = set(["Friday", "FRIDAY", "FRI", "Fri", "Fr", "FR", "fri",
                              "friday", "fr", "fridays", "Fridays", "FRIDAYS",
                              "friyay"])
        self.analysisSet = set(["Analysis", "ANALYSIS", "analysis"])

        self.matchDictionary = {}
        self.validMentorName = set() # Used to see if a valid mentor name was 
                                     # entered in findSites() function

    def __hash__(self):
        return self # only integers are hashed in this program

    def digitcount(n): #counts number of digits in number n
        if n < 0: n *= -1
        elif n == 0 : return 1
        count = 0
        while n > 0:
            n //= 10
            count += 1
        return count

    def secondstime(binarytime): #converts to seconds
        if (iCal.digitcount(binarytime) < 3): #under 60 seconds
            return binarytime
        elif (iCal.digitcount(binarytime) < 5): #under 60 minutes
            secondsdigits = binarytime % 100 #seconds count
            minutesdigits = binarytime // 100 #minutes count
            return minutesdigits * 60 + secondsdigits
        else: #over 60 minutes
            secondsdigits = binarytime % 100
            hourdigit = binarytime // 10000 #hours count
            minutesdigits = (binarytime - hourdigit * 10000) // 100
            return hourdigit * 3600 + minutesdigits * 60 + secondsdigits

    def findInterval(self, start, end):
        # Inputs: start and end times in seconds
        # Outputs: interval corresponding to the start and end times
        startInterval = floor(start / 1800) - 16 #1800 is the number of seconds in 30'
        endInterval = ceil(end / 1800) - 17 # subtract 1 more than above
        
        # Case with an event that has a duration of 0 minutes
        if endInterval < startInterval:
            endInterval = startInterval

        # Case with a class that ends before the monitored timespan
        if endInterval < 0:
            return None
        
        # Case with a class that starts before the monitored timespan
        # but ends at a time within the monitored timespan
        elif startInterval < 0:
            startInterval = 0
        
        # Case with a class that starts after the monitored timespan
        if startInterval > (self.totalIntervals - 1):
            return None

        # Case with a class that starts within the monitored timespan
        # but ends at a time after the monitored timespan
        elif endInterval > (self.totalIntervals - 1):
            endInterval = self.totalIntervals - 1

        return list(range(startInterval, endInterval + 1))

    def read(self):
        for file in listdir(iCalPath):

            # Phase 1: Open the iCal file and retrieve desired data
            currentPath = iCalPath + file
            startTimes, endTimes, weekdate = [], [], []
            currentName = split(currentPath)[1][:-self.extensionReduce]
            sleep(0.1)
            print(currentName)
            iCalType = "SIO"
            zuluTime = False # If TRUE, raise alert!

            with open(currentPath, newline='') as csvfile:
                iCalRowReader = csv.reader(csvfile)
                weeklyFound = False # for SIO schedules to check anomalies 
                tempStart, tempEnd = None, None # Used for non SIO-generated calendars 
                tempEventName = None
                tempCountStartDate = None # Used for non SIO-generated calendars 
                                          # that use COUNT for recurrence
                countInRange = False # True if an event starts in the desired range
                for rowIndex, row in enumerate(iCalRowReader):
                    # print(row) #uncommenting this line shows the entire iCal
                    # Check if iCal file is SIO-generated
                    if rowIndex == 1: 
                        if row[0][:15] != "PRODID:-CMU SIO": # not SIO-generated
                            print("-----------------------------------------")
                            print("%s's schedule NON SIO-GENERATED!" % currentName) 
                            print("-----------------------------------------")
                            iCalType = "NON-SIO"
                    if iCalType == "SIO":
                        if row[0] == "BEGIN:VEVENT":
                            weeklyFound = False
                        elif row[0][:6] == "DTSTAR": # found start time
                            startTimes.append(row[0][16:])
                        elif row[0][:6] == "DTEND:": # found end time
                            endTimes.append(row[0][14:])
                        elif row[0][:5] == "RRULE":
                            weeklyFound = True
                            days = [row[0][len(row[0]) - 2:]] #first date
                            daysInRow = len(row)
                            if daysInRow > 1:
                                additionalDays = row[1:] 
                                days += (additionalDays)
                            weekdate.append(days)
                        elif row[0] == "END:VEVENT":
                            if weeklyFound == False: # Found non-weekly event?!
                                print("-----------------------------------------")
                                print("WARNING: Non-recurring event detected!")
                                print("-----------------------------------------")
                    else: # NON-SIO attempt to find relevant events
                        if row[0][:7] == "PRODID:":
                            print("Exported from: %s" % row[0][7:])
                        elif row[0] == "BEGIN:VEVENT":
                            tempStart, tempEnd, self.deliver = None, None, False
                            tempEventName = None
                            countInRange = False
                            self.alarmDetected = False
                        elif row[0][:7] == "SUMMARY":
                            if self.alarmDetected == False:
                                tempEventName = row[0][8:]
                        elif row[0] == "BEGIN:VALARM":
                            self.alarmDetected = True
                        elif row[0] == "END:VALARM":
                            self.alarmDetected = False
                        elif row[0][:6] == "DTSTAR" or row[0][:5] == "DTEND": 
                            # print(row[0])
                            startIndex = row[0].find(self.year)
                            tempCountStartDate = row[0][startIndex:startIndex + 8]
                            # print(tempCountStartDate)
                            if (tempCountStartDate in self.startDatesReg or
                                tempCountStartDate in self.startDatesMini):
                                countInRange = True
                            if startIndex != -1: 
                                # print(row[0][startIndex:startIndex + 15], "Length: ", len(row[0][startIndex:startIndex + 15]))
                                # Check to see if the event lasts the whole day
                                eventCharacterLength = len(row[0][startIndex:startIndex + 15])
                                if eventCharacterLength == 15:
                                    # print(row[0][startIndex + 9:startIndex + 15])
                                    if row[0][:6] == "DTSTAR":
                                        tempStart = row[0][startIndex + 9:startIndex + 15]
                                        if zuluTime == False:
                                            if (row[0][len(row[0]) - 1]) == "Z":
                                                zuluTime = True
                                                print("WARNING: Zulu time detected!")
                                                print("Add 5 hours when importing into calendar")
                                        # print("start:", tempStart)
                                    elif row[0][:5] == "DTEND":
                                        tempEnd = row[0][startIndex + 9:startIndex + 15]
                                        # print("end:", tempEnd)
                        elif row[0][:5] == "RRULE":
                            freqIndex = row[0].find("FREQ")
                            frequency = row[0][freqIndex + 5:freqIndex + 11]
                            # print(frequency)
                            if frequency == "WEEKLY":
                                untilIndex = row[0].find("UNTIL")
                                if untilIndex != -1: # Make sure that UNTIL is even there
                                    # print("Until:", row[0][untilIndex + 6:untilIndex + 14])
                                    until = row[0][untilIndex + 6:untilIndex + 14]
                                    if (until in self.endDatesReg or
                                        until in self.endDatesMini):
                                        self.deliver = True
                                        days = [row[0][len(row[0]) - 2:]] #first date
                                        daysInRow = len(row)
                                        if daysInRow > 1:
                                            additionalDays = row[1:] 
                                            days += (additionalDays)
                                        weekdate.append(days)
                                else:
                                    countIndex = row[0].find("COUNT")
                                    if countIndex == -1: 
                                        print("Found weekly event with infinite recurrence.")
                                        self.deliver = True
                                        days = [row[0][len(row[0]) - 2:]] #first date
                                        daysInRow = len(row)
                                        if daysInRow > 1:
                                            additionalDays = row[1:] 
                                            days += (additionalDays)
                                        weekdate.append(days)
                                    else:
                                        intervalIndex = row[0].find("INTERVAL")
                                        if intervalIndex == -1: # not biweekly since courses are never biweekly
                                            bydayIndex = row[0].find("BYDAY")
                                            # print("Count:", row[0][countIndex + 6:bydayIndex - 1])
                                            count = int(row[0][countIndex + 6:bydayIndex - 1])
                                            if count > 7: # significant recurrence
                                                if countInRange:
                                                    self.deliver = True
                                                    days = [row[0][len(row[0]) - 2:]] #first date
                                                    daysInRow = len(row)
                                                    if daysInRow > 1:
                                                        additionalDays = row[1:] 
                                                        days += (additionalDays)
                                                    weekdate.append(days)
                        elif row[0] == "END:VEVENT":
                            if self.deliver:
                                tempStart = "T" + tempStart
                                tempEnd = "T" + tempEnd
                                startTimes.append(tempStart)
                                endTimes.append(tempEnd)
                                sleep(0.1)
                                print("Found possible event:", tempEventName)

            # print("startTimes:", startTimes)
            # print("endTimes:", endTimes)
            # print("weekdates", weekdate)
            
            # Phase 2: Remove 'T' and leading zeros
            startTimesFormatted, endTimesFormatted = [], []
            for startTime in startTimes: #loop invariants omitted
                if startTime[0] != "T": print("READ ERROR! Delimiter not 'T'")
                elif startTime[1] == "0": 
                    startTimesFormatted.append(int(startTime[2:]))
                elif startTime[1] != "0":
                    startTimesFormatted.append(int(startTime[1:]))
            for endTime in endTimes:
                if endTime[0] != "T": print("READ ERROR! Delimiter not 'T'")
                elif endTime[1] == "0": 
                    endTimesFormatted.append(int(endTime[2:]))
                elif endTime[1] != "0":
                    endTimesFormatted.append(int(endTime[1:]))
            # print("startTimesFormatted:", startTimesFormatted)   
            # print("endTimesFormatted:", endTimesFormatted)

            # Phase 3: Convert formatted times into seconds after midnight
            startTimeSeconds, endTimesSeconds = [], []
            for startTime in startTimesFormatted:
                startTimeSeconds.append(iCal.secondstime(startTime))
            for endTime in endTimesFormatted:
                endTimesSeconds.append(iCal.secondstime(endTime))
            # print("startSeconds:", startTimeSeconds)   
            # print("endSeconds:", endTimesSeconds)

            # Phase 4: Build the busy intervals
            # Note that the use of the set data structure, of course, assumes 
            # that there are no conflicting intervals; conflicting intervals
            # are discarded.
            mondayIntervalsBusy = set()
            tuesdayIntervalsBusy = set()
            wednesdayIntervalsBusy = set()
            thursdayIntervalsBusy = set()
            fridayIntervalsBusy = set()
            for eventIndex in range(len(weekdate)):
                interval = iCal.findInterval(self,
                                             startTimeSeconds[eventIndex],
                                             endTimesSeconds[eventIndex])
                # print(interval)
                if interval != None:
                    if "MO" in weekdate[eventIndex]:
                        for i in interval:
                            mondayIntervalsBusy.add(i)
                    if "TU" in weekdate[eventIndex]: 
                        for i in interval:
                            tuesdayIntervalsBusy.add(i)
                    if "WE" in weekdate[eventIndex]:
                        for i in interval:
                            wednesdayIntervalsBusy.add(i)
                    if "TH" in weekdate[eventIndex]:
                        for i in interval:
                            thursdayIntervalsBusy.add(i)
                    if "FR" in weekdate[eventIndex]:
                        for i in interval:
                            fridayIntervalsBusy.add(i)
            # print("Monday:", mondayIntervalsBusy)
            # print("Tuesday:", tuesdayIntervalsBusy)
            # print("Wednesday:", wednesdayIntervalsBusy)
            # print("Thursday:", thursdayIntervalsBusy)
            # print("Friday:", fridayIntervalsBusy)
            
            # Phase 4A: Build the free intervals (DEBUGGING PURPOSES)
            # original = set(list(range(self.totalIntervals)))
            # mondayIntervalsFree = original - mondayIntervalsBusy
            # tuesdayIntervalsFree = original - tuesdayIntervalsBusy
            # wednesdayIntervalsFree = original - wednesdayIntervalsBusy
            # thursdayIntervalsFree = original - thursdayIntervalsBusy
            # fridayIntervalsFree = original - fridayIntervalsBusy
            # print("And the free intervals:")
            # print("Monday:", mondayIntervalsFree)
            # print("Tuesday:", tuesdayIntervalsFree)
            # print("Wednesday:", wednesdayIntervalsFree)
            # print("Thursday:", thursdayIntervalsFree)
            # print("Friday:", fridayIntervalsFree)

            # Phase 5: Build the list which gets outputted 
            monday = [currentName] * self.totalIntervals
            tuesday = [currentName] * self.totalIntervals
            wednesday = [currentName] * self.totalIntervals
            thursday = [currentName] * self.totalIntervals
            friday = [currentName] * self.totalIntervals
            
            # Empty space means busy interval
            for i in mondayIntervalsBusy:
                monday[i] = ""
            for i in tuesdayIntervalsBusy:
                tuesday[i] = ""
            for i in wednesdayIntervalsBusy:
                wednesday[i] = ""
            for i in thursdayIntervalsBusy:
                thursday[i] = ""
            for i in fridayIntervalsBusy:
                friday[i] = ""
            # print("Monday:", monday)
            # print("Tuesday:", tuesday)
            # print("Wednesday:", wednesday)
            # print("Thursday:", thursday)
            # print("Friday:", friday)

            # Phase 7: Append student's intervals to master list
            self.masterSchedule.append([monday, tuesday,
                                        wednesday, thursday,
                                        friday])

    def writeToCSV(self): # Run this function if xlsxwriter is not installed
        print("Writing all schedules to CSV")
        currentPath = rootPath + "Template.csv"
        with open(currentPath, newline='') as csvfile:
            templateReader = csv.reader(csvfile)
            for row in templateReader:
                self.header = row
        # print(self.header)
        
        currentPath = rootPath + "Monday.csv"
        with open(currentPath, 'w', newline='') as csvfile:
            iCalRowWriter = csv.writer(csvfile)
            iCalRowWriter.writerow(self.header)
            for row in range(len(self.masterSchedule)):
                iCalRowWriter.writerow(self.masterSchedule[row][0])

        currentPath = rootPath + "Tuesday.csv"
        with open(currentPath, 'w', newline='') as csvfile:
            iCalRowWriter = csv.writer(csvfile)
            iCalRowWriter.writerow(self.header)
            for row in range(len(self.masterSchedule)):
                # print(self.masterSchedule[row][1])
                iCalRowWriter.writerow(self.masterSchedule[row][1])

        currentPath = rootPath + "Wednesday.csv"
        with open(currentPath, 'w', newline='') as csvfile:
            iCalRowWriter = csv.writer(csvfile)
            iCalRowWriter.writerow(self.header)
            for row in range(len(self.masterSchedule)):
                iCalRowWriter.writerow(self.masterSchedule[row][2])

        currentPath = rootPath + "Thursday.csv"
        with open(currentPath, 'w', newline='') as csvfile:
            iCalRowWriter = csv.writer(csvfile)
            iCalRowWriter.writerow(self.header)
            for row in range(len(self.masterSchedule)):
                iCalRowWriter.writerow(self.masterSchedule[row][3])

        currentPath = rootPath + "Friday.csv"
        with open(currentPath, 'w', newline='') as csvfile:
            iCalRowWriter = csv.writer(csvfile)
            iCalRowWriter.writerow(self.header)
            for row in range(len(self.masterSchedule)):
                iCalRowWriter.writerow(self.masterSchedule[row][4])

    def writeToXLSX(self): # xlsxwriter must be installed
        print("Writing all schedules to 'Moneythink Schedules Tabulation.xlsx'")
        workbook = xlsxwriter.Workbook('Moneythink Schedules Tabulation.xlsx')
        mondaySheet = workbook.add_worksheet("Monday")
        tuesdaySheet = workbook.add_worksheet("Tuesday")
        wednesdaySheet = workbook.add_worksheet("Wednesday")
        thursdaySheet = workbook.add_worksheet("Thursday")
        fridaySheet = workbook.add_worksheet("Friday")

        bold = workbook.add_format({'bold': True})

        mondaySheet.freeze_panes(1, 0)
        tuesdaySheet.freeze_panes(1, 0)
        wednesdaySheet.freeze_panes(1, 0)
        thursdaySheet.freeze_panes(1, 0)
        fridaySheet.freeze_panes(1, 0)

        # Retrieve half-hour intervals format from the template
        currentPath = rootPath + "Template.csv"
        with open(currentPath, newline='') as csvfile:
            templateReader = csv.reader(csvfile)
            for row in templateReader:
                self.header = row
        for row in range(len(self.masterSchedule) + 1):
            if row == 0: # Write in the half-hour intervals
                for col in range(self.totalIntervals):
                    mondaySheet.write(0, col, self.header[col], bold)
                    tuesdaySheet.write(0, col, self.header[col], bold)
                    wednesdaySheet.write(0, col, self.header[col], bold)
                    thursdaySheet.write(0, col, self.header[col], bold)
                    fridaySheet.write(0, col, self.header[col], bold)
            elif row > 0: # Write in the names
                for col in range(self.totalIntervals):
                    mondaySheet.write(row, col, self.masterSchedule[row - 1][0][col])
                    tuesdaySheet.write(row, col, self.masterSchedule[row - 1][1][col])
                    wednesdaySheet.write(row, col, self.masterSchedule[row - 1][2][col])
                    thursdaySheet.write(row, col, self.masterSchedule[row - 1][3][col])
                    fridaySheet.write(row, col, self.masterSchedule[row - 1][4][col])

    def matchFromCSV(self):
        currentPath = rootPath + "Site Times Input.csv"
        analysisDetected = False

        with open(currentPath, newline='') as csvfile:
            siteReader = csv.reader(csvfile)
            for rowIndex, row in enumerate(siteReader):
                if rowIndex == 1:
                    if len(self.analysisSet & set(row)) > 0:
                        analysisDetected = True
                        analysisIndices = set()
                        for weekdate in range(len(row)):
                            if row[weekdate] in self.analysisSet:
                                analysisIndices.add(weekdate)
            # print(analysisIndices)

        # These 2D-lists should have as many sublists as Site Times Input.csv
        # has rows
        originalSiteOrder = []
        nonAnalysisSites = [[], [], [], [], []]
        analysisSites = [[], [], [], [], []]
        workWeekDays = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"]
        if analysisDetected: 
            # Then, reformat the CSV to put the analysis sites on the rightmost
            # column
            with open(currentPath, newline='') as csvfile:
                siteReader = csv.reader(csvfile)
                for rowIndex, row in enumerate(siteReader):
                    originalSiteOrder.append(row)
                    for siteIndex, site in enumerate(row):
                        if siteIndex in analysisIndices:
                            # Placeholder column
                            if rowIndex == 0: analysisSites[0].append("")
                            elif rowIndex == 1: analysisSites[1].append("     ")
                            elif rowIndex == 2: analysisSites[2].append("100")
                            elif rowIndex == 3: analysisSites[3].append("101")
                            elif rowIndex == 4: analysisSites[4].append("9999")
                            for i in range(5): # MO, TU, WE, TH, FR, is 5 days
                                if rowIndex == 1: # row 1 category is weekdays
                                    analysisSites[rowIndex].append(workWeekDays[i])
                                else:
                                    analysisSites[rowIndex].append(site)
                        else:
                            nonAnalysisSites[rowIndex].append(site)
            # print("nonAnalysisSites:", nonAnalysisSites)
            # print("analysisSites:", analysisSites)
            newCSVRows = []
            for category in range(len(analysisSites)):
                newCSVRows.append(nonAnalysisSites[category] + analysisSites[category])
            # print("newCSVRows:", newCSVRows)
            with open(currentPath, 'w', newline='') as csvfile:
                newInputWriter = csv.writer(csvfile)
                for row in range(len(analysisSites)):
                    newInputWriter.writerow(newCSVRows[row])

        with open(currentPath, newline='') as csvfile:
            siteReader = csv.reader(csvfile)
            for rowIndex, row in enumerate(siteReader):
                # Determine number of sites
                if rowIndex == 0:
                    siteCount = 0
                    for column in row:
                        if column not in ["Name:", ""]:
                            siteCount += 1
                    siteNames = row[1:]
                    print(str(siteCount) + " sites retrieved.")
                    # print("Site Names:", siteNames)
                # Determine the weekday of each site
                elif rowIndex == 1:
                    siteWeekdays = row[1:]
                    # print("Site Weekdays:", siteWeekdays)
                # Build the original and seconds after midnight start times of each site
                elif rowIndex == 2:
                    siteOriginalStart = row[1:]
                    siteStarts = []
                    for originalStart in siteOriginalStart:
                        currentStartTime = originalStart + "00"
                        siteStarts.append(iCal.secondstime(int(currentStartTime)))
                    # print("Original start times:", siteOriginalStart)
                    # print("Seconds start times:", siteStarts)
                # " end times of each site
                elif rowIndex == 3:
                    siteOriginalEnd = row[1:]
                    siteEnds = []
                    for originalEnd in siteOriginalEnd:
                        currentEndTime = originalEnd + "00"
                        siteEnds.append(iCal.secondstime(int(currentEndTime)))
                    # print("Original end times:", siteOriginalEnd)
                    # print("Seconds end times:", siteEnds)
                # Determine the time it takes to commute to each site
                elif rowIndex == 4:
                    siteTolerances = row[1:]
                    # print("Commuting minutes:", siteTolerances)
                    for minuteIndex, minute in enumerate(siteTolerances):
                        siteTolerances[minuteIndex] = int(siteTolerances[minuteIndex]) * 60
                    # print("Commuting seconds:", siteTolerances)
        adjustedStarts, adjustedEnds = [], []
        for site in range(len(siteNames)):
            adjustedStarts.append(siteStarts[site] - siteTolerances[site])
            adjustedEnds.append(siteEnds[site] + siteTolerances[site])
        # print("Adjusted start times (seconds after midnight):", adjustedStarts)
        # print("Adjusted end times (seconds after midnight):", adjustedEnds)
        siteIntervals = []
        detailedSiteName = []
        for site in range(len(siteNames)):
            siteIntervals.append(iCal.findInterval(self,
                                             adjustedStarts[site],
                                             adjustedEnds[site]))
            detailedSiteName.append("%s %s" % (siteNames[site], siteWeekdays[site]))
        # print("Site Intervals:", siteIntervals)
        # print("Detailed Site Name:", detailedSiteName)

        masterMatches = []
        for site in range(len(siteNames)):
            currentMatches = []
            adjustedStartIndex = siteIntervals[site][0]
            adjustedEndIndex = siteIntervals[site][-1]
            sleep(0.1)
            if detailedSiteName[site] != "      ":
                print("Matching %s" % detailedSiteName[site])
            else:
                print("Performing contingency analysis...")
            if siteWeekdays[site] in self.mondaySet:
                currentPath = rootPath + "Monday.csv"
            elif siteWeekdays[site] in self.tuesdaySet:
                currentPath = rootPath + "Tuesday.csv"
            elif siteWeekdays[site] in self.wednesdaySet:
                currentPath = rootPath + "Wednesday.csv"
            elif siteWeekdays[site] in self.thursdaySet:
                currentPath = rootPath + "Thursday.csv"
            elif siteWeekdays[site] in self.fridaySet:
                currentPath = rootPath + "Friday.csv"
            elif siteWeekdays[site] == "     ": # Blank column placeholder
                currentPath = rootPath + "Friday.csv"
            with open(currentPath, newline='') as csvfile:
                timeReader = csv.reader(csvfile)
                for rowIndex, row in enumerate(timeReader):
                    if site == 1 and rowIndex > 0: # append mentor names to set
                        self.validMentorName |= (set(row) - set(['']))
                    if rowIndex > 0:
                        for name in set(row) - set(['']):
                            currentName = name
                        intervalSet = set(row[adjustedStartIndex:adjustedEndIndex + 1])
                        # print(currentName)
                        # print(intervalSet)
                        # If the set's length is 2, this means that a schedule
                        # is partly free for that interval and partly busy
                        # If the set's length is 1, this either means that
                        # the schedule is entirely free for that interval
                        # or is entirely busy for that interval
                        if len(intervalSet) == 1:
                            if intervalSet != set(['']):
                                currentMatches.append(currentName)
            # print("currentMatches:", currentMatches)
            # Append site matches to matches master list
            masterMatches.append(currentMatches)

        # print("MasterMatches", masterMatches)
        # Create the match dictionary for use when determining
        # what sites an individual mentor can be allotted to
        for siteIndex, site in enumerate(detailedSiteName):
            self.matchDictionary[site] = set(masterMatches[siteIndex])
        # print("Match Dictionary:", self.matchDictionary)

        # Transpose the matches master list for inputting into CSV row-by-row
        # Determine size of square
        maxLen = 0
        for site in masterMatches:
            if len(site) > maxLen:
                maxLen = len(site)

        # Square the list (make the list N by N)
        for i in range(len(masterMatches)):
            currentLen = len(masterMatches[i])
            for j in range(maxLen - currentLen):
                masterMatches[i].append("")
        # print("Squared masterMatches:", masterMatches)

        # Perform the transposition
        transpose = []
        for row in range(maxLen):
            transpose.append([])
            for col in range(len(masterMatches)):
                transpose[row].append(masterMatches[col][row])
        # print("Transposed masterMatches:", transpose)

        currentPath = rootPath + "Matches.csv"
        with open(currentPath, 'w', newline='') as csvfile:
            matchesWriter = csv.writer(csvfile)
            matchesWriter.writerow(detailedSiteName)
            for row in range(len(transpose)):
                matchesWriter.writerow(transpose[row])

        if analysisDetected: 
            currentPath = rootPath + "Site Times Input.csv"
            with open(currentPath, 'w', newline='') as csvfile:
                originalInputWriter = csv.writer(csvfile)
                for row in range(len(originalSiteOrder)):
                    originalInputWriter.writerow(originalSiteOrder[row])

    def findSites(self, mentor):
        if mentor not in self.validMentorName:
            print("Invalid mentor name entered:", mentor)
        else:
            match = set()
            for key in self.matchDictionary:
                if mentor in self.matchDictionary[key]:
                    match.add(key)
            for site in match:
                sleep(.1)
                print(site)
            if len(match) == 0: 
                print(mentor, "cannot be matched with any sites.")