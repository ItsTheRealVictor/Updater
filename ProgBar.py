import datetime
from dateutil import parser
from dateutil import relativedelta

startDateList = []
startTimeList = []
endDateList = []

while True:
    startDate = input("Enter your start date (YYYY-MM-DD) or enter 't' to use today: ")
    if startDate == 't':
        startDate = str(datetime.date.today())
    elif startDate == 'q':
        break
    startDateList.append(startDate)
    startTime = input("Enter the start time: ")
    if startTime == 'q':
        break
    startTimeList.append(startTime)
    userTime = input("Enter your time interval (in hours): ")
    if userTime == 'q':
        break

    """Takes the date and time interval from the user and computes the future date and day. """
    userinputTime = relativedelta.relativedelta(hours=int(userTime)) 
    parsedUserDateInput = parser.parse(startDate).date()  
    future = parsedUserDateInput + userinputTime #calculates the future date

    parsedFutureDate = datetime.datetime.strftime(future, "%Y-%m-%d")#converts the future date into datetime format
    endDateList.append(parsedFutureDate)
    dateTuples = list(zip(startDateList, startTimeList, endDateList))

    parsedUserDateInputDay = datetime.datetime.strftime(parsedUserDateInput, "%A")
    parsedFutureDay = datetime.datetime.strftime(future, "%A") #takes the future date from parsedFutureDate and computes the day 
                                                                    #of the week
    # GetDataFromUser.dataList.insert(5, f"{parsedFutureDate} ({parsedFutureDay})") #adds the future date and day of the week to the dataList
    # print(f"{userTime} hours after {parsedUserDateInputDay} {startDate} is {parsedFutureDay} {parsedFutureDate}")

print(dateTuples)
for start, time, end in dateTuples:
    print(f"""
        starting date and time is {start} {time}
        end date is {end}
    """)
