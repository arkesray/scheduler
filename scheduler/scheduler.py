from tkinter import * 
import xlrd
import datetime
    

def getData(fileLocation):
    
    workbook = xlrd.open_workbook(fileLocation)
    sheet = workbook.sheet_by_index(0)
    data = [[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    return data

def findEvent(time):

    for eventNumber in eventTimings:
        startTime = datetime.datetime.strptime(eventTimings[eventNumber][0], "%H:%M").time()
        endTime = datetime.datetime.strptime(eventTimings[eventNumber][1], "%H:%M").time()
        if  startTime < time and endTime > time:
            return eventNumber
    return 0

def update_labelTime():
    global labelTime
    labelTime.config(text = datetime.datetime.now().time().strftime("%I:%M %p"))
    labelTime.after(10, update_labelTime)

def update_labelDate():
    global labelDate
    labelTime.config(text = datetime.datetime.now().date())
    labelTime.after(1000, update_labelDate)

def update_labelEvent():
    global labelEvent
    global currentEvent
    global eventNames

    eventNumber = findEvent(datetime.datetime.now().time())
    currentEvent = eventNames[eventNumber]
    labelEvent.config(text = currentEvent)
    labelEvent.after(1000, update_labelEvent)


if __name__ == '__main__':

    # fetching the data from the Excel file
    file = "E:\\schedule.xlsx"
    
    # data received in 2D format
    schedule = getData(file)

    # modifying data before use
    eventNames = ["Free-Time"]
    eventTimings = {}

    # storing in python variables list and dict 
    for i in range(1, len(schedule)):
        eventNames.append(schedule[i][0])
        eventTimings[i] = (schedule[i][1], schedule[i][2])
    
    # creating main window
    root = Tk()
    root.title("Event Notifier")
    root.geometry("500x500")    #fixed size
    
    # top farme
    frame = Frame(root, bg = "black")

    labelDate = Label(frame, text = datetime.datetime.now().date(), bg = "black", fg = "white")
    labelDate.pack(side = LEFT)
    labelDate.after(1000, update_labelDate)

    currentEvent = "Free-Time"
    labelEvent = Label(frame, text = currentEvent, bg = "green", fg = "white", width = 25)
    labelEvent.pack(side = LEFT, fill = X)
    labelEvent.after(1000, update_labelEvent)

    labelTime = Label(frame, text = datetime.datetime.now().time().strftime("%I:%M %p"), bg = "black", fg = "white")
    labelTime.pack(side = LEFT)
    labelTime.after(10, update_labelTime)

    frame.pack()

    # bottom frame
    frame2 = Frame(root)

    for i in range(len(schedule)):
        tempText = "\t\t"
        for j in range(len(schedule[0])):
            tempText += schedule[i][j] + "    \t\t"
        if i == 0:
            label = Label(font = "none 10 bold underline", text = "\t" + schedule[0][0] + "\t" + schedule[0][1] + "\t" + schedule[0][2])
        else:
            label = Label(text = tempText)
        label.pack()

    frame2.pack()

    root.mainloop()