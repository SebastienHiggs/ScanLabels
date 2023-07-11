from __future__ import print_function

from win32com.client import Dispatch
import pathlib

import cv2
import csv
import os

import time
import datetime
import threading

def initialiseCOM():
    print("Initialising COM")
    printer_COM = Dispatch('Dymo.DymoAddIn')
    return printer_COM

def initialisePrinter(printer_COM):
    print("Initialising Printer")
    filePath = pathlib.Path("C:/Users/MCIC Makerspace/Desktop/Labels/Code/blank.label")
    printer = printer_COM.getDymoPrinters()

    printer_COM.selectPrinter(printer)
    printer_COM.Open2(filePath)

    printer_label = Dispatch('Dymo.DymoLabels')
    return printer_label

def readFiles():
    folder_path = "CSV_report"
    print(f"Reading which files are in the {folder_path} folder")

    # Get a list of all the files in the folder
    files = os.listdir(folder_path)
    print(f"Found the files: {files}")
    lst = []

    # Print the list of files
    for file in files:
        if file.startswith("report"):
            lst.append(file)
    
    if len(lst) == 1:
        return lst[0]
    elif len(lst) > 1:
        print("Too many files in the folder")
        print(lst)
        print(f"Please only have the downloaded csv from eventbrite in the {folder_path} folder")
    else:
        print("Can't find any files in the folder, please put the csv from eventbrite")
    print("Program is closing")
    time.sleep(10)
    exit()
    

def readCSV(file):
    print(f"Reading CSV file {file}")
    data = []
    with open('CSV_report/' + file) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        for row in csv_reader:
                # Iterate over each value in the row
            row_values = []
            for value in row:
                # Split the value into separate values at each comma
                value_list = value.split(',')
                # Extend the row values list with the new values
                row_values.extend(value_list)
            # Add the row values as a list to the data list
            data.append(row_values)
    orderNo = []
    firstnames = []
    surnames = []
    titles = []
    for i in range(len(data)-1):
        orderNo.append(data[i+1][0])
        firstnames.append(data[i+1][2])
        surnames.append(data[i+1][3])
        titles.append(data[i+1][7])
    return [orderNo,firstnames,surnames,titles]

def initCam():
    print("initialising video capture object")
    # define a video capture object
    vid = cv2.VideoCapture(0)
    detector = cv2.QRCodeDetector()
    return vid, detector

def webcamToText(vid, detector):
    qr_number = ""
    # Capture the video frame by frame
    ret, frame = vid.read()
    data, bbox, straight_qrcode = detector.detectAndDecode(frame)
    if len(data) > 0:
        qr_number = data

    # Display the resulting frame
    cv2.imshow('frame', frame)
    # the 'q' button is set as the
    # quitting button you may use any
    # desired button of your choice
    if cv2.waitKey(1) & 0xFF == ord('q'):
        pass
    return qr_number

def killCam(vid):
    # After the loop release the cap object
    vid.release()
    # Destroy all the windows
    cv2.destroyAllWindows()

def searchList(name_matrix,qr_number):
    for row in name_matrix:
        if row[0] == qr_number:
            return row[2],row[3]

def printName(printer_COM,printer_label,firstname = "",surname = "",title = ""):
    current_time = datetime.datetime.now().strftime('%H:%M:%S')
    writeCSV(firstname,surname, title)
    print(f"Printing {firstname=} {surname=} {title=}")
    printer_label.SetField('TEXT', firstname)
    printer_label.SetField('TEXT_1', surname)
    printer_label.SetField('TEXT_2', title)

    printer_COM.StartPrintJob()
    printer_COM.Print(1,False)
    printer_COM.EndPrintJob()

def writeCSV(firstname,surname, title):
    print(f"Writing {firstname=} {surname=} {title=} to CSV")
    # Get today's date
    today = datetime.date.today()
    # Create the filename with today's date
    filename = f"CSV_report/{today.strftime('%Y-%m-%d')}.csv"
    # Open the file in append mode
    with open(filename, 'a', newline='') as csvfile:
        # Create a CSV writer
        csvwriter = csv.writer(csvfile)
        # Get the current time
        current_time = datetime.datetime.now().strftime('%H:%M:%S')
        # Write the firstname, lastname, and current time to the CSV file
        csvwriter.writerow([firstname, surname, current_time])

def input_thread(printer_COM,printer_label,name_matrix):
    while True:
        print()
        print("Please enter a name if they don't have a ticket:")
        name = input()
        first = name_matrix[1]
        sur = name_matrix[2]
        title = name_matrix[3]
        lower_firstname = [s.lower() for s in first]
        lower_surname = [s.lower() for s in sur]
        if name.lower() in lower_firstname:
            printName(printer_COM,printer_label,first[lower_firstname.index(name.lower())],sur[lower_firstname.index(name.lower())])
        elif name.lower() in lower_surname:
            printName(printer_COM,printer_label,first[lower_surname.index(name.lower())],sur[lower_surname.index(name.lower())])
        else:
            print("Name not found. Enter a blank line to return or enter the name you would like to print regardless")
            inp = input()
            if inp == '':
                pass
            else:
                if len(inp.split()) >= 2:
                    printName(printer_COM,printer_label,inp.split()[0],inp.split()[1])
                else:
                    printName(printer_COM,printer_label,inp.split()[0])

if __name__ == "__main__":
    printer_COM = initialiseCOM()       
    printer_label = initialisePrinter(printer_COM)
    file = readFiles()
    name_matrix = readCSV(file)
    vid, detector = initCam()
    print("Ready to print!")
    input_thread = threading.Thread(target=input_thread, args=(printer_COM,printer_label,name_matrix))
    input_thread.start()
    while True:
        qr_number = webcamToText(vid, detector)
        if qr_number in name_matrix[0]:
            rowNo = name_matrix[0].index(qr_number)
            firstname = name_matrix[1][rowNo]
            surname = name_matrix[2][rowNo]
            title = name_matrix[3][rowNo]
            printName(printer_COM,printer_label,firstname,surname, title)
            time.sleep(1)
        elif qr_number != '':
            print(f"Value {qr_number} not found in data!")
            time.sleep(1)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
        time.sleep(0.05)
    killCam(vid)
        

