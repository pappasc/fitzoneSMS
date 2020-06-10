from ringcentral import SDK
from openpyxl import load_workbook
import phonenumbers
import datetime
import math
import time
import json
import multiprocessing
import sys
import os
from dotenv import load_dotenv
load_dotenv()

def checkWorkbook():
    listName = input("Name of the file: ")
    while not os.path.isfile(listName):
        print("File not found.")
        if input("Try again? (Y/y or otherwise): ") in ["y", "Y", "yes", "Yes"]:
            listName = input("Name of the file: ")
        else:
            return

    try:
        workbookObj = load_workbook(listName)
    except Exception as error:
        print(str(error))
        raise error

    mainSheet = workbookObj.active
    count = 0
    for row in mainSheet.iter_rows(min_row=2, max_col=5, max_row=mainSheet.max_row, values_only=True):
        try:
            firstName = row[1].lower()
            capFirstName = firstName[0].upper() + firstName[1:]
            pNum = str(row[4])
            pNum = getNumberAuto(pNum)
            print(str(count) + ": " + capFirstName + " " + pNum)
            count += 1
        except Exception as error:
            break
    print(str(count))

def getNumber():
    while True:
        userNum = input("What is the phone number (start with area code)?: ")
        parseNum = phonenumbers.parse(str(userNum), "US")
        if phonenumbers.is_valid_number(parseNum):
            return phonenumbers.format_number(parseNum, phonenumbers.PhoneNumberFormat.E164)
        else:
            print("Invalid number. Try again")

def getNumberAuto(pfNum):
    index = 0
    for i in range(len(pfNum)):
        if pfNum[i] == ".":
            index = i
    pNum = pfNum[:index]

    try:
        parseNum = phonenumbers.parse(str(pNum), "US")
        if phonenumbers.is_valid_number(parseNum):
            return phonenumbers.format_number(parseNum, phonenumbers.PhoneNumberFormat.E164)
        else:
            raise Exception("Number not valid")
    except Exception as error:
        print("Error in getNumberAuto: " + str(error))
        raise error

def worker_task(ident, oPhone, fiName, checkList):
    time.sleep(10)
    print("\n****** Worker " + str(ident) + " is starting to recheck " + str(len(checkList)) + " messages ******\n")

    for teMess in checkList:
        print("Worker " + str(ident) + ": Check message " + str(teMess.messId) + " with number: " + teMess.pNum + " to: " + teMess.recipient)
        while True:
            try:
                time.sleep(oPhone.checkInterval)
                oPhone.checkMessageStatus(teMess)
                break
            except Exception as error:
                print("Worker " + str(ident) + ": Error checking message status: " + str(error))
                if str(error) == "Request rate exceeded":
                    oPhone.increaseCheckThrottle()
                    print("API throttle engaged while checking message status. Force sleep for 1 min for worker: " + str(ident))
                    time.sleep(15)
                    print("Worker " + str(ident) + " has 45s remaining")
                    time.sleep(15)
                    print("Worker " + str(ident) + " has 30s remaining")
                    time.sleep(15)
                    print("Worker " + str(ident) + " has 15s remaining")
                    time.sleep(15)
                    print("Worker " + str(ident) + " is resuming")
                    time.sleep(oPhone.checkInterval)
                else:
                    print("Worker " + str(ident) + ": An error occurred checking message status but not API throttle")
                    print("Error checking in sendMessage: " + str(error))
                    raise error
        if teMess.sentResult == 'Sent' or teMess.sentResult == 'Delivered' or teMess.sentResult == 'Received':
            print("Worker " + str(ident) + ": Message " + str(teMess.messId) + " to number " + teMess.pNum + " to: " + teMess.recipient + " was sent")
            continue
        elif teMess.sentResult == 'SendingFailed' or teMess.sentResult == 'DeliveryFailed':
            print("Worker " + str(ident) + ": Message " + str(teMess.messId) + " to number " + teMess.pNum + " to: " + teMess.recipient + " failed")
            if teMess.sendAttempt < 5:
                print("Worker " + str(ident) + ": Attempting resend to " + teMess.pNum + " to: " + teMess.recipient)
                while True:
                    try:
                        time.sleep(oPhone.sendInterval)
                        print("Worker " + str(ident) + ": Resending to number " + teMess.pNum + " to: " + teMess.recipient)
                        oPhone.sendSMS(teMess)
                        #recheckMessages.remove(teMess)
                        checkList.append(teMess)
                        break
                    except Exception as error:
                        print("Worker " + str(ident) + ": Error sending message: " + str(error))
                        if str(error) == "Request rate exceeded":
                            oPhone.increaseSendThrottle()
                            print("Worker " + str(ident) + ": API throttle engaged while sending message. Force sleep for 1 min for worker: " + str(ident))
                            time.sleep(15)
                            print("Worker " + str(ident) + " has 45s remaining")
                            time.sleep(15)
                            print("Worker " + str(ident) + " has 30s remaining")
                            time.sleep(15)
                            print("Worker " + str(ident) + " has 15s remaining")
                            time.sleep(15)
                            print("Worker " + str(ident) + " is resuming")
                            time.sleep(oPhone.sendInterval + 15)
                        else:
                            print("An error occurred resending message not API throttle")
                            print("Error sending in sendMessage: " + str(error))
                            raise error
            else:
                print("Worker " + str(ident) + ": Message " + str(teMess.messId) + " to number " + teMess.pNum + " to: " + teMess.recipient + " absolutely failed")
                f = open(fiName, "a")
                f.write("Failed: " + str(teMess.pNum) + " " + teMess.recipient + "\n")
                f.close()
                #recheckMessages.remove(teMess)
        elif teMess.sentResult == 'Queued':
            if teMess.reQueue > 5:
                print("Worker " + str(ident) + ": Message " + str(teMess.messId) + " to number " + teMess.pNum + " to: " + teMess.recipient + " requeued a lot")
                f = open(fiName, "a")
                f.write("Queued: " + str(teMess.pNum) + " " + teMess.recipient + "\n")
                f.close()
            else:
                print("Worker " + str(ident) + ": Requeuing " +  teMess.pNum + " to: " + teMess.recipient)
                teMess.reQueue += 1
                checkList.append(teMess)
                continue

    print("Worker " + str(ident) + " is finished")

#*****************************************************************************
# Class: studioPhone
# Purpose: Class for creating a phone to use
#*****************************************************************************

class studioPhone:
    def __init__(self, rcClientId, rcClientSecret, rcUn, rcPass, rcServer, rcExt):
        print("Initializing phone")
        self.rcUn = rcUn
        self.rcsdk = SDK(rcClientId, rcClientSecret, rcServer)
        self.platform = self.rcsdk.platform()
        self.platform.login(rcUn, rcExt, rcPass)
        self.sendInterval = 1
        self.checkInterval = 1

    def increaseSendThrottle(self):
        if self.sendInterval < 20:
            self.sendInterval += 1
            print("Send interval increased to " + str(self.sendInterval))

    def increaseCheckThrottle(self):
        if self.checkInterval < 20:
            self.checkInterval += 1
            print("Check interval increased to " + str(self.checkInterval))

    def resetThrottles(self):
        self.sendInterval = 1
        self.checkInterval = 1

    def decreaseSendThrottle(self):
        if self.sendInterval > 1:
            self.sendInterval -= 1

    def decreaseCheckThrottle(self):
        if self.checkInterval > 1:
            self.checkInterval -= 1

    def sendSMS(self, tMess):
        try:
            result = self.platform.post('/restapi/v1.0/account/~/extension/~/sms',
                          {
                              'from' : { 'phoneNumber': self.rcUn },
                              'to'   : [ {'phoneNumber': tMess.pNum} ],
                              'text' : tMess.tMess
                          })
            tMess.sendAttempt += 1
            tMess.reQueue = 0
            tMess.messId = str(json.loads(result.text())['id'])
            tMess.iterSendAttempt = str(json.loads(result.text())['smsSendingAttemptsCount'])
        except Exception as error:
            print("Error in sendSMS " + str(error))
            raise error

    def checkMessageStatus(self, tMess):
        try:
            checkStatus = self.platform.get('/restapi/v1.0/account/~/extension/~/message-store/' + tMess.messId)
            tMess.sentResult = json.loads(checkStatus.text())["messageStatus"]
            return tMess.sentResult
        except Exception as error:
            print("Error in getMessageStatus: " + str(error))
            raise error

    def getMessages(self, messDir, messFromDate):
        try:
            response = self.platform.get('/restapi/v1.0/account/~/extension/~/message-store',
                                {
                                    'messageType': ['SMS'],
                                    'direction': messDir,
                                    'dateFrom': messFromDate

                                })
            if response.response().status_code == 200:
                return response.text()
            else:
                raise Exception(response.error())
        except Exception as error:
            print("Error in getMessages: " + str(error))
            raise error

#*****************************************************************************
# Class: smsMessage
# Purpose: Class for a single message, to be built by the phone
#*****************************************************************************

class smsMessage:
    def __init__(self, pNum, tMess, recipient):
        self.pNum = pNum
        self.tMess = tMess
        self.recipient = recipient
        self.sentResult = "Not"
        self.sendAttempt = 0
        self.iterSendAttempt = 0

#*******************************************************************************
#Class: smsKiosk
#Purpose: Class  handling bulk message sending
#*******************************************************************************

class smsKiosk:
    def __init__(self, rcId, rcCS, rcUn, rcPass, rcServ, rcExt):
        self.oPhone = studioPhone(rcId, rcCS, rcUn, rcPass, rcServ, rcExt)
        numWorkers = 0
        while numWorkers < 1 or numWorkers > 10:
            try:
                numWorkers = int(input("Number of worker processes (1 to half as many clients): "))
            except Exception as error:
                print(str(error))
        self.numWorkers = numWorkers
        self.oPhone.checkInterval = math.ceil((6/5)*numWorkers)

    def sendSingleMessage(self):
        try:
            cuPnum = getNumber()
            uMessage = input("Type the message you want to send: ")
            self.oPhone.sendSMS(cuPnum, uMessage, 0)
        except Exception as error:
            print("Exception in singleMessOption: " + str(error))

    def sendBulkSMS(self):

        listName = input("Name of the member list file: ")
        while not os.path.isfile(listName):
            print("File not found.")
            if input("Try again? (Y/y or otherwise): ") in ["y", "Y", "yes", "Yes"]:
                listName = input("Name of the member list file: ")
            else:
                return

        try:
            workbookObj = load_workbook(listName)
        except Exception as error:
            print(str(error))
            return

        mainSheet = workbookObj.active

        while True:
            messageFileName = input("Name of the message file: ")
            while not os.path.isfile(messageFileName):
                print("File not found.")
                if input("Try again? (Y/y or otherwise): ") in ["y", "Y", "yes", "Yes"]:
                    messageFileName = input("Name of the message file: ")
                else:
                    return
            try:
                messageFile = open(messageFileName, "r")
                messageBody = messageFile.readline()
            except Exception as error:
                print(str(error))
                raise

            print("Message in file:")
            print(messageBody)
            if input("Is this the correct message? (Y/y or otherwise): ") in ["y", "Y", "yes", "Yes"]:
                break

        recheckMessages = []
        for i in range(self.numWorkers):
            recheckMessages.append([])
        successCount = 0
        count = 1

        #Iterate through the spreadsheet
        for row in mainSheet.iter_rows(min_row=2, max_col=5, max_row=mainSheet.max_row, values_only=True):
            if successCount > 5:
                self.oPhone.decreaseSendThrottle()
            try:
                #Get the information from the spreadsheet, parse it, construct message
                firstName = row[1].lower()
                capFirstName = firstName[0].upper() + firstName[1:]
                lastName = row[0].lower()
                capLastName = lastName[0].upper() + lastName[1:]
                fullName = capFirstName + " " + capLastName
                pNum = getNumberAuto(str(row[4]))
            except Exception as error:
                print("Probs hit the end")
                break
            tMessage = messageBody
            textMessage = smsMessage(pNum, tMessage, fullName)
            #Attempt sending the message
            while True:
                try:
                    self.oPhone.sendSMS(textMessage)
                    break
                except Exception as error:
                    #If the message doesn't send we sleep and try again
                    print("Error sending in message: " + str(error))
                    if str(error) == "Request rate exceeded":
                        self.oPhone.increaseSendThrottle()
                        successCount = 0
                        print("API throttle engaged while sending message. Force sleep for 2 min")
                        time.sleep(75)
                        print("45s remaining")
                        time.sleep(15)
                        print("30s remaining")
                        time.sleep(15)
                        print("15s remaining")
                        time.sleep(15)
                        print("Resuming")
                        time.sleep(self.oPhone.sendInterval + 15)
                    else:
                        print("An error occurred sending message but not API throttle")
                        print("Error sending in sendMessage: " + str(error))
                        raise error
            print("Message attempt sent to " + fullName + " at " + pNum)
            recheckMessages[count % self.numWorkers].append(textMessage)
            count = count + 1
            successCount += 1
            time.sleep(self.oPhone.sendInterval)

        print("Bulk text finished!")

        x = datetime.datetime.now()
        fName = x.strftime("%y%b%d_%H_%M_%S_failedMessageNums.out")

        def start(task, *args):
            process = multiprocessing.Process(target=task, args=args)
            process.daemon = True
            process.start()
        for i in range(self.numWorkers):
            start(worker_task, i, self.oPhone, fName, recheckMessages[i])



#*******************************************************************************
#Function: main
#Purpose: Driver
#*******************************************************************************

def main():

    startupPass = True
    RINGCENTRAL_EXTENSION = '101'

    if sys.argv[1] == "sandbox":
        if sys.argv[2] == "t":
            RINGCENTRAL_CLIENTID = os.getenv("TS_CID")
            RINGCENTRAL_CLIENTSECRET = os.getenv("TS_CS")
            RINGCENTRAL_USERNAME = os.getenv("TS_PNUM")
            RINGCENTRAL_PASSWORD = os.getenv("TS_PASS")
            RINGCENTRAL_SERVER = os.getenv("SAND_ADDRESS")
        elif sys.argv[2] == "hv":
            RINGCENTRAL_CLIENTID = os.getenv("HVS_CID")
            RINGCENTRAL_CLIENTSECRET = os.getenv("HVS_CS")
            RINGCENTRAL_USERNAME = os.getenv("HVS_PNUM")
            RINGCENTRAL_PASSWORD = os.getenv("HVS_PASS")
            RINGCENTRAL_SERVER = os.getenv("SAND_ADDRESS")
        elif sys.argv[2] == "h":
            RINGCENTRAL_CLIENTID = os.getenv("HS_CID")
            RINGCENTRAL_CLIENTSECRET = os.getenv("HS_CS")
            RINGCENTRAL_USERNAME = os.getenv("HS_PNUM")
            RINGCENTRAL_PASSWORD = os.getenv("HS_PASS")
            RINGCENTRAL_SERVER = os.getenv("SAND_ADDRESS")
        elif sys.argv[2] == "evc":
            RINGCENTRAL_CLIENTID = os.getenv("EVCS_CID")
            RINGCENTRAL_CLIENTSECRET = os.getenv("EVCS_CS")
            RINGCENTRAL_USERNAME = os.getenv("EVCS_PNUM")
            RINGCENTRAL_PASSWORD = os.getenv("EVCS_PASS")
            RINGCENTRAL_SERVER = os.getenv("SAND_ADDRESS")
        else:
            print("No proper identity named")
            startupPass = False
    elif sys.argv[1] == "t":
        RINGCENTRAL_CLIENTID = os.getenv("T_CID")
        RINGCENTRAL_CLIENTSECRET = os.getenv("T_CS")
        RINGCENTRAL_USERNAME = os.getenv("T_PNUM")
        RINGCENTRAL_PASSWORD = os.getenv("T_PASS")
        RINGCENTRAL_SERVER = os.getenv("PROD_ADDRESS")
    elif sys.argv[1] == "hv":
        RINGCENTRAL_CLIENTID = os.getenv("HV_CID")
        RINGCENTRAL_CLIENTSECRET = os.getenv("HV_CS")
        RINGCENTRAL_USERNAME = os.getenv("HV_PNUM")
        RINGCENTRAL_PASSWORD = os.getenv("HV_PASS")
        RINGCENTRAL_SERVER = os.getenv("PROD_ADDRESS")
    elif sys.argv[1] == "h":
        RINGCENTRAL_CLIENTID = os.getenv("H_CID")
        RINGCENTRAL_CLIENTSECRET = os.getenv("H_CS")
        RINGCENTRAL_USERNAME = os.getenv("H_PNUM")
        RINGCENTRAL_PASSWORD = os.getenv("H_PASS")
        RINGCENTRAL_SERVER = os.getenv("PROD_ADDRESS")
    elif sys.argv[1] == "evc":
        RINGCENTRAL_CLIENTID = os.getenv("EVC_CID")
        RINGCENTRAL_CLIENTSECRET = os.getenv("EVC_CS")
        RINGCENTRAL_USERNAME = os.getenv("EVC_PNUM")
        RINGCENTRAL_PASSWORD = os.getenv("EVC_PASS")
        RINGCENTRAL_SERVER = os.getenv("PROD_ADDRESS")
    else:
        print("No proper identity named")
        startupPass = False

    frontDesk = smsKiosk(RINGCENTRAL_CLIENTID, RINGCENTRAL_CLIENTSECRET, RINGCENTRAL_USERNAME, RINGCENTRAL_PASSWORD, RINGCENTRAL_SERVER, RINGCENTRAL_EXTENSION)

    while True:
        print("\n1. Send a single message\n2. Send a group message\n3. Get messages\n4. Check workbook\n5. Exit process\n")
        uChoice = input("Pick an option: ")
        if int(uChoice) == 1:
            frontDesk.sendSingleMessage()
        elif int(uChoice) == 2:
            try:
                frontDesk.sendBulkSMS()
            except Exception as error:
                print(str(error))
        elif int(uChoice) == 3:
            print(sPhone.getMessages('Outbound', '2020-05-25T22:00:00.000Z'))
        elif int(uChoice) == 4:
            checkWorkbook()
        elif int(uChoice) == 5:
            print("Goodbye!")
            break
        else:
            print("No valid option chosen\n")
        frontDesk.oPhone.resetThrottles()

    del frontDesk

if __name__ == "__main__":
    main()
