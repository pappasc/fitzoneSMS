from ringcentral import SDK
from openpyxl import load_workbook
import phonenumbers
import datetime
import math
import time
import json
import multiprocessing
import zmq
import sys
import os
from dotenv import load_dotenv
load_dotenv()

def checkWorkbook():
    listName = input("Workbook name: ")
    workbookObj = load_workbook(listName)
    mainSheet = workbookObj.active
    count = 1
    for row in mainSheet.iter_rows(min_row=2, max_col=5, max_row=mainSheet.max_row, values_only=True):
        firstName = row[1].lower()
        capFirstName = firstName[0].upper() + firstName[1:]
        pNum = str(row[4])
        print(str(count) + ": " + capFirstName + " " + pNum)
        count += 1
    print(str(mainSheet.max_row))

def getFirstName(nameNum):
    index = 0
    for i in range(len(nameNum)):
        if nameNum[i] == " ":
            index = i
    return nameNum[:index]

def getNumber():
    while True:
        userNum = input("What is the phone number (start with area code)?: ")
        parseNum = phonenumbers.parse(str(userNum), "US")
        if phonenumbers.is_valid_number(parseNum):
            return phonenumbers.format_number(parseNum, phonenumbers.PhoneNumberFormat.E164)
        else:
            print("Invalid number. Try again")

def getNumberAuto(nameNum):
    index = 0
    for i in range(len(nameNum)):
        if nameNum[i] == " ":
            index = i + 1
    pfNum = nameNum[index:]
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

def getNumber():
    while True:
        userNum = input("What is the phone number (start with area code)?: ")
        parseNum = phonenumbers.parse(str(userNum), "US")
        if phonenumbers.is_valid_number(parseNum):
            return phonenumbers.format_number(parseNum, phonenumbers.PhoneNumberFormat.E164)
        else:
            print("Invalid number. Try again")

def client_task(numClients, ident, listName):
    """Basic request-reply client using REQ socket."""
    socket = zmq.Context().socket(zmq.REQ)
    socket.identity = u"Client-{}".format(ident).encode("ascii")
    socket.connect("ipc://frontend.ipc")

    workbookObj = load_workbook(listName)
    mainSheet = workbookObj.active
    count = 1
    for row in mainSheet.iter_rows(min_row=2, max_col=5, max_row=mainSheet.max_row, values_only=True):
        if count % numClients == ident:
            firstName = row[1].lower()
            capFirstName = firstName[0].upper() + firstName[1:]
            pNum = str(row[4])
            socket.send((capFirstName + " " + pNum).encode())
            reply = socket.recv()
            print("{}: {}".format(socket.identity.decode("ascii"),
                          reply.decode("ascii")))
            count = count + 1
        else:
            count = count + 1

def worker_task(ident, oPhone, fName):
    """Worker task, using a REQ socket to do load-balancing."""
    socket = zmq.Context().socket(zmq.REQ)
    socket.identity = u"Worker-{}".format(ident).encode("ascii")
    socket.connect("ipc://backend.ipc")

    # Tell broker we're ready for work
    socket.send(b"READY")

    messageAddress = "Hello "
    messageBody = "Blow up message"

    while True:
        address, empty, request = socket.recv_multipart()
        nameNum = request.decode("ascii")
        fName = getFirstName(nameNum)
        pNum = getNumberAuto(nameNum)
        tMessage = messageBody
        #tMessage = "Test message from worker: {}".format(socket.identity.decode("ascii"))
        # Reaplcing this print statement would be the work fu   nction i.e. send text message to API and get response
        print("{}: {}".format(socket.identity.decode("ascii"),
                              request.decode("ascii")))
        # This is the worker response back that everytning worked fine (you can worry about reliability or
        # error messages in a future release
        try:
            if oPhone.sendSMS(pNum, tMessage, ident):
                outcomeMessage = "Message successfully sent to " + fName + " at " + pNum
            else:
                outcomeMessage = "Message send failed to " + fName + " at " + pNum
                f = open(fName, "a")
                f.write(pNum + "\n")
                f.close()
        except Exception as error:
            print("Failure sending message to " + fName + " at " + pNum)
            print(str(error))
        socket.send_multipart([address, b"", outcomeMessage.encode()])
        # Force worker to sleep for a hot sec
        time.sleep(oPhone.sendInterval)

"""*****************************************************************************
Class: studioPhone
Purpose: Class for creating a phone to use
*****************************************************************************"""

class studioPhone:
    def __init__(self, rcClientId, rcClientSecret, rcUn, rcPass, rcServer, rcExt):
        print("Initializing phone")
        self.rcUn = rcUn
        self.rcsdk = SDK(rcClientId, rcClientSecret, rcServer)
        self.platform = self.rcsdk.platform()
        self.platform.login(rcUn, rcExt, rcPass)
        self.sendInterval = 1
        self.checkInterval = 1

    def increaseSendThrottle(self, workerId):
        self.sendInterval += 1
        print("Send interval increased to " + str(self.sendInterval) + " for worker " + str(workerId))

    def increaseCheckThrottle(self, workerId):
        self.checkInterval += 1
        print("Check interval increased to " + str(self.checkInterval) + " for worker " + str(workerId))

    def resetThrottles(self):
        self.sendInterval = 1
        self.checkInterval = 1

    def sendSMS(self, pNum, tMess, workerId):
        try:
            textMessage = smsMessage(pNum, tMess)
            while not textMessage.sendMessage(self, workerId) and textMessage.sendAttempt < 3:
                continue
            if not textMessage.succSent:
                print("Failed to send message to: " + pNum)
                return False
            del textMessage
            return True
        except Exception as error:
            print("Error in sendSMS " + str(error))
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

"""*****************************************************************************
Class: smsMessage
Purpose: Class for a single message, to be built by the phone
*****************************************************************************"""

class smsMessage:
    def __init__(self, pNum, tMess):
        self.pNum = pNum
        self.tMess = tMess
        self.succSent = False
        self.sendAttempt = 0
        self.retryCounter = 0

    def getMessageStatus(self, oPhone):
        try:
            checkStatus = oPhone.platform.get(self.messURL)
            return checkStatus
        except Exception as error:
            print("Error in getMessageStatus: " + str(error))
            raise error

    def sendMessage(self, oPhone, workerId):
        try:
            result = oPhone.platform.post('/restapi/v1.0/account/~/extension/~/sms',
                          {
                              'from' : { 'phoneNumber': oPhone.rcUn },
                              'to'   : [ {'phoneNumber': self.pNum} ],
                              'text' : self.tMess
                          })
            self.sendAttempt += 1
            self.messURL = json.loads(result.text())['uri']
            while True:
                if self.retryCounter == 10:
                    print("10 check attempts on message to " + self.pNum + " reached on send attempt " + str(self.sendAttempt))
                    self.retryCounter = 0
                    break
                time.sleep(oPhone.checkInterval)
                self.retryCounter += 1
                try:
                    succCheck = self.getMessageStatus(oPhone)
                    succCheckJson = json.loads(succCheck.text())
                    if succCheckJson['messageStatus'] == 'Queued':
                        continue
                    elif succCheckJson['messageStatus'] == 'Sent' or succCheckJson['messageStatus'] == 'Delivered' or succCheckJson['messageStatus'] == 'Received' or succCheckJson['readStatus'] == 'Read':
                        self.succSent = True
                        break
                    elif succCheckJson['messageStatus'] == 'SendingFailed' or succCheckJson['messageStatus'] == 'DeliveryFailed':
                        if succCheckJson['smsSendingAttemptsCount'] >= 3:
                            break
                        else:
                            continue
                    else:
                        raise Exception("Something broke checking message status")
                except Exception as error:
                    print("Error checking in sendMessage: " + str(error))
                    if str(error) == "Request rate exceeded":
                        oPhone.increaseCheckThrottle(workerId)
                        print("API throttle engaged while checking message status. Force sleep for 1 min for worker: " + str(workerId))
                        time.sleep(15)
                        print("Worker " + str(workerId) + " has 45s remaining")
                        time.sleep(15)
                        print("Worker " + str(workerId) + " has 30s remaining")
                        time.sleep(15)
                        print("Worker " + str(workerId) + " has 15s remaining")
                        time.sleep(15)
                        print("Worker " + str(workerId) + " is resuming")
                    else:
                        print("A 400 level error occurred in sendMessage status 403 but not API throttle")
                        print("Error checking in sendMessage: " + str(error))
                        raise Exception(result.error())
            return self.succSent
        except Exception as error:
            print("Error sending in sendMessage: " + str(error))
            if str(error) == "Request rate exceeded":
                oPhone.increaseSendThrottle(workerId)
                print("API throttle engaged while sending message. Force sleep for 1 min for worker: " + str(workerId))
                time.sleep(15)
                print("Worker " + str(workerId) + " has 45s remaining")
                time.sleep(15)
                print("Worker " + str(workerId) + " has 30s remaining")
                time.sleep(15)
                print("Worker " + str(workerId) + " has 15s remaining")
                time.sleep(15)
                print("Worker " + str(workerId) + " is resuming")
                return self.succSent
            else:
                print("A 400 level error occurred in sendMessage status 403 but not API throttle")
                print("Error sending in sendMessage: " + str(error))
                raise Exception(result.error())

"""*****************************************************************************
Class: smsKiosk
Purpose: Class for handling bulk message sending
*****************************************************************************"""

class smsKiosk:
    def __init__(self, oPhone):
        self.studioPhone = oPhone
        numClients = 0
        numWorkers = 0
        while numClients < 1 or numClients > 30:
            try:
                numClients = int(input("Number of client processes (1-30): "))
            except Exception as error:
                print(str(error))
        while numWorkers < 1 or numWorkers > math.ceil(numClients/2):
            try:
                numWorkers = int(input("Number of worker processes (1 to half as many clients): "))
            except Exception as error:
                print(str(error))
        self.numClients = numClients
        self.numWorkers = numWorkers

    def sendSingleMessage(self):
        try:
            cuPnum = getNumber()
            uMessage = input("Type the message you want to send: ")
            self.studioPhone.sendSMS(cuPnum, uMessage, 0)
        except Exception as error:
            print("Exception in singleMessOption: " + str(error))

    def sendBulkSMS(self):
        context = zmq.Context.instance()
        frontend = context.socket(zmq.ROUTER)
        frontend.bind("ipc://frontend.ipc")
        backend = context.socket(zmq.ROUTER)
        backend.bind("ipc://backend.ipc")

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
            return

        mainSheet = workbookObj.active

        x = datetime.datetime.now()
        fName = x.strftime("%y%b%d_%H_%M_%S_failedMessageNums.txt")

        # Start background tasks
        def start(task, *args):
            process = multiprocessing.Process(target=task, args=args)
            process.daemon = True
            process.start()
        for i in range(self.numClients):
            start(client_task, self.numClients, i, listName)
        for i in range(self.numWorkers):
            start(worker_task, i, self.studioPhone, fName)

        # Initialize main loop state
        count = mainSheet.max_row - 1
        upCount = 0
        workers = []
        poller = zmq.Poller()
        # Only poll for requests from backend until workers are available
        poller.register(backend, zmq.POLLIN)

        while True:
            sockets = dict(poller.poll())

            if backend in sockets:
                # Handle worker activity on the backend
                request = backend.recv_multipart()
                worker, empty, client = request[:3]
                if not workers:
                    # Poll for clients now that a worker is available
                    poller.register(frontend, zmq.POLLIN)
                workers.append(worker)
                if client != b"READY" and len(request) > 3:
                    # If client reply, send rest back to frontend
                    empty, reply = request[3:]
                    frontend.send_multipart([client, b"", reply])
                    count -= 1
                    upCount += 1
                    print("Message number " + str(upCount) + " attempted")
                    if not count:
                        break

            if frontend in sockets:
                # Get next client request, route to last-used worker
                client, empty, request = frontend.recv_multipart()
                worker = workers.pop(0)
                backend.send_multipart([worker, b"", client, b"", request])
                if not workers:
                    # Don't poll clients if no workers are available
                    poller.unregister(frontend)

        # Clean up
        backend.close()
        frontend.close()
        context.term()
        print("Bulk text finished!")

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


    sPhone = studioPhone(RINGCENTRAL_CLIENTID, RINGCENTRAL_CLIENTSECRET, RINGCENTRAL_USERNAME, RINGCENTRAL_PASSWORD, RINGCENTRAL_SERVER, RINGCENTRAL_EXTENSION)
    frontDesk = smsKiosk(sPhone)

    while True:
        print("\n1. Send a single message\n2. Send a group message\n3. Get messages\n4. Check workbook\n5. Exit process\n")
        uChoice = input("Pick an option: ")
        sPhone.resetThrottles()
        if int(uChoice) == 1:
            frontDesk.sendSingleMessage()
        elif int(uChoice) == 2:
            frontDesk.sendBulkSMS()
        elif int(uChoice) == 3:
            print(sPhone.getMessages('Outbound', '2020-05-25T22:00:00.000Z'))
        elif int(uChoice) == 4:
            checkWorkbook()
        elif int(uChoice) == 5:
            print("Goodbye!")
            break
        else:
            print("No valid option chosen\n")

    del frontDesk
    del sPhone

if __name__ == "__main__":
    main()
