from ringcentral import SDK
import time
import json
from smsMessageClass import smsMessage

#*******************************************************************************
# Class: studioPhone
# Purpose: Class for creating a phone to use
#*******************************************************************************

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
        while True:
            try:
                time.sleep(self.sendInterval)
                result = self.platform.post('/restapi/v1.0/account/~/extension/~/sms',
                              {
                                  'from' : { 'phoneNumber': self.rcUn },
                                  'to'   : [ {'phoneNumber': tMess.getPnum()} ],
                                  'text' : tMess.getMessageText()
                              })
            except Exception as error:
                print("Error in sendSMS sending in message: " + str(error))
                if str(error) == "Request rate exceeded":
                    self.increaseSendThrottle()
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
                    time.sleep(self.sendInterval + 15)
                else:
                    print("An error occurred sending message but not API throttle")
                    print("Error sending in sendMessage: " + str(error))
                    raise error
            else:
                tMess.increaseSendAttempt()
                tMess.resetReQueue()
                tMess.setMessId(str(json.loads(result.text())['id']))
                tMess.setIterSendAttempt(str(json.loads(result.text())['smsSendingAttemptsCount']))
                break

    def checkMessageStatus(self, tMess):
        while True:
            try:
                time.sleep(self.checkInterval)
                checkStatus = self.platform.get('/restapi/v1.0/account/~/extension/~/message-store/' + tMess.getMessId())
            except Exception as error:
                print("Error checking message status: " + str(error))
                if str(error) == "Request rate exceeded":
                    self.increaseCheckThrottle()
                    print("API throttle engaged while checking message status. Force sleep for 1 min: " + str(ident))
                    time.sleep(15)
                    print("45s remaining")
                    time.sleep(15)
                    print("30s remaining")
                    time.sleep(15)
                    print("15s remaining")
                    time.sleep(15)
                    print("Resuming")
                    time.sleep(self.checkInterval)
                else:
                    print("An error occurred checking message status but not API throttle")
                    print("Error checking message status: " + str(error))
                    raise error
            else:
                tMess.setResult(json.loads(checkStatus.text())["messageStatus"])
                break

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
