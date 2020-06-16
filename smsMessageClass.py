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
        self.reQueue = 0

    def setResult(self, sentResult):
        self.sentResult = sentResult

    def getResult(self):
        return self.sentResult

    def increaseSendAttempt(self):
        self.sendAttempt += 1

    def getSendAttempt(self):
        return self.sendAttempt

    def resetReQueue(self):
        self.reQueue = 0

    def increaseReQueue(self):
        self.reQueue += 1

    def getReQueue(self):
        return self.reQueue

    def setIterSendAttempt(self, iterSendAttempt):
        self.iterSendAttempt = iterSendAttempt

    def getIterSendAttempt(self):
        return self.iterSendAttempt

    def getPnum(self):
        return self.pNum

    def getRecipient(self):
        return self.recipient

    def setMessId(self, messId):
        self.messId = messId

    def getMessId(self):
        return self.messId

    def getMessageText(self):
        return self.tMess
