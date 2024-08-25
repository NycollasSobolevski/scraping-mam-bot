import datetime

class UserInfo:
    login = "thiago.brummer"
    password = "1234"
    
class TicketData:
    def __init__(self, ticketNumber, ticketPass , ticketClient, ticketStartDateTime, adress):
        self.Number = ticketNumber
        self.Pass = ticketPass
        self.StartDateTime = ticketStartDateTime
        self.Client =  ticketClient
        self.Adress = adress
        
    # getDate = lambda date: date.split(" ")[0]
    
    def getStartDate(self):
        date = self.StartDateTime.split(" ")[0].split('-')
        time = self.StartDateTime.split(" ")[1].split(':')
        
        return datetime.datetime(int(date[0]), int(date[1]), int(date[2]), int(time[0]), int(time[1]))
    
    def getStringDate(self):
        return self.getStartDate().date().strftime("%d/%m/%Y")
    def getStringTime(self):
        return self.getStartDate().date().strftime("%H:%M")