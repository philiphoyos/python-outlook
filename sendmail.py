#
import win32com.client
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "Python generated mail"
#newMail.HTMLBody  = htmltext
newMail.Body = """
				Test mail fra Python
				"""
newMail.To = "yourmail@domain.com"
#attachment1 = "c:\\mypic.jpg"
#newMail.Attachments.Add(attachment1)
newMail.Send()
