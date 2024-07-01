# need to install the pywin32 library
# pip install pywin32

import win32com.client

#Create and send appointment
outlook = win32com.client.Dispatch("Outlook.Application")
appt = outlook.CreateItem(1) # AppointmentItem, 0 - Email
appt.Start = "2020-11-20 17:35" # yyyy-MM-dd hh:mm
appt.Subject = "Subject of the meeting"
appt.Duration = 60 # In minutes (60 Minutes)
appt.Location = "Location Name"

# 1 - olMeeting; Changing the appointment to meeting.
# Only after changing the meeting status recipients can be added
appt.MeetingStatus = 1
  
appt.Recipients.Add("<SOME_EMAIL>") # Don't end ; as delimiter

appt.Save()
appt.Send()
