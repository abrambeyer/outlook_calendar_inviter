Option Explicit

 

Sub CREATE_MULTIPLE_APPOINTMENTS()

 

Dim O As Outlook.Application
Set O = New Outlook.Application

 


Dim ONS As Outlook.Namespace
Set ONS = O.GetNamespace("mapi")

 

Dim CAL_FOL As Outlook.Folder
Set CAL_FOL = ONS.GetDefaultFolder(olFolderCalendar)

 

Dim APT As Outlook.AppointmentItem

 

Dim wb As Workbook
Dim ws As Worksheet

 

Set wb = ActiveWorkbook
Set ws = Sheets("For Abram")

 

Dim r As Long
Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient
Dim mtgAttendee As Outlook.Recipient
For r = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    Set APT = CAL_FOL.Items.Add(olAppointmentItem)
    
    With APT
        .MeetingStatus = olMeeting
        .Start = ws.Cells(r, 8).Value + ws.Cells(r, 9).Value
        .End = DateAdd("d", 1, ws.Cells(r, 8).Value) + ws.Cells(r, 10).Value
        .Subject = ws.Cells(r, 1).Value
        .Location = ws.Cells(r, 11).Value
        .Body = "PERFORMANCE CLOSE TASK"
        .AllDayEvent = True
        Set mtgAttendee = .Recipients.Add(ws.Cells(r, 12).Value)
        mtgAttendee.Type = olRequired
        If IsEmpty(ws.Cells(r, 13).Value) = False Then
            Set mtgAttendee = .Recipients.Add(ws.Cells(r, 13).Value)
            mtgAttendee.Type = olRequired
        End If
        If IsEmpty(ws.Cells(r, 14).Value) = False Then
            Set mtgAttendee = .Recipients.Add(ws.Cells(r, 14).Value)
            mtgAttendee.Type = olRequired
        End If
        If IsEmpty(ws.Cells(r, 15).Value) = False Then
            Set mtgAttendee = .Recipients.Add(ws.Cells(r, 15).Value)
            mtgAttendee.Type = olRequired
        End If
        If IsEmpty(ws.Cells(r, 16).Value) = False Then
            Set mtgAttendee = .Recipients.Add(ws.Cells(r, 16).Value)
            mtgAttendee.Type = olRequired
        End If
        .Save
        If (APT.Recipients.ResolveAll) Then
            APT.Send
        Else
            APT.Display
        End If
    End With
        
Next

 

End Sub
 
