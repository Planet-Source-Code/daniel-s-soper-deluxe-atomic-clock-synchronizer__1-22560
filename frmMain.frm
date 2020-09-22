VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atomic Time Synchronizer!"
   ClientHeight    =   2700
   ClientLeft      =   1050
   ClientTop       =   3135
   ClientWidth     =   2895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timDelay 
      Enabled         =   0   'False
      Interval        =   4500
      Left            =   600
      Top             =   2880
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1080
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   15
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   1455
      ScaleWidth      =   58500
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   58500
   End
   Begin VB.Timer timAni 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   8
      X2              =   29
      Y1              =   134
      Y2              =   134
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   128
      X2              =   184
      Y1              =   134
      Y2              =   134
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Synchronization Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'*                                                                               *
'* ATOMIC TIME SYNCHRONIZER! By: Daniel S. Soper... If you like it, vote for it! *
'*                                                                               *
'* NOTES: This program is designed to automatically terminate itself after       *
'*        attempting to synchronize your system clock to the USNO atomic clock.  *
'*        If you compile this code into an EXE, you can simply move the EXE file *
'*        into your "Startup" folder, and then your system will automatically    *
'*        synchronize itself to the atomic time server every time you log on to  *
'*        your computer.                                                         *
'*********************************************************************************

Option Explicit 'Explicitly declare all variables


Private Type SYSTEMTIME 'SystemTime structure for the Win32 API
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION 'TimeZoneInformation structure for the Win32 API
   Bias As Long
   StandardName(0 To 63) As Byte
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long 'Retrieves the time zone settings of this system
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal hSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long 'Blits animation frames to the form

Dim AniFrame As Integer 'Keeps track of which frame of animation to display
Dim TimeServerURL As String 'The URL of the U.S. Naval Observatory's atomic time server
Dim InProcess As Boolean 'Keeps track of whether or not the program is currently trying to sychronize the system clock

Private Sub GetAtomicTime() 'This sub retreives the raw time data file from the USNO atomic time server
    On Error GoTo ErrRtn 'Traps the "RequestTimeout" property of the Internet Transfer Control
    Dim tempData As String 'Holds the data received from the atomic time server
    
    lblStatus.Caption = "Connecting to USNO Atomic Time server..." 'Updates the status label
    DoEvents
    tempData = Inet1.OpenURL(TimeServerURL) 'Request time data from USNO atomic time server
    Call SetAtomicTime(tempData) 'Call SetAtomicTime sub
    Exit Sub
    
ErrRtn: 'This routine is run only if the attempted network request fails
    lblStatus.ForeColor = vbRed
    lblStatus.Caption = "Network request failed!"
    InProcess = False
    Unload Me 'Exit program
End Sub

Private Sub SetAtomicTime(RawData As String) 'Extrapolates the UTC time from the raw data received from the USNO
                                             'atomic time server, and sets the local system's time to the time-zone
                                             'adjusted UTC atomic time
                                             
    Dim X As Integer 'Holds found character positions
    Dim Y As Integer 'Holds found character positions
    Dim tempTime As Variant 'Holds the extrapolated UTC and adjusted times
    
    X = InStr(1, RawData, "Universal") 'Find "Universal" in the raw data ("Universal" indicates UTC time)
    If X > 0 Then 'If "Universal" was found in the raw data
        tempTime = Left$(RawData, X) 'Set "tempTime" equal to the section of the raw data we're interested in
        Y = InStrRev(tempTime, ",") 'Find the first comma in the tempTime data, starting from the back
        If Y > 0 Then 'If a comma was found in the "tempTime" data
            tempTime = CDate(Trim(Mid$(RawData, Y + 1, (X - (Y + 1))))) 'Cast the "tempTime" variable into a date containing the extracted actual UTC time
            Time = tempTime - AdjustTimeForTimeZone 'Set the local system time to the time-zone adjusted UTC atomic time
            lblStatus.ForeColor = RGB(127, 255, 127) 'Change the status label's forecolor to light green
            lblStatus.Caption = "Your system time has been changed to: " & Time & "..." 'Update the status label
            InProcess = False
            DoEvents
            Unload Me 'Exit the program
        Else 'If no comma was found in the "tempTime" data
            lblStatus.ForeColor = vbRed
            lblStatus.Caption = "Received bad data!"
            InProcess = False
            DoEvents
            Unload Me 'Exit the program
        End If
    Else 'If "Universal" was not found in the raw data
        lblStatus.ForeColor = vbRed
        lblStatus.Caption = "Received bad data!"
        InProcess = False
        DoEvents
        Unload Me 'Exit the program
    End If
End Sub

Private Function AdjustTimeForTimeZone() As Single 'Returns the amount of adjustment necessary
                                                   'from UTC time for the current system by checking
                                                   'the system's time zone and daylight savings properties

    Dim TZI As TIME_ZONE_INFORMATION 'Holds the system's time zone information
    Dim DaylightSavingsTime As Boolean 'Holds the system's daylight savings time status
    Dim RetVal As Long 'Return value for calculations

    Call GetTimeZoneInformation(TZI) 'Populate the TimeZoneInformation structure
    If TZI.StandardDate.wMonth = Month(Now) Then 'Check for daylight savings time
        DaylightSavingsTime = Day(Now) < TZI.StandardDate.wDay
    ElseIf TZI.DaylightDate.wMonth = Month(Now) Then
        DaylightSavingsTime = Day(Now) >= TZI.DaylightDate.wDay
    Else
        If TZI.DaylightDate.wMonth < TZI.StandardDate.wMonth Then
            DaylightSavingsTime = Month(Now) > TZI.DaylightDate.wMonth And Month(Now) < TZI.StandardDate.wMonth
        Else
            DaylightSavingsTime = Month(Now) > TZI.DaylightDate.wMonth Or Month(Now) < TZI.StandardDate.wMonth
        End If
    End If
    RetVal = TZI.Bias 'the difference, in minutes, between Coordinated Universal Time (UTC) and local time
    If DaylightSavingsTime = True Then 'Calculate the daylight savings adjustment (if any)
        RetVal = RetVal + TZI.DaylightBias
    Else
        RetVal = RetVal + TZI.StandardBias
    End If
    AdjustTimeForTimeZone = ((RetVal / 60) / 24) 'Calculate and return the final time adjustment
End Function

Private Sub Form_Load() 'Program Startup
    TimeServerURL = "http://tycho.usno.navy.mil/cgi-bin/timer.pl" 'Set the URL of the USNO atomic time server
    lblStatus.Caption = "Initializing..." 'Update the status label
    AniFrame = 0 'Set the current animation frame to 0
    InProcess = True 'Atomic clock synchronization is being attempted
    timAni.Enabled = True 'Enable the time that controls the animation
    Me.Visible = True 'Make sure the form is visible
    Me.Refresh
    Call GetAtomicTime 'Try to retreive the current atomic time
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Request to unload program
    If InProcess = True Then 'Do not allow program to unload if atomic clock synchronization is being attempted
        Cancel = 1
    Else 'Allow a delay so that the user can see the status of the sychronization attempt before closing
        If timDelay.Enabled = False Then
            Cancel = 1
            timDelay.Enabled = True
        End If
    End If
End Sub

Private Sub timAni_Timer() 'Controls the animation displayed onscreen
    Call StretchBlt(Me.hdc, 3, 3, (Me.ScaleWidth - 6), (Label1.Top - 6), picBuffer.hdc, (AniFrame * 130), 0, 130, 97, vbSrcCopy) 'Blit the current frame to the screen
    AniFrame = AniFrame + 1 'Increment the current animation frame
    If AniFrame > 29 Then 'If all of the animation frames have been displayed, start again at frame 0
        AniFrame = 0
    End If
    Me.Refresh 'Refresh the screen
End Sub

Private Sub timDelay_Timer() 'Allows a delay so that the user can see the status of the sychronization attempt before closing
    DoEvents
    Unload Me
End Sub
