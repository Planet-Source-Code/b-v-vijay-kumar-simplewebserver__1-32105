VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   FillColor       =   &H00C0E0FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4080
      Top             =   3240
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1530
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sample WebServer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   2970
      Shape           =   3  'Circle
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   2259
      Shape           =   3  'Circle
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   1546
      Shape           =   3  'Circle
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   810
      Shape           =   3  'Circle
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Author : B.V.Vijay Kumar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   570
      TabIndex        =   2
      Top             =   2670
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Send Data to IE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   635
      TabIndex        =   1
      Top             =   2107
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   1605
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'*
'* Created on Febraury 26,Tuesday,2002
'* Created by B.V.Vijay Kumar
'* My Mail ID : vijay_hrd@yahoo.com
'*
'******************************************************


Private Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim strData As String
Dim strreq As Long
Dim i As Integer
Dim j As Integer
Dim intStart As Long
Dim intEnd As Long
Dim strFilename As String
Dim strFileStream As String
Dim strErrorString As String

Private Sub Form_Load()
    
    Winsock1.LocalPort = 1001
    Winsock1.Listen
    'Create a Elliptical Region
    h = CreateEllipticRgn(0, 0, 300, 300)
    
    'set the Elliptical region
    h = SetWindowRgn(Me.hWnd, h, True)
    
    Timer1.Enabled = False
    
End Sub

Private Sub Label5_Click()
    End
End Sub

Private Sub Timer1_Timer()

    'Animate the Form
    animate

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    Winsock1.Accept requestID
    strreq = requestID
 
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    
    'Get the Header info from the IE, parse it to get requested file name
    Winsock1.GetData strd, vbString
    
    
    strFilename = ParseHeader(strd)
        
    If strFilename = "ERROR" Then
        Winsock1.SendData MessageInHTML("Inavalid Extension")
    ElseIf strFilename = "ANIMATE" Then
        Timer1.Enabled = True
        Winsock1.SendData MessageInHTML("Animation Started")
    ElseIf strFilename = "STOP" Then
        Timer1.Enabled = False
        Winsock1.SendData MessageInHTML("Animation Stopped")
    Else
       strFileStream = GetFileStream(strFilename)
       Winsock1.SendData strFileStream
    End If
    
End Sub

Private Sub Winsock1_SendComplete()
   
    Winsock1.Close
    Winsock1.LocalPort = 1001
    Winsock1.Listen

End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    
    Label1.Caption = "Bytes Sent : " & bytesSent & "  " & " Bytes Remaining : " & bytesRemaining

End Sub

Public Sub animate()
    
    Me.Move i, j, Me.Width, Me.Height
    i = i + 20
    
    If i >= Screen.Width Then i = 0
    j = j + 20
    
    If j >= Screen.Width Then j = 0
    
End Sub

Public Function ParseHeader(ByVal strHeader As String) As String
        
    'This code is to parse the Header sent by the browser and check for
    'an extension and get the filename
    
    'These are the extensions, this can be extended further
    Dim strExtensions(3) As String
    Dim blnFound As Boolean
    
    strExtensions(0) = ".html"
    strExtensions(1) = ".htm"
    strExtensions(2) = ".asp"
    
    'Find for the first / and .html,.asp,.htm etc and get the filename only
    s = strHeader
    intStart = InStr(s, "/")
    If intStart > 0 Then
        s = Mid(s, intStart + 1, Len(s))
    End If
    
    'Check all the Extensions
    blnFound = False
    For i = 0 To UBound(strExtensions, 1) - 1
        intEnd = InStr(s, strExtensions(i))
        If intEnd > 0 Then
            blnFound = True
            Exit For
        End If
    Next
    
    'If the Extension not found then check whether Animation Params called
    If blnFound = True Then
        actstring = Mid(s, 1, intEnd + Len("html"))
        ParseHeader = actstring
        Exit Function
    Else
        actstring = Mid(s, 1, InStr(s, "HTTP") - 1)
    End If
    
    'Check if the parameter is one of ANIMATE or STOP
    If UCase(Trim(actstring)) = "ANIMATE" Then
        ParseHeader = "ANIMATE"
    ElseIf UCase(Trim(actstring)) = "STOP" Then
        ParseHeader = "STOP"
    Else
        ParseHeader = "ERROR"
    End If
    
End Function

Public Function GetFileStream(ByVal vstrFilename As String) As String
    
    On Error Resume Next
    Dim fso As New Scripting.FileSystemObject
    Dim txtTextStream As TextStream
    
    'Get the filename and read the stream and send that stream to Internet Explorer
    Set txtTextStream = fso.OpenTextFile(App.Path & "\" & vstrFilename, ForReading)
    strErrorString = ""
    If Err.Number <> 0 Then
        strErrorString = MessageInHTML(Err.Description)
        GetFileStream = strErrorString
        Exit Function
    End If
    
    GetFileStream = txtTextStream.ReadAll
     
End Function

Public Function MessageInHTML(ByVal strhtmlMessage As String)
    MessageInHTML = "<HTML><BODY><CENTER><FONT SIZE=15>" & strhtmlMessage & "</FONT></CENTER></BODY></HTML>"
    
End Function

