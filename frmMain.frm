VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "VB JSON Deserializer"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptionClass 
      Caption         =   "Run Class"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton OptionModClass 
      Caption         =   "Run Module"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton CommandReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TextFilePath 
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   10695
   End
   Begin MSComDlg.CommonDialog cmndialog 
      Left            =   3240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox CheckVBJSONDeserializer 
      Caption         =   "VBJSONDeserializer"
      Height          =   255
      Left            =   7560
      TabIndex        =   4
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox CheckJson 
      Caption         =   "JSON"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.TextBox txtVBJSONDeserializer 
      Height          =   6375
      Left            =   7560
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   7095
   End
   Begin VB.TextBox txtJSON 
      Height          =   6375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   7215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function FileText(ByVal filename As String) As String
    Dim handle As Integer
     
       ' ensure that the file exists
    If Len(Dir$(filename)) = 0 Then
        Err.Raise 53  ' File not found
    End If
     
       ' open in binary mode
    handle = FreeFile
    Open filename$ For Binary As #handle
       ' read the string and close the file
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
End Function

Private Sub cmdRun_Click()

Dim deserializer As cJSONDeserializer
Dim sw As New StopWatch
Dim elapsedtime As Long
Dim filename As String
Dim filecontents As String
Dim resultstring As String
Dim numberofruns As Integer
Dim items As Variant
Dim i As Integer
Dim totaltime As Long

numberofruns = 5
filename = Me.TextFilePath.Text
filecontents = FileText(filename)

Screen.MousePointer = vbHourglass

Set deserializer = New cJSONDeserializer

If Me.CheckJson.Value = vbChecked Then
    
    totaltime = 0

    For i = 1 To numberofruns
        sw.StartTimer
        Set items = JSON.parse(filecontents)
        elapsedtime = sw.EndTimer
        totaltime = totaltime + elapsedtime
    
        resultstring = Now & vbCrLf & _
                        "Parsing time: " & elapsedtime & " [micro sec]" & vbCrLf & _
                        "items.Count: " & items.Count
        
        Me.txtJSON.Text = Me.txtJSON.Text & vbCrLf & vbCrLf & resultstring
        DoEvents
    Next i
    
    Me.txtJSON.Text = Me.txtJSON.Text & vbCrLf & vbCrLf & "average time: " & Round((totaltime / numberofruns), 2) & " [micro sec]"
    
End If

If Me.CheckVBJSONDeserializer.Value = vbChecked Then
    
    totaltime = 0
    
    For i = 1 To numberofruns
        sw.StartTimer
        If Me.OptionModClass Then
            Set items = VBJSONDeserializer.parse(filecontents)
        Else
            Set items = JSON.parse(filecontents)
        End If
        elapsedtime = sw.EndTimer
        totaltime = totaltime + elapsedtime
        
        resultstring = Now & vbCrLf & _
                        "Parsing time: " & elapsedtime & " [micro sec]" & vbCrLf & _
                        "items.Count: " & items.Count
        
        Me.txtVBJSONDeserializer.Text = Me.txtVBJSONDeserializer.Text & vbCrLf & vbCrLf & resultstring
        DoEvents
    Next i
    
    Me.txtVBJSONDeserializer.Text = Me.txtVBJSONDeserializer.Text & vbCrLf & vbCrLf & "average time: " & Round((totaltime / numberofruns), 2) & " [micro sec]"
    
End If

Screen.MousePointer = vbNormal

End Sub

Private Sub CommandReset_Click()

Me.txtJSON.Text = vbNullString
Me.txtVBJSONDeserializer.Text = vbNullString

End Sub

Private Sub Form_Load()
Me.TextFilePath.Text = App.Path & "\test - 10000.json"
End Sub

Private Sub TextFilePath_DblClick()

On Error GoTo errorHandler

With Me.cmndialog
    .Filter = "JSON Files (*.json)|*.json|Text Files (*.txt)|*.txt"
    .FilterIndex = 1
    .DefaultExt = ".json"
    .CancelError = True
    .DialogTitle = "Open JSON file"
    .ShowOpen
End With

Me.TextFilePath.Text = Me.cmndialog.filename

errorHandler:
Select Case Err.Number
Case 0
Case 32755
    Me.TextFilePath.Text = ""
Case Else
    Debug.Print Err.Number & " " & Err.Description
End Select
End Sub
