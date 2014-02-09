VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.9#0"; "EditCtlsU.ocx"
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
   Begin VB.CheckBox CheckShowNames 
      Caption         =   "Show Names"
      Height          =   255
      Left            =   9600
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.OptionButton OptionClass 
      Caption         =   "Run Class"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton OptionMod 
      Caption         =   "Run Module"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton CommandReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TextFilePath 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox CheckJson 
      Caption         =   "JSON"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin EditCtlsLibUCtl.TextBox txtVBJSONDeserializer 
      Height          =   6375
      Left            =   7440
      TabIndex        =   7
      Top             =   1800
      Width           =   7215
      _cx             =   12726
      _cy             =   11245
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   3075
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FormattingRectangleHeight=   0
      FormattingRectangleLeft=   0
      FormattingRectangleTop=   0
      FormattingRectangleWidth=   0
      HAlignment      =   0
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertSoftLineBreaks=   0   'False
      LeftMargin      =   -1
      MaxTextLength   =   -1
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   -1  'True
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":0000
      Text            =   "frmMain.frx":0020
   End
   Begin EditCtlsLibUCtl.TextBox txtJSON 
      Height          =   6375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   7215
      _cx             =   12726
      _cy             =   11245
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   3075
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FormattingRectangleHeight=   0
      FormattingRectangleLeft=   0
      FormattingRectangleTop=   0
      FormattingRectangleWidth=   0
      HAlignment      =   0
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertSoftLineBreaks=   0   'False
      LeftMargin      =   -1
      MaxTextLength   =   -1
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   -1  'True
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":0040
      Text            =   "frmMain.frx":0060
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function FileText(ByVal filename As String) As String

Dim stmText As New ADODB.Stream

With stmText
    .Open
    .Type = adTypeBinary
    .LoadFromFile filename
    .Type = adTypeText
    .LineSeparator = adLF
    .Charset = "UTF-8"
    Do Until .EOS
        FileText = .ReadText(adReadLine)
    Loop
    .Close
End With

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
Dim item As Dictionary
Dim i As Integer
Dim totaltime As Long
Dim arraytime As Long

numberofruns = 5
filename = Me.TextFilePath.Text
filecontents = FileText(filename)

Screen.MousePointer = vbHourglass

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
        
        Me.txtJSON.AppendText resultstring & vbCrLf, True, True
        
        DoEvents
    Next i
    
    Me.txtJSON.AppendText vbCrLf & "average time: " & Round((totaltime / numberofruns), 2) & " [micro sec]", True, True
    
End If

If Me.CheckVBJSONDeserializer.Value = vbChecked Then
    
    totaltime = 0
    
    For i = 1 To numberofruns
        sw.StartTimer
        If Me.OptionMod Then
            Set items = VBJSONDeserializer.parse(filecontents)
            Debug.Print VBJSONDeserializer.GetParserErrors
        Else
            Set deserializer = New cJSONDeserializer
            Set items = deserializer.parse(filecontents)
            Set deserializer = Nothing
            arraytime = 0
        End If
        elapsedtime = sw.EndTimer
        totaltime = totaltime + elapsedtime
        
        resultstring = Now & vbCrLf & _
                        "Parsing time: " & elapsedtime & " [micro sec]" & vbCrLf & _
                        "items.Count: " & items.Count
        
        Me.txtVBJSONDeserializer.AppendText resultstring & vbCrLf, True, True
        
        DoEvents
    Next i
    
    Me.txtVBJSONDeserializer.AppendText vbCrLf & "average time: " & Round((totaltime / numberofruns), 2) & " [micro sec]" & vbCrLf, True, True
    
    If Me.CheckShowNames.Value = vbChecked Then
        For Each item In items
            Me.txtVBJSONDeserializer.AppendText "Name: " & item.item("Name") & vbCrLf
            DoEvents
        Next item
    End If
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
