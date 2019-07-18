VERSION 5.00
Begin VB.Form GameEditor 
   Caption         =   "DE_CHEAT by EzekieL Centurion [ LooSeR ]"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Procesos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   2295
      Begin VB.ComboBox cboProcess 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   310
         Width           =   1815
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   14
         ToolTipText     =   "Actualizar"
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buskar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   2295
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Start a new search"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Sig."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Continue Searching"
         Top             =   1200
         Width           =   615
      End
      Begin VB.ListBox lstHistory 
         Height          =   1035
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resultados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
      Begin VB.ListBox lstResults 
         Height          =   1425
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tabla de Cheats"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin VB.ListBox lstActiveCheats 
         Height          =   2205
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdFreeze 
         Caption         =   "Frizar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         ToolTipText     =   "Freeze Address"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdPoke 
         Caption         =   "Nuke"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         ToolTipText     =   "Poke Address a value"
         Top             =   640
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Parar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         ToolTipText     =   "Remove Cheat"
         Top             =   1040
         Width           =   735
      End
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   0
      Top             =   4920
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   480
      Top             =   4920
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   100
      Left            =   960
      Top             =   4920
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   1440
      Top             =   4920
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   100
      Left            =   1920
      Top             =   4920
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   100
      Left            =   2400
      Top             =   4920
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   100
      Left            =   2880
      Top             =   4920
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   100
      Left            =   3360
      Top             =   4920
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   100
      Left            =   3840
      Top             =   4920
   End
   Begin VB.Timer freezetimer 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   100
      Left            =   4320
      Top             =   4920
   End
   Begin VB.Frame Frame4 
      Caption         =   "Informacion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2400
      TabIndex        =   0
      Top             =   2640
      Width           =   3615
      Begin VB.Label Label1 
         Caption         =   "PrOgRaMa by EzekieL C."
         BeginProperty Font 
            Name            =   "Inkburrow"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "LooSeR.LooSeR@Hotmail.com [ xDWeb.TK ]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   4560
      Width           =   7815
   End
   Begin VB.Menu mnuResults 
      Caption         =   "Results"
      Visible         =   0   'False
      Begin VB.Menu mnuAddtoactive 
         Caption         =   "Agregar"
      End
   End
   Begin VB.Menu mnuActives 
      Caption         =   "Active"
      Visible         =   0   'False
      Begin VB.Menu mnuaddtotable 
         Caption         =   "Agregar"
      End
   End
End
Attribute VB_Name = "GameEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private freezevalues(100) As Double
Private freezecount As Integer
Private rightclick As Boolean
Private PIDs(1000) As Long

Private Sub RefreshProcessList()
'Reads Process List and Fills combobox (cboProcess)

Dim myProcess As PROCESSENTRY32
Dim mySnapshot As Long

'first clear our combobox
cboProcess.Clear

myProcess.dwSize = Len(myProcess)

'create snapshot
mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

'get first process
ProcessFirst mySnapshot, myProcess
cboProcess.AddItem myProcess.szexeFile ' set exe name
PIDs(cboProcess.ListCount - 1) = myProcess.th32ProcessID ' set PID

'while there are more processes
While ProcessNext(mySnapshot, myProcess)
  cboProcess.AddItem myProcess.szexeFile ' set exe name
  PIDs(cboProcess.ListCount - 1) = myProcess.th32ProcessID ' ' store PID
Wend

End Sub



Private Sub cmdFreeze_Click()

If lstActiveCheats.ListIndex > -1 Then
  Call mnuFreeze_Click
End If

End Sub

Private Sub cmdNext_Click()
'Search Next

Dim sc As Integer
Dim strSearch As String

If lstHistory.ListCount = 0 Then MsgBox "Please use start to search first value.", vbCritical, "CheatMaster": Exit Sub

'check if numeric
If Not IsNumeric(txtSearch) Then MsgBox "Please enter only a numeric value in search field.", vbCritical, "CheatMaster": Exit Sub

'should be less or equal to DWORD &hFFFFFFFF
If Val(txtSearch) > 4294967295# Then MsgBox "Please enter a smaller value in search field.", vbCritical, "CheatMaster": Exit Sub

'convert number to string
strSearch = ConvertNumberToString(Val(txtSearch))

'add to history
lstHistory.AddItem txtSearch

'do next search
sc = DoNextSearch(strSearch)


If sc = 1 Then MsgBox "You got it :)", vbInformation, "CheatMaster": Exit Sub

If sc = 0 Then MsgBox "Not found in memory, please start from beginning.", vbCritical, "CheatMaster"

If sc > 20 Then MsgBox "Found" & Str(sc) & " results. You should search more.", vbInformation, "CheatMaster"

If sc > 0 And sc < 21 Then MsgBox "Please check results box, you can search more to identify exact address", vbInformation, "CheatMaster"


End Sub

Private Sub cmdPoke_Click()
If lstActiveCheats.ListIndex > -1 Then
  Call mnuPoke_Click
End If
End Sub

Private Sub cmdRefresh_Click()
'Refresh Process List
  RefreshProcessList
End Sub

Private Sub cmdStart_Click()
'First Search

Dim sc As Integer
Dim strSearch As String

'clear history
lstHistory.Clear

'check if numeric
If Not IsNumeric(txtSearch) Then MsgBox "Please enter only a numeric value in search field.", vbCritical, "CheatMaster": Exit Sub

'should be less or equal to DWORD &hFFFFFFFF
If Val(txtSearch) > 4294967295# Then MsgBox "Please enter a smaller value in search field.", vbCritical, "CheatMaster": Exit Sub

'convert number to string
strSearch = ConvertNumberToString(Val(txtSearch))

'check combo selected
If cboProcess.ListIndex = -1 Then MsgBox "Please select a process to cheat.", vbCritical, "CheatMaster": Exit Sub

'init cheater
If Not InitProcessCheater(PIDs(cboProcess.ListIndex)) Then MsgBox "Could not open process. sorry :(", vbCritical, "CheatMaster": Exit Sub

'do first search
sc = DoFirstSearch(strSearch)

'add to history
lstHistory.AddItem txtSearch

If sc = 0 Then MsgBox "Not found in memory, please try other values.", vbCritical, "CheatMaster"

If sc > 20 Then MsgBox "Found" & Str(sc) & " results. You should search more.", vbInformation, "CheatMaster"

If sc > 0 And sc < 21 Then MsgBox "Please check results box, you can search more to identify exact address", vbInformation, "CheatMaster"

End Sub

Private Sub cmdStop_Click()

If lstActiveCheats.ListIndex > -1 Then
  Call mnuStop_Click
End If

End Sub




Private Function ConvertNumberToString(number As Double) As String
'converts number to string will be searched in memory

If number < 256 Then ConvertNumberToString = Chr(number): Exit Function

If number < 65536 Then
    ConvertNumberToString = Chr(number And 255) & Chr((number And 65280) / 256)
    Exit Function
End If

b4 = number And 255: number = Int(number / 256)
b3 = number And 255: number = Int(number / 256)
b2 = number And 255: number = Int(number / 256)
b1 = number And 255: number = Int(number / 256)

ConvertNumberToString = Chr(b4) & Chr(b3) & Chr(b2) & Chr(b1)

End Function


Private Sub freezetimer_Timer(Index As Integer)

Dim addr As Long
Dim value As String * 1

addr = CLng(Val(freezetimer(Index).Tag))
value = ConvertNumberToString(freezevalues(Index))

Call WriteProcessMemory(myHandle, addr, value, Len(value), l)

End Sub

Private Sub lstResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 And lstResults.ListIndex > -1 Then PopupMenu mnuResults
End Sub

Private Sub mnuAddtoactive_Click()
  
  If lstActiveCheats.ListCount > 10 Then MsgBox "You cannot add more than 10 cheats.", vbCritical, "CheatMaster": Exit Sub
  strname = InputBox("Please enter name for this cheat", "CheatMaster")
  If strname = "" Then strname = "[no name]" Else strname = "[" & strname & "]"
  lstActiveCheats.AddItem lstResults.List(lstResults.ListIndex) & " :: " & strname
End Sub

Private Sub mnuFreeze_Click()
  
  strValue = InputBox("Please enter value to freeze", "CheatMaster")
  If strValue = "" Or Not IsNumeric(strValue) Then Exit Sub
  X = Split(lstActiveCheats.List(lstActiveCheats.ListIndex), "::")
  freezetimer(lstActiveCheats.ListIndex).Tag = Trim(X(0))
  freezevalues(lstActiveCheats.ListIndex) = Val(strValue)
  freezetimer(lstActiveCheats.ListIndex).Enabled = True

End Sub

Private Sub mnuPoke_Click()
  
  Dim addr As Long
  Dim value As String
    
  strValue = InputBox("Please enter value to freeze", "CheatMaster")
  If strValue = "" Or Not IsNumeric(strValue) Then Exit Sub
  X = Split(lstActiveCheats.List(lstActiveCheats.ListIndex), "::")
  
  addr = CLng(Val(Trim(X(0))))
  value = ConvertNumberToString(Val(strValue))
   
  Call WriteProcessMemory(myHandle, addr, value, Len(value), l)
    
End Sub

Private Sub mnuStop_Click()
  freezetimer(lstActiveCheats.ListIndex).Enabled = False
End Sub
Private Sub Form_Load()
Call FormOnTop(Me.hwnd, True)
 RefreshProcessList
End Sub

