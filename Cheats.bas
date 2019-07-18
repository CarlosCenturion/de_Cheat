Attribute VB_Name = "Cheats"
Public myHandle As Long

Function InitProcessCheater(pid As Long)

pHandle = OpenProcess(&H1F0FFF, False, pid)

If (pHandle = 0) Then
    InitProcessCheater = False
    myHandle = 0
Else
    InitProcessCheater = True
    myHandle = pHandle
End If

End Function



Function DoFirstSearch(s As String) As Integer

Dim c As Integer
Dim addr As Long
Dim buffer As String * 5000
Dim readlen As Long

Open "c:\cheat.mem" For Output As #1: Close #1 ' kill if exists
Open "c:\cheat.mem" For Random As #1 Len = Len(addr)

'count of results
c = 0

For addr = 0 To 40000    ' loop through buffers

Call ReadProcessMemory(myHandle, addr * 5000, buffer, 5000, readlen)

If addr Mod 400 = 0 Then
'update status
  GameEditor.lblStatus.Caption = "Searching %" & Trim(Str(Int(addr / 400)))
  DoEvents
End If

'if read successfull
If readlen > 0 Then
  startpos = 1
  'find all search string in buffer
  While InStr(startpos, buffer, Trim(s)) > 0
    p = (addr) * 5000 + InStr(startpos, buffer, s) - 1 ' position of string
    Put #1, , CLng(p)  ' put address in file for later searches
    c = c + 1 ' increase counter
    If c < 20 Then GameEditor.lstResults.AddItem p
    startpos = InStr(startpos, buffer, Trim(s)) + 1 ' find next position
  Wend
End If

'next buffer
Next addr

'Update status
  GameEditor.lblStatus.Caption = "Search Done."

'close file
Close #1

DoFirstSearch = c

End Function


Function DoNextSearch(s As String)

Dim sc As Integer
Dim addr As Long
Dim buffer As String
Dim readlen As Long

buffer = Space(Len(s))

'open first search results
Open "c:\cheat.mem" For Random As #1 Len = Len(addr)

'clear results
GameEditor.lstResults.Clear

'reset filepointer
fp = 0

'clear counter
sc = 0

'loop until end of file
While Not EOF(1)
    
    'increase file pointer
    fp = fp + 1
    'read address data
    Get #1, fp, addr
    
    'if it is not banned
    If addr <> 0 Then
        
        buffer = Space(Len(s))
        Call ReadProcessMemory(myHandle, addr, buffer, Len(s), readlen)
     
             
        If buffer <> s Then
            'ban address
            Put #1, fp, CLng(0)
        Else
            'add result
            sc = sc + 1
            If sc < 20 Then GameEditor.lstResults.AddItem addr
        End If

    End If


Wend


Close #1


DoNextSearch = sc


End Function

