VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim fNAME As String
  '//Create Unicode folder first before giving a folder name with Unicode Characters!!!
  fNAME = BrowseForFolder(Me.hwnd, "C:\Program Files", ChrW$(&H6B22) & ChrW$(&H8FCE) & " to choose a folder ...", True, True)  '"C:\" & ChrW$(&H6B22) & ChrW$(&H8FCE), , , True)
  If fNAME <> "" Then 'they did not hit cancel
      MsgBox fNAME
  End If
End Sub

Function MsgBox(Prompt As String, _
                Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
                Optional Title As String) As VbMsgBoxResult

'MsgBox Unicode Version, Writen by Dana Seaman (www.cuberActivex.com)
Dim WshShell         As Object

    Set WshShell = CreateObject("WScript.Shell")
    MsgBox = WshShell.Popup(Prompt, 0&, Title, Buttons)
    Set WshShell = Nothing

End Function


