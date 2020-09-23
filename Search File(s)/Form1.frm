VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search File(s)"
   ClientHeight    =   4320
   ClientLeft      =   1590
   ClientTop       =   1410
   ClientWidth     =   9225
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Search results"
      Height          =   3615
      Left            =   2520
      TabIndex        =   7
      Top             =   600
      Width           =   6615
      Begin VB.ListBox resultsList 
         Height          =   3180
         ItemData        =   "Form1.frx":0442
         Left            =   120
         List            =   "Form1.frx":0449
         TabIndex        =   8
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.DriveListBox myDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   165
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Se&arch Now"
      Default         =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox myDir 
      Height          =   3240
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.FileListBox myFile 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Choose drive"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Choose a directory to search"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@@@@@@@@@@@@@@@[[[[by Geovany Villegas]]]]@@@@@@@@@@@@@@@@@@@@
'example of using recursion to search a file, this need a
'FileListBox and a DirListBox basically. Optional Drive control
'If you use this code please don't forget the original author!

'PS: Sorry about my poor english.

Private Sub Command1_Click()

    If Trim(Text2.Text) = "" Then
        MsgBox "You must specify a Filename"
        Exit Sub
    Else
        myFile.Pattern = Text2.Text
    End If

    resultsList.Clear
    myDir.Enabled = False
    myDrive.Enabled = False
    
    DoEvents
    
    Call SearchSubs(myDir.Path)
    
    myDrive.Enabled = True
    myDir.Enabled = True
    myDrive.SetFocus
    
    DoEvents
    
    MsgBox "There is " & resultsList.ListCount & " File(s) in your search", , "myFile(s) found"
End Sub

Private Sub myDrive_Change()
    On Error Resume Next
    
    myDir.Path = myDrive.Drive
End Sub

Private Sub SearchSubs(ByVal lastDir As String)
Dim idxSub As Integer
Dim i As Integer
    
    'sets the active directory and get the last
    'subdirectory's index in this directory
    myDir.Path = lastDir
    idxSub = myDir.ListCount - 1
    
    'runs the procedure again (recursion) while the folder
    'contains more subdirectories
    Do While idxSub > -1
        Call SearchSubs(myDir.List(idxSub))
        myDir.Path = lastDir
        idxSub = idxSub - 1
    Loop
    
    'updates the FileListBox and add the files
    'to the list.
    myFile.Path = lastDir

    For i = 0 To myFile.ListCount - 1
        If Len(myFile.Path) > 3 Then
            resultsList.AddItem myFile.Path & "\" & myFile.List(i)
        Else
            resultsList.AddItem myFile.Path & myFile.List(i)
        End If
    Next i
    
    'refresh the screen
    DoEvents
End Sub
