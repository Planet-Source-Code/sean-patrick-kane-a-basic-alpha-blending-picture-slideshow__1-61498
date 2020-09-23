VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slideshow"
   ClientHeight    =   5880
   ClientLeft      =   14820
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   5760
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      DialogTitle     =   "Slideshow - Save Queue"
      Filter          =   "*.txt |*.txt|"
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Slideshow"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame frmBehavior 
      Caption         =   "Slide Behavior"
      Height          =   855
      Left            =   3240
      TabIndex        =   10
      Top             =   4320
      Width           =   5415
      Begin VB.ComboBox cmbSlides 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   3120
         List            =   "frmMain.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Seconds between slides"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frmQueue 
      Caption         =   "Files in queue"
      Height          =   4095
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton lstLoad 
         Caption         =   "Load queue list"
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save queue list"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove selected files from queue"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   3120
         Width           =   2895
      End
      Begin VB.ListBox lstQueue 
         Height          =   2595
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame frmFiles 
      Caption         =   "Files available for queue"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add selected files to queue"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   5040
         Width           =   2295
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   360
         MultiSelect     =   2  'Extended
         System          =   -1  'True
         TabIndex        =   3
         Top             =   3240
         Width           =   2295
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Make sure VB will return an error if we don't declare ALL our variables

Private Sub cmdAdd_Click()
Dim i As Integer
'Run through the File1 listbox -- if it's set to true, it means it is selected
For i = 0 To File1.ListCount - 1    'Start at 0 because the listbox is 0-indexed (starts at 0, not 1)
    If File1.Selected(i) = True Then 'Yep, it's selected...we need to add it
        lstQueue.AddItem File1.Path & "\" & File1.List(i)   'Add the item to the right side listbox -- this is the listbox used as the queue for the slideshow
    End If
Next i
End Sub

Private Sub cmdRemove_Click()
Dim i As Integer
'Run through them all to see if it's selected as true
For i = lstQueue.ListCount - 1 To 0 Step -1 'if you go backwards, you don't have to reset the listcount -- this is a very neat trick that saves problems/bugs
    If lstQueue.Selected(i) = True Then
        lstQueue.RemoveItem i
    End If
Next i
End Sub

Private Sub cmdRun_Click()
If lstQueue.ListCount <> 0 Then 'If there's at least one item in the queue...
    frmSlideshow.Show   'Start the slideshow
Else    'If there isn't anything in the queue...
    MsgBox "You must add picture files into the queue!", vbCritical, "Error"    'Tell them to go add stuff
End If
End Sub

Private Sub cmdSave_Click()
Dim FF As Long, i As Integer
CommonDialog.ShowSave   'Show the save dialog -- the filename will be saved as CommonDialog.filename when they're done
If CommonDialog.filename <> "" Then 'they didn't hit cancel... -- .filename will equal "" if they hit cancel
    FF = FreeFile   'Find a file number that is free
    Open CommonDialog.filename For Output As FF 'Open it up -- if there was anything in the file beforehand...it will be overwritten
        For i = 0 To lstQueue.ListCount - 1
            Print #FF, lstQueue.List(i) 'Write each queued item onto a new line in the file
        Next i
    Close FF    'Close the file we opened
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path  'If they change the directory listbox, the file listbox will obviously change
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive    'If they change the drive, we need to change the directory listbox
File1.Path = Dir1.Path      'As well as the file listbox
End Sub

Private Sub Form_Load()
cmbSlides.ListIndex = 0 'Set a value for the two combo boxes -- if this wasn't here, they'd start out blank
End Sub

Private Sub lstLoad_Click()
Dim FF As Long, tmpInput As String
CommonDialog.ShowOpen   'The file they want to open is saved as CommonDialog.filename
If CommonDialog.filename <> "" Then 'If they didn't hit cancel...
    lstQueue.Clear  'We're opening a file...so clear anything that is there ahead of time
    FF = FreeFile   'Open a new file number
    Open CommonDialog.filename For Input As FF  'Open up the file
        Do Until EOF(FF)    'Loop until we hit the end of the file (EOF)
            Line Input #FF, tmpInput    'Grab the next line
            lstQueue.AddItem tmpInput   'Add it to the queue
        Loop
    Close FF    'Clean up
End If
End Sub
