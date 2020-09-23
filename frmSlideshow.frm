VERSION 5.00
Begin VB.Form frmSlideshow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Slideshow"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3240
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrChange 
      Enabled         =   0   'False
      Left            =   2280
      Top             =   1200
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmSlideshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currentindex As Integer

Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_Load()
    'Hide the cursor -- we're really just putting the cursor in the bottom right corner, but this causes much fewer problems than using the SetCursor function
    SetCursorPos Screen.Width, Screen.Height
    currentindex = 0 'we're going to start the slideshow at frmmain.lstqueue.list(0)
    'Set the backgrounds to black so they blend into the background of the form
    picDisplay.BackColor = vbBlack
    picTemp.BackColor = vbBlack
    picDraw.BackColor = vbBlack
    'Prepare our form... (make it the size of the screen, and set it in the top left corner)
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.Top = 0
    Me.Left = 0
    'Prepare picDisplay (the picturebox that will contain the picture we want to show the user)...
    picDisplay.Width = Screen.Width
    picDraw.Width = Screen.Width
    picTemp.Width = Screen.Width
    picDisplay.Height = Screen.Height
    picDraw.Height = Screen.Height
    picTemp.Height = Screen.Height
    picDisplay.Left = 0
    picDraw.Left = 0
    picTemp.Left = 0
    picDisplay.Top = 0
    picDraw.Top = 0
    picTemp.Top = 0
    'Set the interval for the form based on the combo box in frmMain
    tmrChange.interval = CInt(frmMain.cmbSlides.Text) * 1000    'There are 1000 milliseconds in one second...so if the user select 1 second, that is equivalent to 1000 milliseconds, and the .interval property uses milliseconds
    'If the user doesn't select 1 second, it takes too long for the timer to take effect...so we're going to "jumpstart" the timer
    DisplayPicture frmMain.lstQueue.List(0), picDisplay
    tmrChange.Enabled = True
    If frmMain.cmbSlides.ListIndex <> 0 Then tmrChange_Timer 'This is the function that runs at each interval on the timer
End Sub

Private Sub picDisplay_DblClick()
Unload Me
End Sub

Private Sub picDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub picDraw_Click()
Unload Me
End Sub

Private Sub picTemp_Click()
Unload Me
End Sub

Private Sub tmrChange_Timer()
Dim BF As BLENDFUNCTION, lBF As Long
Dim alphapercent As Integer
    'Increment currentindex
    currentindex = currentindex + 1
    'At this point, a picture has already been displayed (lstQueue.list(0) if we just started the timer) -- we need to check to see if we're done yet...
    If currentindex = frmMain.lstQueue.ListCount Then
        For alphapercent = 0 To 255 'Try to keep the last image on screen about as long as the others
            Pause 0.2
        Next alphapercent
        picDisplay.Picture = LoadPicture("") 'Show a blank image
        tmrChange.Enabled = False 'We don't want this function running again because we've gone through the entire queue
        picDisplay.Visible = False  'We need to hide this picturebox so they can see the text that we're going to print on the next line
        Me.Print "You have reached the end of the slideshow"    'Note: I can't figure out how to center this...but I'd like to
        Me.Refresh
        Exit Sub 'We're done -- no need to stay in this function
    End If
    
    'If we're at this point, we have a picture displayed, and in order to move on to the next image (the currentindex image), we need to alpha blend it
    DisplayPicture frmMain.lstQueue.List(currentindex), picTemp 'Put our currentindex picture into picTemp -- this will alpha blend over picDisplay
    'Now let's go ahead and do a slow alpha blend -- But first, let's set up the situation...
    picTemp.ScaleMode = vbPixels
    picDisplay.ScaleMode = vbPixels
    BF.BlendOp = AC_SRC_OVER
    BF.BlendFlags = 0
    BF.AlphaFormat = 0
    For alphapercent = 0 To 255 Step 20 'We'll increment the percent that the new image will show until it's showing all the way (0->0%...255->100%)
        BF.SourceConstantAlpha = alphapercent 'Set the new blend percent -- this will change after each iteration in the for loop
        RtlMoveMemory lBF, BF, 4 'Set the bf  variable into memory
        AlphaBlend picDisplay.hdc, 0, 0, picDisplay.ScaleWidth, picDisplay.ScaleHeight, picTemp.hdc, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, lBF 'Actually blend the two pictures together
        picDisplay.Refresh 'The change won't be evident unless we refresh the picturebox
    Next alphapercent
    picDisplay.ScaleMode = 1 'Set the scalemode back to default
    picTemp.ScaleMode = 1
    'Now that we're done alpha blending, just replace the image
    DisplayPicture frmMain.lstQueue.List(currentindex), picDisplay
End Sub

Private Function DisplayPicture(filename As String, Pic1 As PictureBox)
On Error Resume Next    'The error would be on the line 'picDraw.Picture = LoadPicture(filename)' -- it just means that the file doesn't exist or can't be opened...instead of throwing a fit, just move on
Dim resizeratio As Double
    'Before we can display a fresh image...we have to clear the background from the previous image
    Pic1.Picture = LoadPicture("") 'clear the pictureboxes
    picDraw.Picture = LoadPicture("")
    picDraw.Picture = LoadPicture(filename)
    'todo: error check for bad pictures
    
    'We need to figure out if we should resize the image...if the image comes off the screen, we should resize it
    If (picDraw.Picture.Width > Screen.Width) Or (picDraw.Picture.Height > picDraw.Picture.Height) Then 'Doesn't fit...we need to resize
        If picDraw.Picture.Width > picDraw.Picture.Height Then 'If it's wider than it is tall...
            resizeratio = Screen.Width / picDraw.Picture.Width
        Else    'Note: if the picture is too big and it's perfectly square...it will fall into the 'else' statement, but it won't matter because the resizeratio would be the same
            resizeratio = Screen.Height / picDraw.Picture.Height
        End If
    Else
        'The picture fit...just set the ratio to 1 so it won't be resized any
        resizeratio = 1
    End If
    'Right now, resizeratio * both the width and height of the picture will resize the picture so it fits into the screen perfectly
    'We've got everything ready...let's paint the picture onto the picturebox -- the follow function also centers the picture
    Pic1.PaintPicture picDraw.Picture, (Pic1.Width / 2) - (picDraw.Picture.Width * resizeratio / 2), (Pic1.Height / 2) - (picDraw.Picture.Height * resizeratio / 2), picDraw.Picture.Width * resizeratio, picDraw.Picture.Height * resizeratio
    Pic1.Picture = Pic1.Image
End Function
