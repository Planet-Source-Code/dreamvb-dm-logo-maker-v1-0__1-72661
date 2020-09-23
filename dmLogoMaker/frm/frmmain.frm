VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DM Logo Maker"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSave 
      AutoRedraw      =   -1  'True
      Height          =   315
      Left            =   690
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   61
      Top             =   7650
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   9900
      TabIndex        =   60
      Top             =   4305
      Width           =   1680
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   405
      Left            =   9900
      TabIndex        =   59
      Top             =   3780
      Width           =   1680
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   405
      Left            =   9900
      TabIndex        =   58
      Top             =   3255
      Width           =   1680
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800000&
      Height          =   330
      Left            =   9795
      ScaleHeight     =   270
      ScaleWidth      =   1770
      TabIndex        =   56
      Top             =   105
      Width           =   1830
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   57
         Top             =   30
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6990
      Top             =   2370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   2685
      Left            =   9825
      ScaleHeight     =   2625
      ScaleWidth      =   1740
      TabIndex        =   18
      Top             =   435
      Width           =   1800
      Begin VB.PictureBox PicColor 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   30
         ScaleHeight     =   225
         ScaleWidth      =   1650
         TabIndex        =   55
         Top             =   45
         Width           =   1710
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00009900&
         Height          =   210
         Index           =   35
         Left            =   30
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   54
         Top             =   2385
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H0033CC66&
         Height          =   210
         Index           =   34
         Left            =   465
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   53
         Top             =   2385
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H0066FF99&
         Height          =   210
         Index           =   33
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   52
         Top             =   2385
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00CCFFCC&
         Height          =   210
         Index           =   32
         Left            =   1335
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   51
         Top             =   2385
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00CC6600&
         Height          =   210
         Index           =   31
         Left            =   30
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   50
         Top             =   2130
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FF9900&
         Height          =   210
         Index           =   30
         Left            =   465
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   49
         Top             =   2130
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FFCC99&
         Height          =   210
         Index           =   29
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   48
         Top             =   2130
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FFEED2&
         Height          =   210
         Index           =   28
         Left            =   1335
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   47
         Top             =   2130
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00CC6666&
         Height          =   210
         Index           =   27
         Left            =   30
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   46
         Top             =   1875
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FF9999&
         Height          =   210
         Index           =   26
         Left            =   465
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   45
         Top             =   1875
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FFCCCC&
         Height          =   210
         Index           =   25
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   44
         Top             =   1875
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FFCCFF&
         Height          =   210
         Index           =   24
         Left            =   1335
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   43
         Top             =   1875
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H000036D9&
         Height          =   210
         Index           =   23
         Left            =   30
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   42
         Top             =   1635
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H004B79FF&
         Height          =   210
         Index           =   22
         Left            =   465
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   41
         Top             =   1635
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H003399FF&
         Height          =   210
         Index           =   21
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   40
         Top             =   1635
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H009DDBFF&
         Height          =   210
         Index           =   20
         Left            =   1335
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   39
         Top             =   1635
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00006699&
         Height          =   210
         Index           =   19
         Left            =   30
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   38
         Top             =   1395
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H000099CC&
         Height          =   210
         Index           =   18
         Left            =   465
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   37
         Top             =   1395
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H0000CCFF&
         Height          =   210
         Index           =   17
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   36
         Top             =   1395
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H0099FFFF&
         Height          =   210
         Index           =   16
         Left            =   1335
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   35
         Top             =   1395
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FF0000&
         Height          =   210
         Index           =   15
         Left            =   30
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   34
         Top             =   1155
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FF00FF&
         Height          =   210
         Index           =   14
         Left            =   465
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   33
         Top             =   1155
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FFFF00&
         Height          =   210
         Index           =   13
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   32
         Top             =   1155
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   12
         Left            =   1335
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   31
         Top             =   1155
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00C0C0C0&
         Height          =   210
         Index           =   11
         Left            =   30
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   30
         Top             =   900
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H000000FF&
         Height          =   210
         Index           =   10
         Left            =   465
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   29
         Top             =   900
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H0000FF00&
         Height          =   210
         Index           =   9
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   28
         Top             =   900
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H0000FFFF&
         Height          =   210
         Index           =   8
         Left            =   1335
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   27
         Top             =   900
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00800000&
         Height          =   210
         Index           =   7
         Left            =   30
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   26
         Top             =   645
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00800080&
         Height          =   210
         Index           =   6
         Left            =   465
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   25
         Top             =   645
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00808000&
         Height          =   210
         Index           =   5
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   24
         Top             =   645
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00808080&
         Height          =   210
         Index           =   4
         Left            =   1335
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   23
         Top             =   645
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   30
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   22
         Top             =   390
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   450
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   21
         Top             =   390
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00008000&
         Height          =   210
         Index           =   2
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   20
         Top             =   390
         Width           =   405
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00008080&
         Height          =   210
         Index           =   3
         Left            =   1335
         ScaleHeight     =   150
         ScaleWidth      =   345
         TabIndex        =   19
         Top             =   390
         Width           =   405
      End
   End
   Begin VB.ListBox lstSize 
      Height          =   1440
      IntegralHeight  =   0   'False
      Left            =   8535
      TabIndex        =   17
      Top             =   4035
      Width           =   1215
   End
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   11595
      TabIndex        =   16
      Top             =   5565
      Width           =   11655
   End
   Begin VB.ListBox lstFonts 
      Height          =   1440
      IntegralHeight  =   0   'False
      Left            =   5580
      TabIndex        =   14
      Top             =   4035
      Width           =   2865
   End
   Begin VB.CheckBox chkShowB 
      Caption         =   "Show Text Line 2"
      Height          =   225
      Left            =   3840
      TabIndex        =   13
      Top             =   5055
      Value           =   1  'Checked
      Width           =   1590
   End
   Begin VB.CheckBox chkShowA 
      Caption         =   "Show Text Line 1"
      Height          =   225
      Left            =   2175
      TabIndex        =   12
      Top             =   5055
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox TextLine 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3060
      TabIndex        =   11
      Text            =   "Your text here"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox TextLine 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3060
      TabIndex        =   10
      Text            =   "Your text here"
      Top             =   4065
      Width           =   2415
   End
   Begin VB.PictureBox pTitle2 
      BackColor       =   &H00800000&
      Height          =   330
      Left            =   2025
      ScaleHeight     =   270
      ScaleWidth      =   7680
      TabIndex        =   6
      Top             =   3630
      Width           =   7740
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fonts Properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3570
         TabIndex        =   15
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   7
         Top             =   30
         Width           =   1290
      End
   End
   Begin VB.FileListBox imgLoc 
      Height          =   1455
      Left            =   75
      Pattern         =   "*.gif;*.jpg;*.bmp"
      TabIndex        =   5
      Top             =   4020
      Width           =   1875
   End
   Begin VB.PictureBox pTitle1 
      BackColor       =   &H00800000&
      Height          =   330
      Left            =   75
      ScaleHeight     =   270
      ScaleWidth      =   1815
      TabIndex        =   3
      Top             =   3630
      Width           =   1875
      Begin VB.Label lblImage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image Properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   4
         Top             =   30
         Width           =   1440
      End
   End
   Begin VB.PictureBox pArea 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Left            =   75
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   120
      Width           =   9660
      Begin VB.Image imgLogo 
         Height          =   675
         Left            =   4110
         Picture         =   "frmmain.frx":0000
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your text here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   5100
         TabIndex        =   2
         Top             =   1710
         Width           =   1050
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your text here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   2970
         TabIndex        =   1
         Top             =   1410
         Width           =   1050
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text Line 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2145
      TabIndex        =   9
      Top             =   4635
      Width           =   795
   End
   Begin VB.Label lblText1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text Line 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2160
      TabIndex        =   8
      Top             =   4110
      Width           =   795
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TheObject As Object
Private OldX As Integer
Private OldY As Integer
Private tIndex As Integer

Private Sub BuildLogo()
Dim Ret As Long
    
    'Refresh canvas
    pArea.Refresh
    
    With picSave
        .Width = pArea.Width
        .Height = pArea.Height
        'Paste pArea onto PicSave
        Ret = BitBlt(.hDC, 0, 0, pArea.Width, pArea.Height, pArea.hDC, 0, 0, vbSrcCopy)
        .Refresh
    End With
End Sub

Private Function GetDLGName(Optional ShowOpen As Boolean = True, Optional Title As String = "Open", Optional Filter As String)
On Error GoTo CanErr:
    'Show open or save dialog
    With CD1
        .CancelError = True
        .DialogTitle = Title
        .Filter = Filter
        
        If (ShowOpen) Then
           Call .ShowOpen
        Else
           Call .ShowSave
        End If
        
        GetDLGName = .FileName
        .FileName = vbNullString
    End With
    
    Exit Function
CanErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Function GetColor() As Long
On Error GoTo CanErr:
    With CD1
        .CancelError = True
        Call .ShowColor
        GetColor = .Color
    End With
    
    Exit Function
CanErr:
    GetColor = -1
End Function


Private Sub ObjectAction(TheObj As Object, ByVal Action As Integer, Button As Integer, ByVal x As Single, ByVal y As Single)
    Select Case Action
        Case 0
            'Mouse down
            If (Button = vbLeftButton) Then
                'Set Mouse Pos
                OldX = (x \ Screen.TwipsPerPixelX)
                OldY = (y \ Screen.TwipsPerPixelY)
                
                Set TheObject = TheObj
                TheObject.MousePointer = vbSizeAll
            End If
        Case 1
            'Mouse move
            If (Button = vbLeftButton) Then
                TheObject.Left = (TheObject.Left + x \ Screen.TwipsPerPixelX) - OldX
                TheObject.Top = (TheObject.Top + y \ Screen.TwipsPerPixelY) - OldY
                
                'Fix Object positions
                If (TheObject.Left < 0) Then TheObject.Left = 0
                If (TheObject.Top < 0) Then TheObject.Top = 0
                
                If (TheObject.Left > pArea.ScaleWidth - TheObject.Width) Then
                    TheObject.Left = (pArea.ScaleWidth - TheObject.Width)
                End If
                
                If (TheObject.Top > pArea.ScaleHeight - TheObject.Height) Then
                    TheObject.Top = (pArea.ScaleHeight - TheObject.Height)
                End If
                'Refresh canvas
                pArea.Refresh
            End If
        Case 2
            'Mouse up
            TheObject.MousePointer = vbDefault
    End Select
    
End Sub

Private Sub chkShowA_Click()
    'Show/Hide Text 1
    lblA(0).Visible = chkShowA.Value
End Sub

Private Sub chkShowB_Click()
    'Show/Hide Text 2
    lblA(1).Visible = chkShowB.Value
End Sub

Private Sub cmdAbout_Click()
    'Show about message
    MsgBox frmmain.Caption & " Version 1.0" & vbCrLf & "An easy way to make a company logo." & _
    vbCrLf & vbCrLf & "The program is freeware" & vbCrLf & vbCrLf & "Please vote if you like this code...", vbInformation, "About"
End Sub

Private Sub cmdExit_Click()
    'Unload program
    Call Unload(frmmain)
End Sub

Private Sub cmdSave_Click()
Dim lFile As String

    lFile = GetDLGName(False, "Save Logo", "Bitmap Files(*.bmp)|*.bmp|")
    
    If Len(lFile) Then
        'Build the logo
        Call BuildLogo
        'Save the picture
        Call SavePicture(picSave.Image, lFile)
        'Destroy picSave
        Set picSave.Picture = Nothing
    End If
    
End Sub

Private Sub Form_Load()
Dim Count As Integer
    
    imgLoc.Path = FixPath(App.Path) & "samples\"
    
    Call chkShowA_Click
    Call chkShowB_Click
    
    'Load font names
    For Count = 0 To Screen.FontCount - 1
        Call lstFonts.AddItem(Screen.Fonts(Count))
    Next Count
    
    'Add font sizes
    For Count = 8 To 100
        Call lstSize.AddItem("Size: " & Count)
        lstSize.ItemData(lstSize.ListCount - 1) = Count
    Next Count
End Sub

Private Sub imgLoc_Click()
    imgLogo.Picture = LoadPicture(FixPath(imgLoc.Path) & imgLoc.FileName)
End Sub

Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ObjectAction(imgLogo, 0, Button, x, y)
End Sub

Private Sub imgLogo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ObjectAction(imgLogo, 1, Button, x, y)
End Sub

Private Sub imgLogo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ObjectAction(imgLogo, 2, Button, x, y)
End Sub

Private Sub lblA_Click(Index As Integer)
    'Highlight the textbox
    TextLine(tIndex).BackColor = vbWhite
    TextLine(Index).BackColor = &HC0FFFF
    tIndex = Index
End Sub

Private Sub lblA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ObjectAction(lblA(Index), 0, Button, x, y)
End Sub

Private Sub lblA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ObjectAction(lblA(Index), 1, Button, x, y)
End Sub

Private Sub lblA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ObjectAction(lblA(Index), 2, Button, x, y)
End Sub

Private Sub lstFonts_Click()
    'Check if the object is a label
    If LCase(TypeName(TheObject)) = "label" Then
        'Set label font name
        TheObject.FontName = lstFonts.Text
    End If
End Sub

Private Sub lstSize_Click()
    'Check if the object is a label
    If LCase(TypeName(TheObject)) = "label" Then
        'Set label font size
        TheObject.FontSize = lstSize.ItemData(lstSize.ListIndex)
    End If
End Sub

Private Sub pColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = vbLeftButton) Then
        'Set picture backcolor
        PicColor.BackColor = pColor(Index).BackColor
        'Check if the object is a label
        If LCase(TypeName(TheObject)) = "label" Then
            'Set label font name
            TheObject.ForeColor = PicColor.BackColor
        End If
        'Refresh canvas
        Call pArea.Refresh
    End If
End Sub

Private Sub PicColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Ret As Long

    If (Button = vbLeftButton) Then
        'Get color form dialog
        Ret = GetColor()
        If (Ret <> -1) Then
            'Set picture backcolor
            PicColor.BackColor = Ret
            'Check if the object is a label
            If LCase(TypeName(TheObject)) = "label" Then
                'Set label font name
                TheObject.ForeColor = PicColor.BackColor
            End If
            'Refresh canvas
            Call pArea.Refresh
        End If
    End If
End Sub

Private Sub TextLine_Change(Index As Integer)
    lblA(Index).Caption = TextLine(Index).Text
End Sub

Private Sub TextLine_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Refresh canvas
    Call pArea.Refresh
End Sub

Private Sub TextLine_LostFocus(Index As Integer)
    If Len(Trim$(TextLine(Index).Text)) = 0 Then
        TextLine(Index).Text = "Your text here"
    End If
End Sub
