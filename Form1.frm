VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MIDAR's Glowing LCD Font"
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerScript 
      Interval        =   10
      Left            =   270
      Top             =   2160
   End
   Begin VB.PictureBox pictUserInput 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10170
      TabIndex        =   2
      Top             =   6975
      Visible         =   0   'False
      Width           =   10170
      Begin VB.TextBox txtInput 
         Height          =   315
         Left            =   840
         MaxLength       =   256
         TabIndex        =   0
         Top             =   60
         Width           =   3375
      End
      Begin VB.CommandButton btnRefresh 
         Caption         =   "&Refresh"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   315
         Left            =   4290
         TabIndex        =   1
         Top             =   60
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImageListLow 
      Left            =   270
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   15
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   47
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "ascii=32"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":02EF
            Key             =   "ascii=92"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0647
            Key             =   "ascii=45"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0933
            Key             =   "ascii=40"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C8D
            Key             =   "ascii=41"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FC7
            Key             =   "ascii=61"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1327
            Key             =   "ascii=37"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16A2
            Key             =   "ascii=58"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19A6
            Key             =   "ascii=44"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C73
            Key             =   "ascii=47"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1FC3
            Key             =   "ascii=46"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":227B
            Key             =   "ascii=48"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25F8
            Key             =   "ascii=49"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":291C
            Key             =   "ascii=50"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C88
            Key             =   "ascii=51"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2FE8
            Key             =   "ascii=52"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3336
            Key             =   "ascii=53"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":36A1
            Key             =   "ascii=54"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3A19
            Key             =   "ascii=55"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3D5A
            Key             =   "ascii=56"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":40D8
            Key             =   "ascii=57"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4448
            Key             =   "ascii=65"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":47BD
            Key             =   "ascii=66"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4B39
            Key             =   "ascii=67"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4E8B
            Key             =   "ascii=68"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":520B
            Key             =   "ascii=69"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5569
            Key             =   "ascii=70"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":58B9
            Key             =   "ascii=71"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5C24
            Key             =   "ascii=72"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5F89
            Key             =   "ascii=73"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":629A
            Key             =   "ascii=74"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":65EF
            Key             =   "ascii=75"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6965
            Key             =   "ascii=76"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6C96
            Key             =   "ascii=77"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":700D
            Key             =   "ascii=78"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7375
            Key             =   "ascii=79"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":76F1
            Key             =   "ascii=80"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7A4E
            Key             =   "ascii=81"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7DC9
            Key             =   "ascii=82"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8142
            Key             =   "ascii=83"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":84B3
            Key             =   "ascii=84"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":87ED
            Key             =   "ascii=85"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8B5C
            Key             =   "ascii=86"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8EB3
            Key             =   "ascii=87"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9220
            Key             =   "ascii=88"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9599
            Key             =   "ascii=89"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9900
            Key             =   "ascii=90"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListHigh 
      Left            =   870
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   15
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   49
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9C64
            Key             =   "ascii=46"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F4B
            Key             =   "ascii=32"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A23A
            Key             =   "ascii=40"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A593
            Key             =   "ascii=41"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A8DD
            Key             =   "ascii=45"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ABD4
            Key             =   "ascii=91"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AF38
            Key             =   "ascii=93"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B283
            Key             =   "ascii=61"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B5EC
            Key             =   "ascii=48"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B962
            Key             =   "ascii=49"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BCA9
            Key             =   "ascii=50"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C021
            Key             =   "ascii=51"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C38D
            Key             =   "ascii=52"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C6EC
            Key             =   "ascii=53"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CA63
            Key             =   "ascii=54"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CDE6
            Key             =   "ascii=55"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D13D
            Key             =   "ascii=56"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D4BC
            Key             =   "ascii=57"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D831
            Key             =   "ascii=92"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DB93
            Key             =   "ascii=58"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DEA4
            Key             =   "ascii=44"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E18B
            Key             =   "ascii=47"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E4EC
            Key             =   "ascii=37"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E861
            Key             =   "ascii=65"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EBDB
            Key             =   "ascii=66"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EF4E
            Key             =   "ascii=67"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F2B9
            Key             =   "ascii=68"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F629
            Key             =   "ascii=69"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F990
            Key             =   "ascii=70"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FCEA
            Key             =   "ascii=71"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10065
            Key             =   "ascii=72"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":103D6
            Key             =   "ascii=73"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10713
            Key             =   "ascii=74"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10A73
            Key             =   "ascii=75"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10DED
            Key             =   "ascii=76"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1113A
            Key             =   "ascii=77"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":114B2
            Key             =   "ascii=78"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11821
            Key             =   "ascii=79"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11B97
            Key             =   "ascii=80"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11F09
            Key             =   "ascii=81"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1228C
            Key             =   "ascii=82"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1260E
            Key             =   "ascii=83"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12989
            Key             =   "ascii=84"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12CED
            Key             =   "ascii=85"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13059
            Key             =   "ascii=86"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":133BB
            Key             =   "ascii=87"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13733
            Key             =   "ascii=88"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13AAB
            Key             =   "ascii=89"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13E24
            Key             =   "ascii=90"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub PrintVFT(strText As String, objForm As Form, Top As Single, Left As Single, blnGoSlow As Boolean)

    ' ==========================================================
    ' Prints out Vacuum Fluorescent Display to X/Y co-ordinates.
    '
    ' Things to do...
    '   * Create multi-coloured displays (red, green, yellow)
    '   * Multiple Sizes
    '   * Different Fonts, ie. Dot Matrix style font.
    ' ==========================================================
    
    On Error GoTo errTrap
    
    ' Basic error checking.
    If objForm Is Nothing Then Exit Sub
    If strText = "" Then Exit Sub
    
    ' Declare variables.
    Dim intN As Integer
    Dim strSingleCharacter As String
    Dim intASCIIValue As Integer
    Dim strImageKey As String
    Dim sngLeftPosition As Single
    
    ' Convert the text to upper case.
    strText = UCase(strText)
    
    ' Loop through all characters in the text box.
    For intN = 1 To Len(strText)
        
        strSingleCharacter = Mid(strText, intN, 1)
        
        intASCIIValue = Asc(strSingleCharacter)
        
        strImageKey = "ascii=" & intASCIIValue
        
        sngLeftPosition = Left + (intN * ImageListHigh.ImageWidth * Screen.TwipsPerPixelX)
        
        ImageListHigh.ListImages.Item(strImageKey).Draw objForm.hDC, sngLeftPosition, Top, imlTransparent
        
        If blnGoSlow = True Then
            ' Make not work properly with Timer Controls and DoEvent loops.
            Call Sleep(25)
            Me.Refresh
        End If
        
    Next intN
    
    ' Increment Y position down.
    objForm.CurrentY = objForm.CurrentY + (objForm.ImageListHigh.ImageHeight * Screen.TwipsPerPixelY)
    
    Exit Sub
errTrap:
    
    ' Convert any unknown character into a space.
    ImageListHigh.ListImages.Item("ascii=32").Draw objForm.hDC, sngLeftPosition, Top, imlTransparent
    Resume Next
    
End Sub

Private Sub btnRefresh_Click()

    ImageListHigh.MaskColor = vbBlack
    Me.AutoRedraw = True
        
    Call PrintVFT(Me.txtInput.Text, Me, Me.CurrentY, 0, True)
    
    Me.txtInput.SelStart = 0
    Me.txtInput.SelLength = Len(Me.txtInput.Text)
    Me.txtInput.SetFocus
        
    Me.Refresh
    
    If Me.txtInput.Text = "QUIT" Or Me.txtInput.Text = "EXIT" Then
        Call Sleep(750)
        Call PrintVFT("GOODBYE......", Me, Me.CurrentY, 0, True)
        Call Sleep(750)
        Unload Me
    End If
    
End Sub


Private Sub Form_Resize()

    Me.btnRefresh.Left = Me.ScaleWidth - Me.btnRefresh.Width
    
    Me.txtInput.Left = 0
    Me.txtInput.Width = Me.ScaleWidth - Me.btnRefresh.Width - 60
    
End Sub

Private Sub TimerScript_Timer()

    Static s_lngScriptCounter As Long
    
    s_lngScriptCounter = s_lngScriptCounter + 1
    
    Select Case s_lngScriptCounter
        Case 100
            Call PrintVFT(App.ProductName, Me, Me.CurrentY, Me.CurrentX, True)
            Call PrintVFT(" ", Me, Me.CurrentY, Me.CurrentX, True)
            
        Case 150
            Call PrintVFT("by", Me, Me.CurrentY, Me.CurrentX, True)
            Call PrintVFT(" ", Me, Me.CurrentY, Me.CurrentX, True)
            
        Case 200
            Call PrintVFT(App.CompanyName, Me, Me.CurrentY, Me.CurrentX, True)
            Call PrintVFT(" ", Me, Me.CurrentY, Me.CurrentX, True)
        
        Case 400
            Call PrintVFT(App.LegalCopyright, Me, Me.CurrentY, Me.CurrentX, True)
        
        Case 500
            Call PrintVFT("Visit my web site today: www.midar.com.au", Me, Me.CurrentY, Me.CurrentX, True)
        
        Case 600
            Call PrintVFT(" ", Me, Me.CurrentY, Me.CurrentX, True)
                
        Case 900
            Me.Cls
            
        Case 1000
            Call PrintVFT("you type at the bottom of this screen", Me, Me.CurrentY, Me.CurrentX, True)
        
        Case 1200
            Me.pictUserInput.Visible = True
            
            
        Case 1400
            Call PrintVFT("To exit this application simply type: exit", Me, Me.CurrentY, Me.CurrentX, True)
                        
        Case 1600
            Me.txtInput.Text = "EXIT"
            
        Case 1800
            Call PrintVFT("Now it is your turn...", Me, Me.CurrentY, Me.CurrentX, True)
            Me.txtInput.Text = ""
            Me.btnRefresh.Enabled = True
            Me.txtInput.SetFocus
            
        Case 2200
            Me.Cls
            
        Case Is > 3000
            TimerScript.Enabled = False
            s_lngScriptCounter = 0
            
    End Select
    
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


