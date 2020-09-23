VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Updater 1.0 by nerphed"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox list 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3930
      ItemData        =   "frmUpdate.frx":0000
      Left            =   120
      List            =   "frmUpdate.frx":0002
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   3000
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   4480
      Width           =   1095
   End
   Begin VB.Label latver 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label curver 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblConnect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   4480
      Width           =   1095
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()

End Sub

Private Sub Label1_Click()
frmAbout.Show
End Sub

Private Sub lblConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblConnect.BackColor = &HC0C0C0
End Sub

Private Sub lblConnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblConnect.BackColor = &H808080
Dim St As String
Dim fSt As String
Dim i As Long
i = 0
list.AddItem "Retrieving version info."
DownloadFile "version.dat", "vers.dat"
Open "vers.dat" For Input As #21
    Input #21, St
    latver = St
    list.AddItem "Latest version is " & St
    If CurrentVersion = latver Then
        list.AddItem "You have the latest version."
    Else
        list.AddItem "Retrieving file list."
        DownloadFile "files.dat", "update.dat"
        DoEvents
        list.AddItem "Comparing file lists."
        DoFiles
        DoEvents
        Open "update.dat" For Input As #20
            Do Until EOF(20)
                Line Input #20, fSt
                GetSections fSt, ","
                CheckFile Section(1), Section(2)
            Loop
        Close #20
        CreateFileList
        If Dir("version.dat") <> "" Then Kill ("version.dat")
        Open "version.dat" For Output As #24
            Print #24, latver
        Close #24
        list.AddItem "Version has been updated to " & latver & "."
        curver = "Version " & latver
        list.AddItem "Update complete."
    End If
    DoEvents
Close #21
If Dir("vers.dat") <> "" Then Kill "vers.dat"
If Dir("update.dat") <> "" Then Kill "update.dat"
End Sub
