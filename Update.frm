VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Update 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Update Norton Antivirus 200X"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7650
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6960
      Top             =   480
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3000
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1059
            MinWidth        =   1059
            Text            =   "Status:"
            TextSave        =   "Status:"
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   9173
            MinWidth        =   9173
            Text            =   "Idle"
            TextSave        =   "Idle"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "09:49"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1766
            MinWidth        =   1766
            TextSave        =   "18.10.2004"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton butCheckVer 
      Caption         =   "C&heck"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtVerLokal 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtVerServer 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin Norton_AV_Updater.FileDownloader FileDownloader 
      Left            =   6120
      Top             =   2760
      _ExtentX        =   1799
      _ExtentY        =   1667
   End
   Begin VB.Frame Frame1 
      Caption         =   "Progress:"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   5895
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton butUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Norton Antivirus Updater (works also after subscription has expired!) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   7245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      Caption         =   "Version Local:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Version on Server:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1350
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This programs checks for new virus definitions for Norton Antivirus.
'The main reason to use it is that it automates the installation of
'virus definition updates - even if the subscription has expired!
'In the near future I plan to release a version which runs in the
'systray and does everything automatically.
'You can treat this programm as BETA.
'The Download Control is written by BelgiumBoy_007

'** (C) THEINT, mail for comments: theint@hotmail.com **


Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public VirSigFile As String
Public updater


Private Sub butCancel_Click()

butCancel.Enabled = False
butUpdate.Enabled = True

    FileDownloader.Cancel  'cancel download

End Sub

Private Sub butCheckVer_Click()

Dim tempfile As String
Dim getVer As String
Dim VirDefInf As String

tempfile = Environ("TEMP") + "\nv_updater.inf"   'get temp. directory
VirDefInf = Environ("CommonProgramFiles") + "\Symantec Shared\VirusDefs\definfo.dat" 'Path to Norton AV VirusDefs
    
    StatusBar.Panels.Item(2) = "Checking for new Virus Definitions..."
    
    FileDownloader.DownloadFile "http://www.symantec.com/avcenter/download/pages/DE-N95.html", tempfile  'Download index file

Open VirDefInf For Input As #1
While Not EOF(1)
    Line Input #1, getVer
    If InStr(1, getVer, "CurDef") Then txtVerLokal = Mid(getVer, 9, 4) + "-" + Mid(getVer, 13, 2) + "-" + Mid(getVer, 15, 2)
Wend
Close #1

Open tempfile For Input As #1
While Not EOF(1)
    Line Input #1, getVer
    If InStr(1, getVer, "http://definitions.symantec.com/defs/") Then
        pos = InStr(1, getVer, "http://definitions.symantec.com/defs/")
        VirSigFile = Mid(getVer, pos, 57)
        txtVerServer = Mid(VirSigFile, 38, 4) + "-" + Mid(VirSigFile, 42, 2) + "-" + Mid(VirSigFile, 44, 2)
    End If
Wend
Close

'compare date on server and local
If CVDate(txtVerLokal) < CVDate(txtVerServer) Then
    StatusBar.Panels.Item(2) = "Update available!"
    butUpdate.Enabled = True
End If

If CVDate(txtVerLokal) >= CVDate(txtVerServer) Then StatusBar.Panels.Item(2) = "No Update available."

End Sub

Private Sub butUpdate_Click()

Dim tempfile As String
tempfile = Environ("TEMP") + "\$na_updater.exe"

butCancel.Enabled = True
butUpdate.Enabled = False

    StatusBar.Panels.Item(2) = "Downloading new Virus Definitions..."

    FileDownloader.DownloadFile "http://securityresponse.symantec.com/avcenter/download/us-files/" + Right(VirSigFile, 20), tempfile

    StatusBar.Panels.Item(2) = "Updating Virus Definitions..."

RunUpdate

End Sub

Private Sub FileDownloader_DowloadComplete()

    StatusBar.Panels.Item(2) = "Idle"
    Progress.Value = 0
    
butCancel.Enabled = False
butUpdate.Enabled = False

End Sub

Private Sub FileDownloader_DownloadErrors(strError As String)
    
    MsgBox strError

butCancel.Enabled = False
butUpdate.Enabled = False

End Sub

Private Sub FileDownloader_DownloadProgress(intPercent As String)

    Progress.Value = intPercent

End Sub
Function FileExists(ByVal sFileName As String) As Boolean

    Dim sFile As String

    On Error Resume Next

    FileExists = False

    sFile = Dir$(sFileName)
    If (Len(sFile) > 0) And (Err = 0) Then
        FileExists = True
    End If

End Function

Sub RunUpdate()

Dim tempfile As String
tempfile = Environ("TEMP") + "\$na_updater.exe"

        updater = Shell(tempfile, vbHide)

curDate = Date
Date = "1990-01-01"

    AppActivate updater
    SendKeys "{ENTER}", True

Sleep 2000
    Date = curDate

Timer.Enabled = True

End Sub

Private Sub Timer_Timer()

Dim check As String

check = Environ("CommonProgramFiles") + "\Symantec Shared\VirusDefs\" 'Path to Norton AV VirusDefs

check = check + Mid(VirSigFile, 65, 8) + ".0" + Mid(VirSigFile, 75, 2) + "\ZDONE.DAT"
Debug.Print check
If FileExists(check) Then
    Timer.Enabled = False
      Sleep 1000
    AppActivate updater, True
    'SendKeys "{ENTER}", True
    'MsgBox "Update completed!", vbInformation, "Update Complete"
    StatusBar.Panels.Item(2) = "Idle"
    butCancel.Enabled = False
End If

End Sub
