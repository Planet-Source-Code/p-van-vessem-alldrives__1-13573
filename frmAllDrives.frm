VERSION 5.00
Begin VB.Form frmAllDrives 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get all drives on your PC          by Peter van Vessem"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5160
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDrives 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmAllDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' AllDrives
' Written by Peter van Vessem
'
' This can become very handy for those who work with
' networks, or if you want to search for a file
' There are a lot of good reason's why you should use this
' simple drive-location-'trick'

' By using 'filesystemobject' you can do several more things,
' like - copy a folder  fs.copyfolder "c:\test1", "a:\test1", true
'                       where 'true' stands for : overwrite the folder
'                       if already exists
'      - fs.createfile
'      - etc...
' Everything you want to do with a file you can do with filesystemobject

' see for your self in the MSDN library for all the possible uses
'   Index, filesystemobject object
' and there it is
'
' I hope you enjoy this very handy tip

Option Explicit
Dim fs, d
Dim strDrives   As String

Private Sub Form_Load()
    Set fs = CreateObject("scripting.filesystemobject")
' d stands for the drive-letter for instance "A:","C:"
' each drive has its own type -> d.drivetype gives you this
' if you look at the following code you can see the explanation
' between double-quote

    For Each d In fs.drives
        Select Case d.drivetype
            Case 0
                strDrives = strDrives & "Unknown     " & d & vbCrLf
            Case 1
                strDrives = strDrives & "Removable   " & d & vbCrLf
            Case 2
                strDrives = strDrives & "Fixed       " & d & vbCrLf
            Case 3
                strDrives = strDrives & "Remote      " & d & vbCrLf
            Case 4
                strDrives = strDrives & "Cdrom       " & d & vbCrLf
            Case 5
                strDrives = strDrives & "Ramdisk     " & d & vbCrLf
        End Select
    Next
    txtDrives.Text = strDrives
End Sub
