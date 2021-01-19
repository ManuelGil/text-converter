VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text Converter"
   ClientHeight    =   2520
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   4095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOriginal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox txtEncoded 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMain.frx":000C
      Left            =   1200
      List            =   "frmMain.frx":002E
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Encoded:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Original:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Convert:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This method catchs the click event of the "cmdConvert" button.

Private Sub cmdConvert_Click()
    ' In case of error go to "EndSubroutine" label.
    On Error GoTo EndSubroutine
    
    ' Evaluate the item selected.
    Select Case Combo1.ListIndex
    
    Case 0
        ' Convert Hexadecimal to Decimal.
        txtEncoded.Text = HEXToDEC(txtOriginal.Text)
    
    Case 1
        ' Convert Hexadecimal to Decimal.
        txtEncoded.Text = DECToHEX(txtOriginal.Text)
    
    Case 2
        ' Transform Character set to Hexadecimal equivalent.
        txtEncoded.Text = HEXEncode(txtOriginal.Text)
    
    Case 3
        ' Transform Hexadecimal set to Character equivalent.
        txtEncoded.Text = HEXDecode(txtOriginal.Text)
    
    Case 4
        ' Transform Character set to Decimal equivalent.
        txtEncoded.Text = DECEncode(txtOriginal.Text)
    
    Case 5
        ' Transform Decimal set to Character equivalent.
        txtEncoded.Text = DECDecode(txtOriginal.Text)
    
    Case 6
        ' Transform ANSI Text to Multibyte Text.
        txtEncoded.Text = UTF8_Encode(txtOriginal.Text)
    
    Case 7
        ' Transform Multibyte Text to ANSI Text.
        txtEncoded.Text = UTF8_Decode(txtOriginal.Text)
    
    Case 8
        ' Transform Character set to Java notation equivalent.
        txtEncoded.Text = JAVAEncode(txtOriginal.Text)
    
    Case 9
        ' Transform Character set to Visual Basic notation equivalent.
        txtEncoded.Text = VBEncode(txtOriginal.Text)
    
    End Select
    
EndSubroutine:

End Sub

' This method catch the double click event of the "txtOriginal" input.

Private Sub txtOriginal_DblClick()
    ' Clear the input.
    txtOriginal.Text = ""
End Sub

' This method catch the click event of the "mnuFileNew" menu.

Private Sub mnuFileNew_Click()
    ' Define a nwe instance of Main Form.
    Dim fMainForm As New frmMain
    ' Show the new instance.
    fMainForm.Show
End Sub

' This method catch the click event of the "mnuFileQuit" menu.

Private Sub mnuFileQuit_Click()
    ' Ends the program.
    End
End Sub

' This method catch the click event of the "mnuHelpAbout" menu.

Private Sub mnuHelpAbout_Click()
    ' Show the About Form.
    frmAbout.Show vbModal
End Sub

