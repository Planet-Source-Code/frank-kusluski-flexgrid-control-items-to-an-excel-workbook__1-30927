VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Visit My Web Page of Shareware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   " Visit my web page at http://ic.net/~kusluski to download many useful programs "
      Top             =   3720
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Download Excel OCX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   " Free download of Excel OCX! A powerful ActiveX control for exchanging data between VB and Excel via COM technology "
      Top             =   3120
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4048
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Copy FlexGrid Contents to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: Frank Kusluski
'Date Written: 1/18/02
'Please direct any questions/concerns to me at kusluski@mail.ic.net

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Private Sub Command1_Click()
Dim i As Long
Dim n As Long
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number Then
   Err.Clear
   Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add
AppActivate "FlexGrid To Excel Demo"
For i = 0 To 3
    MSFlexGrid1.Row = i
    For n = 0 To 3
        MSFlexGrid1.Col = n
        objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = MSFlexGrid1.Text
    Next
Next
End Sub

Private Sub Command2_Click()
Dim i As Long
i = ShellExecute(Form1.HWnd, "open", "http://download.com.com/3000-2401-10105891.html?tag=lst-0-9", vbNullString, vbNullString, 1)
End Sub

Private Sub Command3_Click()
Dim i As Long
i = ShellExecute(Form1.HWnd, "open", "http://ic.net/~kusluski", vbNullString, vbNullString, 1)
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim n As Integer
Me.Caption = "FlexGrid To Excel Demo"
'Populate the FlexGrid with sample data
With MSFlexGrid1
     .Rows = 1
     .Cols = 4
     'Add field headers
     For i = 1 To 3
         .Col = i
         .Text = "Heading " & i
     Next
     'Add data
     For i = 1 To 3
         .Rows = .Rows + 1
         .Row = .Rows - 1
         .Col = 0
         .Text = "Record " & i
         For n = 1 To 3
             .Col = n
             .Text = "Row " & i & ",Col " & n
         Next
     Next
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objWorkbook = Nothing
Set objExcel = Nothing
End Sub
