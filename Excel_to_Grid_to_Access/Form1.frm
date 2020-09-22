VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Excel to Grid to Access (Database)"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Width           =   1140
   End
   Begin VB.TextBox Textfile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   5175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2040
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
   End
   Begin VB.CommandButton Save_to_database 
      Caption         =   "SAVE TO DATABASE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Load_excel 
      Caption         =   "LOAD EXCEL"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Select Worksheet :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "WorkSheet:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'# Author: Arthur S. Moluñas   #'
'
'# Sañulom Media Corp.         #'
'
'# E-Mail : dsaulom@yahoo.com  #'


'#  Note:
'       To view the transfered data
'       open/check the database
'       The loading of excel to grid may take
'       a second so be patient.
'
'   Have Fun!

Dim XLS As Excel.Application

Dim WBOOK As Excel.Workbook

Dim WSHEET As Excel.Worksheet

Dim RNG As Excel.Range

Dim R As Integer

Dim C As Integer

Dim COUNTER As Integer

Dim cntfileds As Integer

Dim WSheetno As Integer

Dim db As Database

Dim rs As Recordset


Private Sub Combo1_Click()
On Error Resume Next:
    Call load_to_grid
    Save_to_database.Enabled = True
End Sub

Private Sub Load_excel_Click()
On Error Resume Next
'browse xls file
With CommonDialog
    'Set title
    .DialogTitle = "Open Excel Files"
    'Set filename to Null
    .filename = ""
     'Select a filter
    .Filter = "Excel Files (*.xls)" + Chr$(124) + "*.xls" + Chr$(124)
    .ShowOpen
End With
'load filename to text1
Textfile.Text = CommonDialog.filename
'Create a new instance of Excel
Set XLS = CreateObject("Excel.Application")
'Open XLS file
Set WBOOK = XLS.Workbooks.Open(CommonDialog.filename)
    
For WSheetno = 1 To WBOOK.Worksheets.count
   'loads the no. of sheets in combo1
   Combo1.AddItem "Sheet" & (WSheetno)
Next

'close XLS file w/o saving
WBOOK.Close False
'quit excel
XLS.Quit

End Sub

Private Sub load_to_grid()
On Error Resume Next:
If CommonDialog.filename <> "" Then

    Set XLS = CreateObject("Excel.Application")
    Set WBOOK = XLS.Workbooks.Open(CommonDialog.filename)
    'Set the WSHEET variable to the selected worksheet
    Set WSHEET = WBOOK.Worksheets(Combo1.List(Combo1.ListIndex))
    'Get the used range of the current worksheet
    Set RNG = WSHEET.UsedRange
    'load the no. of excel columns to counter
    COUNTER = RNG.Columns.count
    
    'Configure the grid to display data
    MSFlexGrid.Clear
    MSFlexGrid.FixedCols = 0
    MSFlexGrid.FixedRows = 0
    MSFlexGrid.Cols = RNG.Columns.count
    MSFlexGrid.Rows = RNG.Rows.count
      
    'loads data of XLS file to the grid
    For R = 0 To MSFlexGrid.Rows - 1
        MSFlexGrid.Row = R
        For C = 0 To MSFlexGrid.Cols - 1
            MSFlexGrid.Col = C
            MSFlexGrid.Text = WSHEET.Cells(R + 1, C + 1).Value
        Next
    Next
    'close XLS file w/o saving
    WBOOK.Close False
    'quit excel
    XLS.Quit
    
End If
End Sub

Private Sub Save_to_database_Click() 'save to database

On Error Resume Next:
'adds recordset equal to excel column
For cntfields = 1 To COUNTER
    db.Execute ("alter table TEST_TABLE " _
    & "add column " & "A" & cntfields & " text")
Next cntfields


Data1.Refresh
db.Recordsets.Refresh
'open table
Set rs = db.OpenRecordset("TEST_TABLE", dbOpenDynaset)

       For R = 0 To MSFlexGrid.Rows - 1
            MSFlexGrid.Row = R
          
            rs.AddNew 'adds data from grid to the database
            For C = 0 To MSFlexGrid.Cols - 1
                MSFlexGrid.Col = C
                rs.Fields(C) = MSFlexGrid.Text
            Next
            rs.Update
         
        Next

    MsgBox "RECORD SAVE", vbInformation, "Save"
    Save_to_database.Enabled = False
    
End Sub

Private Sub Form_Load()
On Error Resume Next:
'initialize database
Set db = OpenDatabase(App.Path & "\TEST_DB.mdb")
'delete test_table from database
db.Execute ("DROP TABLE TEST_TABLE")
Data1.Refresh
'creates new table named as test_table
db.Execute ("CREATE TABLE TEST_TABLE (A1 TEXT)")
Data1.Refresh
End Sub

