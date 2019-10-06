VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLevelPrint 
   BackColor       =   &H00C0B4A9&
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmLevelPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdPTag 
      Caption         =   "Print Tag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Patient Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox cmbYear 
         Height          =   315
         ItemData        =   "frmLevelPrint.frx":08CA
         Left            =   7080
         List            =   "frmLevelPrint.frx":08CC
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtPName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox txtRefdby 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1800
         Width           =   8175
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cmbPatientID 
         Height          =   315
         ItemData        =   "frmLevelPrint.frx":08CE
         Left            =   1440
         List            =   "frmLevelPrint.frx":08D0
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtCMid 
         Height          =   495
         Left            =   7080
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker BillDate 
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   11730947
         CurrentDate     =   41788
      End
      Begin VB.Label lblPatientID 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Patient ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblBillDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Bill Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblPName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblRefdby 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Refd. by"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Patient Age"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblSex 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Investigation Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   15015
      Begin MSComctlLib.ListView List1 
         Height          =   3855
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6800
         View            =   3
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   12632256
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Sticker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   7560
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdodcSub 
      Height          =   375
      Left            =   7320
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   375
      Left            =   7320
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLevelPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs                              As New ADODB.Recordset
Private rscashmaster                    As New ADODB.Recordset
Private rsCashDetail                    As ADODB.Recordset
Private rsCashDetail1                   As ADODB.Recordset
Private rsATable                        As New ADODB.Recordset
Private rsAMaster                       As New ADODB.Recordset
Private rsCustomerMaster                As New ADODB.Recordset
Dim str                                 As String
Dim Tracer                              As Integer
Private rsTemp2                         As ADODB.Recordset
Dim flagSlNo                            As Integer
Dim strMood                             As String


Private rsRptRtn                        As New ADODB.Recordset
Private rsRptRtn1                       As New ADODB.Recordset
''---------------------------------------------------------------------------
''----Add For Reporting Perpose----------------------------------------------
Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition


Private objReportSub                        As CRPEAuto.Report 'sub
Private objReportDatabaseSub                As CRPEAuto.Database 'sub
Private objReportDatabaseTablesSub          As CRPEAuto.DatabaseTables 'sub
Private objReportDatabaseTableSub           As CRPEAuto.DatabaseTable 'sub
Private objReportFormulaFieldDefinationsSub    As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                         As CRPEAuto.FormulaFieldDefinition


Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
Private rsDailyRpt                          As ADODB.Recordset
'Private Tracer                              As Integer
Private strGroupName                        As String
Dim temp As Double
Dim temp1 As Double
Dim temp2 As Double
''--------------------------------------------------------------------------------



Private Sub cmbPatientID_Click()
AdodcMain.ConnectionString = gstrConnection
'     Call Srch_Pat_ID

AdodcMain.RecordSource = "select [Bill No], [Client Name], Age, Consultant, [Date], Sex  from Cash_Memo_Investigation where CMId='" + txtCMid.text + "' and Year= '" & cmbYear & "' "
'AdodcMain.RecordSource = "select [Bill No], [Client Name], Age, RefdBy, [Date], Sex  from Cash_Memo_Investigation where CMId='" + txtCMid.text & "' "
    AdodcMain.Refresh
    
    If AdodcMain.Recordset.RecordCount > 0 Then
        cmbPatientID.text = AdodcMain.Recordset![Bill No]
        txtPName.text = AdodcMain.Recordset![Client Name]
        txtAge.text = AdodcMain.Recordset!Age
        txtRefdby.text = AdodcMain.Recordset!RefdBy
        BillDate = AdodcMain.Recordset!Date
        txtSex.text = AdodcMain.Recordset!Sex
'        txtKOT = IIf(IsNull(rscashmaster!Kot), "", rscashmaster!Kot)
    
    End If
   Call search
   Call Sex
'End If
End Sub

Private Sub cmbPatientID_GotFocus()
cmbPatientID.BackColor = &HFFFFC0
Call pat_id1
End Sub

Private Sub cmbPatientID_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
 
 
    AdodcMain.ConnectionString = gstrConnection
     Call Srch_Pat_ID

'AdodcMain.RecordSource = "select [Bill No], [Client Name], Age, Consultant, [Date], Sex  from Cash_Memo_Investigation where CMId='" + txtCMid.text + "' and Year= '" & cmbYear & "' "
AdodcMain.RecordSource = "select [Bill No], [Client Name], Age, RefdBy, [Date], Sex  from Cash_Memo_Investigation where CMId='" + txtCMid.text & "' "
    AdodcMain.Refresh
    
    If AdodcMain.Recordset.RecordCount > 0 Then
        cmbPatientID.text = AdodcMain.Recordset![Bill No]
        txtPName.text = AdodcMain.Recordset![Client Name]
'        txtAge.text = AdodcMain.Recordset!Age
        txtRefdby.text = AdodcMain.Recordset!RefdBy
        BillDate = AdodcMain.Recordset!Date
        txtSex.text = AdodcMain.Recordset!Sex
    
    End If
   Call search
   Call Sex
End If
'cmdPrint.SetFocus
End Sub

Private Sub Sex()
'If txtSex.text = False Then
'txtSex.text = "Male"
'Else
'txtSex.text = "Female"
''
' End If
End Sub

Private Sub Srch_Pat_ID()

    Dim My_Rst As New ADODB.Recordset
    Dim IntPat_ID As Double
    
'    My_Rst.Open "Select CMId, [Bill No]  from Cash_Memo_Investigation Where [Bill No]='" + cmbPatientID.text + "' and Year= '" & cmbYear & "'", cn, adOpenStatic, adLockReadOnly
    My_Rst.Open "Select CMId, [Bill No]  from Cash_Memo_Investigation", cn, adOpenStatic, adLockReadOnly
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!CMid
    End If
    txtCMid.text = IntPat_ID
End Sub

Private Sub cmbPatientID_LostFocus()
On Error Resume Next

    If cmbPatientID = "" Then Exit Sub
    If cmbPatientID <> "" Then
        cmbPatientID.TabStop = False
    End If
End Sub

'Private Sub cmbYear_Change()
'
'End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()

Dim i As Integer
cn.Execute "Delete from Level"
    While i <> List1.ListItems.Count
      

            i = i + 1
            
            If List1.ListItems(i).Checked = True Then
            With List1.ListItems.Item(i)
            cn.Execute "insert into Level(Particulars, Rate,Name, CMid) " & _
                       " Values('" & .text & "'," & .SubItems(1) & ",'" & .SubItems(2) & "','" + txtCMid.text + "') "
                  
    End With
       
     End If
     
     Tracer = 0

    Wend
    
    Call FetchData
    Call PrintSticker
    Call FetchData1
'    Call PrintTag
End Sub

Private Sub cmdPrint_Click()
Dim i As Integer
cn.Execute "Delete from Level"
    While i <> List1.ListItems.Count
      

            i = i + 1
            
             If List1.ListItems(i).Checked = True Then
                With List1.ListItems.Item(i)
                   cn.Execute "insert into Level(Particulars, Rate, Name,CMid) " & _
                       " Values('" & .text & "'," & .SubItems(1) & ",'" & .SubItems(2) & "','" + txtCMid.text + "') "
            
    End With
       
     End If
     
          Tracer = 1
     
     Wend
       
    Call FetchData
    Call PrintSticker

End Sub

Public Sub PrintSticker()
    If rsRptRtn.RecordCount < 0 Then
        MsgBox "No Data found for Print ", vbInformation
        Exit Sub
    End If
    
    Set objReportApp = CreateObject("Crystal.CRPE.Application")
    Set objReport = objReportApp.OpenReport(App.Path & "\Print Sticker.rpt")
'    Set objReport = objReportApp.OpenReport(App.Path & "\reports\Print Sticker.rpt")
    
    Set objReportDatabase = objReport.Database
    Set objReportDatabaseTables = objReportDatabase.Tables
    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = objReport.PrintWindowOptions

    ObjPrinterSetting.HasPrintSetupButton = True
    ObjPrinterSetting.HasRefreshButton = True
    ObjPrinterSetting.HasSearchButton = True
    ObjPrinterSetting.HasZoomControl = True

    objReportDatabaseTable.SetPrivateData 3, rsRptRtn
    objReport.DiscardSavedData
    If Tracer = 0 Then
    objReport.Preview "Sticker Printing", , , , , 16777216 Or 524288 Or 65536
    Else
    objReport.PrintOut (False)
    
    Set objReport = Nothing
    Set objReportDatabase = Nothing
    Set objReportDatabaseTables = Nothing
    Set objReportDatabaseTable = Nothing
    
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Bank Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Bank Information Report"
    End Select
    End If
End Sub


Private Sub cmdPTag_Click()
Dim i As Integer
cn.Execute "Delete from Level"
    While i <> List1.ListItems.Count
      

            i = i + 1
            
             If List1.ListItems(i).Checked = True Then
                With List1.ListItems.Item(i)
                   cn.Execute "insert into Level(Particulars, Rate, Name,CMid) " & _
                       " Values('" & .text & "'," & .SubItems(1) & ",'" & .SubItems(2) & "','" + txtCMid.text + "') "
            
    End With
       
     End If
    Tracer = 1
    
    Wend
    
    Call FetchData
    Call PrintTag

End Sub

Public Sub PrintTag()
    If rsRptRtn.RecordCount < 0 Then
        MsgBox "No Data Found for Print ", vbInformation
        Exit Sub
    End If
    
    Set objReportApp = CreateObject("Crystal.CRPE.Application")
    Set objReport = objReportApp.OpenReport(App.Path & "\Level Tag.rpt")
'    Set objReport = objReportApp.OpenReport(App.Path & "\reports\Level Tag.rpt")
    
    Set objReportDatabase = objReport.Database
    Set objReportDatabaseTables = objReportDatabase.Tables
    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = objReport.PrintWindowOptions

    ObjPrinterSetting.HasPrintSetupButton = True
    ObjPrinterSetting.HasRefreshButton = True
    ObjPrinterSetting.HasSearchButton = True
    ObjPrinterSetting.HasZoomControl = True

    objReportDatabaseTable.SetPrivateData 3, rsRptRtn1
    objReport.DiscardSavedData
    If Tracer = 0 Then
    objReport.Preview "Tag Printing", , , , , 16777216 Or 524288 Or 65536
    Else
    objReport.PrintOut (False)
    
    Set objReport = Nothing
    Set objReportDatabase = Nothing
    Set objReportDatabaseTables = Nothing
    Set objReportDatabaseTable = Nothing
    
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Bank Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Bank Information Report"
    End Select
    
    End If
    
End Sub




Private Sub Form_Load()
Call Connect
   ModFunction.StartUpPosition Me
   
        List1.ColumnHeaders.Add , , "Particulars", 6000
        List1.ColumnHeaders.Add , , "Rate"
        List1.ColumnHeaders.Add , , "Department", 4500
        List1.ColumnHeaders.Add , , "CMid", 1000
          
   
   Call pat_id1
   BillDate.Value = Date
   cmbYear.Clear
   cmbYear.AddItem "2015"
   Me.cmbYear.ListIndex = 0
   cmbYear.AddItem "2014"
   cmbYear.AddItem "2013"
   cmbYear.AddItem "2012"
   cmbYear.AddItem "2011"
   
'   Call Srch_Pat_ID
   
End Sub

Private Sub pat_id1()


    Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT TOP 50 [Bill No] FROM Cash_Memo_Investigation where Posted=1 and Year= '" & cmbYear & "' ORDER BY [Bill No] DESC"), cn, adOpenStatic
      cmbPatientID.Clear
    While Not rsTemp2.EOF
        cmbPatientID.AddItem rsTemp2("Bill No")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub

Private Sub search()
On Error Resume Next
Dim strSQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim i As Integer
    
'   rs.Open "select (select Particulars,Rate,Name_Id from Cash_Memo_Sub_Investigation where CMId='" + txtCMid.text + "'", cn, adOpenStatic, adLockReadOnly
 rs.Open "select (select Name=isnull(Name,'') from Inv_Report_type " & _
         "Where Inv_Report_type.Type_Id = Cash_Memo_Sub_Investigation.Department_Id " & _
         "and CMId='" + txtCMid.text + "') as Name,Particulars,Rate from Cash_Memo_Sub_Investigation where CMId='" + txtCMid.text + "'", cn, adOpenStatic, adLockReadOnly
        
        Me.List1.ListItems.Clear
           
    If rs.RecordCount <> 0 Then

          Do Until rs.EOF
               With List1.ListItems.Add
                        .text = rs("Particulars")
                        .SubItems(1) = rs("Rate")
                        .SubItems(2) = rs("Name")
'                        .SubItems(3) = rs("DOB")
              End With
           rs.MoveNext

        Loop

    End If
       
    rs.Close
        
End Sub


Public Function FetchData()
    Set rsRptRtn = New ADODB.Recordset
                  
    rsRptRtn.Open "select Cash_Memo_Investigation.[Bill No], Cash_Memo_Investigation.[Client Name], " & _
                      "Cash_Memo_Investigation.Age, " & _
                      "Cash_Memo_Investigation.Consultant, Cash_Memo_Investigation.Date, " & _
                      "Cash_Memo_Investigation.Sex, Level.Particulars, " & _
                      "Level.Rate,Level.CMid  from Cash_Memo_Investigation  INNER JOIN " & _
                      "Level on Cash_Memo_Investigation.CMId = Level.CMid " & _
                      "WHERE [Level].Name = 'Biochemical Analysis' OR " & _
                      "Level.Name = 'Haematological' OR " & _
                      "Level.Name = 'Immuonolgy' OR " & _
                      "Level.Name = 'Serological' OR " & _
                      "Level.Name = 'Clinical Pathological Report' OR " & _
                      "Level.Name = 'Urine Examination' OR " & _
                      "Level.Name = 'Stool Examination' OR " & _
                      "Level.Name = 'Microbiology' OR " & _
                      "Level.Name = 'Ultrasonogram' OR " & _
                      "Level.Name = 'Hormone Analysis'", cn, adOpenStatic, adLockReadOnly
'End If

End Function


Public Function FetchData1()
    Set rsRptRtn1 = New ADODB.Recordset
                
    rsRptRtn1.Open "select Cash_Memo_Investigation.[Bill No], Cash_Memo_Investigation.[Client Name], " & _
                      "Cash_Memo_Investigation.Age, " & _
                      "Cash_Memo_Investigation.Consultant, Cash_Memo_Investigation.Date, " & _
                      "Cash_Memo_Investigation.Sex, Level.Particulars, " & _
                      "Level.Rate,Level.CMid  from Cash_Memo_Investigation  INNER JOIN " & _
                      "Level on Cash_Memo_Investigation.CMId = Level.CMid " & _
                      "WHERE [Level].Name = 'DIGITAL X-RAY' OR " & _
                      "Level.Name = 'CT SCAN' OR " & _
                      "Level.Name = 'ECG' OR " & _
                      "Level.Name = 'Echocardiogram' OR " & _
                      "Level.Name = 'Memmography' OR " & _
                      "Level.Name = 'ULTRASONOGRAM'", cn, adOpenStatic, adLockReadOnly

End Function



