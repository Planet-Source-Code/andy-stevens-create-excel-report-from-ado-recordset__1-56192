VERSION 5.00
Begin VB.Form frmExcelExportTest 
   Caption         =   "Export To Excel Test"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export To Excel"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmExcelExportTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare module level variables
Private WithEvents m_objExport As clsExcelExport
Attribute m_objExport.VB_VarHelpID = -1

Private Sub cmdClose_Click()

    'Close the form
    Unload Me
    
End Sub

Private Sub cmdExport_Click()

'Declare local variables
Dim cnnADO      As New ADODB.Connection
Dim rstData     As New ADODB.Recordset
Dim strCustomer As String
Dim strSQL      As String
Dim strFileName As String

On Error Resume Next

    'Set up the connection
    With cnnADO
        .ConnectionString = "Provider=SQLOLEDB.1;" & _
                            "Password=;" & _
                            "Persist Security Info=True;" & _
                            "User ID=sa;" & _
                            "Initial Catalog=Northwind;" & _
                            "Data Source=MyServer"
        .Open
    End With
    
    'chose the customer
    strCustomer = "VINET"
    
    'Build the SQL
    strSQL = "SELECT OrderID AS [Order ID], " & _
                    "CustomerID AS Customer, " & _
                    "OrderDate AS [Date Ordered], " & _
                    "RequiredDate AS [Date Required], " & _
                    "ShippedDate AS [Date Shipped], " & _
                    "Freight AS [Freight Cost], " & _
                    "ShipName + ', ' + ShipAddress + ', ' + " & _
                    "ShipCity  + ', ' + ShipPostalCode + ', ' + " & _
                    "ShipCountry AS [Shipping Address]" & _
             "FROM   Orders " & _
             "WHERE  Orders.CustomerID = '" & strCustomer & "'"

    'Create the recordset
    With rstData
        .ActiveConnection = cnnADO
        .Source = strSQL
        .CursorLocation = adUseClient
        .Open
    End With

    'Set the filename
    strFileName = "C:\NWindOrders.xls"
    
    'Create the Excel object
    Set m_objExport = New clsExcelExport
    
    With m_objExport
        
        'General
        .FilePath = strFileName
        .OverWriteExistingFile = True
        .WorkSheetName = strCustomer & " Orders"
        
        'Headers
        .LeftHeader = "NorthWind Inc"
        .CentreHeader = "Orders for " & strCustomer
        .RightHeader = "NorthWind Inc"
        .HeaderRow = 1
        .ShowHeaderRow = True
        .BoldHeaderRow = True
        .ShadeHeaderRow = True
        .RepeatHeaderRow = True
                
        'Column Formats
        .FormatColumn(3) = LongDate
        .FormatColumn(4) = LongDate
        .FormatColumn(5) = LongDate
        .FormatColumn(6) = Money
        
        'Column Totals
        .TotalColumn(6) = True
        .BoldTotals = True
        
        'Footers
        .LeftFooter = FileNameAndPath
        .CentreFooter = PageNumberAndTotal
        .RightFooter = DateAndTimePrinted
        
        'Page Setup
        .AddBorders = True
        .PaperSize = A4
        .Orientation = Landscape
        .CentreHorizontally = True
        .CentreVertically = False
        .FitToOnePage = False
        
        'Create the report
        If .ExportRecordSetToExcel(rstData) Then
            MsgBox "Report exported to " & strFileName
        End If
    
    End With
    
    'Close the objects
    rstData.Close
    cnnADO.Close
    
    'Kill the objects
    Set rstData = Nothing
    Set cnnADO = Nothing
    Set m_objExport = Nothing

End Sub

Public Sub m_objExport_FileExists(ByVal v_strMessage As String)

    'Display the message
    MsgBox v_strMessage, vbExclamation + vbOKOnly, App.Title
    
End Sub
