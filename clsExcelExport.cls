VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExcelExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************
'   Name        :   ExcelExport
'   Purpose     :   Provides functionality to export data to Excel
'   Author      :   Andy Stevens
'   Date        :   September 2004
'   Amendments
'   ----------------------------------------------------------------
'   Initials    Date        Details
'   --------    --------    ----------------------------------------
'*******************************************************************

'*******************************************************************
'
'   Developer Notes
'
'   ----------------------------------------------------------------
'   Property                Description
'   ---------------------   ----------------------------------------
'   FilePath                Path and name of excel file
'   OverWriteExistingFile   TRUE to overwite any existing file
'   WorkSheetName           Name of the worksheet
'   LeftHeader              Text to appear in the left header
'   CentreHeader            Text to appear in the centre header
'   RightHeader             Text to appear in the right header
'   LeftFooter              Option to appear in the left footer
'   CentreFooter            Option to appear in the centre footer
'   RightFooter             Option to appear in the right footer
'   HeaderRow               Number of the row to contain the header
'   ShowHeaderRow           TRUE to show the header row
'   BoldHeaderRow           TRUE to bold the header row
'   ShadeHeaderRow          TRUE to shade the header row
'   RepeatHeaderRow         TRUE to repeat the header rowon all pages
'   AddBorders              TRUE to add borders to the data
'   Orientation             Page orientation i.e Portrait, Landscape
'   CentreHorizontally      TRUE to centre the data horizontally
'   CentreVertically        TRUE to centre the data vertically
'   FitToOnePage            TRUE to 'fit' the data to one page
'   Zoom                    Percentage to Zoom. default is 100
'   PaperSize               Papaer size i.e A3, A4, A5
'   FormatColumn            Column number and format option
'   TotalColumn             Column number and TRUE to add a total
'   BoldTotals              TRUE to bold totals
'
'   ----------------------------------------------------------------
'   Methods                 Description
'   ---------------------   ----------------------------------------
'   ExportRecordsetToExcel  Creates the Excel report based upon the
'                           supplied recordset. Recordset must use a
'                           client side cursor and fields must be
'                           named as required if header is shown.
'*******************************************************************

'Declare module level variables
Private m_strFilePath               As String
Private m_blnOverWriteExistingFile  As Boolean
Private m_strWorkSheetName          As String
Private m_strLeftHeader             As String
Private m_strCentreHeader           As String
Private m_strRightHeader            As String
Private m_bytLeftFooter             As Byte
Private m_bytCentreFooter           As Byte
Private m_bytRightFooter            As Byte
Private m_intHeaderRow              As Integer
Private m_blnShowHeaderRow          As Boolean
Private m_blnBoldHeaderRow          As Boolean
Private m_blnShadeHeaderRow         As Boolean
Private m_blnRepeatHeaderRow        As Boolean
Private m_blnBoldTotals             As Boolean
Private m_bytOrientation            As Byte
Private m_blnCentreHorizontally     As Boolean
Private m_blnCentreVertically       As Boolean
Private m_blnAddBorders             As Boolean
Private m_blnFitToOnePage           As Boolean
Private m_intZoom                   As Integer
Private m_bytPaperSize              As Byte
Private m_blnColumnsFormatted       As Boolean
Private m_blnColumnsTotaled         As Boolean
Private m_dicColumnFormats          As Scripting.Dictionary
Private m_intColumnTotals()         As Integer

'Declare module level enums
Public Enum FooterType
    DatePrinted = 1
    TimePrinted = 2
    DateAndTimePrinted = 3
    FileName = 4
    FileNameAndPath = 5
    PageNumber = 6
    PageNumberAndTotal = 7
End Enum

Public Enum FormatType
    General = 1
    Number = 2
    Money = 3
    Accounting = 4
    Percentage = 5
    Scientific = 6
    Text = 7
    ShortDate = 8
    LongDate = 9
    ShortTime = 10
    LongTime = 11
End Enum

Public Enum SheetOrientation
    Portrait = 1
    Landscape = 2
End Enum

Public Enum SheetSize
    Letter = 1
    A3 = 8
    A4 = 9
    A5 = 11
End Enum

'Declare module level events
Public Event FileExists(ByVal v_strMessage As String)

Public Property Let FilePath(ByVal v_strFilePath As String)
    m_strFilePath = Trim$(v_strFilePath)
End Property

Public Property Get FilePath() As String
    FilePath = m_strFilePath
End Property

Public Property Let OverWriteExistingFile(ByVal v_blnOverWriteExistingFile As Boolean)
    m_blnOverWriteExistingFile = v_blnOverWriteExistingFile
End Property

Public Property Get OverWriteExistingFile() As Boolean
    OverWriteExistingFile = m_blnOverWriteExistingFile
End Property

Public Property Let WorkSheetName(ByVal v_strWorkSheetName As String)
    m_strWorkSheetName = Trim$(v_strWorkSheetName)
End Property

Public Property Get WorkSheetName() As String
    WorkSheetName = m_strWorkSheetName
End Property

Public Property Let LeftHeader(ByVal v_strLeftHeader As String)
    m_strLeftHeader = Trim$(v_strLeftHeader)
End Property

Public Property Get LeftHeader() As String
    LeftHeader = m_strLeftHeader
End Property

Public Property Let CentreHeader(ByVal v_strCentreHeader As String)
    m_strCentreHeader = Trim$(v_strCentreHeader)
End Property

Public Property Get CentreHeader() As String
    CentreHeader = m_strCentreHeader
End Property

Public Property Let RightHeader(ByVal v_strRightHeader As String)
    m_strRightHeader = Trim$(v_strRightHeader)
End Property

Public Property Get RightHeader() As String
    RightHeader = m_strRightHeader
End Property

Public Property Let LeftFooter(ByVal v_bytLeftFooter As FooterType)
    m_bytLeftFooter = v_bytLeftFooter
End Property

Public Property Get LeftFooter() As FooterType
    LeftFooter = m_bytLeftFooter
End Property

Public Property Let CentreFooter(ByVal v_bytCentreFooter As FooterType)
    m_bytCentreFooter = v_bytCentreFooter
End Property

Public Property Get CentreFooter() As FooterType
    CentreFooter = m_bytCentreFooter
End Property

Public Property Let RightFooter(ByVal v_bytRightFooter As FooterType)
    m_bytRightFooter = v_bytRightFooter
End Property

Public Property Get RightFooter() As FooterType
    RightFooter = m_bytRightFooter
End Property

Public Property Let HeaderRow(ByVal v_intHeaderRow As Integer)
    m_intHeaderRow = v_intHeaderRow
End Property

Public Property Get HeaderRow() As Integer
    HeaderRow = m_intHeaderRow
End Property

Public Property Let ShowHeaderRow(ByVal v_blnShowHeaderRow As Boolean)
    m_blnShowHeaderRow = v_blnShowHeaderRow
End Property

Public Property Get ShowHeaderRow() As Boolean
    ShowHeaderRow = m_blnShowHeaderRow
End Property

Public Property Let BoldHeaderRow(ByVal v_blnBoldHeaderRow As Boolean)
    m_blnBoldHeaderRow = v_blnBoldHeaderRow
End Property

Public Property Get BoldHeaderRow() As Boolean
    BoldHeaderRow = m_blnBoldHeaderRow
End Property

Public Property Let ShadeHeaderRow(ByVal v_blnShadeHeaderRow As Boolean)
    m_blnShadeHeaderRow = v_blnShadeHeaderRow
End Property

Public Property Get ShadeHeaderRow() As Boolean
    ShadeHeaderRow = m_blnShadeHeaderRow
End Property

Public Property Let RepeatHeaderRow(ByVal v_blnRepeatHeaderRow As Boolean)
    m_blnRepeatHeaderRow = v_blnRepeatHeaderRow
End Property

Public Property Get RepeatHeaderRow() As Boolean
    RepeatHeaderRow = m_blnRepeatHeaderRow
End Property

Public Property Let AddBorders(ByVal v_blnAddBorders As Boolean)
    m_blnAddBorders = v_blnAddBorders
End Property

Public Property Get AddBorders() As Boolean
    AddBorders = m_blnAddBorders
End Property

Public Property Let Orientation(ByVal v_bytOrientation As SheetOrientation)
    m_bytOrientation = v_bytOrientation
End Property

Public Property Get Orientation() As SheetOrientation
    Orientation = m_bytOrientation
End Property

Public Property Let CentreHorizontally(ByVal v_blnCentreHorizontally As Boolean)
    m_blnCentreHorizontally = v_blnCentreHorizontally
End Property

Public Property Get CentreHorizontally() As Boolean
    CentreHorizontally = m_blnCentreHorizontally
End Property

Public Property Let CentreVertically(ByVal v_blnCentreVertically As Boolean)
    m_blnCentreVertically = v_blnCentreVertically
End Property

Public Property Get CentreVertically() As Boolean
    CentreVertically = m_blnCentreVertically
End Property

Public Property Let FitToOnePage(ByVal v_blnFitToOnePage As Boolean)
    m_blnFitToOnePage = v_blnFitToOnePage
End Property

Public Property Get FitToOnePage() As Boolean
    FitToOnePage = m_blnFitToOnePage
End Property

Public Property Let Zoom(ByVal v_intZoom As Integer)
    m_intZoom = v_intZoom
End Property

Public Property Get Zoom() As Integer
    Zoom = m_intZoom
End Property

Public Property Let PaperSize(ByVal v_bytPaperSize As SheetSize)
    m_bytPaperSize = v_bytPaperSize
End Property

Public Property Get PaperSize() As SheetSize
    PaperSize = m_bytPaperSize
End Property

Friend Property Let ColumnsFormatted(ByVal v_blnColumnsFormatted As Boolean)
    m_blnColumnsFormatted = v_blnColumnsFormatted
End Property

Friend Property Get ColumnsFormatted() As Boolean
    ColumnsFormatted = m_blnColumnsFormatted
End Property

Friend Property Let ColumnsTotaled(ByVal v_blnColumnsTotaled As Boolean)
    m_blnColumnsTotaled = v_blnColumnsTotaled
End Property

Friend Property Get ColumnsTotaled() As Boolean
    ColumnsTotaled = m_blnColumnsTotaled
End Property

Public Property Let FormatColumn(ByVal v_intColumnIndex As Integer, _
                                 ByVal v_intColumnFormat As FormatType)
    
    'Add the item to the dictionary
    m_dicColumnFormats.Add CStr(v_intColumnIndex), v_intColumnFormat
    
    'Set the flag
    If Not Me.ColumnsFormatted Then
        Me.ColumnsFormatted = True
    End If

End Property

Public Property Let TotalColumn(ByVal v_intColumnIndex As Integer, _
                                ByVal v_blnTotalColumn As Boolean)
    
'Declare local variables
Static intColumn As Integer

    If v_blnTotalColumn Then

        'Re-size the array
        ReDim Preserve m_intColumnTotals(intColumn)
        
        'Add the item to the array
        m_intColumnTotals(intColumn) = v_intColumnIndex
        
        'Increment the counter
        intColumn = intColumn + 1
        
        'Set the flag
        If Not Me.ColumnsTotaled Then
            Me.ColumnsTotaled = True
        End If

    End If
    
End Property

Public Property Let BoldTotals(ByVal v_blnBoldTotals As Boolean)
    m_blnBoldTotals = v_blnBoldTotals
End Property

Public Property Get BoldTotals() As Boolean
    BoldTotals = m_blnBoldTotals
End Property

Public Function ExportRecordSetToExcel(ByVal v_rstData As ADODB.Recordset) As Boolean

'*******************************************************************
'   Name        :   ExportRecordsetToExcel
'   Purpose     :   Exports the supplied recordset to Excel
'   Parameters  :   v_rstData : Recordset to export
'   Returns     :   TRUE if recordset exported, FALSE if not
'   Author      :   Andy Stevens
'   Date        :   09 September 2004
'   Amendments
'   ----------------------------------------------------------------
'   Initials    Date        Details
'   --------    --------    ----------------------------------------
'*******************************************************************

'Declare local variables
Dim blnReturn       As Boolean
Dim intAnswer       As VbMsgBoxResult
Dim objExcel        As Excel.Application
Dim objWBook        As Excel.Workbook
Dim objWSheet       As Excel.Worksheet
Dim intRow          As Integer
Dim intCol          As Integer
Dim intRowCount     As Integer
Dim intFieldCount   As Integer
Dim intIndex        As Integer
Dim objRange        As Excel.Range
Dim strKeys         As Variant

'Declare local constants
Const FUNCTION_NAME As String = "ExportRecordsetToExcel"

    'Set default return value
    blnReturn = False
   
    'Do we have a filename?
    If Me.FilePath <> vbNullString Then
   
        'Does the file alreay exist?
        If Dir$(Me.FilePath) > vbNullString Then
        
            'Does the user want to overwrite the existing file?
            If Me.OverWriteExistingFile Then
        
                'Delete the file
                Kill Me.FilePath
            
            Else
            
                'File exists and should not be overwrtten
                RaiseEvent FileExists(Me.FilePath & " already exists." & vbCrLf & _
                                      "Please choose a different filename")
                                
                'Exit function
                ExportRecordSetToExcel = blnReturn
                Exit Function
         
            End If
            
        End If
    
    End If
    
    'Create the Excel object
    Set objExcel = New Excel.Application
    
    'Hide Excel
    objExcel.Visible = False
        
    'Create the workbook
    Set objWBook = objExcel.Workbooks.Add
    
    'Reference the first worksheet
    Set objWSheet = objWBook.Worksheets(1)
    
    'Set the start row
    If Me.HeaderRow > 0 Then
    
        'Start at the requested row
        intRow = Me.HeaderRow
    
    Else
    
        'Default to the first row
        intRow = 1
    
    End If
    
    'Count the fields and rows
    intFieldCount = v_rstData.Fields.Count
    intRowCount = v_rstData.RecordCount
    
    'Does the user want to add a header row?
    If Me.ShowHeaderRow Then
    
        'Create the column headers
        For intIndex = 0 To intFieldCount - 1
            objWSheet.Cells(intRow, intIndex + 1) = v_rstData.Fields(intIndex).Name
        Next
    
    End If
    
    'Initialise the column counter
    intCol = 1
    
    'Does the user want to bold the header row?
    If Me.BoldHeaderRow Then
    
        'Select the column headers
        Set objRange = objWSheet.Range(objWSheet.Cells(intRow, intCol), _
                                       objWSheet.Cells(intRow, intFieldCount))
               
        'Bold the column headers
        objRange.Cells.Font.Bold = True
               
    End If
    
    'Does the user want to shade the header row?
    If Me.ShadeHeaderRow Then
    
        'Select the column headers
        Set objRange = objWSheet.Range(objWSheet.Cells(intRow, intCol), _
                                       objWSheet.Cells(intRow, intFieldCount))
               
        'Shade the column headers
        With objRange.Cells
            .Interior.ColorIndex = 15
            .Interior.Pattern = xlSolid
        End With
    
    End If
    
    'Is the header row visible?
    If Me.ShowHeaderRow Then
    
        'Increment the row counter
        intRow = intRow + 1
        
    End If
    
    'General worksheet
    With objWSheet
    
        'Add the recordset data
        .Cells(intRow, intCol).CopyFromRecordset v_rstData
        
        'Has the user specified any column formats?
        If Me.ColumnsFormatted Then
        
            'Retrieve the keys
            strKeys = m_dicColumnFormats.Keys
            
            'Format the columns
            For intIndex = 0 To UBound(strKeys)
                
                'Retrieve the column number
                intCol = CInt(strKeys(intIndex))
                            
                'Select the fields
                Set objRange = .Range(.Cells(intRow, intCol), _
                                      .Cells(intRowCount + (intRow - 1), intCol))
                
                'Format the cells
                objRange.Cells.NumberFormat = GetFormatType(CInt(m_dicColumnFormats.Item(CStr(intCol))))
                            
            Next intIndex
            
        End If
        
        'Does the user want to add borders
        If Me.AddBorders Then
        
            'Add borders to the data
            With objWSheet.UsedRange.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        
        End If
        
        'Has the user specified any column totals?
        If Me.ColumnsTotaled Then
        
            'Format the columns
            For intIndex = 0 To UBound(m_intColumnTotals)
                
                'Retrieve the column number
                intCol = CInt(m_intColumnTotals(intIndex))
                                                                  
                'Add the total
                .Cells(intRowCount + (intRow), intCol).FormulaR1C1 = "=SUM(R[-" & intRowCount & "]C:R[-1]C)"
                         
                'Does the user wish to bold the total rows?
                If Me.BoldTotals Then
                
                    'Bold the total
                    .Cells(intRowCount + (intRow), intCol).Font.Bold = True
                
                End If
                
                'Add borders to the totals if required
                If Me.AddBorders Then
        
                    'Add borders to the data
                    With objWSheet.Cells(intRowCount + (intRow), intCol).Borders
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
        
                End If
                         
            Next intIndex
            
        End If
        
        'Tidy up the worksheet
        .Columns.AutoFit
        
        'Do we have a worksheet name?
        If Me.WorkSheetName <> vbNullString Then
            
            'Name the worksheet
            .Name = Me.WorkSheetName
        
        End If
        
    End With
    
    'Page setup options
    With objWSheet.PageSetup
        
        'Do we have a left header?
        If Me.LeftHeader <> vbNullString Then
        
            'Add the left header
            .LeftHeader = Me.LeftHeader
            
        End If
        
        'Do we have a centre header?
        If Me.CentreHeader <> vbNullString Then
        
            'Add the centre header
            .CenterHeader = Me.CentreHeader
            
        End If
        
        'Do we have a right header?
        If Me.RightHeader <> vbNullString Then
        
            'Add the right header
            .RightHeader = Me.RightHeader
            
        End If
        
        'Do we have a left footer?
        If Me.LeftFooter > 0 Then
        
            'Add the left header
            .LeftFooter = GetFooterText(Me.LeftFooter)
            
        End If
        
        'Do we have a centre footer?
        If Me.CentreFooter > 0 Then
        
            'Add the left header
            .CenterFooter = GetFooterText(Me.CentreFooter)
            
        End If
        
        'Do we have a centre footer?
        If Me.RightFooter > 0 Then
        
            'Add the left header
            .RightFooter = GetFooterText(Me.RightFooter)
            
        End If
            
        'Do we have an orientation?
        If Me.Orientation > 0 Then
        
            'Set the orientation
            .Orientation = Me.Orientation
            
        End If
        
        'Does the user want to set the paper size?
        If Me.PaperSize > 0 Then
        
            'Set the page size
            .PaperSize = Me.PaperSize
        
        End If
        
        'Does the user want to 'Zoom' the display?
        If Me.Zoom > 0 Then
            
            'Zoom the worksheet
            .Zoom = Me.Zoom
            
        End If
        
        'Does the user want to repeat the header row?
        If Me.RepeatHeaderRow Then
        
            'Repeat the header row on all pages
            .PrintTitleRows = "$" & Me.HeaderRow & ":$" & Me.HeaderRow
        
        End If
            
        'Does the user want to 'fit' to one page?
        If Me.FitToOnePage Then
            
            'Fit the worksheet
            .FitToPagesTall = 1
            .FitToPagesWide = 1
            .Zoom = False
                                    
        End If
        
        'Centre the page if required
        .CenterHorizontally = Me.CentreHorizontally
        .CenterVertically = Me.CentreVertically
        
    End With
     
    'Do we have a file path?
    If Me.FilePath <> vbNullString Then
    
        'Save the workbook
        objWBook.SaveAs Me.FilePath
        
    Else
    
        'Show the file
        objExcel.Visible = True
            
    End If
    
    'All OK
    blnReturn = True

On Error Resume Next

    'Destroy the Range
    If Not objRange Is Nothing Then
        Set objRange = Nothing
    End If
    
    'Destroy the Worksheet
    If Not objWSheet Is Nothing Then
        Set objWSheet = Nothing
    End If
    
    'Do we have a file path?
    If Me.FilePath <> vbNullString Then
    
        'Destroy the Workbook
        If Not objWBook Is Nothing Then
            objWBook.Close
            Set objWBook = Nothing
        End If
        
        'Destroy the Excel object
        If Not objExcel Is Nothing Then
            objExcel.Quit
            Set objExcel = Nothing
        End If
    
    End If

    ExportRecordSetToExcel = blnReturn
    
End Function

Private Function GetFooterText(ByVal v_bytFooterType As FooterType) As String

'*******************************************************************
'   Name        :   GetFooterText
'   Purpose     :   Retrieves the required footer text
'   Parameters  :   v_bytFooterType: Footer type required
'   Returns     :   String representation of required footer type
'   Author      :   Andy Stevens
'   Date        :   09 September 2004
'   Amendments
'   ----------------------------------------------------------------
'   Initials    Date        Details
'   --------    --------    ----------------------------------------
'*******************************************************************

'Declare local variables
Dim strReturn As String

On Error Resume Next

    'Add the left header
    Select Case v_bytFooterType
    
        Case DatePrinted
        
            'Add the print date
            GetFooterText = "&D"
            
        Case TimePrinted
        
            'Add the print date
            GetFooterText = "&T"
            
        Case DateAndTimePrinted
        
            'Add the print date
            GetFooterText = "&D" & " " & "&T"
                    
        Case FileName
        
            'Add the file name
            GetFooterText = "&File"
            
        Case FileNameAndPath
        
            'Add the file path
            GetFooterText = Me.FilePath
    
        Case PageNumber
    
            'Add the page number
            GetFooterText = "Page &P"
    
        Case PageNumberAndTotal
        
            'Add the page number and total pages
            GetFooterText = "Page &P of &N"
    
       Case Else
       
            'Default to nothing
            GetFooterText = vbNullString
    
    End Select
    
End Function

Private Function GetFormatType(ByVal v_bytFormatType As FormatType) As String

'*******************************************************************
'   Name        :   GetFormatType
'   Purpose     :   Retrieves the required format type
'   Parameters  :   v_bytFormatType: Format type required
'   Returns     :   String representation of required fomrat
'   Author      :   Andy Stevens
'   Date        :   13 September 2004
'   Amendments
'   ----------------------------------------------------------------
'   Initials    Date        Details
'   --------    --------    ----------------------------------------
'*******************************************************************

'Declare local variables
Dim strReturn As String

On Error Resume Next

    'Add the left header
    Select Case v_bytFormatType
            
        Case General
        
            'Format as general
            GetFormatType = "General"
    
        Case Number
        
            'Format as number
            GetFormatType = "0.00"
            
        Case Money
        
            'Format as currency
            GetFormatType = "$#,##0.00"
            
        Case Accounting
        
            'Format as accounting
            GetFormatType = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
        Case ShortDate
        
            'Format as short date
            GetFormatType = "dd/mm/yy"
        
        Case LongDate
        
            'Format as long date
            GetFormatType = "dd-mmm-yyyy"
                   
        Case ShortTime
        
            'Format as short time
            GetFormatType = "h:mm"
            
           Case ShortTime
        
            'Format as long time
            GetFormatType = "h:mm:ss AM/PM"
            
        Case Percentage
        
            'Format as percentage
            GetFormatType = "0.00%"
            
        Case Scientific
        
            'Format as scientific
            GetFormatType = "0.00E+00"
            
        Case Text
        
            'Format as text
            GetFormatType = "@"
        
       Case Else
       
            'Default to general
            GetFormatType = "General"
            
    End Select
    
End Function

Private Sub Class_Initialize()

    'Create the column dictionary
    Set m_dicColumnFormats = New Scripting.Dictionary
    
End Sub

Private Sub Class_Terminate()

    'Kill the column dictionary
    If Not m_dicColumnFormats Is Nothing Then
        Set m_dicColumnFormats = Nothing
    End If

End Sub
