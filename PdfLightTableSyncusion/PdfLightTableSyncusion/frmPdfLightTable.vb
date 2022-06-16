#Region "About"
' / ---------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com/webboard
' /
' / Purpose: Create/Generate PDF from DataTable with Syncfusion PDF .NET Library.
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / ---------------------------------------------------------------
#End Region

Imports Syncfusion.Pdf
Imports Syncfusion.Windows.Forms
Imports Syncfusion.Pdf.Parsing
Imports Syncfusion.Windows.Forms.Grid
Imports Syncfusion.Pdf.Graphics
Imports Syncfusion.Pdf.Tables

'/ H E L P
'/ https://help.syncfusion.com/cr/file-formats/Syncfusion.Pdf.Tables.PdfLightTable.html

Public Class frmPdfLightTable
    '/ Customize the application path.
    Dim strPathPDF As String = MyPath(Application.StartupPath)

    ' / --------------------------------------------------------------------------------
    ' / Get my project path
    ' / AppPath = C:\My Project\bin\debug
    ' / Replace "\bin\debug" with "\"
    ' / Return : C:\My Project\
    Function MyPath(ByVal AppPath As String) As String
        '/ Return Value
        MyPath = AppPath.ToLower.Replace("\bin\debug", "\").Replace("\bin\release", "\").Replace("\bin\x86\debug", "\")
        '// Put the \ (BackSlash) at the end.
        If Microsoft.VisualBasic.Right(MyPath, 1) <> Chr(92) Then MyPath = MyPath & Chr(92)
    End Function

    '/ START HERE
    Private Sub frmPdfLightTable_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call SamplePDF()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Create or Generate PDF with Syncfusion.
    Private Sub SamplePDF()
        '/ Create a new PDF Document.
        Dim Doc As New PdfDocument()
        Doc.PageSettings.Size = PdfPageSize.A4
        Doc.PageSettings.Orientation = PdfPageOrientation.Portrait
        '/ Size A5 and Landscape
        'Doc.PageSettings.Size = PdfPageSize.A5
        'Doc.PageSettings.Orientation = PdfPageOrientation.Landscape
        '/ Add a Page.
        Dim Page As PdfPage = Doc.Pages.Add()

        '/ Create a PdfLightTable  
        Dim PdfLightTable As New PdfLightTable
        Try
            '/ Assign data source.
            PdfLightTable.DataSource = GetDataTable()
            '/ Set Column Headers of PdfLightTable
            With PdfLightTable
                .Columns(0).ColumnName = "Primary Key"
                .Columns(1).ColumnName = "รหัสสินค้า"
                .Columns(2).ColumnName = "ชื่อสินค้า"
                .Columns(3).ColumnName = "ราคา"
            End With
            '/ Create PDF graphics for the page
            Dim logo As PdfGraphics = Page.Graphics
            '/ Load the image from the disk.
            Dim image As New PdfBitmap(strPathPDF & "g2gnet.png")
            '/ Draw the image.
            logo.DrawImage(image, 10, 0)
            '/ Create Text Header.
            Dim HeaderFont As PdfFont = New PdfTrueTypeFont(New Font("Century Gothic", 18, FontStyle.Bold), True)
            '/ Create PDF graphics for the page.
            Dim graphics As PdfGraphics = Page.Graphics
            '/ Draw the text.
            graphics.DrawString("PDF Light Table Syncfusion (PDF .NET Library)", HeaderFont, PdfBrushes.Black, New PointF(100, 10))
            '/ Set new Font.
            HeaderFont = New PdfTrueTypeFont(New Font("Arial Unicode MS", 20, FontStyle.Bold), True)    '/ True = Unicode
            graphics.DrawString("ทดสอบภาษาไทย ประถมศึกษาชั้นปีที่ 1", HeaderFont, PdfBrushes.Black, New PointF(130, 35))

            '/ Declare font and define the header style in the table.
            Dim Font As PdfFont = New PdfTrueTypeFont(New Font("Arial Unicode MS", 12, FontStyle.Regular), True)
            Dim MyStyle = New PdfCellStyle(Font, PdfBrushes.Black, PdfPens.Orange)
            Dim HeaderStyle As New PdfCellStyle(Font, PdfBrushes.White, PdfPens.Orange)
            HeaderStyle.BackgroundBrush = PdfBrushes.Orange
            With PdfLightTable
                .Style.HeaderStyle = HeaderStyle
                .Style.DefaultStyle = MyStyle
                '/ Set cell padding, Spacing.
                .Style.CellPadding = 5
                .Style.CellSpacing = 0
            End With

            '/ Create a new string format.
            Dim Format As New PdfStringFormat
            For i = 0 To 2
                '/ Set the text Alignment columns 0- 2.
                Format.Alignment = PdfTextAlignment.Left
                Format.LineAlignment = PdfVerticalAlignment.Middle
                PdfLightTable.Columns(i).StringFormat = Format
            Next
            '/ Set new stringformat.
            Format = New PdfStringFormat
            With Format
                .Alignment = PdfTextAlignment.Right
                .LineAlignment = PdfVerticalAlignment.Middle
            End With
            '/ Add the string format to PdfLightTable column.
            PdfLightTable.Columns(3).StringFormat = Format

            '/ Show header in the table.
            PdfLightTable.Style.ShowHeader = True
            '/ Draw grid to the Page of PDF Document.
            PdfLightTable.Draw(Page, New PointF(10, 76))

            '/ Save the Document
            Doc.Save(strPathPDF & "Output.pdf")
            '/ Close the Document
            Doc.Close(True)
            '/ Open PDF on PDFViewerControl
            Me.PdfViewerControl1.Load(strPathPDF & "Output.pdf", "")    '/ "" No Password

        Catch ex As Exception
            MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
            MessageBoxAdv.MessageFont = New Font("Tahoma", 12, FontStyle.Regular)
            MessageBoxAdv.Show(ex.Message, "รายงานความผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub

    ' / --------------------------------------------------------------------------------
    '// SAMPLE DATATABLE
    Private Function GetDataTable() As DataTable
        Dim DT As New DataTable
        '// Add Columns
        With DT.Columns
            .Add("PK", GetType(Integer))
            .Add("ProductID", GetType(String))
            .Add("ProductName", GetType(String))
            .Add("UnitPrice", GetType(String))
        End With
        '// Add Rows
        With DT.Rows
            '// PK, ProductID, ProductName, UnitPrice
            .Add(1, "01", "Classic Chicken", "45.00")
            .Add(2, "02", "Mexicana", "85.00")
            .Add(3, "03", "Lemon Shrimp", "110.00")
            .Add(4, "04", "Bacon", "90.00")
            .Add(5, "05", "Spicy Shrimp", "120.00")
            .Add(6, "06", "Tex Supreme", "80.00")
            .Add(7, "07", "Fish", "100.00")
            .Add(8, "08", "Pepsi Small", "20.00")
            .Add(9, "09", "Coke Small", "20.00")
            .Add(10, "10", "7Up Small", "20.00")
            .Add(11, "11", "มอคค่า", "85.00")
            .Add(12, "12", "อเมริกาโน่", "95.00")
            .Add(13, "13", "เอ็กซ์เพรสโซ่", "115.00")
            .Add(14, "14", "ชานมไข่มุก", "45.00")
            .Add(15, "15", "เหล้าเป็ก", "99.00")
        End With
        Return DT
    End Function

    Private Sub frmPdfLightTable_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

End Class
