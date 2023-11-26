

Imports myExcel = Microsoft.Office.Interop.Excel

Public Class Excel_Access

    Public _xlApp As myExcel.Application
    Public _xlWorkBooks As myExcel.Workbooks
    Public _xlWorkBook As myExcel.Workbook
    Public _xlWorkSheet As myExcel.Worksheet
    Public _xlWorkSheet_ChartSheet As myExcel.Worksheet
   
    Public _IsOpenWorkFile As Boolean = False

    Public _ChartTitle As String
    Public _LineCount As Integer
    Public _LineName As ArrayList
    Public _LineRange As ArrayList

    Public _X_Axise As String
    Public _X_MaxScale As Double
    Public _X_MinScale As Double
    Public _X_MajorUnit As Double
    Public _X_MinorUniit As Double
    Public _X_logNumber As Integer = 0

    Public _Y_Axise As String
    Public _Y_MaxScale As Double
    Public _Y_MinScale As Double
    Public _Y_MajorUnit As Double
    Public _Y_MinorUniit As Double
    Public _Y_logNumber As Integer = 0

    Public Function OpenFile(xlFilePath As String, xlSheetName As String) As String

        Try
            Dim retStr As String = "Done"
            _xlApp = New myExcel.Application
            _xlApp.DisplayAlerts = False

            _xlWorkBooks = _xlApp.Workbooks
            _xlWorkBook = _xlWorkBooks.Open(xlFilePath)

            _xlWorkSheet = _xlWorkBook.Worksheets(xlSheetName)
            If _xlWorkSheet Is Nothing Then
                _xlWorkBook?.Close()
                retStr = "The sheet not found."
                Return retStr
            End If

            _xlWorkSheet.Activate()
            _IsOpenWorkFile = True
            Return retStr
        Catch ex As Exception

            Return "Open excel file fail."

        End Try

    End Function

    Public Function CloseFile(xlFilePath As String, xlSheetName As String) As String

        Dim retStr As String = "Done"
        Try

            _xlWorkSheet.Columns.AutoFit()
            _xlWorkBook.SaveAs(xlFilePath) 
            _xlWorkBook.Close()

            retStr = "Done"
            _IsOpenWorkFile = False

        Catch ex As Exception

            _xlWorkBook?.Close()
            retStr = "Close excel file fail."

        End Try

        _xlApp.Quit()
        Return retStr

    End Function

    Public Function SetChartInformation(Title As String, LineCnt As Integer, LineName As ArrayList, ByRef LineRange As ArrayList) As String

        Dim retStr As String = "Done"

        _ChartTitle = Title
        _LineCount = LineCnt
        _LineName = LineName
        _LineRange = LineRange

        Return retStr

    End Function

    Public Function SetChart_XParameter(X_Name As String, X_Max As Double, X_Min As Double,
                                        X_MajorUnit As Double, X_MinorUniit As Double, logNumber As Integer) As String

        Dim retStr As String = "Done"
        _X_Axise = X_Name
        _X_MaxScale = X_Max
        _X_MinScale = X_Min
        _X_MajorUnit = X_MajorUnit
        _X_MinorUniit = X_MinorUniit
        _X_logNumber = logNumber

        Return retStr

    End Function

    Public Function SetChart_YParameter(Y_Name As String, Y_Max As Double, Y_Min As Double,
                                        Y_MajorUnit As Double, Y_MinorUniit As Double, logNumber As Integer) As String

        Dim retStr As String = "Done"

        _Y_Axise = Y_Name
        _Y_MaxScale = Y_Max
        _Y_MinScale = Y_Min
        _Y_MajorUnit = Y_MajorUnit
        _Y_MinorUniit = Y_MinorUniit
        _Y_logNumber = logNumber

        Return retStr

    End Function

    Public Function CreateChart(xlSheetName As String) As String

        Try
            Dim retStr As String = "Done"
   
            If _xlWorkBook Is Nothing Then
                Return "Not open file yet."
            End If
            If _IsOpenWorkFile = False Then
                Return "Not open file yet."
            End If

            Dim chartPage As myExcel.Chart
            Dim xlCharts As myExcel.ChartObjects
            Dim myChart As myExcel.ChartObject
            xlCharts = _xlWorkSheet.ChartObjects

            If xlCharts.Count > 0 Then
                xlCharts.Delete()
            End If

            myChart = xlCharts.Add(800, 180, 450, 350) ' Set Chart position and size
            chartPage = myChart.Chart

            With chartPage
                For LineCnt = 1 To _LineCount
                    If LineCnt = 1 Then
                        Dim dataSeries As myExcel.Series = .SeriesCollection.NewSeries()
                        ' Set Y range
                        dataSeries.Values = _xlWorkSheet.Range(_LineRange(0)).ClearContents()
                        ' Set X range
                        dataSeries.XValues = _xlWorkSheet.Range(_LineRange(1)).ClearContents()
                        dataSeries.Name = _LineName(0)    
                        'dataSeries.Border.Color = Color.DodgerBlue
                    End If

                    If LineCnt = 2 Then
                        Dim dataSeries1 As myExcel.Series = .SeriesCollection.NewSeries()
                        dataSeries1.Values = _xlWorkSheet.Range(_LineRange(2)).ClearContents()
                        dataSeries1.XValues = _xlWorkSheet.Range(_LineRange(3)).ClearContents()
                        dataSeries1.Name = _LineName(1)     '
                        'dataSeries1.Border.Color = Color.MediumSeaGreen
                    End If

                    If LineCnt = 3 Then
                        Dim dataSeries2 As myExcel.Series = .SeriesCollection.NewSeries()
                        dataSeries2.Values = _xlWorkSheet.Range(_LineRange(4)).ClearContents()
                        dataSeries2.XValues = _xlWorkSheet.Range(_LineRange(5)).ClearContents()
                        dataSeries2.Name = _LineName(2)     ' 
                        'dataSeries2.Border.Color = Color.Purple
                    End If

                    If LineCnt = 4 Then
                        Dim dataSeries3 As myExcel.Series = .SeriesCollection.NewSeries()
                        dataSeries3.Values = _xlWorkSheet.Range(_LineRange(6)).ClearContents()
                        dataSeries3.XValues = _xlWorkSheet.Range(_LineRange(7)).ClearContents()
                        dataSeries3.Name = _LineName(3)    
                        'dataSeries1.Border.Color = Color.Brown
                    End If

                    If LineCnt = 5 Then
                        Dim dataSeries4 As myExcel.Series = .SeriesCollection.NewSeries()
                        dataSeries4.Values = _xlWorkSheet.Range(_LineRange(8)).ClearContents()
                        dataSeries4.XValues = _xlWorkSheet.Range(_LineRange(9)).ClearContents()
                        dataSeries4.Name = _LineName(4)     
                        'dataSeries2.Border.Color = Color.Coral

                    End If
                    If LineCnt = 6 Then
                        Dim dataSeries5 As myExcel.Series = .SeriesCollection.NewSeries()
                        dataSeries5.Values = _xlWorkSheet.Range(_LineRange(10)).ClearContents()
                        dataSeries5.XValues = _xlWorkSheet.Range(_LineRange(11)).ClearContents()
                        dataSeries5.Name = _LineName(5)    
                        'dataSeries1.Border.Color = Color.DarkGreen

                    End If
                Next







                'set data labels for bars
                '.ApplyDataLabels(myExcel.XlDataLabelsType.xlDataLabelsShowValue)
                .ApplyDataLabels(myExcel.XlDataLabelsType.xlDataLabelsShowLabel, False, False, False, False, True, False, False, False, False)
                'set legend to be displayed or not
                .HasLegend = True
                'set legend location
                .Legend.Position = myExcel.XlLegendPosition.xlLegendPositionRight
                'select chart type
                .ChartType = myExcel.XlChartType.xlXYScatterSmoothNoMarkers

                'chart title
                .HasTitle = True
                .ChartTitle.Text = _ChartTitle '
                .ChartTitle.Font.Size = 14
                .ChartTitle.Font.Bold = False
                'set titles for Axis values and categories
                Dim xlAxisCategory, xlAxisValue As myExcel.Axes

                ' Set X information
                xlAxisCategory = CType(chartPage.Axes(, myExcel.XlAxisGroup.xlPrimary), myExcel.Axes)
                xlAxisCategory.Item(myExcel.XlAxisType.xlCategory).HasTitle = True
                xlAxisCategory.Item(myExcel.XlAxisType.xlCategory).AxisTitle.Characters.Text = _X_Axise '
                xlAxisCategory.Item(myExcel.XlAxisType.xlCategory).AxisTitle.Characters.Font.Bold = False

                ' Set Y information
                xlAxisValue = CType(chartPage.Axes(, myExcel.XlAxisGroup.xlPrimary), myExcel.Axes)
                xlAxisValue.Item(myExcel.XlAxisType.xlValue).HasTitle = True
                xlAxisValue.Item(myExcel.XlAxisType.xlValue).AxisTitle.Characters.Text = _Y_Axise '
                xlAxisValue.Item(myExcel.XlAxisType.xlValue).AxisTitle.Characters.Font.Bold = False
                xlAxisValue.Item(myExcel.XlAxisType.xlValue).MaximumScale = _Y_MaxScale ' 100
                xlAxisValue.Item(myExcel.XlAxisType.xlValue).MinimumScale = _Y_MinScale ' 0.0
                xlAxisValue.Item(myExcel.XlAxisType.xlValue).MajorUnit = _Y_MajorUnit   '25.0
                xlAxisValue.Item(myExcel.XlAxisType.xlValue).MinorUnit = _Y_MinorUniit  '25.0
                'xlAxisValue.Item(myExcel.XlAxisType.xlValue).ScaleType = XlScaleType.xlScaleLogarithmic
                xlAxisValue.Item(myExcel.XlAxisType.xlValue).HasMajorGridlines = True
                xlAxisValue.Item(myExcel.XlAxisType.xlValue).HasMinorGridlines = True

                ' Set X information
                xlAxisValue.Item(myExcel.XlAxisType.xlCategory).MaximumScale = _X_MaxScale  '10.0
                xlAxisValue.Item(myExcel.XlAxisType.xlCategory).MinimumScale = _X_MinScale  '0.0
                xlAxisValue.Item(myExcel.XlAxisType.xlCategory).MajorUnit = _X_MajorUnit    '2.0
                xlAxisValue.Item(myExcel.XlAxisType.xlCategory).MinorUnit = _X_MinorUniit   '2.0
                xlAxisValue.Item(myExcel.XlAxisType.xlCategory).HasMajorGridlines = True
                xlAxisValue.Item(myExcel.XlAxisType.xlCategory).HasMinorGridlines = True

            End With

            Return retStr

        Catch ex As Exception

            Return "Active chart sheet fail."

        End Try

    End Function

    Public Function SetValue(RowIndex As Integer, RowBase As Integer, ColIndex As Integer, ColBase As Integer, ColInteval As Integer, ValueData As Double) As String

        Dim retStr As String = "Done"
        Try
            _xlApp.Cells(RowIndex + RowBase, ColBase + (ColIndex * ColInteval)) = ValueData
        Catch ex As Exception

            retStr = "Set value fail."

        End Try

        Return retStr

    End Function

End Class
