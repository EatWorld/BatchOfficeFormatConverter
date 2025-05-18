Sub 户型分析()

  ' 定义变量

  Dim cell As Range

  Dim invalid As Boolean

  Dim totalUnits As Long

  Dim totalHouses As Long

  Dim totalArea As Double

  Dim houseTypeCounts As Object

  Dim areaCounts As Object

  Dim houseTypes As Variant

  Dim houseType As Variant

  Dim area As Double

  Dim dataRange As Range

  Dim dataArray As Variant

  Dim row As Long

  Dim col As Long

  Dim ws As Worksheet

  Dim outputRow As Long

  Dim i As Long, j As Long

  ' 数据验证

  invalid = False

  For Each cell In Selection

    If Not IsValidContent(cell.Value) Then

      cell.Interior.Color = RGB(255, 0, 0)

      invalid = True

    End If

  Next cell

  If invalid Then

    MsgBox "发现不规范的内容，请修复红色单元格中的内容后再运行代码。"

    Exit Sub

  End If

  ' 创建字典对象，用于存储每种户型的套数

  Set houseTypeCounts = CreateObject("Scripting.Dictionary")

  ' 创建字典对象，用于存储不同档位的套数

  Set areaCounts = CreateObject("Scripting.Dictionary")

  ' 初始化各档位的套数

  areaCounts.Add "50以下", 0

  areaCounts.Add "50-60", 0

  areaCounts.Add "60-70", 0

  areaCounts.Add "70-80", 0

  areaCounts.Add "80-100", 0

  areaCounts.Add "100-110", 0

  areaCounts.Add "110-120", 0

  areaCounts.Add "120-134", 0

  areaCounts.Add "135", 0

  areaCounts.Add "135以上", 0

  ' 获取选中的单元格区域

  Set dataRange = Selection

  ' 一次性将选中的单元格内容读取到数组中

  dataArray = dataRange.Value

  ' 遍历数组的行

  For row = LBound(dataArray, 1) To UBound(dataArray, 1)

    ' 遍历数组的列

    For col = LBound(dataArray, 2) To UBound(dataArray, 2)

      Dim cellValue As Variant  ' 添加变量声明

      cellValue = dataArray(row, col)

      ' 如果数组元素（单元格内容）不为空，则进行处理

      If cellValue <> "" Then

        ' 总户数加1

        totalUnits = totalUnits + 1

        ' 清理单元格内容格式

        Dim cleanedCell As String

        cleanedCell = CleanDataFormat(cellValue)

        ' 使用Split函数将单元格内容按照"/"分割成数组，得到每个户型的面积值（字符串类型）

        houseTypes = Split(cleanedCell, "/")

        ' 遍历每个户型面积值（字符串类型）

        For Each houseType In houseTypes

          ' 将户型面积值转换为双精度浮点数类型（Double）

          area = CDbl(houseType)

          If area <> 0 Then

            ' 总套数加1

            totalHouses = totalHouses + 1

            ' 累加总面积值

            totalArea = totalArea + area

            ' 更新各档位的套数

            Select Case area

              Case Is < 50

                areaCounts("50以下") = areaCounts("50以下") + 1

              Case 50 To 60

                areaCounts("50-60") = areaCounts("50-60") + 1

              Case 60 To 70

                areaCounts("60-70") = areaCounts("60-70") + 1

              Case 70 To 80

                areaCounts("70-80") = areaCounts("70-80") + 1

              Case 80 To 100

                areaCounts("80-100") = areaCounts("80-100") + 1

              Case 100 To 110

                areaCounts("100-110") = areaCounts("100-110") + 1

              Case 110 To 120

                areaCounts("110-120") = areaCounts("110-120") + 1

              Case 120 To 134

                areaCounts("120-134") = areaCounts("120-134") + 1

              Case 135

                areaCounts("135") = areaCounts("135") + 1

              Case Is > 135

                areaCounts("135以上") = areaCounts("135以上") + 1

            End Select

            ' 如果字典中不存在该户型，则添加该户型并设置套数为1；否则将该户型的套数加1。

            If Not houseTypeCounts.Exists(area) Then

              houseTypeCounts.Add area, 1

            Else

              houseTypeCounts(area) = houseTypeCounts(area) + 1

            End If

          End If

        Next houseType

      End If

    Next col

  Next row

  ' 获取当前工作簿的名称，而不是工作表名称

  Dim currentWorkbookName As String

  currentWorkbookName = ThisWorkbook.Name

  ' 如果文件名包含扩展名，去掉扩展名

  If InStr(currentWorkbookName, ".") > 0 Then

    currentWorkbookName = Left(currentWorkbookName, InStrRev(currentWorkbookName, ".") - 1)

  End If

  

  ' 创建新工作簿

  Dim newWB As Workbook

  Set newWB = Workbooks.Add

  Set ws = newWB.Sheets(1)

  ws.Name = "户型统计分析"

  ' 设置工作表默认样式

  With ws

    .Tab.Color = RGB(70, 130, 180)

    .Cells.Font.Name = "微软雅黑"

    .Cells.Font.Size = 11

  End With

  

  ' 写入总体信息

  ws.Range("A1:G1").Merge

  ws.Range("A1").Value = "户型分析结果"

  With ws.Range("A1")

    .Font.Size = 20

    .Font.Bold = True

    .HorizontalAlignment = xlCenter

    .VerticalAlignment = xlCenter

  End With

  ws.Rows(1).RowHeight = 40

  ' 写入总体信息到第二行

  ws.Range("A2:G2").Merge

  ws.Range("A2").Value = "一共有" & totalUnits & "户，一共有" & totalHouses & "套房屋，房屋总面积为" & Round(totalArea, 2) & "㎡。"

  With ws.Range("A2")

    .Font.Size = 16

    .Font.Bold = True

    .HorizontalAlignment = xlCenter

  End With

  ' 创建左侧户型表格

  ws.Range("A3").Value = "户型面积"

  ws.Range("B3").Value = "套数"

  ws.Range("C3").Value = "占比" ' 新增占比列

  ' 设置表头样式

  With ws.Range("A3:C3")

    .Font.Bold = True

    .Font.Size = 16

    .Interior.Color = RGB(220, 230, 241) ' 浅蓝色表头

    .Borders.LineStyle = xlContinuous

    .Borders.Weight = xlThin

    .HorizontalAlignment = xlCenter

  End With

  ' 设置基本格式

  ws.Cells.Font.Size = 16

  ws.Cells.RowHeight = 30

  ws.Columns("A:C").ColumnWidth = 20

  ' 填充户型数据

  outputRow = 4

  Dim sortedKeys As Variant

  sortedKeys = houseTypeCounts.Keys()

  Call BubbleSort(sortedKeys)

  For i = LBound(sortedKeys) To UBound(sortedKeys)

    ws.Cells(outputRow, 1).Value = sortedKeys(i)

    ws.Cells(outputRow, 2).Value = houseTypeCounts(sortedKeys(i))

    ' 计算并显示占比

    If totalHouses > 0 Then

      ws.Cells(outputRow, 3).Value = Format(houseTypeCounts(sortedKeys(i)) / totalHouses, "0.0%")

    Else

      ws.Cells(outputRow, 3).Value = "0.0%"

    End If

    outputRow = outputRow + 1

  Next i

  ' 为户型数据表格添加边框

  With ws.Range("A4:C" & (outputRow - 1))

    .Borders.LineStyle = xlContinuous

    .Borders.Weight = xlThin

    .HorizontalAlignment = xlCenter

  End With

  ' 创建右侧档位表格

  ws.Range("E3").Value = "档位"

  ws.Range("F3").Value = "套数"

  ws.Range("G3").Value = "占比" ' 新增占比列

  ' 设置表头样式

  With ws.Range("E3:G3")

    .Font.Bold = True

    .Font.Size = 16

    .Interior.Color = RGB(220, 230, 241) ' 浅蓝色表头

    .Borders.LineStyle = xlContinuous

    .Borders.Weight = xlThin

    .HorizontalAlignment = xlCenter

  End With

  ' 记录档位表格行数，方便后续确定图表位置

  Dim areaStartRow As Long

  Dim areaEndRow As Long

  areaStartRow = 4

  ' 档位顺序数组

  Dim areaOrder As Variant

  areaOrder = Array("50以下", "50-60", "60-70", "70-80", "80-100", "100-110", "110-120", "120-134", "135", "135以上")

  ' 填充档位数据

  For i = LBound(areaOrder) To UBound(areaOrder)

    ws.Cells(areaStartRow + i, 5).Value = areaOrder(i)

    ws.Cells(areaStartRow + i, 6).Value = areaCounts(areaOrder(i))

    ' 计算并显示占比

    If totalHouses > 0 Then

      ws.Cells(areaStartRow + i, 7).Value = Format(areaCounts(areaOrder(i)) / totalHouses, "0.0%")

    Else

      ws.Cells(areaStartRow + i, 7).Value = "0.0%"

    End If

    ' 设置零套数为灰色

    If areaCounts(areaOrder(i)) = 0 Then

      ws.Range(ws.Cells(areaStartRow + i, 5), ws.Cells(areaStartRow + i, 7)).Font.Color = RGB(169, 169, 169) ' 灰色

    End If

  Next i

  areaEndRow = areaStartRow + UBound(areaOrder)

  ' 为档位数据表格添加边框

  With ws.Range("E4:G" & areaEndRow)

    .Borders.LineStyle = xlContinuous

    .Borders.Weight = xlThin

    .HorizontalAlignment = xlCenter

  End With

  ' 设置列宽

  ws.Columns("A").ColumnWidth = 20

  ws.Columns("B").ColumnWidth = 20

  ws.Columns("C").ColumnWidth = 20 ' 占比列宽度

  ws.Columns("D").ColumnWidth = 6 ' 留出间隔

  ws.Columns("E").ColumnWidth = 20

  ws.Columns("F").ColumnWidth = 20

  ws.Columns("G").ColumnWidth = 20 ' 占比列宽度

  ws.Columns("H").ColumnWidth = 6 ' 留出图表间隔

  ' 添加网格线

  With ws.Range("A3:G15")

    .Borders(xlInsideHorizontal).LineStyle = xlContinuous

    .Borders(xlInsideVertical).LineStyle = xlContinuous

    .Borders(xlEdgeTop).LineStyle = xlContinuous

    .Borders(xlEdgeBottom).LineStyle = xlContinuous

    .Borders(xlEdgeLeft).LineStyle = xlContinuous

    .Borders(xlEdgeRight).LineStyle = xlContinuous

  End With

  ' 创建图表

  Call AddCharts(ws, outputRow - 1, areaEndRow + 2)

  ' 格式化整个工作表

  ws.Cells.VerticalAlignment = xlCenter

  ' 设置窗口视图

  ws.Activate

  ws.Range("A1").Select

  ' 生成汇总信息

  Dim summaryMessage As String

  summaryMessage = "一共有" & totalUnits & "户，一共有" & totalHouses & "套房屋，房屋总面积为" & Round(totalArea, 2) & "㎡。" & vbCrLf

  ' 使用已排序的键数组

  For i = LBound(sortedKeys) To UBound(sortedKeys)

    Dim percentage As String

    If totalHouses > 0 Then

      percentage = Format(houseTypeCounts(sortedKeys(i)) / totalHouses, "0.0%")

    Else

      percentage = "0.0%"

    End If

    summaryMessage = summaryMessage & vbCrLf & sortedKeys(i) & "户型有" & houseTypeCounts(sortedKeys(i)) & "套，占比" & percentage & "。" & vbCrLf

  Next i

  

  summaryMessage = summaryMessage & vbCrLf & "档位统计："

  

  For i = LBound(areaOrder) To UBound(areaOrder)

    Dim areaPercentage As String

    If totalHouses > 0 Then

      areaPercentage = Format(areaCounts(areaOrder(i)) / totalHouses, "0.0%")

    Else

      areaPercentage = "0.0%"

    End If

    summaryMessage = summaryMessage & vbCrLf & areaOrder(i) & "档位有" & areaCounts(areaOrder(i)) & "套，占比" & areaPercentage & "。"

  Next i

  

  MsgBox summaryMessage

End Sub

  

Function CleanDataFormat(content As Variant) As String

  If IsEmpty(content) Then

    CleanDataFormat = ""

    Exit Function

  End If

  Dim cleanedContent As String

  cleanedContent = CStr(content)

  ' 替换常见分隔符为标准分隔符"/"

  cleanedContent = Replace(cleanedContent, ",", "/")

  cleanedContent = Replace(cleanedContent, "、", "/")

  cleanedContent = Replace(cleanedContent, " ", "/")

  ' 移除非数字字符（保留小数点和分隔符）

  Dim result As String, i As Integer

  For i = 1 To Len(cleanedContent)

    If Mid(cleanedContent, i, 1) Like "[0-9./]" Then

      result = result & Mid(cleanedContent, i, 1)

    End If

  Next i

  ' 处理多余的分隔符

  While InStr(result, "//") > 0

    result = Replace(result, "//", "/")

  Wend

  ' 处理首尾的分隔符

  If Left(result, 1) = "/" Then result = Mid(result, 2)

  If Right(result, 1) = "/" Then result = Left(result, Len(result) - 1)

  CleanDataFormat = result

End Function

  

Function IsValidContent(content As Variant) As Boolean

  If content = "" Then

    IsValidContent = True

    Exit Function

  End If

  ' 先清理格式

  Dim cleanedContent As String

  cleanedContent = CleanDataFormat(content)

  If cleanedContent = "" Then

    IsValidContent = False

    Exit Function

  End If

  

  Dim houseTypes As Variant

  houseTypes = Split(cleanedContent, "/")

  Dim houseType As Variant

  For Each houseType In houseTypes

    If Not IsNumeric(houseType) Then

      IsValidContent = False

      Exit Function

    End If

  Next houseType

  IsValidContent = True

End Function

  

' 添加图表函数

Sub AddCharts(ws As Worksheet, lastRow As Long, startRow As Long)

  ' 创建一个专用的数据区域，确保只有一个系列

  Dim chartDataStartRow As Long

  Dim r As Long

  Dim lastDataRow As Long

  Dim chartObj As ChartObject

  Dim pieChartObj As ChartObject

  Dim totalCount As Long

  Dim i As Long, pointValue As Long

  Dim percentage As Double

  ' 在工作表最右侧创建辅助数据区域

  chartDataStartRow = 3  ' 改为从第3行开始，与表格对齐

  ' 准备图表数据区域 - 放在工作表最右边

  ws.Cells(chartDataStartRow, 20).Value = "户型面积"

  ws.Cells(chartDataStartRow, 21).Value = "套数"

  ws.Cells(chartDataStartRow, 22).Value = "百分比"

  ' 计算总套数

  totalCount = 0

  For i = 4 To lastRow  ' 从第4行开始，因为数据现在从第4行开始

    totalCount = totalCount + ws.Cells(i, 2).Value

  Next i

  ' 复制数据

  r = chartDataStartRow + 1

  For i = 4 To lastRow  ' 从第4行开始

    ws.Cells(r, 20).Value = ws.Cells(i, 1).Value

    ws.Cells(r, 21).Value = ws.Cells(i, 2).Value

    If totalCount > 0 Then

      percentage = Round((ws.Cells(i, 2).Value / totalCount) * 100, 1)

    Else

      percentage = 0

    End If

    ws.Cells(r, 22).Value = percentage

    r = r + 1

  Next i

  lastDataRow = r - 1

  ' 创建柱状图 - 调整位置到表格右侧，从I列开始

  Set chartObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Top:=ws.Rows(3).Top, Width:=800, Height:=400)

  ' 设置柱状图基本属性

  With chartObj.Chart

    .ChartType = xlColumnClustered

    .HasTitle = True

    .ChartTitle.Text = "各户型套数分布"

    .ChartTitle.Font.Size = 16

    .ChartTitle.Font.Bold = True

    .HasLegend = False

    .SetSourceData Source:=ws.Range(ws.Cells(chartDataStartRow + 1, 21), ws.Cells(lastDataRow, 21))

    .SeriesCollection(1).XValues = ws.Range(ws.Cells(chartDataStartRow + 1, 20), ws.Cells(lastDataRow, 20))

    .SeriesCollection(1).Interior.Color = RGB(153, 102, 51)

    ' 设置数据标签

    With .SeriesCollection(1)

      .HasDataLabels = True

      With .DataLabels

        .ShowValue = True

        .Font.Bold = True

        .Font.Size = 12

        .Font.Color = RGB(0, 0, 0)

        .Position = 1  ' 使用1代表上方位置

      End With

    End With

    ' 设置类别轴（X轴）

    With .Axes(xlCategory)

      .TickLabels.Font.Bold = True

      .TickLabels.Font.Size = 12

      .TickLabelPosition = xlTickLabelPositionLow

      .MajorTickMark = xlOutside

      .MinorTickMark = xlNone

      .AxisBetweenCategories = True

      .CrossesAt = 1

      .Crosses = xlAutomatic

    End With

    ' 设置数值轴（Y轴）

    With .Axes(xlValue)

      .HasMajorGridlines = True

      .TickLabels.Font.Size = 12

      .MinimumScale = 0  ' 设置Y轴最小值为0

      .MaximumScale = .MaximumScale * 1.15  ' 增加更多空间用于显示数据标签

      .CrossesAt = 0  ' Y轴从0开始

      .Crosses = xlAutomatic

    End With

  End With

  ' 创建饼图

  Set pieChartObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Top:=ws.Rows(17).Top, Width:=400, Height:=300)

  ' 设置饼图基本属性

  With pieChartObj.Chart

    .ChartType = xlPie

    .HasTitle = True

    .ChartTitle.Text = "档位套数占比"

    .ChartTitle.Font.Size = 16

    .ChartTitle.Font.Bold = True

    ' 直接使用原始数据源中的非零数据

    .SetSourceData Source:=ws.Range("F4:F13")  ' 套数列

    .SeriesCollection(1).XValues = ws.Range("E4:E13")  ' 档位列

    ' 设置数据标签

    .SeriesCollection(1).HasDataLabels = True

    With .SeriesCollection(1).DataLabels

      .NumberFormat = "0.0%"  ' 设置百分比格式

      .Position = 2  ' 使用数字2代表居中位置

      .Font.Size = 12  ' 设置字体大小

      .Font.Bold = True  ' 设置为粗体

    End With

    ' 不显示图例

    .HasLegend = False

  End With

  ' 隐藏辅助数据（确保在所有图表创建完成后执行）

  With ws.Range(ws.Cells(chartDataStartRow, 20), ws.Cells(lastDataRow, 22))

    .Font.Color = RGB(255, 255, 255)  ' 白色字体

    .Interior.Color = RGB(255, 255, 255)  ' 白色背景

  End With

End Sub

  

' 冒泡排序算法，用于对字典中的键进行排序。

Sub BubbleSort(arr)

  Dim i As Long, j As Long, temp As Variant

  For i = LBound(arr) To UBound(arr) - 1

    For j = i + 1 To UBound(arr)

      If arr(i) > arr(j) Then

        temp = arr(i)

        arr(i) = arr(j)

        arr(j) = temp

      End If

    Next j

  Next i

End Sub