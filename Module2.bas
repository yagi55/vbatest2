Attribute VB_Name = "Module2"
'visa5_test1_xxx_送受.
'抵抗値の値を追記対応
Sub 送受電グラフ3()

    lastRowNum = Cells(Rows.Count, 1).End(xlUp).Row

    Tick_time = InputBox("グラフ時間間隔(分）を入力")
    Tick_time = Tick_time / 24 / 60
    Start_time = InputBox("開始時間(13:10 -> 1310 入力）自動は0入力 ")
    Stop_time = InputBox("終了時間(13:30 -> 1330 入力）自動は0入力")
    
    If Start_time <> 0 Then
        Start_time_h = Start_time \ 100
        Start_time_s = Start_time Mod 100
        
        Minimum_time = Fix(Cells(2, 3))
        Minimum_time = Start_time_h / 24 + Start_time_s / 24 / 60 + Minimum_time
    Else
        Minimum_time = Cells(2, 3)
    End If
    
    If Stop_time <> 0 Then
        Stop_time_h = Stop_time \ 100
        Stop_time_s = Stop_time Mod 100
        
        Maximum_time = Fix(Cells(2, 3))
        Maximum_time = Stop_time_h / 24 + Stop_time_s / 24 / 60 + Maximum_time
        Maximum_time = WorksheetFunction.RoundUp(Maximum_time, 5)
    Else
        Maximum_time = Cells(lastRowNum, 3)
    End If
    
    Set Chart0b = ActiveSheet.ChartObjects.Add(0, 150, 800, 200)
    Set Chart1b = ActiveSheet.ChartObjects.Add(0, 360, 800, 200)
    Set Chart2b = ActiveSheet.ChartObjects.Add(0, 570, 800, 200)
    
    
    With Chart0b.Chart
        .ChartType = xlXYScatterSmoothNoMarkers
        .SetSourceData Range(Cells(1, 3), Cells(lastRowNum, 24))
        
        For i = 1 To 9
            .SeriesCollection(1).Delete
        Next

        For i = 1 To 9
            .SeriesCollection(3).Delete
        Next
               
        .Axes(xlValue, 1).HasTitle = True
        .Axes(xlValue, 1).AxisTitle.Text = "Power[W]"
        .Axes(xlValue, 1).MinimumScale = 0
        .Axes(xlValue, 1).MaximumScale = 5000
    
        
        .Axes(xlCategory, 1).HasTitle = True
        .Axes(xlCategory, 1).AxisTitle.Text = "time"
        
     
        .Axes(xlCategory, 1).MinimumScale = Minimum_time
        .Axes(xlCategory, 1).MaximumScale = Maximum_time
        
        '.Axes(xlCategory, 1).TickLabelPosition = xlLow
        .Axes(xlCategory, 1).HasMajorGridlines = True
        .Axes(xlCategory, 1).MajorUnit = Tick_time
        .Axes(xlCategory, 1).TickLabels.NumberFormatLocal = "yyyy/m/d h:mm:ss"
        
    End With
    
    lastRowNum_ppa = Cells(Rows.Count, 30).End(xlUp).Row
    
    With Chart0b.Chart.SeriesCollection.NewSeries
        .ChartType = xlXYScatterSmoothNoMarkers
        .XValues = Range(Cells(2, 30), Cells(lastRowNum_ppa, 30))
        .Values = Range(Cells(2, 32), Cells(lastRowNum_ppa, 32))
        .Name = Cells(1, 32)
        '.Name = "transmit power"
      
    End With

'   Chart1bの描画(効率、力率）

    With Chart1b.Chart
        .ChartType = xlXYScatterSmoothNoMarkers
        .SetSourceData Range(Cells(1, 30), Cells(lastRowNum_ppa, 37))
        
        For i = 1 To 3
            .SeriesCollection(1).Delete
        Next

               
        .Axes(xlValue, 1).HasTitle = True
        .Axes(xlValue, 1).AxisTitle.Text = "Power Factor, Efficiency"
        .Axes(xlValue, 1).MinimumScale = 0
        .Axes(xlValue, 1).MaximumScale = 1
    
        
        .Axes(xlCategory, 1).HasTitle = True
        .Axes(xlCategory, 1).AxisTitle.Text = "time"
        
     
        .Axes(xlCategory, 1).MinimumScale = Minimum_time
        .Axes(xlCategory, 1).MaximumScale = Maximum_time
        
        '.Axes(xlCategory, 1).TickLabelPosition = xlLow
        .Axes(xlCategory, 1).HasMajorGridlines = True
        .Axes(xlCategory, 1).MajorUnit = Tick_time
        .Axes(xlCategory, 1).TickLabels.NumberFormatLocal = "yyyy/m/d h:mm:ss"
        
    End With

'   Chart2bの描画

    With Chart2b.Chart
        .ChartType = xlXYScatterSmoothNoMarkers
        .SetSourceData Range(Cells(1, 30), Cells(lastRowNum_ppa, 45))
        
        For i = 1 To 3
            .SeriesCollection(1).Delete
        Next


        For i = 1 To 9
            .SeriesCollection(3).Delete
        Next

               
        .Axes(xlValue, 1).HasTitle = True
        .Axes(xlValue, 1).AxisTitle.Text = "Power Factor, Efficiency"
        .Axes(xlValue, 1).MinimumScale = 0
        .Axes(xlValue, 1).MaximumScale = 1
    
        
        .Axes(xlCategory, 1).HasTitle = True
        .Axes(xlCategory, 1).AxisTitle.Text = "time"
        
     
        .Axes(xlCategory, 1).MinimumScale = Minimum_time
        .Axes(xlCategory, 1).MaximumScale = Maximum_time
        
        '.Axes(xlCategory, 1).TickLabelPosition = xlLow
        .Axes(xlCategory, 1).HasMajorGridlines = True
        .Axes(xlCategory, 1).MajorUnit = Tick_time
        .Axes(xlCategory, 1).TickLabels.NumberFormatLocal = "yyyy/m/d h:mm:ss"
        
        .SeriesCollection(3).AxisGroup = 2
        .Axes(xlValue, 2).HasTitle = True
        .Axes(xlValue, 2).AxisTitle.Text = "Resistance"
        '.Axes(xlValue, 2).AxisTitle.Left = 780
        .Axes(xlValue, 2).MinimumScale = 0
        .Axes(xlValue, 2).MaximumScale = 100
        '.Axes(xlValue, 2).MajorUnit = 6
        
    End With


End Sub


Sub 充電装置グラフ調整()
    With ActiveSheet.ChartObjects(1).Chart.PlotArea
        .Width = 700
        .Left = 10
    End With
    
    With ActiveSheet.ChartObjects(2).Chart.PlotArea
       .Width = 700
       .Left = 10
    End With

    With ActiveSheet.ChartObjects(3).Chart.PlotArea
       .Width = 700
       .Left = 10
    End With
    
End Sub


Sub 解析1グラフ調整()
    With ActiveSheet.ChartObjects(1).Chart.PlotArea
        .Width = 650
        .Left = 30
    End With
    
    With ActiveSheet.ChartObjects(2).Chart.PlotArea
       .Width = 650
       .Left = 30
    End With
    
End Sub

' --------------------------------------------------------------------------------
' 以下の処理は電流設定ごとの送受電電力および電流設定ごとの効率および力率の解析を実施。
' 解析結果は新規のワークシートに出力するのてグラフ描画はSub 解析1グラフ()で実行する。
' 本当は別ファイルにした方が良かったかも・・
Sub 解析1()

    Dim current(25) As Variant
    Dim t_power(25) As Variant
    Dim r_power(25) As Variant
    Dim ach_power(25) As Variant
    Dim ch_power(25) As Variant
    Dim c_count(25) As Variant
    Dim eff1(25) As Variant
    Dim eff2(25) As Variant
    Dim eff3(25) As Variant
    Dim Powf(25) As Variant
    Dim dum As Variant
    Dim hensu(25) As Variant
        hensu(1) = 624
        hensu(2) = 759
        hensu(3) = 895
        hensu(4) = 1031
        hensu(5) = 1166
        hensu(6) = 1302
        hensu(7) = 1438
        hensu(8) = 1573
        hensu(9) = 1709
        hensu(10) = 1845
        hensu(11) = 1980
        hensu(12) = 2116
        hensu(13) = 2252
        hensu(14) = 2387
        hensu(15) = 2523
        hensu(16) = 2659
        hensu(17) = 2794
        hensu(18) = 2930
        hensu(19) = 3066
        hensu(20) = 3201
        hensu(21) = 3337
        hensu(22) = 3405

    Start_row = InputBox("開始行")
    End_row = InputBox("最後の行")

    'Start_time_s = Start_time Mod 100
    'Mnimum_time = Fix(Cells(2, 3))
    'Minimum_time = Start_time_h / 24 + Start_time_s / 24 / 60 + Minimum_time

    For n = 1 To 25
        Call zeros(current(n), t_power(n), r_power(n), ach_power(n), ch_power(n), c_count(n))
        Call zeros(eff1(n), eff2(n), eff3(n), Powf(n), dum, dum)
    Next n
    
    Dim num As Variant
    For num = 1 To 22
        For i = Start_row To End_row
            If (Cells(i, 38) = hensu(num)) Then
                current(num) = current(num) + Cells(i, 42)
                t_power(num) = t_power(num) + Cells(i, 32)
                r_power(num) = r_power(num) + Cells(i, 39)
                ach_power(num) = ach_power(num) + Cells(i, 40)
                ch_power(num) = ch_power(num) + Cells(i, 41)
                eff1(num) = eff1(num) + Cells(i, 35)
                eff2(num) = eff2(num) + Cells(i, 36)
                eff3(num) = eff3(num) + Cells(i, 37)
                Powf(num) = Powf(num) + Cells(i, 34)
                c_count(num) = c_count(num) + 1
            End If
        Next

        If c_count(num) <> 0 Then
            current(num) = current(num) / c_count(num)
            t_power(num) = t_power(num) / c_count(num)
            r_power(num) = r_power(num) / c_count(num)
            ach_power(num) = ach_power(num) / c_count(num)
            ch_power(num) = ch_power(num) / c_count(num)
            eff1(num) = eff1(num) / c_count(num)
            eff2(num) = eff2(num) / c_count(num)
            eff3(num) = eff3(num) / c_count(num)
            Powf(num) = Powf(num) / c_count(num)
        Else
            current(num) = Null
            t_power(num) = Null
            r_power(num) = Null
            ach_power(num) = Null
            ch_power(num) = Null
            eff1(num) = Null
            eff2(num) = Null
            eff3(num) = Null
            Powf(num) = Null
        End If
    Next

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Set ws1 = ActiveSheet

    Worksheets.Add after:=Worksheets(Worksheets.Count)
    Set ws2 = ActiveSheet
    ws1.Activate

    Application.ScreenUpdating = False
    ws2.Cells(1, 1) = "current"
    ws2.Cells(1, 2) = "t_power"
    ws2.Cells(1, 3) = "r_power"
    ws2.Cells(1, 4) = "ach_power"
    ws2.Cells(1, 5) = "ch_power"
    ws2.Cells(1, 6) = "効率1"
    ws2.Cells(1, 7) = "効率2"
    ws2.Cells(1, 8) = "効率3"
    ws2.Cells(1, 9) = "力率"

    ws2.Cells(1, 11) = "開始行"
    ws2.Cells(1, 12) = Start_row
    ws2.Cells(2, 11) = "停止行"
    ws2.Cells(2, 12) = End_row

    For n = 1 To 22
        ws2.Cells(1 + n, 1) = current(n)
        ws2.Cells(1 + n, 2) = t_power(n)
        ws2.Cells(1 + n, 3) = r_power(n)
        ws2.Cells(1 + n, 4) = ach_power(n)
        ws2.Cells(1 + n, 5) = ch_power(n)
        
        ws2.Cells(1 + n, 6) = eff1(n)
        ws2.Cells(1 + n, 7) = eff2(n)
        ws2.Cells(1 + n, 8) = eff3(n)
        ws2.Cells(1 + n, 9) = Powf(n)
    Next

End Sub

Sub zeros(a, b, c, d, e, f)
    a = 0
    b = 0
    c = 0
    d = 0
    e = 0
    f = 0
End Sub

'電流設定ごとの送受電電力および電流設定ごとの効率および力率の解析を実施
Sub 解析1グラフ()

    Set Chart0b = ActiveSheet.ChartObjects.Add(500, 50, 450, 350)
    Set Chart1b = ActiveSheet.ChartObjects.Add(975, 50, 450, 350)
    lastRowNum = Cells(Rows.Count, 1).End(xlUp).Row
 
    With Chart0b.Chart
        .ChartType = xlXYScatter
        .SetSourceData Range(Cells(1, 1), Cells(lastRowNum, 5))
        .HasTitle = True
        .ChartTitle.Text = "電流設定ごとの送受電電力"
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 20
            .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark2
        End With
        
        .Axes(xlValue, 1).HasTitle = True
        .Axes(xlValue, 1).AxisTitle.Text = "電力[W]"
        .Axes(xlValue, 1).MinimumScale = 0
        .Axes(xlValue, 1).MaximumScale = 4500
    
        .Axes(xlCategory, 1).HasTitle = True
        .Axes(xlCategory, 1).AxisTitle.Text = "電流[A]"
        
     
        .Axes(xlCategory, 1).MinimumScale = 5
        .Axes(xlCategory, 1).MaximumScale = 30
        
        '.Axes(xlCategory, 1).TickLabelPosition = xlLow
        .Axes(xlCategory, 1).HasMajorGridlines = True
        .Axes(xlCategory, 1).MajorUnit = 5
        
        .Legend.Position = xlLegendPositionBottom
        
    End With
    'Dim ChartObj    As Object

    Set ChartOb = ActiveSheet.ChartObjects(1)
    
    With ChartOb.Chart.SeriesCollection(2) '系列2を指定
        .Select
        '.MarkerStyle = xlMarkerStyleCircle
        .Trendlines.Add '系列2の近似曲線を追加
        .Trendlines(1).Select
        If .Type = xlPolynomial Then .Order = 1
          
        With Selection.Border
           .ColorIndex = 15 ' 線の色:灰
           .Weight = xlThin ' 線の種類：薄線
           .LineStyle = xlDot ' 線のスタイル：点線
           
        End With
    End With
    
    With ChartOb.Chart.SeriesCollection(3) '系列3を指定
        .Select
        .Trendlines.Add '系列3の近似曲線を追加
        .Trendlines(1).Select
        If .Type = xlPolynomial Then .Order = 1
        With Selection.Border
           .ColorIndex = 15 ' 線の色:灰色
           .Weight = xlThin ' 線の種類：薄線
           .LineStyle = xlDot ' 線のスタイル：点線
        End With
    End With
    
    With ChartOb.Chart.SeriesCollection(4) '系列4を指定
        .Select
        .MarkerStyle = xlMarkerStyleCircle
        .Trendlines.Add '系列4の近似曲線を追加
        
        .Trendlines(1).Select
        If .Type = xlPolynomial Then .Order = 1
        With Selection.Border
           .ColorIndex = 15 ' 線の色:灰色
           .Weight = xlThin ' 線の種類：薄線
           .LineStyle = xlDot ' 線のスタイル：点線
        End With
    End With

    With Chart1b.Chart
        .ChartType = xlXYScatter
        .SetSourceData Union(Range(Cells(1, 1), Cells(lastRowNum, 1)), Range(Cells(1, 6), Cells(lastRowNum, 9)))

        .HasTitle = True
        .ChartTitle.Text = "電流設定ごとの効率および力率"
         With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 20
            .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark2
        End With
        
        
        .Axes(xlValue, 1).HasTitle = True
        .Axes(xlValue, 1).AxisTitle.Text = "効率/力率"
        .Axes(xlValue, 1).MinimumScale = 0
        .Axes(xlValue, 1).MaximumScale = 1
    
        .Axes(xlCategory, 1).HasTitle = True
        .Axes(xlCategory, 1).AxisTitle.Text = "電流[A]"
        
     
        .Axes(xlCategory, 1).MinimumScale = 5
        .Axes(xlCategory, 1).MaximumScale = 30
        
        '.Axes(xlCategory, 1).TickLabelPosition = xlLow
        .Axes(xlCategory, 1).HasMajorGridlines = True
        .Axes(xlCategory, 1).MajorUnit = 5
       
       .Legend.Position = xlLegendPositionBottom
       
    End With
  
  
    With Chart1b.Chart.SeriesCollection(4) '系列4を指定
        .Select
        .MarkerStyle = xlMarkerStyleCircle
    End With
  
End Sub

