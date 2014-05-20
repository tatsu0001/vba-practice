Attribute VB_Name = "GraphTest"
Option Compare Database
Option Explicit


Public Sub Test()
    Dim db As DbConnection
    Set db = New DbConnection
    db.OpenConnection "FileDSN=" & CurrentProject.Path & "\GraphData.dsn;" _
        & "DefaultDir=" & CurrentProject.Path & ";"
    
    Dim records As ADODB.Recordset
    Set records = db.Execute("SELECT * FROM [Book1.csv]")
    
    ' チャート
    Dim graphChart As Object
    Set graphChart = CreateObject("MSGraph.Chart.8")
    With graphChart
        ' 積み上げ縦棒グラフ
        .ChartType = 52
        .HasTitle = True
        
        .ChartTitle.Text = "サンプルグラフその1"
        With .ChartTitle.Font
            .Name = "Meiryo"
            .Size = 8
            .Bold = True
            .Color = RGB(255, 0, 0)
        End With
        
        ' サイズはとりあえず2倍
        .Height = .Height * 2
        .Width = .Width * 2
        
        With .Application.DataSheet
            ' データ項目名
            .Cells.Clear
            .Cells(1, 1).Value = "言語"
            .Cells(2, 1).Value = "種別"
            .Cells(3, 1).Value = "Step(現在)"
            .Cells(4, 1).Value = "Step(計画)"
            .Cells(5, 1).Value = "進捗(現在)"
            .Cells(6, 1).Value = "進捗(計画)"
            
            ' データ1-1
            .Cells(1, 2).Value = "C#"
            .Cells(2, 2).Value = "モジュール"
            .Cells(3, 2).Value = 5
            .Cells(4, 2).Value = 10
            .Cells(5, 2).Value = "10%"
            .Cells(6, 2).Value = "100%"
            
            ' データ1-2
            .Cells(1, 3).Value = ""
            .Cells(2, 3).Value = "製品"
            .Cells(3, 3).Value = 40
            .Cells(4, 3).Value = 50
            .Cells(5, 3).Value = "90%"
            .Cells(6, 3).Value = "100%"
            
            ' データ2-1
            .Cells(1, 4).Value = "Java"
            .Cells(2, 4).Value = "モジュール"
            .Cells(3, 4).Value = 3
            .Cells(4, 4).Value = 10
            .Cells(5, 4).Value = "30%"
            .Cells(6, 4).Value = "100%"
            
            ' データ2-2
            .Cells(1, 5).Value = ""
            .Cells(2, 5).Value = "製品"
            .Cells(3, 5).Value = 4
            .Cells(4, 5).Value = 20
            .Cells(5, 5).Value = "5%"
            .Cells(6, 5).Value = "100%"
            
            ' データ3-1
            .Cells(1, 6).Value = "Ruby"
            .Cells(2, 6).Value = "モジュール"
            .Cells(3, 6).Value = 3
            .Cells(4, 6).Value = 30
            .Cells(5, 6).Value = "10%"
            .Cells(6, 6).Value = "100%"
            
            ' データ3-2
            .Cells(1, 7).Value = ""
            .Cells(2, 7).Value = "製品"
            .Cells(3, 7).Value = 10
            .Cells(4, 7).Value = 40
            .Cells(5, 7).Value = "25%"
            .Cells(6, 7).Value = "100%"
        End With
        
        ' データの並びは列方向
        .Application.PlotBy = 1
        
        ' 種別の列はデータとして不要なので削除
        .SeriesCollection(1).Delete
        
        ' 凡例の位置はグラフの下
        .Legend.Position = -4107
        With .Legend.Font
            .Name = "Meiryo"
            .Size = 8
            .Bold = True
            .Color = RGB(0, 0, 0)
        End With
        
        With .SeriesCollection(3)
            .ChartType = 65
            .AxisGroup = 2
            .MarkerStyle = 3
        End With
        
        With .SeriesCollection(4)
            .ChartType = 65
            .AxisGroup = 2
            .MarkerStyle = 3
        End With
        
        ' 第2軸の表示設定
        .Axes(2, 2).MajorTickMark = 3
        .Axes(2, 2).MinimumScale = "0%"
        .Axes(2, 2).MaximumScale = "110%"
        
        ' 縦棒グラフの表示間隔
        .ChartGroups(1).GapWidth = 20
        
        ' 縦棒グラフの1列目についてグラデーションを適用
        ' パターンタイプの値については
        ' http://msdn.microsoft.com/ja-jp/library/ff864036.aspx
        .SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
        .SeriesCollection(1).Points(1).Fill.Patterned 16
        .SeriesCollection(1).Points(3).Fill.Patterned 16
        .SeriesCollection(1).Points(5).Fill.Patterned 16
        
        .SeriesCollection(2).Interior.Color = RGB(0, 0, 255)
        .SeriesCollection(2).Points(1).Fill.Patterned 16
        .SeriesCollection(2).Points(3).Fill.Patterned 16
        .SeriesCollection(2).Points(5).Fill.Patterned 16
        
        .PlotArea.Interior.Color = RGB(255, 255, 255)
    End With
    
    '数値軸のフォント設定
    Dim axis As Object
    For Each axis In graphChart.Axes
        With axis.TickLabels.Font
            .Name = "Meiryo"
            .Size = 8
            .Bold = True
            .Color = RGB(0, 0, 0)
        End With
    Next
    
    Dim i As Object
    For Each i In graphChart.SeriesCollection
        ' チャート上の数値のフォント
        With i
            .HasDataLabels = True
            With .DataLabels.Font
                .Name = "Meiryo"
                .Size = 8
                .Bold = True
                .Color = RGB(0, 0, 0)
            End With
        End With
    Next
    
    
    
    Debug.Print graphChart.ChartGroups.Count
    
    graphChart.Application.Update
    graphChart.Export CurrentProject.Path & "\data.png"
    graphChart.Application.Visible = True
    graphChart.Application.Quit
End Sub
