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
    
    ' �`���[�g
    Dim graphChart As Object
    Set graphChart = CreateObject("MSGraph.Chart.8")
    With graphChart
        ' �ςݏグ�c�_�O���t
        .ChartType = 52
        .HasTitle = True
        
        .ChartTitle.Text = "�T���v���O���t����1"
        With .ChartTitle.Font
            .Name = "Meiryo"
            .Size = 8
            .Bold = True
            .Color = RGB(255, 0, 0)
        End With
        
        ' �T�C�Y�͂Ƃ肠����2�{
        .Height = .Height * 2
        .Width = .Width * 2
        
        With .Application.DataSheet
            ' �f�[�^���ږ�
            .Cells.Clear
            .Cells(1, 1).Value = "����"
            .Cells(2, 1).Value = "���"
            .Cells(3, 1).Value = "Step(����)"
            .Cells(4, 1).Value = "Step(�v��)"
            .Cells(5, 1).Value = "�i��(����)"
            .Cells(6, 1).Value = "�i��(�v��)"
            
            ' �f�[�^1-1
            .Cells(1, 2).Value = "C#"
            .Cells(2, 2).Value = "���W���[��"
            .Cells(3, 2).Value = 5
            .Cells(4, 2).Value = 10
            .Cells(5, 2).Value = "10%"
            .Cells(6, 2).Value = "100%"
            
            ' �f�[�^1-2
            .Cells(1, 3).Value = ""
            .Cells(2, 3).Value = "���i"
            .Cells(3, 3).Value = 40
            .Cells(4, 3).Value = 50
            .Cells(5, 3).Value = "90%"
            .Cells(6, 3).Value = "100%"
            
            ' �f�[�^2-1
            .Cells(1, 4).Value = "Java"
            .Cells(2, 4).Value = "���W���[��"
            .Cells(3, 4).Value = 3
            .Cells(4, 4).Value = 10
            .Cells(5, 4).Value = "30%"
            .Cells(6, 4).Value = "100%"
            
            ' �f�[�^2-2
            .Cells(1, 5).Value = ""
            .Cells(2, 5).Value = "���i"
            .Cells(3, 5).Value = 4
            .Cells(4, 5).Value = 20
            .Cells(5, 5).Value = "5%"
            .Cells(6, 5).Value = "100%"
            
            ' �f�[�^3-1
            .Cells(1, 6).Value = "Ruby"
            .Cells(2, 6).Value = "���W���[��"
            .Cells(3, 6).Value = 3
            .Cells(4, 6).Value = 30
            .Cells(5, 6).Value = "10%"
            .Cells(6, 6).Value = "100%"
            
            ' �f�[�^3-2
            .Cells(1, 7).Value = ""
            .Cells(2, 7).Value = "���i"
            .Cells(3, 7).Value = 10
            .Cells(4, 7).Value = 40
            .Cells(5, 7).Value = "25%"
            .Cells(6, 7).Value = "100%"
        End With
        
        ' �f�[�^�̕��т͗����
        .Application.PlotBy = 1
        
        ' ��ʂ̗�̓f�[�^�Ƃ��ĕs�v�Ȃ̂ō폜
        .SeriesCollection(1).Delete
        
        ' �}��̈ʒu�̓O���t�̉�
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
        
        ' ��2���̕\���ݒ�
        .Axes(2, 2).MajorTickMark = 3
        .Axes(2, 2).MinimumScale = "0%"
        .Axes(2, 2).MaximumScale = "110%"
        
        ' �c�_�O���t�̕\���Ԋu
        .ChartGroups(1).GapWidth = 20
        
        ' �c�_�O���t��1��ڂɂ��ăO���f�[�V������K�p
        ' �p�^�[���^�C�v�̒l�ɂ��Ă�
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
    
    '���l���̃t�H���g�ݒ�
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
        ' �`���[�g��̐��l�̃t�H���g
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
