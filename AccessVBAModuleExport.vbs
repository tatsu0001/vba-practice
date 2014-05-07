Option Explicit


' エクスポートの共通処理
Class ExportModuleBase
    Private fileSystem
    Private exportBasePath
    Private outputDirName
    Private dateTime
    Private exportDirPath

    ' 初期化
    Private Sub Class_Initialize()
        Dim currentDateTime
        outputDirName = "export"
        currentDateTime = Date
        dateTime = GetDateTimeString(currentDateTime)

        Set fileSystem = WScript.CreateObject("Scripting.FileSystemObject")
        exportBasePath = fileSystem.GetParentFolderName(WScript.ScriptFullName)
    End Sub

    Public Property Let ExportBase(path)
        exportBasePath = path
    End Property

    Public Property Let OutputDir(name)
        outputDirName = name
    End Property

    Private Function GetDateTimeString(dateTime)
        Dim dateTimeStr

        dateTimeStr = Year(Now())
        dateTimeStr = dateTimeStr & Right("0" & Month(Now()) , 2)
        dateTimeStr = dateTimeStr & Right("0" & Day(Now()) , 2)
        dateTimeStr = dateTimeStr & "-"
        dateTimeStr = dateTimeStr & Right("0" & Hour(Now()) , 2)
        dateTimeStr = dateTimeStr & Right("0" & Minute(Now()) , 2)
        dateTimeStr = dateTimeStr & Right("0" & Second(Now()) , 2)

        GetDateTimeString = dateTimeStr
    End Function

    Public Function CreateDir(path)
        Dim parentPath
        Dim dirPath

        ' 指定ディレクトリの親ディレクトリから再帰的に作成
        parentPath = fileSystem.GetParentFolderName(path)
        If fileSystem.FolderExists(parentPath) Then
            If Not fileSystem.FolderExists(path) Then
                fileSystem.CreateFolder(path)
                WScript.Echo "create export dir " & path
            End If
            dirPath = path 
        Else
            dirPath = CreateDir(parentPath)
            dirPath = CreateDir(path)
        End If

        CreateDir = dirPath
    End Function

    ' エクスポート先のベースディレクトリを作成
    Private Function CreateExportDir()
        Dim path

        path = exportBasePath & "\" & outputDirName & "\" & dateTime
        CreateExportDir = CreateDir(path)
    End Function

    Private Function InputVBAProjectFile()
        Dim path
        path = InputBox("VBAプロジェクトファイルのパスを入力して下さい。")
        WScript.Echo "vba project file = " & path
        InputVBAProjectFile = path
    End Function

    ' エクスポート処理実行
    Public Function TryExport(selector)
        WScript.Echo "start TryExport."

        Dim exporter
        Dim vbaProject
        Dim exportDir

        vbaProject = InputVBAProjectFile()
        ' reuqired selector object implements SelectBy
        Set exporter = selector.SelectBy(vbaProject)
        exportDir = CreateExportDir()

        ' エクスポートオブジェクトの処理を呼ぶだけ
        ' required export object implements ExportBase property, and TryExport function
         exporter.ExportBase = Me
         Set TryExport = exporter.TryExport(vbaProject, exportDir)
    End Function
End Class


Class SelectBySuffix
    Public Function SelectBy(path)
        Dim suffix

        suffix = GetSuffix(path)
        Set SelectBy = GetExporter(suffix)
    End Function

    Private Function GetSuffix(path)
        Dim index
        index = InStrRev(path, ".")
        GetSuffix = Mid(path, index + 1, Len(path))
    End Function

    Private Function GetExporter(suffix)
        WScript.Echo "export target file suffix = " & suffix

        Dim exporter
        Select Case suffix
            Case "adp", "accdb"
                Set exporter = New AccessExport
            Case Else
                Set exporter = New NotExport
        End Select

        exporter.Suffix = suffix
        Set GetExporter = exporter
    End Function
End Class

Class NotExport
    Public Property Let ExportBase(baseObj)
    End Property
    
    Public Function TryExport(vbaProject, exportDir)
        WScript.Echo "Unknown vba file type. Not Export."
        Set TryExport = Me
    End Function
End Class

' Accessプロジェクトファイルからモジュールをエクスポートする
Class AccessExport
    Private exportUtil
    Private strSuffix

    Public Property Let ExportBase(baseObj)
        Set exportUtil = baseObj
    End Property

    Public Property Let Suffix(value)
        strSuffix = value
    End Property

    Public Function TryExport(vbaProject, exportDir)
        Dim access 
        Dim module

        Wscript.Echo "AccessProject Start TryExport."
    
        ' Accdb or Adpファイルオープン
        Set access = CreateObject("Access.Application")
        If strSuffix = "adp" Then
            access.OpenAccessProject(vbaProject)
        ElseIf strSuffix = "accdb" Then
            access.OpenCurrentDatabase(vbaProject)
        End If

        ' VBAモジュールのExport
        Dim moduleSuffix
        For Each module In access.VBE.ActiveVBProject.VBComponents
            moduleSuffix = GetModuleSuffix(module)

            WScript.Echo " -- " & module.Name & " --" 
            WScript.Echo " type " & module.Type  

            If Not moduleSuffix = "" Then
                module.Export (exportDir & "/" & module.Name & moduleSuffix)
                WScript.Echo " export " & module.Name
            End If
        Next

        ' マクロ
        WScript.Echo "start export macros."
        ExportNotVBAModules access, access.CurrentProject.AllMacros, exportDir, "mcr"

        access.Quit()

        Set TryExport = Me
    End Function

    ' VBAモジュール以外のExport
    Private Sub ExportNotVBAModule(ByRef access, ByRef module, ByVal exportDir, ByVal suffix)
        'Dim exportDirBySuffix
        Dim exportModuleFilePath

        'exportDirBySuffix = exportDir & "\" & suffix
        'exportUtil.CreateDir(exportDirBySuffix)

        'exportModuleFilePath = exportDirBySuffix & "\" & module.Name & "." & suffix
        exportModuleFilePath = exportDir & "\" & module.Name & "." & suffix
        access.SaveAsText module.Type, module.Name, exportModuleFilePath

        WScript.Echo "export " & module.Name & " -> " & exportModuleFilePath
    End Sub

    Private Sub ExportNotVBAModules(ByRef access, ByRef modules, ByVal exportDir, ByVal suffix)
        Dim moduleObj

        If modules.Count > 0 Then
            For Each moduleObj In modules
                ExportNotVBAModule access, moduleObj, exportDir, suffix
            Next
        End If
    End Sub

    Private Function GetModuleSuffix(ByRef module)
        Dim moduleSuffix 
        moduleSuffix = ""

        If module.Type = 1 Then
            ' 標準モジュール
            moduleSuffix = ".bas"
        ElseIf module.Type = 2 Then
            ' クラスモジュール
            moduleSuffix = ".cls"
        ElseIf module.Type = 100 Then
            If InStr(module.Name, "Form_") = 1 Then
                ' フォーム
                moduleSuffix = ".frm"
            ElseIf InStr(module.Name, "Report_") = 1 Then
                ' レポート
                moduleSuffix = ".rpt"
            End If
        End If

        GetModuleSuffix = moduleSuffix
    End Function

End Class
    
Sub Main()
    Dim exportBase
    Dim selector
    Dim exporter
    
    Set exportBase = New ExportModuleBase 
    Set selector = New SelectBySuffix
    Set exporter = exportBase.TryExport(selector)
End Sub

Main

