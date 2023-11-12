
' 普通模块代码
Option Explicit
' 主运行方法 Main
Sub Main()
    ' 初始化用户界面，准备数据填充
    UserForm1.Show 0
    InitializeUserForm
    ' 显示用户界面
    CheckMappingAndUpdateLabel
    FillMappingListBoxAndIdentifiers
End Sub

' 功能：获取用于存储设置的文本文件的完整路径。
' 在这个函数中：
' 1. 使用 Application.SldWorks 获取 SolidWorks 的应用程序对象。
' 2. 获取当前运行的宏的路径，并从中提取目录路径。
' 3. 拼接目标文件名，即存储设置的文本文件名。
' 4. 返回完整的文件路径。
Function GetFilePath() As String
    Dim swApp As Object
    Dim currentMacroPath As String
    Dim directoryPath As String
    Dim targetFileName As String
    
    ' 获取SolidWorks应用对象
    Set swApp = Application.SldWorks
    
    ' 获取当前宏的路径
    currentMacroPath = swApp.GetCurrentMacroPathName
    directoryPath = Left(currentMacroPath, InStrRev(currentMacroPath, "\"))
    targetFileName = "智能命名系统data.txt"
    
    ' 返回构建的文件路径
    GetFilePath = directoryPath & targetFileName
End Function
' 此子程序用于读取映射行并填充列表框，同时填充文本框显示完整的代号行内容
Public Sub FillMappingListBoxAndIdentifiers()
    Dim filePath As String
    Dim fileContent As String
    Dim lines As Variant
    Dim i As Integer
    
    filePath = GetFilePath()
    If filePath <> "" And Dir(filePath) <> "" Then
        fileContent = GetFileContent(filePath)
        lines = Split(fileContent, vbCrLf)
        
        ' 清空列表框
        UserForm1.ListBox_映射.Clear
        
        ' 遍历每一行
        For i = LBound(lines) To UBound(lines)
            If InStr(lines(i), "映射|") > 0 Then
                ' 只向列表框添加以"映射"开头的行
                UserForm1.ListBox_映射.AddItem lines(i)
            ElseIf InStr(lines(i), "钣金件代号|") > 0 Then
                ' 显示完整的钣金件代号行
                UserForm1.TextBox_钣金件代号.text = lines(i)
            ElseIf InStr(lines(i), "车床件代号|") > 0 Then
                ' 显示完整的车床件代号行
                UserForm1.TextBox_车床件代号.text = lines(i)
            ElseIf InStr(lines(i), "亚克力代号|") > 0 Then
                ' 显示完整的亚克力代号行
                UserForm1.TextBox_亚克力代号.text = lines(i)
            ElseIf InStr(lines(i), "机加件标识|") > 0 Then
                ' 显示完整的机加件标识行
                UserForm1.TextBox_机加件标识.text = lines(i)
            End If
        Next i
    Else
        MsgBox "配置文件不存在，请检查。"
    End If
End Sub






Public Function GetFileContent(ByVal filePath As String) As String
    Dim fileNumber As Integer
    Dim lineContent As String
    Dim contentBuilder As String
    fileNumber = FreeFile

    On Error GoTo ErrorHandler
    Open filePath For Input As fileNumber
    
    ' 初始化内容构建器
    contentBuilder = ""
    
    ' 循环读取每一行直到文件结束
    Do Until EOF(fileNumber)
        Line Input #fileNumber, lineContent
        contentBuilder = contentBuilder & lineContent & vbCrLf
    Loop
    
    Close fileNumber
    GetFileContent = contentBuilder
    Exit Function
    
ErrorHandler:
    MsgBox "读取文件时发生错误: " & Err.Description, vbCritical, "错误"
    Close fileNumber
    GetFileContent = ""
End Function

' 检查当前装配体名称是否在映射列表中，并更新Label_映射信息的状态
Public Sub CheckMappingAndUpdateLabel()
    Dim swApp As Object
    Dim swModel As Object
    Dim assemblyName As String
    Dim assemblyBaseName As String ' 添加变量来存储基础装配体名称
    Dim foundMapping As Boolean
    Dim mappingEntry As String
    Dim mappingParts As Variant
    Dim i As Integer
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    foundMapping = False
    
    If Not swModel Is Nothing Then
        assemblyName = swModel.GetTitle ' 获取装配体名称
        ' 检查是否存在下划线，并获取下划线之前的名称
        If InStr(assemblyName, "_") > 0 Then
            assemblyBaseName = Left(assemblyName, InStr(assemblyName, "_") - 1)
        Else
            assemblyBaseName = assemblyName
        End If
        Debug.Print "基础装配体名称: " & assemblyBaseName ' 输出基础装配体名称
        
        For i = 0 To UserForm1.ListBox_映射.ListCount - 1
            mappingEntry = UserForm1.ListBox_映射.List(i) ' 获取映射条目
            mappingParts = Split(mappingEntry, "|") ' 分割映射条目
            ' 检查映射条目是否包含基础装配体名称
            If UBound(mappingParts) >= 2 And Trim(mappingParts(1)) = assemblyBaseName Then
                foundMapping = True
                UserForm1.TextBox_代号.Value = mappingParts(2) ' 设置映射字符
                Exit For
            End If
        Next i
        
        ' 更新Label_映射信息的显示
        With UserForm1.Label_映射信息
            If foundMapping Then
                .caption = "映射成功"
                .ForeColor = RGB(0, 128, 0) ' 绿色
            Else
                .caption = "暂无映射，请确认"
                .ForeColor = RGB(255, 0, 0) ' 红色
                UserForm1.TextBox_代号.Value = "" ' 清除之前的值
            End If
        End With
    Else
        MsgBox "未打开装配体文档。"
    End If
End Sub


' 将文本追加到文件
Public Sub AppendTextToFile(ByVal filePath As String, ByVal text As String)
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open filePath For Append As fileNumber
    Print #fileNumber, text
    Close fileNumber
End Sub



' 这个子程序用于获取当前装配体上级目录下所有文件夹的名称
' 获取当前装配体所在的上级目录下所有子文件夹的名称，并填充到列表框中
Sub GetFoldersAndFillList()
    Dim swApp As Object
    Dim swModel As Object
    Dim asmPath As String
    Dim parentPath As String
    Dim subFolder As Object
    Dim fso As Object
    Dim folder As Object

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not swModel Is Nothing Then
        asmPath = swModel.GetPathName
        If asmPath <> "" Then
            ' 获取文件所在的目录
            parentPath = fso.GetParentFolderName(asmPath)
            ' 获取上级目录
            Set folder = fso.GetFolder(parentPath).parentFolder

            ' 清空现有列表
            UserForm1.ListBox_文件夹.Clear

            ' 获取上级目录下的所有子目录，并添加到列表框中
            For Each subFolder In folder.SubFolders
                UserForm1.ListBox_文件夹.AddItem subFolder.path

            Next subFolder
        Else
            MsgBox "装配体文件没有路径信息。可能是未保存的新文件。"
        End If
    Else
        MsgBox "未打开装配体文档。"
    End If
End Sub



' 这个子程序用于捕获选中的组件信息并显示在用户窗体控件中
Sub CaptureSelectedComponentInfo()
    Dim swApp As Object
    Dim swModel As Object
    Dim swSelMgr As Object
    Dim swComponent As Object
    Dim path As String
    Dim fileName As String
    Dim fso As Object
    Dim fileExtension As String
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If swSelMgr.GetSelectedObjectCount2(-1) > 0 Then
        Set swComponent = swSelMgr.GetSelectedObjectsComponent4(1, -1)
        If Not swComponent Is Nothing Then
            path = swComponent.GetPathName
            fileName = fso.GetBaseName(path)
            fileExtension = fso.GetExtensionName(path)
            
            UserForm1.Label_选中组件路径.caption = path
            UserForm1.TextBox_选中组件名称.Value = fileName
            
            Select Case LCase(fileExtension)
                Case "sldasm"
                    UserForm1.Label_组件类型.caption = "装配体"
                Case "sldprt"
                    UserForm1.Label_组件类型.caption = "零件"
                Case Else
                    UserForm1.Label_组件类型.caption = "未知"
            End Select
            
            If fso.FileExists(Replace(path, fileExtension, "SLDDRW")) Then
                UserForm1.CheckBox_包含工程图.Value = True
            Else
                UserForm1.CheckBox_包含工程图.Value = False
            End If
        Else
            'MsgBox "未选中任何组件。"
        End If
    Else
        'MsgBox "未选中任何组件。"
    End If
End Sub

' 这个子程序用于初始化用户窗体，设置初始状态
Sub InitializeUserForm()
    ' 清空所有控件
    UserForm1.TextBox_代号.Value = ""
    UserForm1.ListBox_文件夹.Clear
    UserForm1.Label_选中组件路径.caption = ""
    UserForm1.TextBox_选中组件名称.Value = ""
    'UserForm1.Label_组件类型.caption = ""
    UserForm1.CheckBox_包含工程图.Value = False
    
    ' 调用函数填充数据
    'Call ConvertAssemblyNameToUpper
    Call GetFoldersAndFillList
End Sub
'获取文件夹下所有文件
Public Function GetFilesInDirectory(directory As String) As Collection
    Dim folder As folder
    Dim file As file
    Dim fs As New FileSystemObject
    Dim files As New Collection

    Set folder = fs.GetFolder(directory)

    For Each file In folder.files
        files.Add file.path
    Next file

    Set GetFilesInDirectory = files
End Function
' 这个子程序用于写入用户修改的钣金件代号、车床件代号、亚克力代号和机加件标识到文本文件中
Public Sub WriteIdentifiers()
    Dim filePath As String
    Dim fileContent As String
    Dim lines As Variant
    Dim i As Integer
    Dim outputLines As String
    Dim identifierFound(1 To 4) As Boolean

    filePath = GetFilePath()
    If filePath <> "" And Dir(filePath) <> "" Then
        fileContent = GetFileContent(filePath)
        lines = Split(fileContent, vbCrLf)

        ' 初始化标识符是否已找到的数组
        identifierFound(1) = False
        identifierFound(2) = False
        identifierFound(3) = False
        identifierFound(4) = False

        ' 遍历文件中的每一行，更新对应的代号行
        For i = LBound(lines) To UBound(lines)
            If InStr(lines(i), "钣金件代号|") > 0 Then
                lines(i) = UserForm1.TextBox_钣金件代号.text
                identifierFound(1) = True
            ElseIf InStr(lines(i), "车床件代号|") > 0 Then
                lines(i) = UserForm1.TextBox_车床件代号.text
                identifierFound(2) = True
            ElseIf InStr(lines(i), "亚克力代号|") > 0 Then
                lines(i) = UserForm1.TextBox_亚克力代号.text
                identifierFound(3) = True
            ElseIf InStr(lines(i), "机加件标识|") > 0 Then
                lines(i) = UserForm1.TextBox_机加件标识.text
                identifierFound(4) = True
            End If
            outputLines = outputLines & lines(i) & vbCrLf
        Next i

        ' 如果有代号未在文件中找到，则添加它们
        If Not identifierFound(1) Then
            outputLines = outputLines & UserForm1.TextBox_钣金件代号.text & vbCrLf
        End If
        If Not identifierFound(2) Then
            outputLines = outputLines & UserForm1.TextBox_车床件代号.text & vbCrLf
        End If
        If Not identifierFound(3) Then
            outputLines = outputLines & UserForm1.TextBox_亚克力代号.text & vbCrLf
        End If
        If Not identifierFound(4) Then
            outputLines = outputLines & UserForm1.TextBox_机加件标识.text & vbCrLf
        End If

        ' 写入文件
        WriteToFile filePath, outputLines

        ' 反馈给用户
        MsgBox "代号保存成功！"
    Else
        MsgBox "配置文件不存在，请检查。"
    End If
End Sub




' 这个子程序用于将更新后的全部文本内容写入到指定的文件路径
Private Sub WriteToFile(ByVal filePath As String, ByVal content As String)
    Dim fileNum As Integer
    fileNum = FreeFile

    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
End Sub
