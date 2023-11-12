下面SolidWorks vba的主要作用是
1. 获取txt文本中以 映射| 开头的行,然后将参数显示在UserForm1的"ListBox_映射"中
2. 当用户点击"CommandButton_写入映射"时,会将参数写入回txt文本,并且重新读取,显示在"ListBox_映射"中
现在我想新增一个写入读取,保存的功能
1. 读取:
	1.1:钣金件代号|文本    展示在Userform1的"TextBox_钣金件代号"中
	1.2:车床件代号|文本    展示在Userform1的"TextBox_车床件代号"中
	1.3:亚克力代号|文本    展示在Userform1的"TextBox_亚克力代号"中

2. 写入:
	当用户修改了这三个文本框的值,并且点击按钮"CommandButton_代号保存",将值保存回txt文本,并且重新读取展示
	
注意:由于这些读取写入在同一个txt文本,所有我们应该以什么方式把它们分开,两种类型的数据不能混在一起,或者说以分隔符还是其它什么方式?
请问你是否完全明白我的意思?请问我的描述是否清楚,有没有模糊的地方?请问还需要我提供其它信息给你吗?如果需要请提出来.
为了确保你完全理解我的需求,请将我的需求详细阐述,请先不要给出代码,我们先沟通清楚再继续写程序

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


' 读取映射文件并填充ListBox_映射
Public Sub FillMappingListBox()
    Dim filePath As String
    Dim fileContent As String
    Dim mappingLines As Variant
    Dim i As Integer
    
    filePath = GetFilePath()
    If filePath <> "" And Dir(filePath) <> "" Then
        fileContent = GetFileContent(filePath)
        mappingLines = Split(fileContent, vbCrLf)
        
        With UserForm1.ListBox_映射
            .Clear
            For i = LBound(mappingLines) To UBound(mappingLines)
                If mappingLines(i) <> "" Then
                    .AddItem mappingLines(i)
                End If
            Next i
        End With
        CheckMappingAndUpdateLabel
    Else
        MsgBox "映射文件不存在，请检查。"
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





根据您提供的信息，我理解您想要实现以下功能：

1. **读取功能**：
   您想要从一个文本文件中读取特定前缀的行，并显示在不同的文本框中。具体来说，这些行以以下格式开始：
   - 钣金件代号|
   - 车床件代号|
   - 亚克力代号|

   相应的文本将被显示在UserForm1上的三个文本框中。

2. **写入功能**：
   当用户在这些文本框中更改了文本并点击一个命名为"CommandButton_代号保存"的按钮时，应用程序将更新文本文件中的相应行，并重新读取显示在文本框中。

3. 您关注数据的组织方式，即如何在文本文件中区分不同类型的数据，以避免混淆。对于这一点，通常我们可以使用不同的标识符（就像您已经使用的"映射|"那样），或者在文件中使用不同的部分来区分不同的数据。

确保数据的高内聚性和低耦合性，我建议继续使用标识符来区分不同类型的数据。每个类型的数据可以有一个唯一的标识符，例如：

- 映射|映射参数
- 钣金件代号|代号
- 车床件代号|代号
- 亚克力代号|代号

这样，当读取或写入数据时，您的程序可以检查每一行的前缀，确定它属于哪种类型，并进行相应的操作。

如果您同意这个方案，我们可以继续讨论如何实现它。或者，如果您有其他的组织方式或者需要澄清的地方，请告诉我。


我的意思是会不会写入后变成这样?因为有时候写入的顺序是不固定的

- 映射|映射参数
- 钣金件代号|代号
- 车床件代号|代号
- 亚克力代号|代号
- 映射|映射参数
- 映射|映射参数

应该也不行吧,因为保存映射的和代号的按钮不是同一个,所以要么是保存映射,要么是保存代号


好的,就按照你这个方法,请问还需要我提供其它信息给你吗?如果需要请提出来,如果你已经认为自己可以写出完整的程序了,那么请提供所有已经修改的子程序的完整代码

注意1:我们使用的是SolidWorks vba编程
注意2:请不要遗漏必要的 变量声明
注意3:请不要遗漏我们原有的功能。请新增或者更改
注意4:请在每一个子程序的开头注释这段子程序的作用



完整代码是这样吗?

' 普通模块代码
Option Explicit
' 在模块级别声明一个字典，用于存储文件中的所有数据
Dim DataDictionary As Object
' 在用户表单初始化或宏开始时调用这个子程序来填充字典
Public Sub InitializeDataDictionary()
    Set DataDictionary = CreateObject("Scripting.Dictionary")
    DataDictionary.Add "映射", New Collection
    DataDictionary.Add "钣金件代号", ""
    DataDictionary.Add "车床件代号", ""
    DataDictionary.Add "亚克力代号", ""
    LoadFileToDictionary GetFilePath()
End Sub
' 根据标识符读取文件并填充字典
Public Sub LoadFileToDictionary(filePath As String)
    Dim fileContent As String
    Dim lines As Variant
    Dim line As Variant
    Dim lineParts As Variant
    Dim identifier As String
    Dim data As String

    fileContent = GetFileContent(filePath)
    If fileContent = "" Then Exit Sub

    lines = Split(fileContent, vbCrLf)

    For Each line In lines
        If line <> "" Then
            lineParts = Split(line, "|")
            identifier = Trim(lineParts(0))
            data = Trim(lineParts(1))
            
            ' 根据标识符将数据添加到正确的字典键
            Select Case identifier
                Case "映射"
                    DataDictionary("映射").Add data
                Case "钣金件代号"
                    DataDictionary("钣金件代号") = data
                Case "车床件代号"
                    DataDictionary("车床件代号") = data
                Case "亚克力代号"
                    DataDictionary("亚克力代号") = data
            End Select
        End If
    Next
End Sub
' 重写文件
Public Sub WriteDictionaryToFile(filePath As String)
    Dim key As Variant
    Dim value As Variant
    Dim i As Integer
    Dim fileContent As String

    fileContent = ""
    
    ' 为映射部分构建文件内容
    For Each value In DataDictionary("映射")
        fileContent = fileContent & "映射|" & value & vbCrLf
    Next value
    
    ' 添加其他部分
    fileContent = fileContent & "钣金件代号|" & DataDictionary("钣金件代号") & vbCrLf
    fileContent = fileContent & "车床件代号|" & DataDictionary("车床件代号") & vbCrLf
    fileContent = fileContent & "亚克力代号|" & DataDictionary("亚克力代号") & vbCrLf
    
    ' 写入文件
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open filePath For Output As fileNumber
    Print #fileNumber, fileContent
    Close fileNumber
End Sub

' 主运行方法 Main
Sub Main()
    ' 初始化用户界面，准备数据填充
    UserForm1.Show 0
    InitializeUserForm
    ' 显示用户界面
    CheckMappingAndUpdateLabel
End Sub

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


' 读取映射文件并填充ListBox_映射
Public Sub FillMappingListBox()
    Dim filePath As String
    Dim fileContent As String
    Dim mappingLines As Variant
    Dim i As Integer
    
    filePath = GetFilePath()
    If filePath <> "" And Dir(filePath) <> "" Then
        fileContent = GetFileContent(filePath)
        mappingLines = Split(fileContent, vbCrLf)
        
        With UserForm1.ListBox_映射
            .Clear
            For i = LBound(mappingLines) To UBound(mappingLines)
                If mappingLines(i) <> "" Then
                    .AddItem mappingLines(i)
                End If
            Next i
        End With
        CheckMappingAndUpdateLabel
    Else
        MsgBox "映射文件不存在，请检查。"
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
                UserForm1.TextBox_代号.value = mappingParts(2) ' 设置映射字符
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
                UserForm1.TextBox_代号.value = "" ' 清除之前的值
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
            UserForm1.TextBox_选中组件名称.value = fileName
            
            Select Case LCase(fileExtension)
                Case "sldasm"
                    UserForm1.Label_组件类型.caption = "装配体"
                Case "sldprt"
                    UserForm1.Label_组件类型.caption = "零件"
                Case Else
                    UserForm1.Label_组件类型.caption = "未知"
            End Select
            
            If fso.FileExists(Replace(path, fileExtension, "SLDDRW")) Then
                UserForm1.CheckBox_包含工程图.value = True
            Else
                UserForm1.CheckBox_包含工程图.value = False
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
    UserForm1.TextBox_代号.value = ""
    UserForm1.ListBox_文件夹.Clear
    UserForm1.Label_选中组件路径.caption = ""
    UserForm1.TextBox_选中组件名称.value = ""
    UserForm1.Label_组件类型.caption = ""
    UserForm1.CheckBox_包含工程图.value = False
    
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


是的你很聪明,你的理解完全正确。请问还需要我提供其它信息给你吗?如果需要请提出来,如果你已经认为自己可以写出完整的程序了,那么请提供所有已经修改的子程序的完整代码

注意1:我们使用的是SolidWorks vba编程
注意2:请不要遗漏必要的 变量声明
注意3:请不要遗漏我们原有的功能。请新增或者更改
注意4:请在每一个子程序的开头注释这段子程序的作用

这个子程序能不再加一个分隔符
比如:
	钣金件代号|双滑块2|BJ
	车床件代号|双滑块4|CC
	亚克力代号|双滑块|YKL
	
这第二个分隔符的内容,我后续需要用去其它的调用

Public Sub ReadIdentifiers()
    Dim filePath As String
    Dim fileContent As String
    Dim lines As Variant
    Dim i As Integer

    filePath = GetFilePath()
    If filePath <> "" And Dir(filePath) <> "" Then
        fileContent = GetFileContent(filePath)
        lines = Split(fileContent, vbCrLf)

        ' 遍历每一行来查找和设置相应的代号
        For i = LBound(lines) To UBound(lines)
            If InStr(lines(i), "钣金件代号|") > 0 Then
                UserForm1.TextBox_钣金件代号.text = Split(lines(i), "|")(1)
            ElseIf InStr(lines(i), "车床件代号|") > 0 Then
                UserForm1.TextBox_车床件代号.text = Split(lines(i), "|")(1)
            ElseIf InStr(lines(i), "亚克力代号|") > 0 Then
                UserForm1.TextBox_亚克力代号.text = Split(lines(i), "|")(1)
            End If
        Next i
    Else
        MsgBox "配置文件不存在，请检查。"
    End If
End Sub


怎么回事,点击写入后,多出来一格数据

钣金件代号|钣金件代号|TBXLA01|BJ
车床件代号|车床件代号|TBXLA01|CC
亚克力代号|亚克力代号|TBXLA01|YKL


' 这个子程序用于写入用户修改的钣金件代号、车床件代号和亚克力代号到文本文件中
Public Sub WriteIdentifiers()
    Dim filePath As String
    Dim fileContent As String
    Dim lines As Variant
    Dim i As Integer
    Dim outputLines As String
    Dim identifierFound(1 To 3) As Boolean

    filePath = GetFilePath()
    If filePath <> "" And Dir(filePath) <> "" Then
        fileContent = GetFileContent(filePath)
        lines = Split(fileContent, vbCrLf)

        ' 初始化标识符是否已找到的数组
        identifierFound(1) = False
        identifierFound(2) = False
        identifierFound(3) = False

        ' 遍历文件中的每一行，更新对应的代号行
        For i = LBound(lines) To UBound(lines)
            Select Case True
                Case InStr(lines(i), "钣金件代号|") > 0
                    lines(i) = "钣金件代号|" & UserForm1.TextBox_钣金件代号.text
                    identifierFound(1) = True
                Case InStr(lines(i), "车床件代号|") > 0
                    lines(i) = "车床件代号|" & UserForm1.TextBox_车床件代号.text
                    identifierFound(2) = True
                Case InStr(lines(i), "亚克力代号|") > 0
                    lines(i) = "亚克力代号|" & UserForm1.TextBox_亚克力代号.text
                    identifierFound(3) = True
            End Select
            outputLines = outputLines & lines(i) & vbCrLf
        Next i

        ' 如果有代号未在文件中找到，则添加它们
        If Not identifierFound(1) Then
            outputLines = outputLines & "钣金件代号|" & UserForm1.TextBox_钣金件代号.text & vbCrLf
        End If
        If Not identifierFound(2) Then
            outputLines = outputLines & "车床件代号|" & UserForm1.TextBox_车床件代号.text & vbCrLf
        End If
        If Not identifierFound(3) Then
            outputLines = outputLines & "亚克力代号|" & UserForm1.TextBox_亚克力代号.text & vbCrLf
        End If

        ' 写入文件
        WriteToFile filePath, outputLines

        ' 反馈给用户
        MsgBox "代号保存成功！"
    Else
        MsgBox "配置文件不存在，请检查。"
    End If
End Sub

非常好，现在我还需要增加一个数据读取和写入
机加件标识|X
然后展示在"TextBox_机加件标识"中
点击写入也是一样的写入回txt文本
请问你是否完全明白我的意思?请问我的描述是否清楚,有没有模糊的地方?请问还需要我提供其它信息给你吗?如果需要请提出来.
为了确保你完全理解我的需求,请将我的需求详细阐述,请先不要给出代码,我们先沟通清楚再继续写程序
是的你很聪明,你的理解完全正确。请问还需要我提供其它信息给你吗?如果需要请提出来,如果你已经认为自己可以写出完整的程序了,那么请提供所有已经修改的子程序的完整代码

注意1:我们使用的是SolidWorks vba编程
注意2:请不要遗漏必要的 变量声明
注意3:请不要遗漏我们原有的功能。请新增或者更改
注意4:为了防止你遗忘了代码,我再次同步给你

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
            End If
        Next i
    Else
        MsgBox "配置文件不存在，请检查。"
    End If
End Sub

' 这个子程序用于写入用户修改的钣金件代号、车床件代号和亚克力代号到文本文件中
Public Sub WriteIdentifiers()
    Dim filePath As String
    Dim fileContent As String
    Dim lines As Variant
    Dim i As Integer
    Dim outputLines As String
    Dim identifierFound(1 To 3) As Boolean

    filePath = GetFilePath()
    If filePath <> "" And Dir(filePath) <> "" Then
        fileContent = GetFileContent(filePath)
        lines = Split(fileContent, vbCrLf)

        ' 初始化标识符是否已找到的数组
        identifierFound(1) = False
        identifierFound(2) = False
        identifierFound(3) = False

        ' 遍历文件中的每一行，更新对应的代号行
        For i = LBound(lines) To UBound(lines)
            If InStr(lines(i), "钣金件代号|") > 0 Then
                lines(i) = UserForm1.TextBox_钣金件代号.text ' 直接使用文本框中的文本
                identifierFound(1) = True
            ElseIf InStr(lines(i), "车床件代号|") > 0 Then
                lines(i) = UserForm1.TextBox_车床件代号.text ' 直接使用文本框中的文本
                identifierFound(2) = True
            ElseIf InStr(lines(i), "亚克力代号|") > 0 Then
                lines(i) = UserForm1.TextBox_亚克力代号.text ' 直接使用文本框中的文本
                identifierFound(3) = True
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

        ' 写入文件
        WriteToFile filePath, outputLines

        ' 反馈给用户
        MsgBox "代号保存成功！"
    Else
        MsgBox "配置文件不存在，请检查。"
    End If
End Sub



















我现在有一个SolidWorks vba程序需要新增一个功能
当用户点击UserForm1中不同的optionbuttn时,将"TextBox_钣金件代号","TextBox_车床件代号","TextBox_亚克力代号"中值第一组数据,也就是第一个分隔符后的数据展示在UserForm1的"TextBox_代号"

在之前我们已经写完了代码,将txt数据的值展示在"TextBox_钣金件代号","TextBox_车床件代号","TextBox_亚克力代号"
	举例,下面是列表框分别展示的值
	钣金件代号|TBXLA01|BJ
	车床件代号|TBXLA02|CC
	亚克力代号|TBXLA03|YKL
	
	那么用户点击"OptionButton_钣金件"时,将TBXLA01展示在“TextBox_代号”
	那么用户点击"OptionButton_车床件"时,将TBXLA02展示在“TextBox_代号”
	那么用户点击"OptionButton_亚克力"时,将TBXLA03展示在“TextBox_代号”
	
请问你是否完全明白我的意思?请问我的描述是否清楚,有没有模糊的地方?请问还需要我提供其它信息给你吗?如果需要请提出来.
为了确保你完全理解我的需求,请将我的需求详细阐述,请先不要给出代码,我们先沟通清楚再继续写程序

请注意,"UpdateSelection"子程序我们不用管,我们需要新增一个子程序,然后在调用"UpdateSelection"后调用,也就是说需要调用两个程序
下面是点击"OptionButton"的代码
Private Sub OptionButton_钣金件_Click()
    If Not isUpdating Then Call UpdateSelection("钣金件")
End Sub

Private Sub OptionButton_车床件_Click()
    If Not isUpdating Then Call UpdateSelection("车床件")
End Sub

Private Sub OptionButton_亚克力_Click()
    If Not isUpdating Then Call UpdateSelection("亚克力")
End Sub


用户点击“OptionButton_钣金件”，应从“TextBox_钣金件代号”中提取“TBXLA01”并显示在“TextBox_代号”中。
用户点击“OptionButton_车床件”，应从“TextBox_车床件代号”中提取“TBXLA02”并显示在“TextBox_代号”中。
用户点击“OptionButton_亚克力”，应从“TextBox_亚克力代号”中提取“TBXLA03”并显示在“TextBox_代号”中。
每个文本框中的数据格式似乎都遵循相同的模式：“[分类名称]|[代号]|[缩写]”，您希望提取的是位于第一个分隔符“|”后面的代号部分。

您提到已经存在一个名为“UpdateSelection”的子程序，且不需要修改它。相反，您需要另一个新的子程序，在调用“UpdateSelection”后调用，来处理和显示数据。

请确认我对需求的理解是否准确，或者是否有其他的要点和细节需要考虑。如果我的理解准确，我们可以继续讨论如何实现这一功能。

是的你很聪明,你的理解完全正确。请问还需要我提供其它信息给你吗?如果需要请提出来,如果你已经认为自己可以写出完整的程序了,那么请提供所有已经修改的子程序的完整代码

注意1:我们使用的是SolidWorks vba编程
注意2:请不要遗漏必要的 变量声明
注意3:请不要遗漏我们原有的功能。请新增或者更改
注意4:请在每一个子程序的开头注释这段子程序的作用


非常好,我们成功了
现在下面子程序是"ListBox_文件"的文件中获取前缀的,现在请帮我将下面子程序"UpdateFileTypeTextBox"改为:
当用户点击不同的OptionButton时,将第二个数据传递到"prefixCharacter"中作为前缀
	举例,下面是列表框分别展示的值
	钣金件代号|TBXLA01|BJ
	车床件代号|TBXLA02|CC
	亚克力代号|TBXLA03|YKL
	
	那么用户点击"OptionButton_钣金件"时,将"BJ"传递到"prefixCharacter"中作为前缀
	那么用户点击"OptionButton_车床件"时,将"CC"传递到"prefixCharacter"中作为前缀
	那么用户点击"OptionButton_亚克力"时,将"YKL"传递到"prefixCharacter"中作为前缀
	
请问你是否完全明白我的意思?请问我的描述是否清楚,有没有模糊的地方?请问还需要我提供其它信息给你吗?如果需要请提出来.
为了确保你完全理解我的需求,请将我的需求详细阐述,请先不要给出代码,我们先沟通清楚再继续写程序
	
Private Sub UpdateFileTypeTextBox()
    Dim maxNumber As Integer
    Dim currentNumber As Integer
    Dim i As Integer
    Dim fileName As String
    Dim baseCode As String
    Dim prefixAndNumber As Variant
    Dim prefixCharacter As String
    
    baseCode = UserForm1.TextBox_代号.text
    maxNumber = 0 ' 初始化最大数字为0
    prefixCharacter = "" ' 初始化前缀字符为空字符串

    Debug.Print "开始分析文件名列表..."

    ' 遍历ListBox_文件中的每个文件名
    For i = 0 To UserForm1.ListBox_文件.ListCount - 1
        fileName = UserForm1.ListBox_文件.List(i)
        ' 检查文件名是否以TextBox_代号的内容开始，并跟随数字
        If fileName Like baseCode & "*[0-9]*" Then
            ' 提取数字部分和前缀字符
            prefixAndNumber = ExtractPrefixAndNumberFromFilename(fileName, baseCode)
            If IsNumeric(prefixAndNumber(1)) Then
                currentNumber = CInt(prefixAndNumber(1))
                ' 更新最大数字和前缀字符
                If currentNumber > maxNumber Then
                    maxNumber = currentNumber
                    prefixCharacter = prefixAndNumber(0) '储存数字前缀,比如X
                End If
                Debug.Print "当前文件名: " & fileName & "; 提取数字: " & currentNumber
            End If
        End If
    Next i

    ' 如果找到符合条件的数字，则设置TextBox_类型的值
    If maxNumber > 0 Then
        UserForm1.TextBox_类型.text = prefixCharacter & Format(maxNumber + 1, "00")
    Else
        UserForm1.TextBox_类型.text = "" ' 如果没有找到，保持TextBox_类型为空
    End If

    Debug.Print "最大数字: " & maxNumber
    Debug.Print "TextBox_类型设置为: " & UserForm1.TextBox_类型.text
End Sub

1. 我们是根据"OptionButton"当前的value是否为true来决定的
2. 请注意，这个值不是固定的，而是根据第二个参数来取值的
2. 还有一个"OptionButton_机加件",当这个"OptionButton"value为true时,将从"TextBox_机加件标识"中获取第一个参数作为变量传递,注意了,这个不是第二个参数
	举例:TextBox_机加件标识中的值为:"机加件标识|X",那么就是将X传递到变量了
请问你是否完全明白我的意思?请问我的描述是否清楚,有没有模糊的地方?请问还需要我提供其它信息给你吗?如果需要请提出来.

有一点需要澄清一下,不是用户点击"OptionButton"来确实"prefixCharacter"的值。
而是根据当前哪一个"OptionButton"的value是为true来确定的

是的你很聪明,你的理解完全正确。请问还需要我提供其它信息给你吗?如果需要请提出来,如果你已经认为自己可以写出完整的程序了,那么请提供所有已经修改的子程序的完整代码

注意1:我们使用的是SolidWorks vba编程
注意2:请不要遗漏必要的 变量声明
注意3:请不要遗漏我们原有的功能。请新增或者更改
注意4:请在每一个子程序的开头注释这段子程序的作用

为了防止你遗忘,我再次把子程序"UpdateFileTypeTextBox"同步给你
Private Sub UpdateFileTypeTextBox()
    Dim maxNumber As Integer
    Dim currentNumber As Integer
    Dim i As Integer
    Dim fileName As String
    Dim baseCode As String
    Dim prefixAndNumber As Variant
    Dim prefixCharacter As String
    
    baseCode = UserForm1.TextBox_代号.text
    maxNumber = 0 ' 初始化最大数字为0
    prefixCharacter = "" ' 初始化前缀字符为空字符串

    Debug.Print "开始分析文件名列表..."

    ' 遍历ListBox_文件中的每个文件名
    For i = 0 To UserForm1.ListBox_文件.ListCount - 1
        fileName = UserForm1.ListBox_文件.List(i)
        ' 检查文件名是否以TextBox_代号的内容开始，并跟随数字
        If fileName Like baseCode & "*[0-9]*" Then
            ' 提取数字部分和前缀字符
            prefixAndNumber = ExtractPrefixAndNumberFromFilename(fileName, baseCode)
            If IsNumeric(prefixAndNumber(1)) Then
                currentNumber = CInt(prefixAndNumber(1))
                ' 更新最大数字和前缀字符
                If currentNumber > maxNumber Then
                    maxNumber = currentNumber
                    prefixCharacter = prefixAndNumber(0) '储存数字前缀,比如X
                End If
                Debug.Print "当前文件名: " & fileName & "; 提取数字: " & currentNumber
            End If
        End If
    Next i

    ' 如果找到符合条件的数字，则设置TextBox_类型的值
    If maxNumber > 0 Then
        UserForm1.TextBox_类型.text = prefixCharacter & Format(maxNumber + 1, "00")
    Else
        UserForm1.TextBox_类型.text = "" ' 如果没有找到，保持TextBox_类型为空
    End If

    Debug.Print "最大数字: " & maxNumber
    Debug.Print "TextBox_类型设置为: " & UserForm1.TextBox_类型.text
End Sub



我不是说了吗?这个前缀不是绝对值,你这里为什么还是把它设为绝对值了?我们应该和"OptionButton_机加件"一样,指定的第三个参数获取值
这里的举例不是让你把BJ,CC,YKL作为变量传递,而是把第二个分隔符,也就是第二个数据作为变量传递,因为这个值经常修改的
	举例,下面是列表框分别展示的值
	钣金件代号|TBXLA01|BJ
	车床件代号|TBXLA02|CC
	亚克力代号|TBXLA03|YKL
	
	那么用户点击"OptionButton_钣金件"时,将"BJ"传递到"prefixCharacter"中作为前缀
	那么用户点击"OptionButton_车床件"时,将"CC"传递到"prefixCharacter"中作为前缀
	那么用户点击"OptionButton_亚克力"时,将"YKL"传递到"prefixCharacter"中作为前缀


' 这个函数检查当前选中的OptionButton并返回相应的前缀字符
Function GetSelectedPrefix() As String
    Dim prefixCharacter As String
    
    ' 检查每个OptionButton的value属性，确定哪一个是True
    If UserForm1.OptionButton_钣金件.Value = True Then
        prefixCharacter = "BJ"  ' 钣金件的前缀
    ElseIf UserForm1.OptionButton_车床件.Value = True Then
        prefixCharacter = "CC"  ' 车床件的前缀
    ElseIf UserForm1.OptionButton_亚克力.Value = True Then
        prefixCharacter = "YKL" ' 亚克力的前缀
    ElseIf UserForm1.OptionButton_机加件.Value = True Then
        ' 从TextBox_机加件标识中获取标识符，假设格式是"机加件标识|X"
        prefixCharacter = Split(UserForm1.TextBox_机加件标识.Text, "|")(1)
    Else
        prefixCharacter = "" ' 如果没有选中的按钮，返回空字符串
    End If
    
    GetSelectedPrefix = prefixCharacter
End Function
prefixCharacter = Split(UserForm1.TextBox_亚克力代号.text, "|")(1)


' 这个函数检查当前选中的OptionButton并返回相应的前缀字符


下面这个SolidWorks 子程序提取的数字是不是不对啊?
这段提取中最大值明明是11,为什么程序找到的最大值却是9?
	当前文件名: TBXLA01-CC-09[主动轮轴].SLDPRT; 提取数字: 9
	 13 
	当前文件名: TBXLA01-CC-10[从动轮轴].SLDPRT; 提取数字: 0
	 13 
	当前文件名: TBXLA01-CC-11[标签转轴].SLDPRT; 提取数字: 1
	最大数字: 9
	TextBox_类型设置为: -10
	
' 这个子程序更新TextBox_类型的值，基于ListBox_文件中的文件名以及当前选中的OptionButton
Private Sub UpdateFileTypeTextBox()
    Dim maxNumber As Integer
    Dim currentNumber As Integer
    Dim i As Integer
    Dim fileName As String
    Dim baseCode As String
    Dim prefixAndNumber As Variant
    Dim prefixCharacter As String
    
    baseCode = UserForm1.TextBox_代号.text
    maxNumber = 0 ' 初始化最大数字为0
    
    ' 获取当前选中的OptionButton的前缀字符
    prefixCharacter = GetSelectedPrefix()

    Debug.Print "开始分析文件名列表..."

    ' 遍历ListBox_文件中的每个文件名
    For i = 0 To UserForm1.ListBox_文件.ListCount - 1
        fileName = UserForm1.ListBox_文件.List(i)
        ' 检查文件名是否以TextBox_代号的内容开始，并跟随数字
        If fileName Like baseCode & "*[0-9]*" Then
            ' 提取数字部分和前缀字符
            prefixAndNumber = ExtractPrefixAndNumberFromFilename(fileName, baseCode)
            If IsNumeric(prefixAndNumber(1)) Then
                currentNumber = CInt(prefixAndNumber(1))
                ' 更新最大数字
                If currentNumber > maxNumber Then
                    maxNumber = currentNumber
                    ' 不再更新prefixCharacter，因为它是由GetSelectedPrefix提供
                End If
                Debug.Print "当前文件名: " & fileName & "; 提取数字: " & currentNumber
            End If
        End If
    Next i

    ' 如果找到符合条件的数字，则设置TextBox_类型的值
    If maxNumber > 0 Then
        UserForm1.TextBox_类型.text = prefixCharacter & Format(maxNumber + 1, "00")
    Else
        UserForm1.TextBox_类型.text = "" ' 如果没有找到，保持TextBox_类型为空
    End If

    Debug.Print "最大数字: " & maxNumber
    Debug.Print "TextBox_类型设置为: " & UserForm1.TextBox_类型.text
End Sub


' 这个函数从文件名中提取前缀字符和数字
Private Function ExtractPrefixAndNumberFromFilename(fileName As String, baseCode As String) As Variant
    Dim startPos As Integer
    Dim endPos As Integer
    Dim numberPart As String
    Dim prefixCharacter As String
    Dim result(1) As String
    
    ' 查找baseCode后的第一个数字的位置
    startPos = InStr(fileName, baseCode) + Len(baseCode)
    While Not Mid(fileName, startPos, 1) Like "[0-9]" And startPos <= Len(fileName)
        startPos = startPos + 1
    Wend
    ' 如果没有找到数字，则退出函数
    If startPos > Len(fileName) Then
        result(0) = ""
        result(1) = ""
        ExtractPrefixAndNumberFromFilename = result
        Exit Function
    End If
    
    ' 提取前缀字符（如果存在）
    prefixCharacter = Mid(fileName, startPos - 1, 1)
    If Not prefixCharacter Like "[0-9]" Then
        result(0) = prefixCharacter
        startPos = startPos + 1 ' 跳过前缀字符
    End If
    Debug.Print startPos
    ' 查找数字部分的结束位置（即下一个非数字字符的位置）
    endPos = startPos
    While Mid(fileName, endPos, 1) Like "[0-9]" And endPos <= Len(fileName)
        endPos = endPos + 1
    Wend

    ' 提取数字部分
    numberPart = Mid(fileName, startPos, endPos - startPos)
    result(1) = numberPart
    
    ExtractPrefixAndNumberFromFilename = result
End Function