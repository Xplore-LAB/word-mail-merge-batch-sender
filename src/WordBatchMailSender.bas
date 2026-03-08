Attribute VB_Name = "模块3"
Option Explicit

Sub BatchMailMergeSend()

    Dim ds As MailMergeDataSource
    Dim recCount As Long, i As Long
    Dim validCount As Long, blockedCount As Long, warnCount As Long
    Dim attachCount As Long, missingAttachCount As Long
    Dim successCount As Long, failCount As Long, draftCount As Long, skipCount As Long
    
    Dim OutlookApp As Object, OutlookNS As Object, OutlookAccounts As Object, SendAccount As Object
    Dim OutlookMail As Object, objDoc As Object
    Dim fso As Object
    Dim mergedDoc As Document
    
    Dim defaultSender As String, defaultSubject As String
    Dim subjectText As String
    Dim sendMode As String, sendModeText As String
    
    Dim accList As String, accChoice As String
    Dim accIndex As Long, defaultAccIndex As Long, j As Long
    
    Dim hasEmailField As Boolean, hasAttachmentField As Boolean
    Dim emailAddr As String, filePath As String
    
    Dim overallStatus As String, issueType As String, issueDetail As String
    Dim suggestion As String, canContinue As String
    
    Dim confirmMsg As String
    Dim userChoice As VbMsgBoxResult
    Dim fatalStop As Boolean
    Dim issueLog As String
    
    ' 默认值：可按需要修改
    defaultSender = "your_account@example.com"
    defaultSubject = "请查收本次通知"
    
    recCount = 0
    validCount = 0
    blockedCount = 0
    warnCount = 0
    attachCount = 0
    missingAttachCount = 0
    successCount = 0
    failCount = 0
    draftCount = 0
    skipCount = 0
    fatalStop = False
    issueLog = ""
    
    hasEmailField = False
    hasAttachmentField = False
    
    ' =========================
    ' 1. 基础环境检查
    ' =========================
    On Error Resume Next
    
    If ActiveDocument.Path = "" Then
        MsgBox "请先将文档保存为“启用宏的 Word 文档（.docm）”后再运行。" & vbCrLf & vbCrLf & _
               "建议解决方案：" & vbCrLf & _
               "1. 点击“文件”→“另存为”" & vbCrLf & _
               "2. 选择“启用宏的 Word 文档（*.docm）”" & vbCrLf & _
               "3. 保存后重新运行。", vbCritical, "环境检查失败"
        Exit Sub
    End If
    
    Set ds = ActiveDocument.MailMerge.DataSource
    If Err.Number <> 0 Or ds.ConnectString = "" Then
        MsgBox "未检测到数据源。" & vbCrLf & vbCrLf & _
               "建议解决方案：" & vbCrLf & _
               "1. 打开“邮件”选项卡" & vbCrLf & _
               "2. 点击“选择收件人”重新绑定 Excel 数据源" & vbCrLf & _
               "3. 确认数据源绑定成功后再运行。", vbCritical, "数据源检查失败"
        Exit Sub
    End If
    
    ' 检查邮箱字段是否存在（支持多种别名）
    Err.Clear
    emailAddr = GetFieldValueSafe(ds, _
                                  "收件邮箱", "邮箱", "邮箱地址", _
                                  "Email", "EmailAddress", "Email_Address", _
                                  "Email Addresses", "Email_Addresses")
    If Err.Number = 0 Then
        If HasAnyField(ds, _
                       "收件邮箱", "邮箱", "邮箱地址", _
                       "Email", "EmailAddress", "Email_Address", _
                       "Email Addresses", "Email_Addresses") Then
            hasEmailField = True
        End If
    End If
    
    ' 检查附件字段是否存在（支持多种别名）
    If HasAnyField(ds, _
                   "附件路径", "附件", "文件路径", _
                   "Attachment", "AttachmentPath", "FilePath") Then
        hasAttachmentField = True
    End If
    
    On Error GoTo 0
    
    If Not hasEmailField Then
        MsgBox "未找到邮箱字段。" & vbCrLf & vbCrLf & _
               "程序支持以下任意一个表头名称：" & vbCrLf & _
               "收件邮箱 / 邮箱 / 邮箱地址 / Email / EmailAddress / Email_Address / Email Addresses / Email_Addresses" & vbCrLf & vbCrLf & _
               "建议解决方案：" & vbCrLf & _
               "请检查 Excel 表头名称，并重新绑定数据源。", vbCritical, "字段检查失败"
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' =========================
    ' 2. 初始化 Outlook
    ' =========================
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNS = OutlookApp.GetNamespace("MAPI")
    Set OutlookAccounts = OutlookNS.Accounts
    
    If OutlookAccounts Is Nothing Or OutlookAccounts.Count = 0 Then
        MsgBox "未检测到可用的 Outlook 账号。" & vbCrLf & vbCrLf & _
               "建议解决方案：" & vbCrLf & _
               "1. 请先打开 Outlook" & vbCrLf & _
               "2. 确认已登录至少一个邮箱账号" & vbCrLf & _
               "3. 再重新运行本程序。", vbCritical, "无可用发件账号"
        Exit Sub
    End If
    
    ' =========================
    ' 3. 选择发件账号
    ' =========================
    defaultAccIndex = 1
    For j = 1 To OutlookAccounts.Count
        If LCase(Trim(OutlookAccounts.Item(j).SmtpAddress)) = LCase(Trim(defaultSender)) Then
            defaultAccIndex = j
            Exit For
        End If
    Next j
    
    If OutlookAccounts.Count = 1 Then
        accIndex = 1
        Set SendAccount = OutlookAccounts.Item(1)
    Else
        accList = "检测到以下 Outlook 发件账号：" & vbCrLf & vbCrLf
        For j = 1 To OutlookAccounts.Count
            accList = accList & j & ". " & OutlookAccounts.Item(j).SmtpAddress
            If j = defaultAccIndex Then accList = accList & " （默认）"
            accList = accList & vbCrLf
        Next j
        
        accChoice = InputBox(accList & vbCrLf & "请输入要使用的发件账号序号：", _
                             "选择发件账号", _
                             CStr(defaultAccIndex))
        
        If Trim(accChoice) = "" Then Exit Sub
        
        If Not IsNumeric(accChoice) Then
            MsgBox "输入的账号序号无效，请输入数字。", vbCritical, "账号选择错误"
            Exit Sub
        End If
        
        accIndex = CLng(accChoice)
        If accIndex < 1 Or accIndex > OutlookAccounts.Count Then
            MsgBox "账号序号超出范围。", vbCritical, "账号选择错误"
            Exit Sub
        End If
        
        Set SendAccount = OutlookAccounts.Item(accIndex)
    End If
    
    ' =========================
    ' 4. 标题与模式
    ' =========================
    subjectText = InputBox("请输入本次邮件标题：", "设置邮件标题", defaultSubject)
    If Trim(subjectText) = "" Then
        MsgBox "邮件标题不能为空，已取消。", vbExclamation, "标题为空"
        Exit Sub
    End If
    
    sendMode = InputBox("请选择发送模式：" & vbCrLf & _
                        "1 = 直接发送" & vbCrLf & _
                        "2 = 仅生成草稿", _
                        "发送模式", _
                        "2")
    If Trim(sendMode) = "" Then Exit Sub
    
    If sendMode <> "1" And sendMode <> "2" Then
        MsgBox "发送模式输入无效，请输入 1 或 2。", vbCritical, "模式选择错误"
        Exit Sub
    End If
    
    If sendMode = "1" Then
        sendModeText = "直接发送"
    Else
        sendModeText = "仅生成草稿"
    End If
    
    ' =========================
    ' 5. 发送前扫描并汇总
    ' =========================
    recCount = GetRecordCount(ds)
    If recCount <= 0 Then
        MsgBox "未读取到任何邮件合并记录。", vbExclamation, "无记录"
        Exit Sub
    End If
    
    For i = 1 To recCount
        ds.ActiveRecord = i
        
        emailAddr = GetFieldValueSafe(ds, _
                                      "收件邮箱", "邮箱", "邮箱地址", _
                                      "Email", "EmailAddress", "Email_Address", _
                                      "Email Addresses", "Email_Addresses")
        filePath = ""
        If hasAttachmentField Then
            filePath = GetFieldValueSafe(ds, _
                                         "附件路径", "附件", "文件路径", _
                                         "Attachment", "AttachmentPath", "FilePath")
        End If
        
        EvaluateRecordStatus emailAddr, filePath, hasAttachmentField, fso, _
                             overallStatus, issueType, issueDetail, suggestion, canContinue
        
        If Trim(emailAddr) <> "" Then validCount = validCount + 1
        
        If Trim(filePath) <> "" Then
            attachCount = attachCount + 1
            If Not fso.FileExists(filePath) Then missingAttachCount = missingAttachCount + 1
        End If
        
        Select Case overallStatus
            Case "PASS"
                ' 无需处理
            Case "WARNING"
                warnCount = warnCount + 1
            Case "BLOCKED"
                blockedCount = blockedCount + 1
        End Select
    Next i
    
    confirmMsg = "请确认本次发送信息：" & vbCrLf & vbCrLf & _
                 "发送模式：" & sendModeText & vbCrLf & _
                 "发件账号：" & SendAccount.SmtpAddress & vbCrLf & _
                 "邮件标题：" & subjectText & vbCrLf & vbCrLf & _
                 "扫描结果：" & vbCrLf & _
                 "总记录数：" & recCount & vbCrLf & _
                 "有效邮箱数：" & validCount & vbCrLf & _
                 "带附件记录数：" & attachCount & vbCrLf & _
                 "警告记录数：" & warnCount & vbCrLf & _
                 "阻断记录数：" & blockedCount & vbCrLf & _
                 "附件缺失记录数：" & missingAttachCount & vbCrLf & vbCrLf & _
                 "说明：" & vbCrLf & _
                 "1. 警告记录一般是附件缺失，可由你判断是否继续；" & vbCrLf & _
                 "2. 阻断记录一般是邮箱为空，无法正常发送；" & vbCrLf & _
                 "3. 执行过程中，程序会对异常逐条提示你处理。" & vbCrLf & vbCrLf & _
                 "是否继续执行？"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion + vbDefaultButton2, "发送前确认") = vbNo Then
        MsgBox "已取消执行。", vbInformation, "操作取消"
        Exit Sub
    End If
    
    ' =========================
    ' 6. 逐条执行
    ' =========================
    For i = 1 To recCount
        
        ds.ActiveRecord = i
        
        emailAddr = GetFieldValueSafe(ds, _
                                      "收件邮箱", "邮箱", "邮箱地址", _
                                      "Email", "EmailAddress", "Email_Address", _
                                      "Email Addresses", "Email_Addresses")
        filePath = ""
        If hasAttachmentField Then
            filePath = GetFieldValueSafe(ds, _
                                         "附件路径", "附件", "文件路径", _
                                         "Attachment", "AttachmentPath", "FilePath")
        End If
        
        EvaluateRecordStatus emailAddr, filePath, hasAttachmentField, fso, _
                             overallStatus, issueType, issueDetail, suggestion, canContinue
        
        If overallStatus = "BLOCKED" Then
            userChoice = MsgBox( _
                "第 " & i & " 条记录无法正常执行。" & vbCrLf & vbCrLf & _
                "问题类型：" & issueType & vbCrLf & _
                "问题详情：" & issueDetail & vbCrLf & vbCrLf & _
                "建议解决方案：" & vbCrLf & suggestion & vbCrLf & vbCrLf & _
                "选择“是”=跳过当前记录继续" & vbCrLf & _
                "选择“否”=终止全部任务", _
                vbYesNo + vbExclamation, _
                "发现阻断问题")
            
            If userChoice = vbYes Then
                skipCount = skipCount + 1
                issueLog = issueLog & "【跳过】第 " & i & " 条：" & issueType & vbCrLf
                GoTo NextRecord
            Else
                fatalStop = True
                issueLog = issueLog & "【终止】第 " & i & " 条：" & issueType & "，用户终止任务。" & vbCrLf
                Exit For
            End If
        ElseIf overallStatus = "WARNING" Then
            userChoice = MsgBox( _
                "第 " & i & " 条记录存在警告。" & vbCrLf & vbCrLf & _
                "问题类型：" & issueType & vbCrLf & _
                "问题详情：" & issueDetail & vbCrLf & vbCrLf & _
                "建议解决方案：" & vbCrLf & suggestion & vbCrLf & vbCrLf & _
                "选择“是”=继续执行当前记录" & vbCrLf & _
                "选择“否”=终止全部任务", _
                vbYesNo + vbQuestion, _
                "发现警告")
            
            If userChoice = vbNo Then
                fatalStop = True
                issueLog = issueLog & "【终止】第 " & i & " 条：WARNING，用户终止任务。" & vbCrLf
                Exit For
            End If
            
            If issueType = "附件缺失" Then
                filePath = ""
                issueLog = issueLog & "【警告后继续】第 " & i & " 条：附件缺失，改为无附件执行。" & vbCrLf
            End If
        End If
        
        On Error GoTo SendErr
        
        ActiveDocument.MailMerge.Destination = wdSendToNewDocument
        ActiveDocument.MailMerge.DataSource.FirstRecord = ds.ActiveRecord
        ActiveDocument.MailMerge.DataSource.LastRecord = ds.ActiveRecord
        ActiveDocument.MailMerge.Execute Pause:=False
        
        Set mergedDoc = ActiveDocument
        Set OutlookMail = OutlookApp.CreateItem(0)
        
        With OutlookMail
            .Display
            Set .SendUsingAccount = SendAccount
            .To = emailAddr
            .Subject = subjectText
            
            Set objDoc = .GetInspector.WordEditor
            mergedDoc.Range.Copy
            objDoc.Range.Paste
            
            If Trim(filePath) <> "" Then
                .Attachments.Add filePath
            End If
            
            If sendMode = "1" Then
                .Send
                successCount = successCount + 1
            Else
                .Save
                draftCount = draftCount + 1
            End If
        End With
        
        mergedDoc.Close SaveChanges:=wdDoNotSaveChanges
        Set mergedDoc = Nothing
        Set OutlookMail = Nothing
        Set objDoc = Nothing
        
        On Error GoTo 0
        GoTo NextRecord
        
SendErr:
        userChoice = MsgBox( _
            "第 " & i & " 条记录处理失败。" & vbCrLf & _
            "收件人：" & emailAddr & vbCrLf & _
            "错误信息：" & Err.Description & vbCrLf & vbCrLf & _
            "建议解决方案：" & vbCrLf & _
            "1. 检查 Outlook 是否已正常打开；" & vbCrLf & _
            "2. 检查发件账号是否可正常发信；" & vbCrLf & _
            "3. 检查当前记录数据是否异常；" & vbCrLf & _
            "4. 必要时先用“仅生成草稿”模式测试。" & vbCrLf & vbCrLf & _
            "选择“是”=跳过当前记录继续" & vbCrLf & _
            "选择“否”=终止全部任务", _
            vbYesNo + vbCritical, _
            "发送失败")
        
        On Error Resume Next
        If Not mergedDoc Is Nothing Then mergedDoc.Close SaveChanges:=wdDoNotSaveChanges
        Set mergedDoc = Nothing
        Set OutlookMail = Nothing
        Set objDoc = Nothing
        On Error GoTo 0
        
        If userChoice = vbYes Then
            failCount = failCount + 1
            issueLog = issueLog & "【失败】第 " & i & " 条：" & Err.Description & vbCrLf
        Else
            failCount = failCount + 1
            fatalStop = True
            issueLog = issueLog & "【终止】第 " & i & " 条：" & Err.Description & "，用户终止任务。" & vbCrLf
            Exit For
        End If
        
NextRecord:
    Next i
    
    ' =========================
    ' 7. 结果汇总
    ' =========================
    confirmMsg = "任务已结束。" & vbCrLf & vbCrLf & _
                 "模式：" & sendModeText & vbCrLf & _
                 "成功发送数：" & successCount & vbCrLf & _
                 "生成草稿数：" & draftCount & vbCrLf & _
                 "失败数：" & failCount & vbCrLf & _
                 "跳过数：" & skipCount
    
    If fatalStop Then
        confirmMsg = confirmMsg & vbCrLf & vbCrLf & "任务状态：用户中途终止。"
    Else
        confirmMsg = confirmMsg & vbCrLf & vbCrLf & "任务状态：执行完成。"
    End If
    
    If issueLog <> "" Then
        confirmMsg = confirmMsg & vbCrLf & vbCrLf & "问题记录（部分）：" & vbCrLf & Left(issueLog, 1800)
    End If
    
    MsgBox confirmMsg, vbInformation, "执行结果汇总"

End Sub


Private Sub EvaluateRecordStatus(ByVal emailAddr As String, _
                                 ByVal filePath As String, _
                                 ByVal hasAttachmentField As Boolean, _
                                 ByVal fso As Object, _
                                 ByRef overallStatus As String, _
                                 ByRef issueType As String, _
                                 ByRef issueDetail As String, _
                                 ByRef suggestion As String, _
                                 ByRef canContinue As String)

    If Trim(emailAddr) = "" Then
        overallStatus = "BLOCKED"
        issueType = "邮箱为空"
        issueDetail = "当前记录未提供收件人邮箱地址，无法发送。"
        suggestion = "请在数据源中补全邮箱地址后再发送。"
        canContinue = "否"
        Exit Sub
    End If
    
    If Trim(filePath) <> "" Then
        If Not fso.FileExists(filePath) Then
            overallStatus = "WARNING"
            issueType = "附件缺失"
            issueDetail = "附件路径已填写，但文件不存在。"
            suggestion = "请检查附件路径、文件名，或确认是否允许无附件发送。"
            canContinue = "是"
            Exit Sub
        End If
    End If
    
    overallStatus = "PASS"
    issueType = "无"
    issueDetail = "检查通过。"
    suggestion = "无需处理。"
    canContinue = "是"
End Sub


Private Function GetFieldValueSafe(ds As MailMergeDataSource, ParamArray fieldNames() As Variant) As String
    Dim i As Long
    Dim val As String
    
    On Error Resume Next
    For i = LBound(fieldNames) To UBound(fieldNames)
        Err.Clear
        val = ds.DataFields(CStr(fieldNames(i))).Value
        If Err.Number = 0 Then
            GetFieldValueSafe = val
            Exit Function
        End If
    Next i
    On Error GoTo 0
    
    GetFieldValueSafe = ""
End Function


Private Function HasAnyField(ds As MailMergeDataSource, ParamArray fieldNames() As Variant) As Boolean
    Dim i As Long
    Dim tempVal As String
    
    On Error Resume Next
    For i = LBound(fieldNames) To UBound(fieldNames)
        Err.Clear
        tempVal = ds.DataFields(CStr(fieldNames(i))).Value
        If Err.Number = 0 Then
            HasAnyField = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
    
    HasAnyField = False
End Function


Private Function GetRecordCount(ds As MailMergeDataSource) As Long
    Dim recCount As Long
    
    On Error Resume Next
    ds.ActiveRecord = wdFirstRecord
    recCount = 0
    
    Do
        recCount = recCount + 1
        ds.ActiveRecord = wdNextRecord
    Loop Until ds.ActiveRecord = recCount Or recCount > 5000
    
    ds.ActiveRecord = wdFirstRecord
    On Error GoTo 0
    
    GetRecordCount = recCount
End Function

