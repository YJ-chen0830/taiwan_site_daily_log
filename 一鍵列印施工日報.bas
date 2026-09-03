Attribute VB_Name = "Module1"
Sub PrintDailyReports()
    '======================================================
    ' 一鍵列印某段時間區間的「施工日報」
    ' 用法：Alt+F8 開啟巨集清單 -> 選 PrintDailyReports -> 執行
    ' 會依序詢問起始日期、結束日期，然後把「施工日報」工作表
    ' G4(填表日期)依序改成範圍內的每一天並列印，印完會還原成
    ' 原本的日期，不影響你目前正在看的那一天。
    '======================================================
    Dim wsReport As Worksheet
    Dim wsDB As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim projStart As Date
    Dim projDays As Long
    Dim d As Date
    Dim ans As String
    Dim originalDate As Variant

    On Error Resume Next
    Set wsReport = ThisWorkbook.Worksheets("施工日報")
    Set wsDB = ThisWorkbook.Worksheets("日資料庫")
    On Error GoTo 0
    If wsReport Is Nothing Or wsDB Is Nothing Then
        MsgBox "找不到「施工日報」或「日資料庫」工作表，請確認工作表名稱是否被改過。", vbExclamation
        Exit Sub
    End If

    projStart = wsDB.Range("B2").Value
    projDays = wsDB.Range("D2").Value

    ans = InputBox("請輸入起始日期 (格式: yyyy/mm/dd)" & vbCrLf & _
                   "開工日期：" & Format(projStart, "yyyy/mm/dd") & vbCrLf & _
                   "預定完工日：" & Format(projStart + projDays - 1, "yyyy/mm/dd"), _
                   "一鍵列印施工日報 - 起始日期", Format(projStart, "yyyy/mm/dd"))
    If ans = "" Then Exit Sub
    If Not IsDate(ans) Then
        MsgBox "日期格式錯誤，請重新執行一次。", vbExclamation
        Exit Sub
    End If
    startDate = CDate(ans)

    ans = InputBox("請輸入結束日期 (格式: yyyy/mm/dd)", _
                   "一鍵列印施工日報 - 結束日期", Format(startDate, "yyyy/mm/dd"))
    If ans = "" Then Exit Sub
    If Not IsDate(ans) Then
        MsgBox "日期格式錯誤，請重新執行一次。", vbExclamation
        Exit Sub
    End If
    endDate = CDate(ans)

    If endDate < startDate Then
        MsgBox "結束日期不可以早於起始日期。", vbExclamation
        Exit Sub
    End If

    If startDate < projStart Or endDate > projStart + projDays - 1 Then
        If MsgBox("所選日期超出「開工日期~預定完工日」的範圍，超出範圍的那幾天會是空白報表，仍要繼續嗎？", _
                  vbYesNo + vbQuestion, "提醒") = vbNo Then
            Exit Sub
        End If
    End If

    If MsgBox("即將列印 " & Format(startDate, "yyyy/mm/dd") & " 至 " & Format(endDate, "yyyy/mm/dd") & _
              "，共 " & (endDate - startDate + 1) & " 天的施工日報，確定要送出列印嗎？", _
              vbYesNo + vbQuestion, "確認列印") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    originalDate = wsReport.Range("G4").Value

    For d = startDate To endDate
        wsReport.Range("G4").Value = d
        wsReport.Calculate
        wsReport.PrintOut Copies:=1, Collate:=True
    Next d

    wsReport.Range("G4").Value = originalDate
    wsReport.Calculate
    Application.ScreenUpdating = True

    MsgBox "列印完成，共 " & (endDate - startDate + 1) & " 天。", vbInformation, "完成"
End Sub
