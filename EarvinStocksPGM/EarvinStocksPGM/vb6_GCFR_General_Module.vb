'***************************************************************************************************
'* 本模組所有的函式都有作用 -- 20080910 --
'***************************************************************************************************

Option Explicit
Option Base 1
Dim sngMax(8) As Single, sngMin(8) As Single

'***************************************************************************************************
'* 說    明: 從指定的檔案名稱中讀出以日為單位的基本資料
'* 輸入參數: strFileName 每日股價資料檔案名稱
'* 輸出參數: 讀檔成功回傳true；失敗回傳false
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Function subReadData(ByVal inputFileName As String) As Boolean
    Dim path As String
    Dim intFileNo As Integer
    Dim openFile As String
    
    On Error GoTo ERR_HANDLE
    
    ' [WAIT-TO-DO]
    openFile = FILEPATH & inputFileName
    
    path = GetAppPath
    openFile = path + "myPGMs\DATA\dat\" & inputFileName
    
    
    If Dir(openFile) = "" Then
        MsgBox "股票資料不存在", vbOKOnly
        subReadData = False
        blnOpenFileSuccess = False
    Else
        ' Reading the data of personal stock
        intFileNo = FreeFile()
        Open openFile For Binary As #intFileNo ' Opening the individual stock
        subReadData = True
      
        '---日股市資料相關資訊 --------------------------------------
        gintDayIndex = LOF(intFileNo) / Len(gudtMydata)
        gsngEndIndex = gintDayIndex
        gsngStartIndex = Int(gsngEndIndex - frmEarvinStocks.Width / gsngBarWidth)  ' 視窗寬度 / K-Bar寬度
        ReDim gudtStockDay(gintDayIndex)  ' 日股市資料
        ReDim gudtIndexDay(gintDayIndex)  ' 日技術指標資料
             
        ' Read the data of personal stock
        Get #intFileNo, , gudtStockDay
        Close #intFileNo
                
        '--- 週股市資料相關資訊 --------------------------------------
        If Int(gintDayIndex / 3) = 0 Then   ' 除以3應該只是為了大約計算配置週資料array的空間
            gintWeekIndex = 1
        Else
            gintWeekIndex = gintDayIndex / 3
        End If
        
        ReDim gudtStockWeek(gintWeekIndex)   ' 週股市資料
        ReDim gudtIndexWeek(gintWeekIndex)   ' 週技術指標資料

        '--- 月股市資料相關資訊 --------------------------------------
'        If Int(gintWeekIndex / 3) = 0 Then  ' 除以3應該只是為了大約計算配置月資料array的空間
        If Int(gintWeekIndex / 4) = 0 Then  ' 20230615 應該是除以4比較合理
            gintMonthIndex = 1
        Else
            gintMonthIndex = gintWeekIndex / 4
        End If
      
        ReDim gudtStockMonth(gintMonthIndex)    ' 月股市資料
        ReDim gudtIndexMonth(gintMonthIndex)    ' 月技術指標資料
  
        ' 呼叫計算指標函式 : 產生日指標資料
        Call subCalculateIndex(gudtStockDay, gudtIndexDay, gintDayIndex)
        ' To general the weekly stock data from daily stock data
        Call subGenerateWeek(gintDayIndex, gudtStockWeek)
        ' 呼叫計算指標函式 : 產生週指標資料
        Call subCalculateIndex(gudtStockWeek, gudtIndexWeek, gintWeekIndex)
        ' To general the monthly stock data from daily stock data
        Call subGenerateMonth(gintWeekIndex, gudtStockMonth)
        ' 呼叫計算指標函式 : 產生月指標資料
        Call subCalculateIndex(gudtStockMonth, gudtIndexMonth, gintMonthIndex)
      
        If frmEarvinStocks.lblDateType.Caption = "日線" Then
            gsngEndIndex = gintDayIndex
        ElseIf frmEarvinStocks.lblDateType.Caption = "週線" Then
            gsngEndIndex = gintWeekIndex
        ElseIf frmEarvinStocks.lblDateType.Caption = "月線" Then
            gsngEndIndex = gintMonthIndex
        End If
        gsngStartIndex = Int(gsngEndIndex - frmEarvinStocks.Width / gsngBarWidth)
    End If
    
    Exit Function
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.subReadData()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Function


'***************************************************************************************************
'* 說    明: 執行計算所有的技術指標的程序
'* 輸入參數: Stockdata  股價資料
'*           IndexData  技術指標資料
'*           intStockNo 資料筆數
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subCalculateIndex(ByRef udtStock() As StockData, _
                             ByRef udtIndex() As IndexData, _
                             ByVal intStockNo As Integer)
    
    On Error GoTo ERR_HANDLE
    
    Call subAverage(udtStock, udtIndex, intStockNo, MAP_0, 0, True)
    Call subAverage(udtStock, udtIndex, intStockNo, MAP_1, 1, True)
    Call subAverage(udtStock, udtIndex, intStockNo, MAP_2, 2, True)
    Call subAverage(udtStock, udtIndex, intStockNo, MAP_3, 3, True)
    Call subAverage(udtStock, udtIndex, intStockNo, MAP_4, 4, True)
    Call subAverage(udtStock, udtIndex, intStockNo, MAP_5, 5, True)
    Call subAverage(udtStock, udtIndex, intStockNo, MAP_6, 6, True)
    Call subAverage(udtStock, udtIndex, intStockNo, MAP_7, 7, True)
  
    Call subPSY(udtStock, udtIndex, intStockNo, PSY_No)
    Call subEMA(udtStock, udtIndex, intStockNo, EMA_NO)
    Call subWMS(udtStock, udtIndex, intStockNo, WMS_No)
    Call subRWMS(udtStock, udtIndex, intStockNo, WMS_No)
    Call subRSI(udtStock, udtIndex, intStockNo, RSI_S, True)
    Call subRSI(udtStock, udtIndex, intStockNo, RSI_L, False)
    Call subStochRSI(udtStock, udtIndex, intStockNo)
    Call subWRSI(udtStock, udtIndex, intStockNo, WRSI)
    Call subMAWRSI(udtStock, udtIndex, intStockNo, MAWRSI, SMAorEMA_WRSI)
    Call subKD(udtStock, udtIndex, intStockNo, KD_No)
    Call subMACD(udtStock, udtIndex, intStockNo, MACD_No, EMA_S, EMA_L, False)
    Call subBias(udtStock, udtIndex, intStockNo, Bias_No)
    
    Call subAverage(udtStock, udtIndex, intStockNo, MAV_1, 1, False)
    Call subAverage(udtStock, udtIndex, intStockNo, MAV_2, 2, False)
    Call subAverage(udtStock, udtIndex, intStockNo, MAV_3, 3, False)
    Call subAverage(udtStock, udtIndex, intStockNo, MAV_4, 4, False)
    Call subAverage(udtStock, udtIndex, intStockNo, MAV_5, 5, False)
    Call subGenerateAcc(gintDayIndex, gudtVolAcc)
    Call subVolumnRatio(gintDayIndex, gudtStockDay)
    Call subShortTerm(gintDayIndex, gudtStockDay, intHeadOrLow)
  
  
    '=======================================================================
    '* GCFR Model: 計算技術指標
    '=======================================================================
    '--- 計算收盤價 ---
    Call subAverage2(udtStock, udtIndex, intStockNo, 3, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 4, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 5, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 6, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 8, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 10, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 12, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 20, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 24, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 30, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 60, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 72, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 120, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 144, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 240, True)
    Call subAverage2(udtStock, udtIndex, intStockNo, 288, True)
    '--- 計算成交量 ---
    Call subAverage2(udtStock, udtIndex, intStockNo, 3, False)
    Call subAverage2(udtStock, udtIndex, intStockNo, 5, False)
    Call subAverage2(udtStock, udtIndex, intStockNo, 6, False)
    Call subAverage2(udtStock, udtIndex, intStockNo, 10, False)
    Call subAverage2(udtStock, udtIndex, intStockNo, 12, False)
    Call subAverage2(udtStock, udtIndex, intStockNo, 20, False)
    Call subAverage2(udtStock, udtIndex, intStockNo, 24, False)
    '--- 其它技術指標 ---
    Call subPSY(udtStock, udtIndex, intStockNo, PSY_No)
    Call subWMS(udtStock, udtIndex, intStockNo, WMS_No)
    Call subRSI(udtStock, udtIndex, intStockNo, RSI_S, True)
    Call subRSI(udtStock, udtIndex, intStockNo, RSI_L, False)
    Call subKD(udtStock, udtIndex, intStockNo, KD_No)
    Call subMACD(udtStock, udtIndex, intStockNo, MACD_No, EMA_S, EMA_L, False)
    Call subBias(udtStock, udtIndex, intStockNo, Bias_No)
    
    Call subVR(udtStock, udtIndex, intStockNo, 12)
    Call subMASlope(udtStock, udtIndex, intStockNo, 60)
'   Call subMADistance(udtStock, udtIndex, intStockNo, 60)
    Call subMAPDis(udtStock, udtIndex, intStockNo)
    Call subLBias(udtStock, udtIndex, intStockNo, 60)
    Call subSBias(udtStock, udtIndex, intStockNo, 10)
    Call subLSBias(udtStock, udtIndex, intStockNo)
    Call subBullList(udtStock, udtIndex, intStockNo)
    Call subBearList(udtStock, udtIndex, intStockNo)
    Call subIsInRange(udtStock, udtIndex, intStockNo)
    Call subUpDownDays(udtStock, udtIndex, intStockNo)
    '=======================================================================
    '* GCFR Model: 計算技術指標     --- END ---
    '=======================================================================

    '=======================================================================
    '* 產生高、低點日期
    '=======================================================================
    Call subGenLowPointDate(udtStock, udtIndex, intStockNo)
    Call subGenHighPointDate(udtStock, udtIndex, intStockNo)
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.subCalculateIndex()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: 從日的基本資料中，產生週的基本資料
'* 輸入參數: gintDayIndex  日股市資料筆數
'*           gudtStockWeek 週股市資料
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subGenerateWeek(ByVal gintDayIndex As Integer, _
                           ByRef gudtStockWeek() As StockData)
    Dim i As Integer
    Dim lngYear As Long
    Dim lngMonth As Long
    Dim lngDay As Long
    Dim datDate As Date
    Dim lngTemp As Long
    Dim intWeekDay As Integer
    Dim intWeekFlag As Integer
    Dim blnFirst As Boolean
    
    i = 1
    gintWeekIndex = 1
    intWeekFlag = 0
    blnFirst = True
    
    On Error GoTo ERR_HANDLE
    
    '*** Weekday() 參數的意義 ***********************************************
    ' FirstDayOfWeek.System    0  系統設定中指定的每週第一天
    ' FirstDayOfWeek.Sunday    1  星期日 (預設值)
    ' FirstDayOfWeek.Monday    2  星期一 (符合 ISO 標準 8601 的第 3.17 節)
    ' FirstDayOfWeek.Tuesday   3  星期二
    ' FirstDayOfWeek.Wednesday 4  星期三
    ' FirstDayOfWeek.Thursday  5  星期四
    ' FirstDayOfWeek.Friday    6  星期五
    ' FirstDayOfWeek.Saturday  7  星期六
    '************************************************************************
    
    
    While i <= gintDayIndex
        lngTemp = gudtStockDay(i).sngDate
        lngYear = Int(lngTemp / 10000)   ' 取出年
        lngMonth = Int((lngTemp - lngYear * 10000) / 100)  ' 取出月
        lngDay = lngTemp - lngYear * 10000 - lngMonth * 100   ' 取出日
        lngYear = lngYear + 1911
        datDate = DateSerial(lngYear, lngMonth, lngDay) ' 轉換成日期
        intWeekDay = Weekday(datDate) ' 計算今天是星期幾
      
        If intWeekFlag < intWeekDay Then
            If blnFirst Then
                gudtStockWeek(gintWeekIndex).sngStartprice = gudtStockDay(i).sngStartprice
                gudtStockWeek(gintWeekIndex).sngHighPrice = gudtStockDay(i).sngHighPrice
                gudtStockWeek(gintWeekIndex).sngLowPrice = gudtStockDay(i).sngLowPrice
                blnFirst = False
            End If
            ' 最高價
            If gudtStockWeek(gintWeekIndex).sngHighPrice < gudtStockDay(i).sngHighPrice Then
                gudtStockWeek(gintWeekIndex).sngHighPrice = gudtStockDay(i).sngHighPrice
            End If
            ' 最低價
            If gudtStockWeek(gintWeekIndex).sngLowPrice > gudtStockDay(i).sngLowPrice Then
                gudtStockWeek(gintWeekIndex).sngLowPrice = gudtStockDay(i).sngLowPrice
            End If
            gudtStockWeek(gintWeekIndex).sngDate = gudtStockDay(i).sngDate
            gudtStockWeek(gintWeekIndex).sngEndprice = gudtStockDay(i).sngEndprice
            gudtStockWeek(gintWeekIndex).sngVol = gudtStockWeek(gintWeekIndex).sngVol + gudtStockDay(i).sngVol
            gudtStockWeek(gintWeekIndex).sngAcc = gudtStockWeek(gintWeekIndex).sngAcc + gudtStockDay(i).sngAcc      ' 融資
            gudtStockWeek(gintWeekIndex).sngTome = gudtStockWeek(gintWeekIndex).sngTome + gudtStockDay(i).sngTome   ' 融券
         
            gudtStockWeek(gintWeekIndex).sngForeignStock = gudtStockWeek(gintWeekIndex).sngForeignStock + gudtStockDay(i).sngForeignStock                   ' 外資庫存
            gudtStockWeek(gintWeekIndex).sngSitAndCbStock = gudtStockWeek(gintWeekIndex).sngSitAndCbStock + gudtStockDay(i).sngSitAndCbStock                ' 投信庫存
            gudtStockWeek(gintWeekIndex).sngSelfEmployedStock = gudtStockWeek(gintWeekIndex).sngSelfEmployedStock + gudtStockDay(i).sngSelfEmployedStock    ' 自營商庫存
            gudtStockWeek(gintWeekIndex).sngLegalPersonStock = gudtStockWeek(gintWeekIndex).sngLegalPersonStock + gudtStockDay(i).sngLegalPersonStock       ' 法人庫存
         
            intWeekFlag = intWeekDay
        Else
            intWeekFlag = intWeekDay
            gintWeekIndex = gintWeekIndex + 1
            gudtStockWeek(gintWeekIndex).sngDate = gudtStockDay(i).sngDate
            gudtStockWeek(gintWeekIndex).sngStartprice = gudtStockDay(i).sngStartprice
            gudtStockWeek(gintWeekIndex).sngHighPrice = gudtStockDay(i).sngHighPrice
            gudtStockWeek(gintWeekIndex).sngLowPrice = gudtStockDay(i).sngLowPrice
            gudtStockWeek(gintWeekIndex).sngEndprice = gudtStockDay(i).sngEndprice
            gudtStockWeek(gintWeekIndex).sngVol = gudtStockWeek(gintWeekIndex).sngVol + gudtStockDay(i).sngVol
            gudtStockWeek(gintWeekIndex).sngAcc = gudtStockWeek(gintWeekIndex).sngAcc + gudtStockDay(i).sngAcc      ' 融資
            gudtStockWeek(gintWeekIndex).sngTome = gudtStockWeek(gintWeekIndex).sngTome + gudtStockDay(i).sngTome   ' 融券
        
            gudtStockWeek(gintWeekIndex).sngForeignStock = gudtStockWeek(gintWeekIndex).sngForeignStock + gudtStockDay(i).sngForeignStock                   ' 外資庫存
            gudtStockWeek(gintWeekIndex).sngSitAndCbStock = gudtStockWeek(gintWeekIndex).sngSitAndCbStock + gudtStockDay(i).sngSitAndCbStock                ' 投信庫存
            gudtStockWeek(gintWeekIndex).sngSelfEmployedStock = gudtStockWeek(gintWeekIndex).sngSelfEmployedStock + gudtStockDay(i).sngSelfEmployedStock    ' 自營商庫存
            gudtStockWeek(gintWeekIndex).sngLegalPersonStock = gudtStockWeek(gintWeekIndex).sngLegalPersonStock + gudtStockDay(i).sngLegalPersonStock       ' 法人庫存
        
        End If
        i = i + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.subGenerateWeek()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: 從週的基本資料中，產生月的基本資料
'* 輸入參數: gintDayIndex 週股市資料筆數
'*           gudtStockWeek 月股市資料
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081223 調整格式及增加註解
'***************************************************************************************************
Public Sub subGenerateMonth(ByVal gintWeekIndex As Integer, _
                            ByRef gudtStockMonth() As StockData)
    Dim i As Integer
    Dim intMonthFlag As Integer
    
    i = 1
    gintMonthIndex = 0
    
    On Error GoTo ERR_HANDLE
    
    While i <= gintWeekIndex
        With gudtStockWeek(i)
            intMonthFlag = .sngDate \ 100
            If gintMonthIndex = 0 Then
                gintMonthIndex = gintMonthIndex + 1
                gudtStockMonth(gintMonthIndex).sngDate = intMonthFlag
                gudtStockMonth(gintMonthIndex).sngStartprice = .sngStartprice
                gudtStockMonth(gintMonthIndex).sngHighPrice = .sngHighPrice
                gudtStockMonth(gintMonthIndex).sngLowPrice = .sngLowPrice
                gudtStockMonth(gintMonthIndex).sngEndprice = .sngEndprice
                gudtStockMonth(gintMonthIndex).sngVol = .sngVol
                gudtStockMonth(gintMonthIndex).sngAcc = .sngAcc     ' 融資
                gudtStockMonth(gintMonthIndex).sngTome = .sngTome   ' 融券
                gudtStockMonth(gintMonthIndex).sngForeignStock = .sngForeignStock               ' 外資庫存
                gudtStockMonth(gintMonthIndex).sngSitAndCbStock = .sngSitAndCbStock             ' 投信庫存
                gudtStockMonth(gintMonthIndex).sngSelfEmployedStock = .sngSelfEmployedStock     ' 自營商庫存
                gudtStockMonth(gintMonthIndex).sngLegalPersonStock = .sngLegalPersonStock       ' 法人庫存
            Else
                If gudtStockMonth(gintMonthIndex).sngDate <> intMonthFlag Then
                    gintMonthIndex = gintMonthIndex + 1
                    gudtStockMonth(gintMonthIndex).sngDate = intMonthFlag
                    gudtStockMonth(gintMonthIndex).sngStartprice = .sngStartprice
                    gudtStockMonth(gintMonthIndex).sngHighPrice = .sngHighPrice
                    gudtStockMonth(gintMonthIndex).sngLowPrice = .sngLowPrice
                    gudtStockMonth(gintMonthIndex).sngEndprice = .sngEndprice
                    gudtStockMonth(gintMonthIndex).sngVol = .sngVol
                    gudtStockMonth(gintMonthIndex).sngAcc = .sngAcc      ' 融資
                    gudtStockMonth(gintMonthIndex).sngTome = .sngTome    ' 融券
                    gudtStockMonth(gintMonthIndex).sngForeignStock = .sngForeignStock           ' 外資庫存
                    gudtStockMonth(gintMonthIndex).sngSitAndCbStock = .sngSitAndCbStock         ' 投信庫存
                    gudtStockMonth(gintMonthIndex).sngSelfEmployedStock = .sngSelfEmployedStock ' 自營商庫存
                    gudtStockMonth(gintMonthIndex).sngLegalPersonStock = .sngLegalPersonStock   ' 法人庫存
                Else
                    If gudtStockMonth(gintMonthIndex).sngHighPrice < .sngHighPrice Then
                        gudtStockMonth(gintMonthIndex).sngHighPrice = .sngHighPrice
                    End If
                    If gudtStockMonth(gintMonthIndex).sngLowPrice > .sngLowPrice Then
                        gudtStockMonth(gintMonthIndex).sngLowPrice = .sngLowPrice
                    End If
                    gudtStockMonth(gintMonthIndex).sngEndprice = .sngEndprice
                    gudtStockMonth(gintMonthIndex).sngVol = gudtStockMonth(gintMonthIndex).sngVol + .sngVol
                    gudtStockMonth(gintMonthIndex).sngAcc = gudtStockMonth(gintMonthIndex).sngAcc + .sngAcc      ' 融資
                    gudtStockMonth(gintMonthIndex).sngTome = gudtStockMonth(gintMonthIndex).sngTome + .sngTome   ' 融券
                    gudtStockMonth(gintMonthIndex).sngTome = gudtStockMonth(gintMonthIndex).sngForeignStock + .sngForeignStock              ' 外資庫存
                    gudtStockMonth(gintMonthIndex).sngTome = gudtStockMonth(gintMonthIndex).sngSelfEmployedStock + .sngSelfEmployedStock    ' 投信庫存
                    gudtStockMonth(gintMonthIndex).sngTome = gudtStockMonth(gintMonthIndex).sngSelfEmployedStock + .sngSelfEmployedStock    ' 自營商庫存
                    gudtStockMonth(gintMonthIndex).sngTome = gudtStockMonth(gintMonthIndex).sngLegalPersonStock + .sngLegalPersonStock      ' 法人庫存
                End If
            End If
        End With
        i = i + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.subGenerateMonth()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: 在顯示畫面的右方顯示一些技術指標的值
'* 輸入參數: gudtStock     股價資料
'*           gudtIndex     技術指標資料
'*           intTotalIndex 資料總筆數
'*           sngIndex      目前畫面所在位置(第??筆)
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub DrawStockIndexes(ByRef gudtStock() As StockData, _
                            ByRef gudtIndex() As IndexData, _
                            ByVal intTotalIndex As Integer, _
                            ByVal sngIndex As Single)
    Dim strUpDown As String
    Dim sngTemp As Single   ' 記錄要顯示的指標文字的頂點座標
    Dim flag_str As String
    Dim sngCount As Integer
    Dim i As Integer
 
    On Error GoTo ERR_HANDLE
    
    If sngIndex <= intTotalIndex And sngIndex >= 1 Then
        With gudtStock(sngIndex)
            If sngIndex > 1 Then
                If (.sngEndprice - gudtStock(sngIndex - 1).sngEndprice) > 0 Then
                    strUpDown = " ▲"
                    frmEarvinStocks.lblStockSts.ForeColor = QBColor(12)
                ElseIf (.sngEndprice - gudtStock(sngIndex - 1).sngEndprice) < 0 Then
                    strUpDown = " ▼"
                    frmEarvinStocks.lblStockSts.ForeColor = QBColor(10)
                Else
                    strUpDown = ""
                End If
            ' 20090108 GRG_PS_Add未實作
'''            If GRG_PS_Add(sngIndex) <> 0 Then
'''               strUpDown = strUpDown & Str(Format((.sngEndprice - gudtStock(sngIndex - 1).sngEndprice), "#.00")) & "  GRG_Pre=" & GRG_PS_Add(sngIndex)
'''            Else
'''               strUpDown = strUpDown & Str(Format((.sngEndprice - gudtStock(sngIndex - 1).sngEndprice), "#.00"))
'''            End If
                strUpDown = strUpDown & Str(Format((.sngEndprice - gudtStock(sngIndex - 1).sngEndprice), "#.00"))
            End If
            Focus_Today = sngIndex
            frmEarvinStocks.lblStockInf.ForeColor = QBColor(11)
            frmEarvinStocks.lblStockInf.Caption = Str(.sngDate) & " 開" & Str(.sngStartprice) & " 高" & Str(.sngHighPrice) & _
                " 低" & Str(.sngLowPrice) & " 收" & Str(.sngEndprice) & "   量" & Str(.sngVol) & "億"
            If sngIndex <> 1 Then
                frmEarvinStocks.lblStockSts.Caption = strUpDown & "  " & Round((.sngEndprice - gudtStock(sngIndex - 1).sngEndprice) / gudtStock(sngIndex - 1).sngEndprice * 100, 2) & "%"
            Else
                frmEarvinStocks.lblStockSts.Caption = strUpDown
            End If
        End With
 
        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand + 20
        frmEarvinStocks.Line (frmEarvinStocks.ScaleWidth - gsngRightLevel, sngTemp)-(frmEarvinStocks.ScaleWidth - 0.01 * gsngRightLevel, sngTemp), RGB(200, 100, 0)
        sngTemp = sngTemp - 20
 
        With frmEarvinStocks.lblDateType
            .Top = sngTemp
            .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
            .Width = gsngRightLevel * 0.96
            .Caption = frmEarvinStocks.cboStocksType.Text
            sngTemp = sngTemp - .Height
        End With
 
        frmEarvinStocks.Line (frmEarvinStocks.ScaleWidth - gsngRightLevel, sngTemp)-(frmEarvinStocks.ScaleWidth - 0.01 * gsngRightLevel, sngTemp), RGB(200, 100, 0)
        sngTemp = sngTemp - 20
 
        ' 先畫格子
        frmEarvinStocks.DrawStyle = vbSolid
        With frmEarvinStocks
            frmEarvinStocks.Line (.ScaleWidth - gsngRightLevel, 0.98 * gsngBottomFrame)-(.ScaleWidth - 0.01 * gsngRightLevel, 0.98 * gsngBottomFrame), RGB(200, 100, 0)
            gsngYshift = gsngBottomFrame
            For sngCount = gbytFrameNum To 2 Step -1
                gsngYshift = gsngYshift + mudtFrame(sngCount).sngHeight
                frmEarvinStocks.Line (.ScaleWidth - gsngRightLevel, gsngYshift)-(.ScaleWidth - 0.01 * gsngRightLevel, gsngYshift), RGB(200, 100, 0)
            Next
        End With
 
        frmEarvinStocks.lblMAP0.Visible = False
        frmEarvinStocks.lblMAP1.Visible = False
        frmEarvinStocks.lblMAP2.Visible = False
        frmEarvinStocks.lblMAP3.Visible = False
        frmEarvinStocks.lblMAP4.Visible = False
        frmEarvinStocks.lblMAP5.Visible = False
 
        frmEarvinStocks.lbl12RSI.Visible = False
        frmEarvinStocks.lbl6RSI.Visible = False
        frmEarvinStocks.lblStochRSI.Visible = False
        frmEarvinStocks.lblAcc.Visible = False
        frmEarvinStocks.lblBias.Visible = False
        frmEarvinStocks.lblQuantity.Visible = False
        frmEarvinStocks.lblD.Visible = False
        frmEarvinStocks.lblK.Visible = False
        frmEarvinStocks.lblMACD.Visible = False
        frmEarvinStocks.lblDIF.Visible = False
        frmEarvinStocks.lblCy.Visible = False
        frmEarvinStocks.lblPSY.Visible = False
        frmEarvinStocks.lblSignAcc.Visible = False
        frmEarvinStocks.lblSignTome.Visible = False
        frmEarvinStocks.lblTome.Visible = False
        frmEarvinStocks.lblWMS.Visible = False
        frmEarvinStocks.lblRWMS.Visible = False
        frmEarvinStocks.lbl6WRSI.Visible = False
        frmEarvinStocks.lbl65MAWRSI.Visible = False

        '==== 2004/3/9 DFTS ====
        frmEarvinStocks.lblTrend.Visible = False
        frmEarvinStocks.lblQM.Visible = False
        frmEarvinStocks.lblDiff1.Visible = False
        frmEarvinStocks.lblDiff2.Visible = False
        frmEarvinStocks.lblDC.Visible = False
        frmEarvinStocks.lblWLST.Visible = False
        frmEarvinStocks.lblWCY.Visible = False
        frmEarvinStocks.lblWLW.Visible = False
        
        ' 20180825
        frmEarvinStocks.lblForeignStock.Visible = False
        frmEarvinStocks.lblLegalPersonStock.Visible = False
        frmEarvinStocks.lblSelfEmployedStock.Visible = False
        frmEarvinStocks.lblSignSitAndCbStock.Visible = False
        frmEarvinStocks.lblSignForeignStock.Visible = False
        frmEarvinStocks.lblSignLegalPersonStock.Visible = False
        frmEarvinStocks.lblSignSelfEmployedStock.Visible = False
        frmEarvinStocks.lblSignSitAndCbStock.Visible = False
        

        '***********決定顯示MAP或Tether_Line或Reverse Engineering RSI     2003/5/26****************
        If frmEarvinStocks.mnuKmap.HelpContextID = 0 Then
            ' 移動平均線跟移動平均量的值一直都顯示
            With frmEarvinStocks.lblMAP0
                .FontSize = 10
                .ForeColor = QBColor(5)
                .Top = sngTemp
                .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                .Width = gsngRightLevel * 0.96
                .Visible = True
                If sngIndex <> 1 Then
                    If gudtIndex(sngIndex).sngMAP0 > gudtIndex(sngIndex - 1).sngMAP0 Then
                        flag_str = "↑"
                    ElseIf gudtIndex(sngIndex).sngMAP0 < gudtIndex(sngIndex - 1).sngMAP0 Then
                        flag_str = "↓"
                    Else
                        flag_str = "–"
                    End If
                End If
              
                If MAP_0 <> 0 Then
                    .Caption = MAP_0 & "MAP=" & Str(Round(gudtIndex(sngIndex).sngMAP0, 2)) & flag_str
                Else
                    .Caption = ""
                End If
                sngTemp = sngTemp - .Height
            End With

            With frmEarvinStocks.lblMAP1
                .FontSize = 10
                .ForeColor = QBColor(3)
                .Top = sngTemp
                .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                .Width = gsngRightLevel * 0.96
                .Visible = True
                If sngIndex <> 1 Then
                    If gudtIndex(sngIndex).sngMAP1 > gudtIndex(sngIndex - 1).sngMAP1 Then
                        flag_str = "↑"
                    ElseIf gudtIndex(sngIndex).sngMAP1 < gudtIndex(sngIndex - 1).sngMAP1 Then
                        flag_str = "↓"
                    Else
                        flag_str = "–"
                    End If
                End If
              
                If MAP_1 <> 0 Then
                    .Caption = MAP_1 & "MAP=" & Str(Round(gudtIndex(sngIndex).sngMAP1, 2)) & flag_str
                Else
                    .Caption = ""
                End If
                sngTemp = sngTemp - .Height
            End With

            With frmEarvinStocks.lblMAP2
                .FontSize = 10
                .ForeColor = QBColor(9)
                .Top = sngTemp
                .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                .Width = gsngRightLevel * 0.96
                .Visible = True
                If sngIndex <> 1 Then
                    If gudtIndex(sngIndex).sngMAP2 > gudtIndex(sngIndex - 1).sngMAP2 Then
                        flag_str = "↑"
                    ElseIf gudtIndex(sngIndex).sngMAP2 < gudtIndex(sngIndex - 1).sngMAP2 Then
                        flag_str = "↓"
                    Else
                        flag_str = "–"
                    End If
                End If
                If MAP_2 <> 0 Then
                    .Caption = MAP_2 & "MAP=" & Str(Round(gudtIndex(sngIndex).sngMAP2, 2)) & flag_str
                Else
                    .Caption = ""
                End If
                sngTemp = sngTemp - .Height
            End With

            With frmEarvinStocks.lblMAP3
                .FontSize = 10
                .ForeColor = QBColor(13)
                .Top = sngTemp
                .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                .Width = gsngRightLevel * 0.96
                .Visible = True
                If sngIndex <> 1 Then
                    If gudtIndex(sngIndex).sngMAP3 > gudtIndex(sngIndex - 1).sngMAP3 Then
                        flag_str = "↑"
                    ElseIf gudtIndex(sngIndex).sngMAP3 < gudtIndex(sngIndex - 1).sngMAP3 Then
                        flag_str = "↓"
                    Else
                        flag_str = "–"
                    End If
                End If
            
                If MAP_3 <> 0 Then
                    .Caption = MAP_3 & "MAP=" & Str(Round(gudtIndex(sngIndex).sngMAP3, 2)) & flag_str
                Else
                    .Caption = ""
                End If
                sngTemp = sngTemp - .Height
            End With

            With frmEarvinStocks.lblMAP4
                .FontSize = 10
                .ForeColor = QBColor(11)
                .Top = sngTemp
                .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                .Width = gsngRightLevel * 0.96
                .Visible = True
                If sngIndex <> 1 Then
                    If gudtIndex(sngIndex).sngMAP4 > gudtIndex(sngIndex - 1).sngMAP4 Then
                        flag_str = "↑"
                    ElseIf gudtIndex(sngIndex).sngMAP4 < gudtIndex(sngIndex - 1).sngMAP4 Then
                        flag_str = "↓"
                    Else
                        flag_str = "–"
                    End If
                End If
            
                If MAP_4 <> 0 Then
                    .Caption = MAP_4 & "MAP=" & Str(Round(gudtIndex(sngIndex).sngMAP4, 2)) & flag_str
                Else
                    .Caption = ""
                End If
                sngTemp = sngTemp - .Height
            End With

            With frmEarvinStocks.lblMAP5
                .FontSize = 10
                .ForeColor = QBColor(7)
                .Top = sngTemp
                .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                .Width = gsngRightLevel * 0.96
                .Visible = True
                If sngIndex <> 1 Then
                    If gudtIndex(sngIndex).sngMAP5 > gudtIndex(sngIndex - 1).sngMAP5 Then
                        flag_str = "↑"
                    ElseIf gudtIndex(sngIndex).sngMAP5 < gudtIndex(sngIndex - 1).sngMAP5 Then
                        flag_str = "↓"
                    Else
                        flag_str = "–"
                    End If
                End If
            
                If MAP_5 <> 0 Then
                    .Caption = MAP_5 & "MAP=" & Str(Round(gudtIndex(sngIndex).sngMAP5, 2)) & flag_str
                Else
                    .Caption = ""
                End If
                sngTemp = sngTemp - .Height
            End With
        End If
     
        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand + 20 - (frmEarvinStocks.lblDateType.Height + mudtFrame(1).sngHeight) / 2
        frmEarvinStocks.Line (frmEarvinStocks.ScaleWidth - gsngRightLevel, sngTemp)-(frmEarvinStocks.ScaleWidth - 0.01 * gsngRightLevel, sngTemp), RGB(200, 100, 0)
        sngTemp = sngTemp - 20
     
        '*** 顯示「成交量」******************************************************************************
        With frmEarvinStocks.lblMAV1
            .FontSize = 9
            .ForeColor = QBColor(14)
            .Top = sngTemp
            .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
            .Width = gsngRightLevel * 0.96
            .Visible = True
            If sngIndex <> 1 Then
                If gudtIndex(sngIndex).sngMAV1 > gudtIndex(sngIndex - 1).sngMAV1 Then
                    flag_str = "↑"
                ElseIf gudtIndex(sngIndex).sngMAV1 < gudtIndex(sngIndex - 1).sngMAV1 Then
                    flag_str = "↓"
                Else
                    flag_str = "–"
                End If
            End If
        
            If MAV_1 <> 0 Then
                .Caption = MAV_1 & "MAV=" & Str(Round(gudtIndex(sngIndex).sngMAV1, 2)) & flag_str
            Else
                .Caption = ""
            End If
            sngTemp = sngTemp - .Height
        End With
    
        With frmEarvinStocks.lblMAV2
            .FontSize = 9
            .ForeColor = QBColor(9)
            .Top = sngTemp
            .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
            .Width = gsngRightLevel * 0.96
            .Visible = True
            If sngIndex <> 1 Then
                If gudtIndex(sngIndex).sngMAV2 > gudtIndex(sngIndex - 1).sngMAV2 Then
                    flag_str = "↑"
                ElseIf gudtIndex(sngIndex).sngMAV2 < gudtIndex(sngIndex - 1).sngMAV2 Then
                    flag_str = "↓"
                Else
                    flag_str = "–"
                End If
            End If
        
            If MAV_2 <> 0 Then
                .Caption = MAV_2 & "MAV=" & Str(Round(gudtIndex(sngIndex).sngMAV2, 2)) & flag_str
            Else
                .Caption = ""
            End If
            sngTemp = sngTemp - .Height
        End With

        With frmEarvinStocks.lblMAV3
            .FontSize = 9
            .ForeColor = QBColor(13)
            .Top = sngTemp
            .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
            .Width = gsngRightLevel * 0.96
            .Visible = True
            If sngIndex <> 1 Then
                If gudtIndex(sngIndex).sngMAV3 > gudtIndex(sngIndex - 1).sngMAV3 Then
                    flag_str = "↑"
                ElseIf gudtIndex(sngIndex).sngMAV3 < gudtIndex(sngIndex - 1).sngMAV3 Then
                    flag_str = "↓"
                Else
                    flag_str = "–"
                End If
            End If
        
            If MAV_3 <> 0 Then
                .Caption = MAV_3 & "MAV=" & Str(Round(gudtIndex(sngIndex).sngMAV3, 2)) & flag_str
            Else
                .Caption = ""
            End If
            sngTemp = sngTemp - .Height
        End With

        With frmEarvinStocks.lblMAV4
            .FontSize = 9
            .ForeColor = QBColor(11)
            .Top = sngTemp
            .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
            .Width = gsngRightLevel * 0.96
            .Visible = True
          
            If sngIndex <> 1 Then
                If gudtIndex(sngIndex).sngMAV4 > gudtIndex(sngIndex - 1).sngMAV4 Then
                    flag_str = "↑"
                ElseIf gudtIndex(sngIndex).sngMAV4 < gudtIndex(sngIndex - 1).sngMAV4 Then
                    flag_str = "↓"
                Else
                    flag_str = "–"
                End If
            End If
        
            If MAV_4 <> 0 Then
                .Caption = MAV_4 & "MAV=" & Str(Round(gudtIndex(sngIndex).sngMAV4, 2)) & flag_str
            Else
                .Caption = ""
            End If
            sngTemp = sngTemp - .Height
        End With

        With frmEarvinStocks.lblMAV5
            .FontSize = 9
            .ForeColor = QBColor(7)
            .Top = sngTemp
            .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
            .Width = gsngRightLevel * 0.96
            .Visible = True
          
            If sngIndex <> 1 Then
                If gudtIndex(sngIndex).sngMAV5 > gudtIndex(sngIndex - 1).sngMAV5 Then
                    flag_str = "↑"
                ElseIf gudtIndex(sngIndex).sngMAV5 < gudtIndex(sngIndex - 1).sngMAV5 Then
                    flag_str = "↓"
                Else
                    flag_str = "–"
                End If
            End If
        
            If MAV_5 <> 0 Then
                .Caption = MAV_5 & "MAV=" & Str(Round(gudtIndex(sngIndex).sngMAV5, 2)) & flag_str
            Else
                .Caption = ""
            End If
            sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand - mudtFrame(1).sngHeight
        End With

        ' 格子已經先畫好了，所以下面這行去掉
'       frmEarvinStocks.Line (frmEarvinStocks.ScaleWidth - gsngRightlevel, sngTemp)-(frmEarvinStocks.ScaleWidth - 0.01 * gsngRightlevel, sngTemp), RGB(200, 100, 0)
        sngTemp = sngTemp - 20

        ' 有選XX指標才顯示XX指標值
        For sngCount = 1 To gbytFrameNum
            Select Case mudtFrame(sngCount).bytAttribute
'               Case mKmap
 
                Case mQuantityMap
                    sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                    For i = 1 To sngCount
                        sngTemp = sngTemp - mudtFrame(i).sngHeight
                    Next
                    sngTemp = sngTemp - 20
            
                Case mKDmap
                    With frmEarvinStocks.lblK
                        .FontSize = 10
                        .ForeColor = QBColor(12)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngK > gudtIndex(sngIndex - 1).sngK Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngK < gudtIndex(sngIndex - 1).sngK Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = KD_No & "K=" & Str(Round(gudtIndex(sngIndex).sngK, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lblD
                        .FontSize = 10
                        .ForeColor = QBColor(9)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngD > gudtIndex(sngIndex - 1).sngD Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngD < gudtIndex(sngIndex - 1).sngD Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = KD_No & "D=" & Str(Round(gudtIndex(sngIndex).sngD, 2)) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
            
                Case mRSImap
                    With frmEarvinStocks.lbl6RSI
                        .FontSize = 10
                        .ForeColor = QBColor(7)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngRSI_S > gudtIndex(sngIndex - 1).sngRSI_S Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngRSI_S < gudtIndex(sngIndex - 1).sngRSI_S Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = RSI_S & "RSI=" & Str(Round(gudtIndex(sngIndex).sngRSI_S, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lbl12RSI
                        .FontSize = 10
                        .ForeColor = QBColor(9)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngRSI_L > gudtIndex(sngIndex - 1).sngRSI_L Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngRSI_L < gudtIndex(sngIndex - 1).sngRSI_L Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = RSI_L & "RSI=" & Str(Round(gudtIndex(sngIndex).sngRSI_L, 2)) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
            
                Case mStochRSImap
                    With frmEarvinStocks.lblStochRSI
                        .FontSize = 10
                        .ForeColor = QBColor(7)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngStochRSI > gudtIndex(sngIndex - 1).sngStochRSI Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngStochRSI < gudtIndex(sngIndex - 1).sngStochRSI Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = StochRSI_No & "StochRSI=" & Str(Round(gudtIndex(sngIndex).sngStochRSI, 2)) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
            
                Case mMACDmap
                    With frmEarvinStocks.lblMACD
                        .FontSize = 10
                        .ForeColor = QBColor(12)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngMACD > gudtIndex(sngIndex - 1).sngMACD Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngMACD < gudtIndex(sngIndex - 1).sngMACD Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = MACD_No & "MACD=" & Str(Round(gudtIndex(sngIndex).sngMACD, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lblDIF
                        .FontSize = 10
                        .ForeColor = QBColor(10)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngDIF > gudtIndex(sngIndex - 1).sngDIF Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngDIF < gudtIndex(sngIndex - 1).sngDIF Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "DIF=" & Str(Round(gudtIndex(sngIndex).sngDIF, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lblCy
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngDIF_MACD > gudtIndex(sngIndex - 1).sngDIF_MACD Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngDIF_MACD < gudtIndex(sngIndex - 1).sngDIF_MACD Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        If gudtIndex(sngIndex).sngDIF_MACD >= 0 Then
                            .ForeColor = QBColor(12)
                        Else
                            .ForeColor = QBColor(10)
                        End If
                        .Caption = "Cy=" & Str(Round(gudtIndex(sngIndex).sngDIF_MACD, 2)) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
        
                Case mWMSmap
                    With frmEarvinStocks.lblWMS
                        .FontSize = 10
                        .ForeColor = QBColor(11)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngWMS > gudtIndex(sngIndex - 1).sngWMS Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngWMS < gudtIndex(sngIndex - 1).sngWMS Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = WMS_No & "WMS=" & Str(Round(gudtIndex(sngIndex).sngWMS, 2)) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
               
                Case mRWMSmap
                    With frmEarvinStocks.lblRWMS
                        .FontSize = 10
                        .ForeColor = QBColor(11)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngRWMS > gudtIndex(sngIndex - 1).sngRWMS Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngRWMS < gudtIndex(sngIndex - 1).sngRWMS Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = RWMS_No & "WMSC=" & Str(Round(gudtIndex(sngIndex).sngRWMS, 2)) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
            
                Case mWRSImap
                    With frmEarvinStocks.lbl6WRSI
                        .FontSize = 10
                        .ForeColor = QBColor(7)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngWRSI > gudtIndex(sngIndex - 1).sngWRSI Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngWRSI < gudtIndex(sngIndex - 1).sngWRSI Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = WRSI & "wrsi=" & Str(Round(gudtIndex(sngIndex).sngWRSI, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lbl65MAWRSI
                        .FontSize = 10
                        .ForeColor = QBColor(11)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngMAWRSI > gudtIndex(sngIndex - 1).sngMAWRSI Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngMAWRSI < gudtIndex(sngIndex - 1).sngMAWRSI Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = MAWRSI & "wrsi=" & Str(Round(gudtIndex(sngIndex).sngMAWRSI, 2)) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
         
                Case mPSYmap
                    With frmEarvinStocks.lblPSY
                        .FontSize = 10
                        .ForeColor = QBColor(12)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngPSY > gudtIndex(sngIndex - 1).sngPSY Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngPSY < gudtIndex(sngIndex - 1).sngPSY Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = PSY_No & "PSY=" & Str(Round(gudtIndex(sngIndex).sngPSY, 2)) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
      
                Case mEMAmap
                    With frmEarvinStocks.lblEMA
                        .FontSize = 10
                        .ForeColor = QBColor(12)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngEMA > gudtIndex(sngIndex - 1).sngEMA Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngEMA < gudtIndex(sngIndex - 1).sngEMA Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = EMA_NO & "EMA=" & Str(Round(gudtIndex(sngIndex).sngEMA, 2)) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
      
                Case mBiasmap
                    With frmEarvinStocks.lblBias
                        .FontSize = 10
                        .ForeColor = QBColor(14)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngBias > gudtIndex(sngIndex - 1).sngBias Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngBias < gudtIndex(sngIndex - 1).sngBias Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = Bias_No & "Bias=" & Str(Round(gudtIndex(sngIndex).sngBias, 2)) & "%" & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
          
                Case mQuantityMap
                    With frmEarvinStocks.lblQuantity
                        .FontSize = 10
                        .ForeColor = QBColor(14)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngVol > gudtStock(sngIndex - 1).sngVol Then
                                flag_str = "↑"
                            ElseIf gudtStock(sngIndex).sngVol < gudtStock(sngIndex - 1).sngVol Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "Vol=" & Str(Round(gudtStock(sngIndex).sngVol, 2)) & "%" & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
       
                Case mAccMap
                    With frmEarvinStocks.lblAcc
                        .FontSize = 10
                        .ForeColor = QBColor(12)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngAcc > gudtStock(sngIndex - 1).sngAcc Then
                                flag_str = "↑"
                            ElseIf gudtStock(sngIndex).sngAcc < gudtStock(sngIndex - 1).sngAcc Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "融資=" & Str(Round(gudtStock(sngIndex).sngAcc, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                        With frmEarvinStocks.lblSignAcc
                            .FontSize = 10
                            .Top = sngTemp
                            .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                            .Width = gsngRightLevel * 0.96
                            .Visible = True
                            If sngIndex <> 1 Then
                                If gudtStock(sngIndex).sngAcc >= gudtStock(sngIndex - 1).sngAcc Then
                                    .ForeColor = QBColor(12)
                                    flag_str = "+"
                                ElseIf gudtStock(sngIndex).sngAcc < gudtStock(sngIndex - 1).sngAcc Then
                                    .ForeColor = QBColor(10)
                                    flag_str = ""
                                End If
                                .Caption = flag_str & Round(CSng(gudtStock(sngIndex).sngAcc) - CSng(gudtStock(sngIndex - 1).sngAcc), 2)
                            Else
                                .Caption = ""
                            End If
                            sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                            For i = 1 To sngCount
                                sngTemp = sngTemp - mudtFrame(i).sngHeight
                            Next
                            sngTemp = sngTemp - 20
                        End With
                    End With
        
                Case mTomeMap
                    With frmEarvinStocks.lblTome
                        .FontSize = 10
                        .ForeColor = QBColor(10)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngTome > gudtStock(sngIndex - 1).sngTome Then
                                flag_str = "↑"
                            ElseIf gudtStock(sngIndex).sngTome < gudtStock(sngIndex - 1).sngTome Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "融券=" & Str(Round(gudtStock(sngIndex).sngTome, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lblSignTome
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngTome >= gudtStock(sngIndex - 1).sngTome Then
                                .ForeColor = QBColor(12)
                                flag_str = "+"
                            ElseIf gudtStock(sngIndex).sngTome < gudtStock(sngIndex - 1).sngTome Then
                                .ForeColor = QBColor(10)
                                flag_str = ""
                            End If
                            .Caption = flag_str & Round(CSng(gudtStock(sngIndex).sngTome) - CSng(gudtStock(sngIndex - 1).sngTome), 2)
                        Else
                            .Caption = ""
                        End If
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
                       
                '****** 20180825 START ***********************************
                Case mForeignStockMap
                    With frmEarvinStocks.lblForeignStock
                        .FontSize = 10
                        .ForeColor = QBColor(10)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngForeignStock > gudtStock(sngIndex - 1).sngForeignStock Then
                                flag_str = "↑"
                            ElseIf gudtStock(sngIndex).sngForeignStock < gudtStock(sngIndex - 1).sngForeignStock Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "外資=" & Str(Round(gudtStock(sngIndex).sngForeignStock, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lblSignForeignStock
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngForeignStock >= gudtStock(sngIndex - 1).sngForeignStock Then
                                .ForeColor = QBColor(12)
                                flag_str = "+"
                            ElseIf gudtStock(sngIndex).sngForeignStock < gudtStock(sngIndex - 1).sngForeignStock Then
                                .ForeColor = QBColor(10)
                                flag_str = ""
                            End If
                            .Caption = flag_str & Round(CSng(gudtStock(sngIndex).sngForeignStock) - CSng(gudtStock(sngIndex - 1).sngForeignStock), 2)
                        Else
                            .Caption = ""
                        End If
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
                       
                Case mSitAndCbStockMap
                    With frmEarvinStocks.lblSitAndCbStock
                        .FontSize = 10
                        .ForeColor = QBColor(10)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngSitAndCbStock > gudtStock(sngIndex - 1).sngSitAndCbStock Then
                                flag_str = "↑"
                            ElseIf gudtStock(sngIndex).sngSitAndCbStock < gudtStock(sngIndex - 1).sngSitAndCbStock Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "投信=" & Str(Round(gudtStock(sngIndex).sngSitAndCbStock, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lblSignSitAndCbStock
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngSitAndCbStock >= gudtStock(sngIndex - 1).sngSitAndCbStock Then
                                .ForeColor = QBColor(12)
                                flag_str = "+"
                            ElseIf gudtStock(sngIndex).sngSitAndCbStock < gudtStock(sngIndex - 1).sngSitAndCbStock Then
                                .ForeColor = QBColor(10)
                                flag_str = ""
                            End If
                            .Caption = flag_str & Round(CSng(gudtStock(sngIndex).sngSitAndCbStock) - CSng(gudtStock(sngIndex - 1).sngSitAndCbStock), 2)
                        Else
                            .Caption = ""
                        End If
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
                       
                Case mSelfEmployedStockMap
                    With frmEarvinStocks.lblSelfEmployedStock
                        .FontSize = 10
                        .ForeColor = QBColor(10)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngSelfEmployedStock > gudtStock(sngIndex - 1).sngSelfEmployedStock Then
                                flag_str = "↑"
                            ElseIf gudtStock(sngIndex).sngSelfEmployedStock < gudtStock(sngIndex - 1).sngSelfEmployedStock Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "自營商=" & Str(Round(gudtStock(sngIndex).sngSelfEmployedStock, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lblSignSelfEmployedStock
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngSelfEmployedStock >= gudtStock(sngIndex - 1).sngSelfEmployedStock Then
                                .ForeColor = QBColor(12)
                                flag_str = "+"
                            ElseIf gudtStock(sngIndex).sngSelfEmployedStock < gudtStock(sngIndex - 1).sngSelfEmployedStock Then
                                .ForeColor = QBColor(10)
                                flag_str = ""
                            End If
                            .Caption = flag_str & Round(CSng(gudtStock(sngIndex).sngSelfEmployedStock) - CSng(gudtStock(sngIndex - 1).sngSelfEmployedStock), 2)
                        Else
                            .Caption = ""
                        End If
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
                       
                Case mLegalPersonStockMap
                    With frmEarvinStocks.lblLegalPersonStock
                        .FontSize = 10
                        .ForeColor = QBColor(10)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngLegalPersonStock > gudtStock(sngIndex - 1).sngLegalPersonStock Then
                                flag_str = "↑"
                            ElseIf gudtStock(sngIndex).sngLegalPersonStock < gudtStock(sngIndex - 1).sngLegalPersonStock Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "法人=" & Str(Round(gudtStock(sngIndex).sngLegalPersonStock, 2)) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lblSignLegalPersonStock
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtStock(sngIndex).sngLegalPersonStock >= gudtStock(sngIndex - 1).sngLegalPersonStock Then
                                .ForeColor = QBColor(12)
                                flag_str = "+"
                            ElseIf gudtStock(sngIndex).sngLegalPersonStock < gudtStock(sngIndex - 1).sngLegalPersonStock Then
                                .ForeColor = QBColor(10)
                                flag_str = ""
                            End If
                            .Caption = flag_str & Round(CSng(gudtStock(sngIndex).sngLegalPersonStock) - CSng(gudtStock(sngIndex - 1).sngLegalPersonStock), 2)
                        Else
                            .Caption = ""
                        End If
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
                    
                '****** 20180825 END   ***********************************
                             
                Case m_Trendmap                     '' 2004/3/9
                    With frmEarvinStocks.lblTrend
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        If (gudtIndex(sngIndex).sngTrend > 0) Then
                            .ForeColor = RGB(255, 50, 100)
                        Else
                            .ForeColor = RGB(50, 200, 100)
                        End If
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngTrend > gudtIndex(sngIndex - 1).sngTrend Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngTrend < gudtIndex(sngIndex - 1).sngTrend Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "Trend = " & gudtIndex(sngIndex).sngTrend & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With

                Case m_QMmap
                    With frmEarvinStocks.lblQM
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .ForeColor = RGB(255, 100, 0)
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngQM > gudtIndex(sngIndex - 1).sngQM Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngQM < gudtIndex(sngIndex - 1).sngQM Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "QM = " & gudtIndex(sngIndex).sngQM & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
            
                Case m_Diffmap
                    With frmEarvinStocks.lblDiff1      '------- Diff1 -------
                        .FontSize = 10
                        .ForeColor = QBColor(12)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngDIFF1 > gudtIndex(sngIndex - 1).sngDIFF1 Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngDIFF1 < gudtIndex(sngIndex - 1).sngDIFF1 Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "Diff1=" & Str(gudtIndex(sngIndex).sngDIFF1) & flag_str
                        sngTemp = sngTemp - .Height
                    End With
                    With frmEarvinStocks.lblDiff2
                        .FontSize = 10
                        .ForeColor = QBColor(9)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngDIFF2 > gudtIndex(sngIndex - 1).sngDIFF2 Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngDIFF2 < gudtIndex(sngIndex - 1).sngDIFF2 Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "Diff2=" & Str(gudtIndex(sngIndex).sngDIFF2) & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With

                Case m_DCmap
                    With frmEarvinStocks.lblDC
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .ForeColor = RGB(255, 50, 100)
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngDC > gudtIndex(sngIndex - 1).sngDC Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngDC < gudtIndex(sngIndex - 1).sngDC Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "DC = " & gudtIndex(sngIndex).sngDC & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With

                Case m_WLSTmap
                    With frmEarvinStocks.lblWLST      '------- Diff1 -------
                        .FontSize = 10
                        .ForeColor = QBColor(13)
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngWLST > gudtIndex(sngIndex - 1).sngWLST Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngWLST < gudtIndex(sngIndex - 1).sngWLST Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "WLST=" & Str(gudtIndex(sngIndex).sngWLST) & flag_str
                        sngTemp = sngTemp - .Height
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
        
                Case m_WCYmap
                    With frmEarvinStocks.lblWCY
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        If (gudtIndex(sngIndex).sngWCY > 0) Then
                            .ForeColor = RGB(255, 50, 100)
                        Else
                            .ForeColor = RGB(50, 200, 100)
                        End If
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngWCY > gudtIndex(sngIndex - 1).sngWCY Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngWCY < gudtIndex(sngIndex - 1).sngWCY Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "WCY = " & gudtIndex(sngIndex).sngWCY & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With

                Case m_WLWmap
                    With frmEarvinStocks.lblWLW
                        .FontSize = 10
                        .Top = sngTemp
                        .Left = frmEarvinStocks.ScaleWidth - 0.98 * gsngRightLevel
                        .Width = gsngRightLevel * 0.96
                        If (gudtIndex(sngIndex).sngWLW > 0) Then
                            .ForeColor = RGB(255, 0, 0)
                        Else
                            .ForeColor = RGB(0, 255, 0)
                        End If
                        .Visible = True
                        If sngIndex <> 1 Then
                            If gudtIndex(sngIndex).sngWLW > gudtIndex(sngIndex - 1).sngWLW Then
                                flag_str = "↑"
                            ElseIf gudtIndex(sngIndex).sngWLW < gudtIndex(sngIndex - 1).sngWLW Then
                                flag_str = "↓"
                            Else
                                flag_str = "–"
                            End If
                        End If
                        .Caption = "WLW = " & gudtIndex(sngIndex).sngWLW & flag_str
                        sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                        For i = 1 To sngCount
                            sngTemp = sngTemp - mudtFrame(i).sngHeight
                        Next
                        sngTemp = sngTemp - 20
                    End With
            
                Case Else
                    sngTemp = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand
                    For i = 1 To sngCount
                        sngTemp = sngTemp - mudtFrame(i).sngHeight
                    Next
                    sngTemp = sngTemp - 20
            End Select
        Next
        ' 最右邊那條線
        With frmEarvinStocks
            frmEarvinStocks.Line (Abs(.ScaleWidth) - 0.01 * gsngRightLevel, gsngBottomFrame)-(Abs(.ScaleWidth) - 0.01 * gsngRightLevel, Abs(.ScaleHeight) - gsngTopFrame - gsngTopCommand), RGB(200, 100, 0)
        End With
        sngTemp = sngTemp - 60
    End If
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.DrawStockIndexes()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: 取得要顯示的股票清單
'* 輸入參數: stockType 取得股票清單的方式 ( "file" or "directory" )
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081223 調整格式及增加註解
'***************************************************************************************************
Public Sub GetStockList(stockType As String)
    Dim i As Integer
    Dim path As String
   
    On Error GoTo ERR_HANDLE
    
    path = GetAppPath
    frmEarvinStocks.cboStocks.Clear
    frmEarvinStocks.cboStocks.AddItem "Taiex", 0
    If stockType = "file" Then
        '======================================
        '* 以讀取檔案方式取得要處理的個股     *
        '======================================
        ' Read StockNo file form "taiwan_stkno.bat" file
        Dim openFile As String
        Dim lineData As String
        Dim splitData As Variant
        ' 路徑寫死，不好!!要再調整!!
'        openFile = "C:\myData\EarvinStockPGMs\myPGMs\DATA\config\操作個股.txt"
        path = GetAppPath
        openFile = path + "myPGMs\DATA\config\操作個股.txt"
        i = 1
        
        Open openFile For Input As #100
        While Not EOF(100)
            Line Input #100, lineData
'            splitData = Split(lineData, " ")
'            frmEarvinStocks.cboStocks.AddItem splitData(1), i
            frmEarvinStocks.cboStocks.AddItem lineData, i
            i = i + 1
        Wend
        Close #100
    ElseIf stockType = "directory" Then
        '=================================
        '* 以目錄方式取得要處理的個股    *
        '=================================
        ' 取得指定路徑下的檔案
        Dim thePath As String
'        thePath = "C:\myData\EarvinStockPGMs\myPGMs\DATA\dat\"
        thePath = GetAppPath + "myPGMs\DATA\dat\"
        Dim fileName As String
        Dim fileInfoList As Variant
        Dim stockName As String
        
        ' 20230309 將取得檔案名稱，先寫入list後再新增至combo-list元件(反轉檔案順序)
'        fileName = Dir(thePath) ' 取得檔案名稱(一次返回一個檔案名稱)
'        Do While fileName <> ""
'            fileInfoList = Split(fileName, ".")
'            stockName = fileInfoList(0) + ", " + GetStockName(fileInfoList(0))
'            frmEarvinStocks.cboStocks.AddItem stockName, i
'            fileName = Dir() ' 再次調用Dir函數,此時可以不帶參數
'        Loop
        Dim stockList(5000) As String   ' 儲存取得之檔案名稱(最多可儲存5000筆)
        i = 1
        fileName = Dir(thePath)
        Do While fileName <> ""
            stockList(i) = fileName
            i = i + 1
            fileName = Dir() ' 再次調用Dir函數,此時可以不帶參數
        Loop
        Dim j As Integer
        For j = 1 To UBound(stockList)
            If stockList(j) <> "" Then
                fileInfoList = Split(stockList(j), ".")
                stockName = fileInfoList(0) + ", " + GetStockName(fileInfoList(0))
                frmEarvinStocks.cboStocks.AddItem stockName
            Else
                j = UBound(stockList)
            End If
             Debug.Print (Str(j))
        Next
    End If
'   frmEarvinStocks.cboStocks.Text = "Taiex"
    frmEarvinStocks.cboStocks.Text = SHOW_DEFAULT_STOCK
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.GetStockList()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: 設定存放個股資料的路徑(取得程式存放路徑的上1層(ex: D:\A\B --> D:\A\))
'* 輸入參數: 無
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081223 調整格式及增加註解
'***************************************************************************************************
Public Function GetAppPath() As String
    Dim blnFlag As Boolean
    Dim path As String
    Dim strChar As String
    Dim intUpDir As Integer
    Const upLevel = 2 ' 往上??層
    Dim i As Integer
    
    blnFlag = False
    intUpDir = 0
    path = App.path
    
    On Error GoTo ERR_HANDLE
   
    For i = Len(path) To 1 Step -1
        strChar = Mid(path, i, 1)
        If strChar = "\" Then
            intUpDir = intUpDir + 1
        End If
        If intUpDir = upLevel Then ' 往上1層
            path = Mid(path, 1, i)
            Exit For
        End If
    Next
    GetAppPath = path
    
    Exit Function
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.GetAppPath()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Function


Public Function GetStockName(ByVal stockCode As String)
    Dim stockName As String
    Dim path As String
    Dim openFile As String
    Dim lineData As String
    Dim splitData As Variant
    Dim i As Integer
   
    On Error GoTo ERR_HANDLE
    
    If stockCode = "Taiex" Then
        stockName = "加權指數"
    Else
        ' 路徑的寫法要再調整，這樣還是寫死!! // 考慮把資料寫到table
        path = GetAppPath
        openFile = path + "myPGMs\DATA\config\股票代號名稱對照表.txt"
        i = 1
   
        Open openFile For Input As #100
        While Not EOF(100)
            Line Input #100, lineData
            splitData = Split(lineData, " ")
         
            If splitData(0) = stockCode Then
                stockName = splitData(1)
                GoTo END_WHILE
            End If
         
            i = i + 1
        Wend

END_WHILE:
      
        Close #100
   
    End If

    GetStockName = stockName
    
    Exit Function
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.GetStockName()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Function


'***************************************************************************************************************************
'* 說    明:
'*    依傳入的MAP轉換成對應的數值做split動作      --** 20071214 OK **--
'* 輸入參數:
'*    indexName: 技術指標名稱
'* 輸出參數: 無
'* 版    本:
'*    2.00  20080916 Earvin   新增
'* 備    註:
'*    1. sngSector = 1 表示為高點區段；sngSector = 0.5 表示為低點區段
'***************************************************************************************************************************
Private Function getIndexDays(ByVal indexName As String) As Integer
    If indexName = "3MAP" Then
        getIndexDays = 3
    ElseIf indexName = "4MAP" Then
        getIndexDays = 4
    ElseIf indexName = "5MAP" Then
        getIndexDays = 5
    ElseIf indexName = "6MAP" Then
        getIndexDays = 6
    ElseIf indexName = "8MAP" Then
        getIndexDays = 8
    ElseIf indexName = "10MAP" Then
        getIndexDays = 10
    ElseIf indexName = "12MAP" Then
        getIndexDays = 12
    ElseIf indexName = "20MAP" Then
        getIndexDays = 20
    ElseIf indexName = "24MAP" Then
        getIndexDays = 24
    ElseIf indexName = "30MAP" Then
        getIndexDays = 30
    ElseIf indexName = "60MAP" Then
        getIndexDays = 60
    ElseIf indexName = "72MAP" Then
        getIndexDays = 72
    ElseIf indexName = "120MAP" Then
        getIndexDays = 120
    ElseIf indexName = "144MAP" Then
        getIndexDays = 144
    ElseIf indexName = "240MAP" Then
        getIndexDays = 240
    ElseIf indexName = "288MAP" Then
        getIndexDays = 288
    ElseIf indexName = "K" Then
        getIndexDays = 600
    ElseIf indexName = "D" Then
        getIndexDays = 601
    Else
'      getIndexDays = 288
        Err.Raise 100
    End If
End Function



'***************************************************************************************************************************
'* 說    明:
'*    分割出每個Sector: use MAP      --** 20071214 OK **--
'* 輸入參數:
'*    udtStock() 每日股價資料(原始資料)
'*    udtIndex() 每日股價資料(技術指標)
'*    intStockNo 股價資料筆數
'*    indexName1 傳入K -- 起日
'*    indexName2 傳入D -- 迄日
'* 輸出參數: 無
'* 版    本:
'*    2.00  20080916 Earvin   新增
'* 備    註:
'*    1. sngSector = 1 表示為高點區段；sngSector = 0.5 表示為低點區段
'***************************************************************************************************************************
Public Sub setSectorByKD(ByRef udtStock() As StockData, _
                        ByRef udtIndex() As IndexData, _
                        ByVal intStockNo As Integer, _
                        ByVal indexDay As String)
    Dim signalFlag As Single       ' 值=1表高點、值=0.5表低點
    Dim indexDay1 As Integer     ' 將傳入的MAP字串轉換為數值
    Dim indexDay2 As Integer     ' 將傳入的MAP字串轉換為數值
    Dim nowIndex1 As Single   ' 條件一：目前資料的MAP值
    Dim prevIndex1 As Single   ' 條件一：前一筆資料的MAP值
    Dim nowIndex2 As Single   ' 條件二：目前資料的MAP值
    Dim prevIndex2 As Single   ' 條件二：前一筆資料的MAP值
    Dim i As Integer
      
    On Error GoTo ERR_HANDLE
    
    '--------------------------------------------------------------------
    '* Initialize Variables
    '--------------------------------------------------------------------
   
    Call subKD(udtStock, udtIndex, intStockNo, indexDay)

    indexDay1 = getIndexDays("K")
    indexDay2 = getIndexDays("D")
    signalFlag = HIGH_SIGNAL
    gintSector = 0
   
    For i = 2 To intStockNo
        With udtIndex(i)
            ' 儲存所選擇的第一個條件的值
            Select Case indexDay1
                Case 600    ' K value
                    nowIndex1 = .sngK
                    prevIndex1 = udtIndex(i - 1).sngK
                Case 601    ' D value
                    nowIndex1 = .sngD
                    prevIndex1 = udtIndex(i - 1).sngD
            End Select
            '--- 儲在所選擇的第二個條件的值 ---
            Select Case indexDay2
                Case 600    ' K value
                    nowIndex2 = .sngK
                    prevIndex2 = udtIndex(i - 1).sngK
                Case 601    ' D value
                    nowIndex2 = .sngD
                    prevIndex2 = udtIndex(i - 1).sngD
            End Select
            
            '--------------------------------------------------------------------
            '* 記錄該點是屬於高點或低點
            '--------------------------------------------------------------------
            If (nowIndex1 > nowIndex2) Or _
                ((nowIndex1 = nowIndex2) And (prevIndex1 < prevIndex2)) Then
                signalFlag = HIGH_SIGNAL     ' 高點
            Else
                signalFlag = LOW_SIGNAL   ' 低點
            End If
            .sngSector = signalFlag
        End With
    Next
   
    Exit Sub
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.setSectorByKD()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    End Select
End Sub


'***************************************************************************************************************************
'* 說    明:
'*    分割出每個Sector -- use MACD  --** 20071218 OK **--
'* 輸入參數:
'*    udtStock 每日股價資料(原始資料)
'*    udtIndex 每日股價資料(技術指標)
'*    intStockNo 股價資料筆數
'*    strDayNo 區間內最少天數(小於此值者需合併)
'*    strStkDis 區間內最高與最低點差距值(小於此值者需合併)
'*    indexName1 傳入的MAP字串--起日
'*    indexName2 傳入的MAP字串--迄日
'* 輸出參數: 無
'* 版    本:
'*    2.00  20080916 Earvin   新增
'* 備    註:
'*    1. sngSector = 1 表示為高點區段；sngSector = 0.5 表示為低點區段
'***************************************************************************************************************************
Public Sub setSectorByMACD(ByRef udtStock() As StockData, _
                        ByRef udtIndex() As IndexData, _
                        ByVal intStockNo As Integer)
    Dim signalFlag As Single       ' 值=1表高點、值=0.5表低點
    Dim i As Integer
      
    On Error GoTo ERR_HANDLE
    
    '--------------------------------------------------------------------
    '* Initialize Variables
    '--------------------------------------------------------------------
    signalFlag = 1
    gintSector = 0
    
    For i = 2 To intStockNo
        With udtIndex(i)
            If udtStock(i).sngDate >= 851015 Then
                If .sngDIF_MACD >= 0 Then
                    signalFlag = HIGH_SIGNAL    ' 高點
                Else
                    signalFlag = LOW_SIGNAL     ' 低點
                End If
                .sngSector = signalFlag
            End If
        End With
    Next
        
    Exit Sub
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.setSectorByMACD()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    End Select
End Sub


'***************************************************************************************************************************
'* 說    明:
'*    分割出每個Sector: use MAP
'* 輸入參數:
'*    udtStock() 每日股價資料(原始資料)
'*    udtIndex() 每日股價資料(技術指標)
'*    intStockCount 股價資料筆數
'*    splitDay 切割區間的天數
'* 輸出參數: 無
'* 版    本:
'*    1.00  20090214 Earvin   新增
'* 備    註:
'*    1. sngSector = 1 表示為高點區段；sngSector = 0.5 表示為低點區段
'***************************************************************************************************************************
Public Sub setSectorByFixedDay(ByRef udtStock() As StockData, _
                        ByRef udtIndex() As IndexData, _
                        ByVal intStockCount As Integer, _
                        ByVal intSplitDay As Integer)
    Dim sngSignalFlag As Single   ' 值=1表高點、值=0.5表低點
    Dim intDayCount As Integer
    Dim i As Integer
      
    On Error GoTo ERR_HANDLE
    
    '--------------------------------------------------------------------
    '* Initialize Variables
    '--------------------------------------------------------------------
   
    sngSignalFlag = LOW_SIGNAL
    intDayCount = 1
   
    For i = 1 To intStockCount
        With udtIndex(i)
            .sngSector = sngSignalFlag
            If intDayCount >= intSplitDay Then
                If sngSignalFlag = HIGH_SIGNAL Then
                    sngSignalFlag = LOW_SIGNAL
                Else
                    sngSignalFlag = HIGH_SIGNAL
                End If
                intDayCount = 1
            Else
                intDayCount = intDayCount + 1
            End If
        End With
    Next
   
    Exit Sub
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.setSectorByFixedDay()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    End Select
End Sub









































'Not Finished -- 921212
'***********************************************************************************
'* Quick Sort Method
'***********************************************************************************
Public Sub QuickSort(udtStkData() As StockData, intStart As Integer, _
                    intEnd As Integer)
    Dim intPosOfSplitter As Integer
    
    On Error GoTo ERR_HANDLE
    
    If intEnd > intStart Then
        Partition udtStkData(), intStart, intEnd
        QuickSort udtStkData(), intStart, intPosOfSplitter - 1
        QuickSort udtStkData(), intPosOfSplitter + 1, intEnd
    End If
        
    Exit Sub
    
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.QuickSort()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'Not Finished -- 921212
'***********************************************************************************
'* Partition Method use for Qucik Sort Method
'***********************************************************************************
Private Sub Partition(udtStkData() As StockData, intStart As Integer, _
                        intEnd As Integer)
    Dim intSplitPos As Integer, intNewStart As Integer
    Dim i As Integer
    
End Sub


'***************************************************************************************************************************
'* 說    明:
'*    將分割出的Sector其漲跌幅小於指定門檻刪除之
'* 輸入參數:
'*    udtStock 股市原始資料
'*    udtIndex 股市技術指標
'*    intStockNo 資料筆數
'*    strFluctuation 最低漲跌幅門檻
'* 輸出參數: 無
'* 版    本:
'*    2.00  20080916 Earvin   新增
'* 備    註:
'*    1. sngSector = 1 表示為高點區段；sngSector = 0.5 表示為低點區段
'***************************************************************************************************************************
Public Sub MeetFluctuation(ByRef udtStock() As StockData, _
                            ByRef udtIndex() As IndexData, _
                            ByVal intStockNo As Integer, _
                            ByVal strFluctuation As String)
    Dim sngHigh As Single
    Dim sngLow As Single
    Dim sngStart As Single
    Dim sngFluctuation As Single
    Dim sngSector As Single
    Dim intBegPos As Integer
    Dim intEndPos As Integer
    Dim i As Integer, j As Integer
    
    On Error GoTo ERR_HANDLE
    
    '--- Initialize Variables ---
    sngFluctuation = 0
    sngHigh = -1
    sngLow = 99999
    sngStart = 0
    sngSector = udtIndex(1).sngSector
    intBegPos = 1
    intEndPos = 1
    
    For i = 2 To intStockNo
        With udtIndex(i)
            If .sngSector <> sngSector Then
                sngStart = udtStock(intBegPos).sngEndprice  ' 該區段的開盤價
                If sngSector = 1 Then
                    For j = intBegPos To intEndPos
                        '--- 找高點 ---
                        If udtStock(j).sngEndprice > sngHigh Then
                            sngHigh = udtStock(j).sngEndprice
                        End If
                    Next
                Else
                    For j = intBegPos To intEndPos
                        '--- 找低點 ---
                        If udtStock(j).sngEndprice < sngLow Then
                            sngLow = udtStock(j).sngEndprice
                        End If
                    Next
                End If
                '--- 判斷最高點與起始價的差距幅度 ---
                If sngSector = 1 Then
                    sngFluctuation = (sngHigh - sngStart) / sngStart
                Else
                    sngFluctuation = (sngStart - sngLow) / sngStart
                End If
                '--- 小於門檻值，則清成0 ---
                If sngFluctuation < CSng(strFluctuation / 100) Then
                    For j = intBegPos To intEndPos
                        udtIndex(j).sngSector = 0
                    Next
                End If
                '--- Reset the variables and continue next sector ---
                sngFluctuation = 0
                sngHigh = -1
                sngLow = 99999
                sngStart = 0
                intBegPos = i
                intEndPos = i
                sngSector = udtIndex(i).sngSector
            Else
                intEndPos = i
            End If
        End With
    Next
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.MeetFluctuation()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************************************
'* 說    明:
'*    分割出每個Sector: use MAP      --** 20071214 Finished **--
'* 輸入參數:
'*    udtStock() 每日股價資料(原始資料)
'*    udtIndex() 每日股價資料(技術指標)
'*    intStockNo 股價資料筆數
'*    strDayNo 區間內最少天數(小於此值者需合併)
'*    strStkDis 區間內最高與最低點差距值(小於此值者需合併)
'*    indexName1 傳入的MAP字串--起日
'*    indexName2 傳入的MAP字串--迄日
'* 輸出參數: 無
'* 版    本:
'*    2.00  20080916 Earvin   新增
'* 備    註:
'*    1. sngSector = 1 表示為高點區段；sngSector = 0.5 表示為低點區段
'***************************************************************************************************************************
Public Sub setSectorByMAP(ByRef udtStock() As StockData, _
                        ByRef udtIndex() As IndexData, _
                        ByVal intStockNo As Integer, _
                        ByVal indexName1 As String, _
                        ByVal indexName2 As String)
    Dim signalFlag As Single    ' 值=1表高點、值=0.5表低點
    Dim indexDay1 As Integer    ' 將傳入的MAP字串轉換為數值
    Dim indexDay2 As Integer    ' 將傳入的MAP字串轉換為數值
    Dim nowIndex1 As Single     ' 條件一：目前資料的MAP值
    Dim prevIndex1 As Single    ' 條件一：前一筆資料的MAP值
    Dim nowIndex2 As Single     ' 條件二：目前資料的MAP值
    Dim prevIndex2 As Single    ' 條件二：前一筆資料的MAP值
    Dim intDayNo As Integer     ' 區間內最少天數(小於此值者需合併)
    Dim intStkDis As Integer    ' 區間內最高與最低點差距值(小於此值者需合併)
    Dim i As Integer
      
    On Error GoTo ERR_HANDLE
    
    '--------------------------------------------------------------------
    '* Initialize Variables
    '--------------------------------------------------------------------
    indexDay1 = getIndexDays(indexName1)
    indexDay2 = getIndexDays(indexName2)
    signalFlag = 1
    gintSector = 0
   
    For i = 2 To intStockNo
        With udtIndex(i)
            '--------------------------------------------------------------------
            '* 目前是用均線來做split sector
            '* 而選擇的均線日期長短會影到split結果
            '* 1.長期均線可omit掉短期小波動，但對快速而短期的大波動
            '*   卻無法反應
            '* 2.短期均線可抓到較多的波動，但因split sector過多，而
            '*   無法抓到整個趨勢的最高點
            '*
            '* 所以這個split 的動作可以再改進，以兼顧長短期之優點，
            '* 去掉其各自的缺點
            '--------------------------------------------------------------------
            ' 儲存所選擇的第一個條件的值
            Select Case indexDay1
                Case 3      ' 3MAP
                    nowIndex1 = .sngP3
                    prevIndex1 = udtIndex(i - 1).sngP3
                Case 4      ' 4MAP
                    nowIndex1 = .sngP4
                    prevIndex1 = udtIndex(i - 1).sngP4
                Case 5      ' 5MAP
                    nowIndex1 = .sngP5
                    prevIndex1 = udtIndex(i - 1).sngP5
                Case 6      ' 6MAP
                    nowIndex1 = .sngP6
                    prevIndex1 = udtIndex(i - 1).sngP6
                Case 8      ' 8MAP
                    nowIndex1 = .sngP8
                    prevIndex1 = udtIndex(i - 1).sngP8
                Case 10     ' 10MAP
                    nowIndex1 = .sngP10
                    prevIndex1 = udtIndex(i - 1).sngP10
                Case 12     ' 12MAP
                    nowIndex1 = .sngP12
                    prevIndex1 = udtIndex(i - 1).sngP12
                Case 20     ' 20MAP
                    nowIndex1 = .sngp20
                    prevIndex1 = udtIndex(i - 1).sngp20
                Case 24     ' 24MAP
                    nowIndex1 = .sngP24
                    prevIndex1 = udtIndex(i - 1).sngP24
                Case 30     ' 30MAP
                    nowIndex1 = .sngP30
                    prevIndex1 = udtIndex(i - 1).sngP30
                Case 60     ' 60MAP
                    nowIndex1 = .sngP60
                    prevIndex1 = udtIndex(i - 1).sngP60
                Case 72     ' 72MAP
                    nowIndex1 = .sngP72
                    prevIndex1 = udtIndex(i - 1).sngP72
                Case 120    ' 120MAP
                    nowIndex1 = .sngP120
                    prevIndex1 = udtIndex(i - 1).sngP120
                Case 144    ' 144MAP
                    nowIndex1 = .sngP144
                    prevIndex1 = udtIndex(i - 1).sngP144
                Case 240    ' 240MAP
                    nowIndex1 = .sngP240
                    prevIndex1 = udtIndex(i - 1).sngP240
                Case 288    ' 288MAP
                    nowIndex1 = .sngP288
                    prevIndex1 = udtIndex(i - 1).sngP288
                Case Else
'               nowIndex1 = .sngP288
'               prevIndex1 = udtIndex(i - 1).sngP288
                    Err.Raise 100
            End Select
            '--- 儲在所選擇的第二個條件的值 ---
            Select Case indexDay2
                Case 3      ' 3MAP
                    nowIndex2 = .sngP3
                    prevIndex2 = udtIndex(i - 1).sngP3
                Case 4      ' 4MAP
                    nowIndex2 = .sngP4
                    prevIndex2 = udtIndex(i - 1).sngP4
                Case 5      ' 5MAP
                    nowIndex2 = .sngP5
                    prevIndex2 = udtIndex(i - 1).sngP5
                Case 6      ' 6MAP
                    nowIndex2 = .sngP6
                    prevIndex2 = udtIndex(i - 1).sngP6
                Case 8      ' 8MAP
                    nowIndex2 = .sngP8
                    prevIndex2 = udtIndex(i - 1).sngP8
                Case 10     ' 10MAP
                    nowIndex2 = .sngP10
                    prevIndex2 = udtIndex(i - 1).sngP10
                Case 12     ' 12MAP
                    nowIndex2 = .sngP12
                    prevIndex2 = udtIndex(i - 1).sngP12
                Case 20     ' 20MAP
                    nowIndex2 = .sngp20
                    prevIndex2 = udtIndex(i - 1).sngp20
                Case 24     ' 24MAP
                    nowIndex2 = .sngP24
                    prevIndex2 = udtIndex(i - 1).sngP24
                Case 30     ' 30MAP
                    nowIndex2 = .sngP30
                    prevIndex2 = udtIndex(i - 1).sngP30
                Case 60     ' 60MAP
                    nowIndex2 = .sngP60
                    prevIndex2 = udtIndex(i - 1).sngP60
                Case 72     ' 72MAP
                    nowIndex2 = .sngP72
                    prevIndex2 = udtIndex(i - 1).sngP72
                Case 120    ' 120MAP
                    nowIndex2 = .sngP120
                    prevIndex2 = udtIndex(i - 1).sngP120
                Case 144    ' 144MAP
                    nowIndex2 = .sngP144
                    prevIndex2 = udtIndex(i - 1).sngP144
                Case 240    ' 240MAP
                    nowIndex2 = .sngP240
                    prevIndex2 = udtIndex(i - 1).sngP240
                Case 288    ' 288MAP
                    nowIndex2 = .sngP288
                    prevIndex2 = udtIndex(i - 1).sngP288
                Case Else
'               nowIndex2 = .sngP288
'               prevIndex2 = udtIndex(i - 1).sngP288
                    Err.Raise 100
            End Select
            
            '--------------------------------------------------------------------
            '* 記錄該點是屬於高點或低點
            '--------------------------------------------------------------------
            If (nowIndex1 > nowIndex2) Or _
                ((nowIndex1 = nowIndex2) And (prevIndex1 < prevIndex2)) Then
                signalFlag = HIGH_SIGNAL     ' 高點
            Else
                signalFlag = LOW_SIGNAL   ' 低點
            End If
            .sngSector = signalFlag
        End With
    Next
   
    Exit Sub
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.setSectorByMAP()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    End Select
End Sub



'-----------------20080918
'***********************************************************************************
'* 依傳入目前資料位置找到對應的日期
'* Input Param.
'*   udtStock  ：每日股價資料
'*   intStockNo：目前資料筆數
'*   strStkDate：傳入的日期
'* Return : 傳回該日期在資料集中的位置 (Type：Integer)
'***********************************************************************************
'***************************************************************************************************************************
'* 說    明:
'*    分割出每個Sector: use MAP      --** 20071214 Finished **--
'* 輸入參數:
'*    udtStock() 每日股價資料(原始資料)
'*    udtIndex() 每日股價資料(技術指標)
'*    intStockNo 股價資料筆數
'*    strDayNo 區間內最少天數(小於此值者需合併)
'*    strStkDis 區間內最高與最低點差距值(小於此值者需合併)
'*    indexName1 傳入的MAP字串--起日
'*    indexName2 傳入的MAP字串--迄日
'* 輸出參數: 無
'* 版    本:
'*    2.00  20080916 Earvin   新增
'* 備    註:
'*    1. sngSector = 1 表示為高點區段；sngSector = 0.5 表示為低點區段
'***************************************************************************************************************************
Public Function ReturnIndexPos(ByRef udtStock() As StockData, _
                                ByVal intStockNo As Integer, _
                                ByVal strStkDate As String) As Integer
    Dim intStkPos As Integer
    Dim i As Integer
    
    On Error GoTo ERR_HANDLE
    
    For i = 1 To intStockNo
        '--- 找到符合的資料，傳回後結束該函式 ---
        If udtStock(i).sngDate >= strStkDate Then
            ReturnIndexPos = i
            Exit Function
        End If
    Next
      
    Err.Raise 10000, "ReturnIndexPos", "找不到指定的日期資料"

    Exit Function
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.ReturnIndexPos()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Function


'***********************************************************************************
'* 計算在每個sector區間的GRG -- use MAP
'* Input Param.
'*   udtStock  ：每日股價資料
'*   udtIndex  ：儲存的指數資料
'*   intStockNo：資料筆數
'*   strStartDT：傳入的資料起始日期
'*   strEndDT  ：傳入的資料結束日期
'***********************************************************************************
Public Sub GetHighLowPoints(ByRef udtStock() As StockData, _
                            ByRef udtIndex() As IndexData, _
                            ByVal intStockNo As Integer, _
                            ByVal strStartDT As String, _
                            ByVal strEndDT As String)
    Dim intPos As Integer           ' 目前計算的位置
    Dim intStartPos As Integer      ' 用來記錄每個sector的起始資料的位置
    Dim intEndPos As Integer        ' 用來記錄每個sector的最後一筆資料的位置
    Dim intTmpStockNo As Integer    ' 記錄所傳入的結束日期來取得在資料(udtStock)中的位置
    Dim signalFlag As Single        ' Sector 內容 (1 : 高點；0.5 : 低點)
    Dim udtGRG() As udtGRGIndex     ' 儲存要做聚類的資料
    Dim udtGroup() As Boolean       ' 記錄是那幾筆資料聚成高(低)點
    Dim intGRGLine As Integer       ' 記錄目前是那筆資料為母數列
    Dim blnFlag As Boolean          ' 判斷此次計算是否結束
    Dim intMainPos As Integer       ' 母數列的位置
    Dim intClusterPos As Integer    ' 儲存cluster 位置
    Dim sngHighLowDate As Single    ' 記錄計算的聚類資料中的最高點或最低點是那一天
    Dim i As Integer, j As Integer, k As Integer
    '--- 設定Output File, For Evaluation Use ---
    Dim intOutputFile As Integer
    Dim OutputFileName As String
    Dim strOutData As String

    On Error GoTo ERR_HANDLE
    
    '--- Initialize Variables ---
    ReDim gudtClusterDat(100)   ' 記錄每個sector最後聚類結果的值 (最多可記100個)
    intClusterPos = 1           ' 記錄目前gudtClusterDat陣列用的指標
    
    '--- Open file to record the results ---
    OutputFileName = "ClusterPattern.csv"
    intOutputFile = FreeFile()
    Open OutputFileName For Output As #intOutputFile
    
    '----------------------------------------------------------------------------
    '--- Write Header ---
    strOutData = "sector起日,sector迄日,sector最高點日期," & _
                "sector最高點日期之收盤價,sector最低點日期," & _
                "sector最低點日期之收盤價,使用聚類的日期," & _
                "GRG,parm1-name,parm1,parm2-name,parm2,parm3-name,parm3," & _
                "parm4-name,parm4,parm5-name,parm5,parm6-name,parm6," & _
                "parm7-name,parm7,parm8-name,parm8,Select-Count"
    Print #intOutputFile, strOutData
    '----------------------------------------------------------------------------
    
    '--- 傳入起始日期來取得在資料(udtStock)中的位置 ---
    intGRGLine = ReturnIndexPos(udtStock, intStockNo, strStartDT)
    '--- 傳入結束日期來取得在資料(udtStock)中的位置 ---
    intTmpStockNo = ReturnIndexPos(udtStock, intStockNo, strEndDT)
    
    i = intGRGLine
    signalFlag = udtIndex(i).sngSector
    '--- 儲存所要計算高點(低點)的sector的值，並找出起始位置 ---
    While signalFlag <> gsngHighLow
        i = i + 1
        If i > intTmpStockNo Then
            Err.Raise 10001, "(GetHighLowPoints)", "找不到符合資料，則產生錯誤訊息"
        End If
        signalFlag = udtIndex(i).sngSector
    Wend
    '--- Assign start line position ---
    intPos = i
    intStartPos = i
    intEndPos = i
    
    '--------------------------------------------------------------------
    '* 計算在 strStartDT 及 strEndDT 間資料聚類
    '--------------------------------------------------------------------
    While intEndPos < intTmpStockNo
        '--- 尋找每個sector起迄點 ---
        intStartPos = intEndPos
        While (signalFlag = udtIndex(intEndPos).sngSector) And intEndPos < intStockNo
            intEndPos = intEndPos + 1
        Wend
        
        '--- 將該區間的值儲在至GRG variables，最後一筆存放母數列的資料 ---
        intMainPos = intEndPos - intStartPos + 1    ' 計算該sector的筆數 (此值也表示母數列將來存放的位置)
        
        '---------------------------------------------------------------------------------------
        '--- 記錄每個sector的起始日、截止日 -- 920923 ---
        strOutData = udtStock(intStartPos).sngDate & "," & udtStock(intEndPos).sngDate
        k = FindHighLowDate(udtStock, intStartPos, intEndPos, 1)        ' sector間的高點
        strOutData = strOutData & "," & udtStock(k).sngDate & "," & udtStock(k).sngEndprice
        k = FindHighLowDate(udtStock, intStartPos, intEndPos, 0.5)      ' sector間的低點
        strOutData = strOutData & "," & udtStock(k).sngDate & "," & udtStock(k).sngEndprice
        '---------------------------------------------------------------------------------------
        
        ReDim udtGRG(intMainPos)    ' 儲存要做聚類的資料
        ReDim udtGroup(intMainPos)  ' 儲存聚類的結果 (在所設定的條件下，那幾筆最後會聚在一起)
 
        '--- Assign 區間內各項指標至udtGRG (最多可選擇8個指標) ---
        k = intStartPos
        For j = 1 To intMainPos - 1
            udtGRG(j).sngDate = udtStock(k).sngDate
            udtGRG(j).sngParam1 = GetFactorValue(udtIndex, k, 1)
            udtGRG(j).sngParam2 = GetFactorValue(udtIndex, k, 2)
            udtGRG(j).sngParam3 = GetFactorValue(udtIndex, k, 3)
            udtGRG(j).sngParam4 = GetFactorValue(udtIndex, k, 4)
            udtGRG(j).sngParam5 = GetFactorValue(udtIndex, k, 5)
            udtGRG(j).sngParam6 = GetFactorValue(udtIndex, k, 6)
            udtGRG(j).sngParam7 = GetFactorValue(udtIndex, k, 7)
            udtGRG(j).sngParam8 = GetFactorValue(udtIndex, k, 8)
            k = k + 1
        Next
        
        '--- 找出要做母數列(該sector的最高點或最低點)的資料assing至最後一列 ---
        i = FindHighLowDate(udtStock, intStartPos, intEndPos, signalFlag)
        '--- Assign 母數列到最後一列 ---
        With udtGRG(intMainPos)
            sngHighLowDate = udtStock(i).sngDate
            
            
            .sngDate = udtStock(i).sngDate
            .sngParam1 = GetFactorValue(udtIndex, i, 1)
            .sngParam2 = GetFactorValue(udtIndex, i, 2)
            .sngParam3 = GetFactorValue(udtIndex, i, 3)
            .sngParam4 = GetFactorValue(udtIndex, i, 4)
            .sngParam5 = GetFactorValue(udtIndex, i, 5)
            .sngParam6 = GetFactorValue(udtIndex, i, 6)
            .sngParam7 = GetFactorValue(udtIndex, i, 7)
            .sngParam8 = GetFactorValue(udtIndex, i, 8)
        End With
        '--- Preprocessing ---
        Call PreProcess(udtGRG, intMainPos)
        '--- Calcuate GRG ---
        Call CalGRG(udtGRG, udtGroup, intMainPos, sngHighLowDate)
        '--------------------------------------------------------------------
        '* 記錄下這次聚類的結果 -- 920914
        '--------------------------------------------------------------------
        If udtGRG(intMainPos).sngGRG <> -1 Then
            gudtClusterDat(intClusterPos) = udtGRG(intMainPos)
            ' 將聚類結果的attribute值還原
            Call RestoreData(gudtClusterDat, intClusterPos)
            
            '-----------------------------------------------------------------------------------------------
            '* Write-out, File-name : ccu.csv
            '* 再做評估時，以此檔案來做GRG -- 921212
            '-----------------------------------------------------------------------------------------------
            ' 20080228 marked for compiler...
'''            strOutData = strOutData & "," & _
'''                        gudtClusterDat(intClusterPos).sngDate & "," & _
'''                        gudtClusterDat(intClusterPos).sngGRG & "," & _
'''                        frmGeryExp.txtAttrName(0).Text & "," & gudtClusterDat(intClusterPos).sngParam1 & "," & _
'''                        frmGeryExp.txtAttrName(1).Text & "," & gudtClusterDat(intClusterPos).sngParam2 & "," & _
'''                        frmGeryExp.txtAttrName(2).Text & "," & gudtClusterDat(intClusterPos).sngParam3 & "," & _
'''                        frmGeryExp.txtAttrName(3).Text & "," & gudtClusterDat(intClusterPos).sngParam4 & "," & _
'''                        frmGeryExp.txtAttrName(4).Text & "," & gudtClusterDat(intClusterPos).sngParam5 & "," & _
'''                        frmGeryExp.txtAttrName(5).Text & "," & gudtClusterDat(intClusterPos).sngParam6 & "," & _
'''                        frmGeryExp.txtAttrName(6).Text & "," & gudtClusterDat(intClusterPos).sngParam7 & "," & _
'''                        frmGeryExp.txtAttrName(7).Text & "," & gudtClusterDat(intClusterPos).sngParam8 & "," & _
'''                        gintSelCount
            Print #intOutputFile, strOutData
            '-----------------------------------------------------------------------------------------------
            
            intClusterPos = intClusterPos + 1
        
            '---------------------------------------------------------------------------
            '* 930324: 將該sector的資料的GRG值Assign回udtindex()供程式display the value
            '---------------------------------------------------------------------------
            i = 1
            gsngCompFa = frmGeryExp.txtGRGThreshold.Text
            gsngCompFa2 = frmGeryExp.txtGRGThreshold2.Text
            
            For intPos = intStartPos To intEndPos - 1
                '--- 920923 : 將未滿足的點之GRG值清成0 ####### ---
                If udtGRG(i).sngGRG >= gsngCompFa Then
                    udtIndex(intPos).sngGRG = udtGRG(i).sngGRG
                ElseIf udtGRG(i).sngGRG <= gsngCompFa2 Then
                    udtIndex(intPos).sngGRG = udtGRG(i).sngGRG
                Else
                    udtIndex(intPos).sngGRG = 0
                End If
                i = i + 1
            Next
        End If
        '--- 應要判斷是算高點或是低點，以便抓取下一個符合的sector運算 ---
        signalFlag = udtIndex(intEndPos).sngSector
        '--- 儲存第一個sector值 ---
        While signalFlag <> gsngHighLow
            intEndPos = intEndPos + 1
            signalFlag = udtIndex(intEndPos).sngSector
            If intEndPos > intTmpStockNo Then
                '--- 已找完所有sector，並無符合資料可供計算 ---
                signalFlag = gsngHighLow
            End If
        Wend
    Wend

    Close #intOutputFile
    
'    '--------------------------------------------------------------------
'    '* 計算90年以後的聚類狀況 -- 920914
'    '* 920917：問題
'    '*         如何將未來的資料與計算出來的高、低點pattern做比較??
'    '* 920922：目前用GRGa的方式在做，但是效果不好…
'    '--------------------------------------------------------------------
'    Call CalculateHighLowPoints(udtStock, udtIndex, intStockNo, gudtClusterDat, gintClusterCnt)
       
    Exit Sub
    
ERR_HANDLE:
    Select Case Err.Number
        Case 6
            MsgBox "[GCFR_General_Module.GetHighLowPoints()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
            Resume Next
        Case 10001
            MsgBox "[GCFR_General_Module.GetHighLowPoints()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
        Case Else
            MsgBox "[GCFR_General_Module.GetHighLowPoints()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
   
    Err.Clear
    Resume Next
End Sub


'***********************************************************************************
'* 計算在每個sector區間的GRG (use Days)
'* Input Param.
'*   udtStock  ：每日股價資料
'*   udtIndex  ：儲存每日股價資料的指數資料
'*   intStockNo：資料筆數
'*   strStartDT：傳入的資料起始日期
'*   strEndDT  ：傳入的資料結束日期
'***********************************************************************************
Public Sub GetHighLowPoints2(ByRef udtStock() As StockData, _
                            ByRef udtIndex() As IndexData, _
                            ByVal intStockNo As Integer, _
                            ByVal strStartDT As String, _
                            ByVal strEndDT As String)
    Dim intPos As Integer       ' 目前計算的位置
    Dim intStartPos As Integer
    Dim intEndPos As Integer
    Dim intTmpStockNo As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim signalFlag As Single       ' Sector 內容
    Dim udtGRG() As udtGRGIndex
    Dim udtGroup() As Boolean
    Dim intGRGLine As Integer   ' 記錄目前是那筆資料為母數列
    Dim blnFlag As Boolean      ' 判斷此次計算是否結束
    Dim blnOK As Boolean
    Dim intMainPos As Integer   ' 母數列的位置
    Dim sngGRG() As Single      ' 儲在各數列的GRG value
    Dim intOutputFile As Integer
    Dim OutputFileName As String
    Dim strData As String

    On Error GoTo ERR_HANDLE
      
    OutputFileName = "ClusterPattern.csv"
    intOutputFile = FreeFile()
    Open OutputFileName For Output As #intOutputFile
    '--- 傳入起始日期來取得在資料(udtStock)中的位置 ---
    intGRGLine = ReturnIndexPos(udtStock, intStockNo, strStartDT)
    '--- 傳入結束日期來取得在資料(udtStock)中的位置 ---
    intTmpStockNo = ReturnIndexPos(udtStock, intStockNo, strEndDT)
    i = intGRGLine
    signalFlag = udtIndex(i).sngSector
    '--- assign start line position ---
    intPos = i
    intStartPos = i
    intEndPos = i
    
    '--------------------------------------------------------------------
    '* 計算在 strStartDT 及 strEndDT 間資料聚類
    '--------------------------------------------------------------------
'    While intPos < intTmpStockNo
    While intEndPos < intTmpStockNo
        '--- 尋找每個sector起迄點 ---
        intStartPos = intEndPos
        While (signalFlag = udtIndex(intEndPos).sngSector) And intEndPos < intStockNo
            intEndPos = intEndPos + 1
        Wend
        
        '--- 將該區間的值儲在至GRG variables，最後一筆存放母數列的資料 ---
        intMainPos = intEndPos - intStartPos + 1
        ReDim udtGRG(intMainPos)
        ReDim sngGRG(intMainPos)
        ReDim udtGroup(intMainPos)  ' save the group
 
        '--- assign 區間內各項指標至udtGRG ---
        k = intStartPos
        For j = 1 To intMainPos - 1
            udtGRG(j).sngDate = udtStock(k).sngDate
            udtGRG(j).sngParam1 = GetFactorValue(udtIndex, k, 1)
            udtGRG(j).sngParam2 = GetFactorValue(udtIndex, k, 2)
            udtGRG(j).sngParam3 = GetFactorValue(udtIndex, k, 3)
            udtGRG(j).sngParam4 = GetFactorValue(udtIndex, k, 4)
            udtGRG(j).sngParam5 = GetFactorValue(udtIndex, k, 5)
            udtGRG(j).sngParam6 = GetFactorValue(udtIndex, k, 6)
            udtGRG(j).sngParam7 = GetFactorValue(udtIndex, k, 7)
            udtGRG(j).sngParam8 = GetFactorValue(udtIndex, k, 8)
            k = k + 1
        Next
        '--- 92705 : 找出要做母數列的資料assing至最後一列 ---
        i = FindHighLowDate(udtStock, intStartPos, intEndPos, signalFlag)
        '--- Assign 母數列到最後一列 ---
        With udtGRG(intMainPos)
            .sngDate = udtStock(i).sngDate
            .sngParam1 = GetFactorValue(udtIndex, i, 1)
            .sngParam2 = GetFactorValue(udtIndex, i, 1)
            .sngParam3 = GetFactorValue(udtIndex, i, 1)
            .sngParam4 = GetFactorValue(udtIndex, i, 1)
            .sngParam5 = GetFactorValue(udtIndex, i, 1)
            .sngParam6 = GetFactorValue(udtIndex, i, 1)
            .sngParam7 = GetFactorValue(udtIndex, i, 1)
            .sngParam8 = GetFactorValue(udtIndex, i, 1)
        End With
        '--- Preprocessing ---
        Call PreProcess(udtGRG, intMainPos)
        '--- Calcuate GRG --- 930810 has bug
        Call CalGRG(udtGRG, udtGroup, intMainPos, signalFlag)
        '--- 記錄下這次聚類的結果 -- 920702 ---
        Dim strOutData As String
        strOutData = ""
        For j = 1 To intMainPos - 1
            If udtGroup(j) = True Then
                k = intStartPos + j - 1
                strOutData = udtStock(k).sngDate & vbTab
                strOutData = strOutData & udtStock(k).sngDate & vbTab
            End If
        Next
        Print #intOutputFile, strOutData
        '--- Assign回udtindex() ---
        i = 1
        For intPos = intStartPos To intEndPos - 1
            udtIndex(intPos).sngGRG = udtGRG(i).sngGRG
            i = i + 1
        Next
        signalFlag = udtIndex(intEndPos).sngSector
    Wend

    Close #intOutputFile
    
    Exit Sub
    
ERR_HANDLE:
    Select Case Err.Number
        Case 6
            MsgBox "[GCFR_General_Module.GetHighLowPoints2()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
            Err.Clear
            Resume Next
        Case 10001
            MsgBox "[GCFR_General_Module.GetHighLowPoints2()] -- " & Err.Number & ":沒有符合sector資料可供計算", vbOKOnly
        Case Else
            MsgBox "[GCFR_General_Module.GetHighLowPoints2()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
'            Resume Next
    End Select
End Sub


' 930410 : CY index use
' -- 921213 -- 目前主要使用之Method
'***********************************************************************************
'* 對資料做前處理
'* Input Param.
'*   udtGRG  ：要處理的資料
'*   intCount：要處理的資料的筆數
'***********************************************************************************
Public Sub PreProcess(ByRef udtGRG() As udtGRGIndex, ByVal intCount As Integer)
    Dim i As Integer

    On Error GoTo ERR_HANDLE
    
'''    '--- Initialize Variables ---
'''    For i = 1 To 8
'''        sngMax(i) = -99999
'''        sngMin(i) = 99999
'''    Next
'''
'''    '--------------------------------------------------------------------
'''    '* Find each Max value and Min value in each parameter
'''    '--------------------------------------------------------------------
'''    For i = 1 To intCount
'''        With udtGRG(i)
'''            If frmGeryExp.chkSelAttr(0).Value = "1" Then
'''                If sngMax(1) < .sngParam1 Then
'''                    sngMax(1) = .sngParam1
'''                End If
'''                If sngMin(1) > .sngParam1 Then
'''                    sngMin(1) = .sngParam1
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(1).Value = "1" Then
'''                If sngMax(2) < .sngParam2 Then
'''                    sngMax(2) = .sngParam2
'''                End If
'''                If sngMin(2) > .sngParam2 Then
'''                    sngMin(2) = .sngParam2
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(2).Value = "1" Then
'''                If sngMax(3) < .sngParam3 Then
'''                    sngMax(3) = .sngParam3
'''                End If
'''                If sngMin(3) > .sngParam3 Then
'''                    sngMin(3) = .sngParam3
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(3).Value = "1" Then
'''                If sngMax(4) < .sngParam4 Then
'''                    sngMax(4) = .sngParam4
'''                End If
'''                If sngMin(4) > .sngParam4 Then
'''                    sngMin(4) = .sngParam4
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(4).Value = "1" Then
'''                If sngMax(5) < .sngParam5 Then
'''                    sngMax(5) = .sngParam5
'''                End If
'''                If sngMin(5) > .sngParam5 Then
'''                    sngMin(5) = .sngParam5
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(5).Value = "1" Then
'''                If sngMax(6) < .sngParam6 Then
'''                    sngMax(6) = .sngParam6
'''                End If
'''                If sngMin(6) > .sngParam6 Then
'''                    sngMin(6) = .sngParam6
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(6).Value = "1" Then
'''                If sngMax(7) < .sngParam7 Then
'''                    sngMax(7) = .sngParam7
'''                End If
'''                If sngMin(7) > .sngParam7 Then
'''                    sngMin(7) = .sngParam7
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(7).Value = "1" Then
'''                If sngMax(8) < .sngParam8 Then
'''                    sngMax(8) = .sngParam8
'''                End If
'''                If sngMin(8) > .sngParam8 Then
'''                    sngMin(8) = .sngParam8
'''                End If
'''            End If
'''        End With
'''    Next
'''
'''    '--------------------------------------------------------------------
'''    '* 初值化
'''    '--------------------------------------------------------------------
'''    For i = 1 To intCount
'''        With udtGRG(i)
'''            '--------------------------------------------------------------------
'''            '* gintSelFactorsMethod = 1 表「望大」
'''            '* gintSelFactorsMethod = 2 表「望小」
'''            '* gintSelFactorsMethod = 3 表「望目」
'''            '--------------------------------------------------------------------
'''            If frmGeryExp.chkSelAttr(0).Value = "1" Then
'''                If gintSelFactorsMethod(1) = 1 Then
'''                    .sngParam1 = (.sngParam1 - sngMin(1)) / (sngMax(1) - sngMin(1))
'''                ElseIf gintSelFactorsMethod(1) = 2 Then
'''                    .sngParam1 = (sngMax(1) - .sngParam1) / (sngMax(1) - sngMin(1))
'''                Else
'''                    .sngParam1 = 1 - Abs(.sngParam1 - udtGRG(intCount).sngParam1) / (IIf((sngMax(1) - udtGRG(intCount).sngParam1 > udtGRG(intCount).sngParam1 - sngMin(1)), sngMax(1) - udtGRG(intCount).sngParam1, udtGRG(intCount).sngParam1 - sngMin(1)))
'''                End If
'''            End If
'''
'''            If frmGeryExp.chkSelAttr(1).Value = "1" Then
'''                If gintSelFactorsMethod(2) = 1 Then
'''                    .sngParam2 = (.sngParam2 - sngMin(2)) / (sngMax(2) - sngMin(2))
'''                ElseIf gintSelFactorsMethod(2) = 2 Then
'''                    .sngParam2 = (sngMax(2) - .sngParam2) / (sngMax(2) - sngMin(2))
'''                Else
'''                    .sngParam2 = 1 - Abs(.sngParam2 - udtGRG(intCount).sngParam2) / (IIf((sngMax(2) - udtGRG(intCount).sngParam2 > udtGRG(intCount).sngParam2 - sngMin(2)), sngMax(2) - udtGRG(intCount).sngParam2, udtGRG(intCount).sngParam2 - sngMin(2)))
'''                End If
'''            End If
'''
'''            If frmGeryExp.chkSelAttr(2).Value = "1" Then
'''                If gintSelFactorsMethod(3) = 1 Then
'''                    .sngParam3 = (.sngParam3 - sngMin(3)) / (sngMax(3) - sngMin(3))
'''                ElseIf gintSelFactorsMethod(3) = 2 Then
'''                    .sngParam3 = (sngMax(3) - .sngParam3) / (sngMax(3) - sngMin(3))
'''                Else
'''                    .sngParam3 = 1 - Abs(.sngParam3 - udtGRG(intCount).sngParam3) / (IIf((sngMax(3) - udtGRG(intCount).sngParam3 > udtGRG(intCount).sngParam3 - sngMin(3)), sngMax(3) - udtGRG(intCount).sngParam3, udtGRG(intCount).sngParam3 - sngMin(3)))
'''                End If
'''            End If
'''
'''            If frmGeryExp.chkSelAttr(3).Value = "1" Then
'''                If gintSelFactorsMethod(4) = 1 Then
'''                    .sngParam4 = (.sngParam4 - sngMin(4)) / (sngMax(4) - sngMin(4))
'''                ElseIf gintSelFactorsMethod(4) = 2 Then
'''                    .sngParam4 = (sngMax(4) - .sngParam4) / (sngMax(4) - sngMin(4))
'''                Else
'''                    .sngParam4 = 1 - Abs(.sngParam4 - udtGRG(intCount).sngParam4) / (IIf((sngMax(4) - udtGRG(intCount).sngParam4 > udtGRG(intCount).sngParam4 - sngMin(4)), sngMax(4) - udtGRG(intCount).sngParam4, udtGRG(intCount).sngParam4 - sngMin(4)))
'''                End If
'''            End If
'''
'''            If frmGeryExp.chkSelAttr(4).Value = "1" Then
'''                If gintSelFactorsMethod(5) = 1 Then
'''                    .sngParam5 = (.sngParam5 - sngMin(5)) / (sngMax(5) - sngMin(5))
'''                ElseIf gintSelFactorsMethod(5) = 2 Then
'''                    .sngParam5 = (sngMax(5) - .sngParam5) / (sngMax(5) - sngMin(5))
'''                Else
'''                    .sngParam5 = 1 - Abs(.sngParam5 - udtGRG(intCount).sngParam5) / (IIf((sngMax(5) - udtGRG(intCount).sngParam5 > udtGRG(intCount).sngParam5 - sngMin(5)), sngMax(5) - udtGRG(intCount).sngParam5, udtGRG(intCount).sngParam5 - sngMin(5)))
'''                End If
'''            End If
'''
'''            If frmGeryExp.chkSelAttr(5).Value = "1" Then
'''                If gintSelFactorsMethod(6) = 1 Then
'''                    .sngParam6 = (.sngParam6 - sngMin(6)) / (sngMax(6) - sngMin(6))
'''                ElseIf gintSelFactorsMethod(6) = 2 Then
'''                    .sngParam6 = (sngMax(6) - .sngParam6) / (sngMax(6) - sngMin(6))
'''                Else
'''                    .sngParam6 = 1 - Abs(.sngParam6 - udtGRG(intCount).sngParam6) / (IIf((sngMax(6) - udtGRG(intCount).sngParam6 > udtGRG(intCount).sngParam6 - sngMin(6)), sngMax(6) - udtGRG(intCount).sngParam6, udtGRG(intCount).sngParam6 - sngMin(6)))
'''                End If
'''            End If
'''
'''            If frmGeryExp.chkSelAttr(6).Value = "1" Then
'''                If gintSelFactorsMethod(7) = 1 Then
'''                    .sngParam7 = (.sngParam7 - sngMin(7)) / (sngMax(7) - sngMin(7))
'''                ElseIf gintSelFactorsMethod(7) = 2 Then
'''                    .sngParam7 = (sngMax(7) - .sngParam7) / (sngMax(7) - sngMin(7))
'''                Else
'''                    .sngParam7 = 1 - Abs(.sngParam7 - udtGRG(intCount).sngParam7) / (IIf((sngMax(7) - udtGRG(intCount).sngParam7 > udtGRG(intCount).sngParam7 - sngMin(7)), sngMax(7) - udtGRG(intCount).sngParam7, udtGRG(intCount).sngParam7 - sngMin(7)))
'''                End If
'''            End If
'''
'''            If frmGeryExp.chkSelAttr(7).Value = "1" Then
'''                If gintSelFactorsMethod(8) = 1 Then
'''                    .sngParam8 = (.sngParam8 - sngMin(8)) / (sngMax(8) - sngMin(8))
'''                ElseIf gintSelFactorsMethod(8) = 2 Then
'''                    .sngParam8 = (sngMax(8) - .sngParam8) / (sngMax(8) - sngMin(8))
'''                Else
'''                    .sngParam8 = 1 - Abs(.sngParam8 - udtGRG(intCount).sngParam8) / (IIf((sngMax(8) - udtGRG(intCount).sngParam8 > udtGRG(intCount).sngParam8 - sngMin(8)), sngMax(8) - udtGRG(intCount).sngParam8, udtGRG(intCount).sngParam8 - sngMin(8)))
'''                End If
'''            End If
'''        End With
'''    Next
    
    Exit Sub
    
    
ERR_HANDLE:
    Select Case Err.Number
        Case 6
            Debug.Print "[GCFR_General_Module.PreProcess()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
        Case Else
            MsgBox "[GCFR_General_Module.PreProcess()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
    
    Err.Clear
    Resume Next
End Sub



'***************************************************************************************************
'* 說    明:
'*    找出該Sector的最高(低)點
'* 輸入參數:
'*    udtStock 要處理的資料(存放每個Sector的資料)
'*    intStartPos 該Sector的起始位置
'*    intEndPos 該Sector的結束位置
'*    sngFlag 值=1表是找高點；值=0.5表示找低點
'* 輸出參數:
'*    Integer 回傳高(低)點的位置
'* 版    本:
'*    2.00: 20080913 Earvin   New
'***************************************************************************************************
Public Function FindHighLowDate(ByRef udtStock() As StockData, _
                            ByVal intStartPos As Integer, _
                            ByVal intEndPos As Integer, _
                            ByVal sngFlag As Single) As Integer

    Dim i As Integer
    Dim sngMaxValue As Single, sngMinValue As Single
    Dim intPos As Integer
       
    On Error GoTo ERR_HANDLE
    
    '*** Initialize Variables ***
    sngMaxValue = 0
    sngMinValue = 99999
    intPos = 0
    
    For i = intStartPos To intEndPos
        '--- sngFlag = 1 --> 找高點；sngFalg = 0.5 --> 找低點 ---
        If sngFlag = 1 Then ' 高點
            If udtStock(i).sngEndprice > sngMaxValue Then
                sngMaxValue = udtStock(i).sngEndprice
                intPos = i
            End If
        Else    ' 低點
            If udtStock(i).sngEndprice < sngMinValue Then
                sngMinValue = udtStock(i).sngEndprice
                intPos = i
            End If
        End If
    Next
    
    FindHighLowDate = intPos
    
    Exit Function
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            Debug.Print "[GCFR_General_Module.FindHighLowDate()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
    
    Err.Clear
    Resume Next
End Function


''''***********************************************************************************
''''* 找出該Sector的最高(低)點
''''* input param:
''''*   udtStock   ：要處理的資料(存放每個Sector的資料)
''''*   intStartPos：該Sector的起始位置
''''*   intEndPos  ：該Sector的結束位置
''''*   signalFlag    ：值=1表是找高點；值=0.5表示找低點
''''*   Return     ：回傳高(低)點的位置
''''***********************************************************************************
'''Public Function FindHighLowDate(ByRef udtStock() As Stockdata, _
'''                            ByVal intStartPos As Integer, _
'''                            ByVal intEndPos As Integer, _
'''                            ByVal signalFlag As Single) As Integer
'''
'''    Dim i As Integer
'''    Dim sngMaxValue As Single, sngMinValue As Single
'''    Dim intPos As Integer
'''
'''    On Error GoTo ERR_HANDLE
'''
'''    '--- Initialize Variables ---
'''    sngMaxValue = 0
'''    sngMinValue = 99999
'''    intPos = 0
'''
'''    For i = intStartPos To intEndPos
'''        '--- signalFlag = 1 --> 找高點；sngFalg = 0.5 --> 找低點 ---
'''        If signalFlag = 1 Then ' 高點
'''            If udtStock(i).sngEndprice > sngMaxValue Then
'''                sngMaxValue = udtStock(i).sngEndprice
'''                intPos = i
'''            End If
'''        Else    ' 低點
'''            If udtStock(i).sngEndprice < sngMinValue Then
'''                sngMinValue = udtStock(i).sngEndprice
'''                intPos = i
'''            End If
'''        End If
'''    Next
'''
'''    FindHighLowDate = intPos
'''
'''    Exit Function
'''
'''ERR_HANDLE:
'''    Select Case Err.Number
'''        Case Else
'''            Debug.Print "[Method: FindHighLowDate()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
'''    End Select
'''
'''    Err.Clear
'''    Resume Next
'''End Function


'***********************************************************************************
'* Get attributes value
'* Input Param:
'*   udtIndex    : 儲存的指數資料
'*   intPos      : 要讀取資料的位置
'*   intSelFactor: 選擇的attribute
'* Return        : 所選擇之資料之attribute的值
'***********************************************************************************
Public Function GetFactorValue(ByRef udtIndex() As IndexData, _
                                ByVal intPos As Integer, _
                                intSelFactor) As Single
    On Error GoTo ERR_HANDLE
    
    '--- 大於選擇的attribute數目，均回傳99999 ---
    If intSelFactor > gintSelCount Then
        GetFactorValue = 99999
        Exit Function
    End If
    
    If gstrSelFactors(intSelFactor) = "BIAS" Then
        GetFactorValue = udtIndex(intPos).sngBias
    ElseIf gstrSelFactors(intSelFactor) = "MACD" Then
        GetFactorValue = udtIndex(intPos).sngMACD
    ElseIf gstrSelFactors(intSelFactor) = "PSY" Then
        GetFactorValue = udtIndex(intPos).sngPSY
    ElseIf gstrSelFactors(intSelFactor) = "RSI_L" Then
        GetFactorValue = udtIndex(intPos).sngRSI_L
    ElseIf gstrSelFactors(intSelFactor) = "RSI_S" Then
        GetFactorValue = udtIndex(intPos).sngRSI_S
    ElseIf gstrSelFactors(intSelFactor) = "WMS" Then
        GetFactorValue = udtIndex(intPos).sngWMS
    ElseIf gstrSelFactors(intSelFactor) = "K" Then
        GetFactorValue = udtIndex(intPos).sngK
    ElseIf gstrSelFactors(intSelFactor) = "D" Then
        GetFactorValue = udtIndex(intPos).sngD
'    ElseIf gstrSelFactors(intSelFactor) = "KDDis" Then
'        GetFactorValue = udtIndex(intPos).sngKDDis
    ElseIf gstrSelFactors(intSelFactor) = "MAPDis" Then
        GetFactorValue = udtIndex(intPos).sngMAPDis
'    ElseIf gstrSelFactors(intSelFactor) = "RSIDis" Then
'        GetFactorValue = udtIndex(intPos).sngRSIDis
    ElseIf gstrSelFactors(intSelFactor) = "MAP3" Then
        GetFactorValue = udtIndex(intPos).sngP3
    ElseIf gstrSelFactors(intSelFactor) = "MAP4" Then
        GetFactorValue = udtIndex(intPos).sngP4
    ElseIf gstrSelFactors(intSelFactor) = "MAP5" Then
        GetFactorValue = udtIndex(intPos).sngP5
    ElseIf gstrSelFactors(intSelFactor) = "MAP6" Then
        GetFactorValue = udtIndex(intPos).sngP6
    ElseIf gstrSelFactors(intSelFactor) = "MAP8" Then
        GetFactorValue = udtIndex(intPos).sngP8
    ElseIf gstrSelFactors(intSelFactor) = "MAP10" Then
        GetFactorValue = udtIndex(intPos).sngP10
    ElseIf gstrSelFactors(intSelFactor) = "MAP12" Then
        GetFactorValue = udtIndex(intPos).sngP12
    ElseIf gstrSelFactors(intSelFactor) = "MAP24" Then
        GetFactorValue = udtIndex(intPos).sngP24
    ElseIf gstrSelFactors(intSelFactor) = "MAP30" Then
        GetFactorValue = udtIndex(intPos).sngP30
'    ElseIf gstrSelFactors(intSelFactor) = "MAP48" Then
'        GetFactorValue = udtIndex(intPos).sngP48
    ElseIf gstrSelFactors(intSelFactor) = "MAP72" Then
        GetFactorValue = udtIndex(intPos).sngP72
    ElseIf gstrSelFactors(intSelFactor) = "MAP144" Then
        GetFactorValue = udtIndex(intPos).sngP144
    ElseIf gstrSelFactors(intSelFactor) = "MAP288" Then
        GetFactorValue = udtIndex(intPos).sngP288
    ElseIf gstrSelFactors(intSelFactor) = "DIF" Then
        GetFactorValue = udtIndex(intPos).sngDIF
    ElseIf gstrSelFactors(intSelFactor) = "DIF_MACD" Then
        GetFactorValue = udtIndex(intPos).sngDIF_MACD
'    ElseIf gstrSelFactors(intSelFactor) = "TAPI" Then
'        GetFactorValue = udtIndex(intPos).sngTAPI
'    ElseIf gstrSelFactors(intSelFactor) = "OBV" Then
'        GetFactorValue = udtIndex(intPos).sngOBV
    ElseIf gstrSelFactors(intSelFactor) = "VR" Then
        GetFactorValue = udtIndex(intPos).sngVR
    ElseIf gstrSelFactors(intSelFactor) = "CY" Then
        GetFactorValue = Str(Round(udtIndex(intPos).sngDIF_MACD, 2))
'    ElseIf gstrSelFactors(intSelFactor) = "AR" Then
'        GetFactorValue = udtIndex(intPos).sngAR
'    ElseIf gstrSelFactors(intSelFactor) = "BR" Then
'        GetFactorValue = udtIndex(intPos).sngBR
    ElseIf gstrSelFactors(intSelFactor) = "VOL24" Then
        GetFactorValue = udtIndex(intPos).sngVol24
    Else
        Err.Raise 10002
    End If
    
    Exit Function
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.GetFactorValue()] -- " & Err.Number & ":" & Err.Description & " -- 找不到對應不到的指標", vbOKOnly
End Function


'***********************************************************************************
'* 計算第intCount列資料的GRG
'* input param:
'*   udtGRG  ：要處理的資料(存放每個Sector的資料)
'*   udtGroup：maybe no use ... 920907
'*   intCount：此次要計算(udtGRG)的筆數，第intCount存放的是母數列的資料
'*
'* 920728：此函式是對個別的factors一起做GRG、做聚類
'* 920907：以GRC做為判斷終止條件
'***********************************************************************************
'Public Sub CalGRG2(ByRef udtGRG() As udtGRGIndex, ByRef udtGroup() As Boolean, _
'                    ByVal intCount As Integer)
'    Dim sngMax(8) As Single
'    Dim sngMin(8) As Single
'    Dim blnFlag As Boolean          ' 此次Sector是否已找到相同聚類 (找到=false)
'    Dim blnChangeFlag As Boolean    ' 比較前後兩次聚類是否相同 (相同=true)
'    Dim udtGRGDiff() As udtGRGIndex
'    Dim sngGRGPrev() As Single
'    Dim i As Integer, j As Integer
'    Dim intLoop As Integer
'    Dim sngSumRo As Single
'
'    On Error GoTo ERR_HANDLE
'
'    blnFlag = True
'    ReDim udtGRGDiff(intCount)
'    ReDim sngGRGPrev(intCount)
'
'    While blnFlag
'        '--- initialize variables ---
'        For i = 1 To 8
'            sngMax(i) = 0
'            sngMin(i) = 1
'        Next
'
'        '--- 計算差序列 (intCount儲放母數列的資料) ---
'        For i = 1 To intCount - 1
'            udtGRGDiff(i).sngParam1 = Abs(udtGRG(intCount).sngParam1 - udtGRG(i).sngParam1)
'            udtGRGDiff(i).sngParam2 = Abs(udtGRG(intCount).sngParam2 - udtGRG(i).sngParam2)
'            udtGRGDiff(i).sngParam3 = Abs(udtGRG(intCount).sngParam3 - udtGRG(i).sngParam3)
'            udtGRGDiff(i).sngParam4 = Abs(udtGRG(intCount).sngParam4 - udtGRG(i).sngParam4)
'            udtGRGDiff(i).sngParam5 = Abs(udtGRG(intCount).sngParam5 - udtGRG(i).sngParam5)
'            udtGRGDiff(i).sngParam6 = Abs(udtGRG(intCount).sngParam6 - udtGRG(i).sngParam6)
'            udtGRGDiff(i).sngParam7 = Abs(udtGRG(intCount).sngParam7 - udtGRG(i).sngParam7)
'            udtGRGDiff(i).sngParam8 = Abs(udtGRG(intCount).sngParam8 - udtGRG(i).sngParam8)
'        Next
'        '--- 將母數列的值清成0 ---
'        With udtGRGDiff(i)
'            .sngGRG = 0
'            .sngParam1 = 0
'            .sngParam2 = 0
'            .sngParam3 = 0
'            .sngParam4 = 0
'            .sngParam5 = 0
'            .sngParam6 = 0
'            .sngParam7 = 0
'            .sngParam8 = 0
'        End With
'        '-------------------------------------------------------
'        '* 找出各個Attribue的差序列最大值及最小值
'        '* 920705 --> 若不足8個Attribute要做處理…
'        '-------------------------------------------------------
'        Dim sngMax As Single    ' 930810 has bug
'        Dim sngMin As Single    ' 930810 has bug
'        For i = 1 To intCount - 1
'            With udtGRGDiff(i)
'                If .sngParam1 > sngMax(1) Then
'                    sngMax = .sngParam1
'                End If
'                If .sngParam1 < sngMin(1) Then
'                    sngMin = .sngParam1
'                End If
'                If .sngParam2 > sngMax(2) Then
'                    sngMax = .sngParam2
'                End If
'                If .sngParam2 < sngMin(2) Then
'                    sngMin = .sngParam2
'                End If
'                If .sngParam3 > sngMax(3) Then
'                    sngMax = .sngParam3
'                End If
'                If .sngParam3 < sngMin(3) Then
'                    sngMin = .sngParam3
'                End If
'                If .sngParam4 > sngMax(4) Then
'                    sngMax = .sngParam4
'                End If
'                If .sngParam4 < sngMin(4) Then
'                    sngMin = .sngParam4
'                End If
'                If .sngParam5 > sngMax(5) Then
'                    sngMax = .sngParam5
'                End If
'                If .sngParam5 < sngMin(5) Then
'                    sngMin = .sngParam5
'                End If
'                If .sngParam6 > sngMax(6) Then
'                    sngMax = .sngParam6
'                End If
'                If .sngParam6 < sngMin(6) Then
'                    sngMin = .sngParam6
'                End If
'                If .sngParam7 > sngMax(7) Then
'                    sngMax = .sngParam7
'                End If
'                If .sngParam7 < sngMin(7) Then
'                    sngMin = .sngParam7
'                End If
'                If .sngParam8 > sngMax(8) Then
'                    sngMax = .sngParam8
'                End If
'                If .sngParam8 < sngMin(8) Then
'                    sngMin = .sngParam8
'                End If
'            End With
'        Next
'        '--- 計算GRC ---
'        For i = 1 To intCount - 1
'            '--------------------------------------------------------------------
'            '* 採用下列方式，gsngRo 會對本運算的影響
'            '* 故可藉由設定gsngRo 值來分散GRG的值
'            '* 920727; 應利用udtGroup來判斷處理幾個attributes -- not implement
'            '--------------------------------------------------------------------
'            With udtGRGDiff(i)
'                If gintSelCount = 1 Then
'                    If (.sngParam1 + gsngRo * sngMax(1)) = 0 Then
'                        .sngParam1 = 0
'                    Else
'                        .sngParam1 = (sngMin(1) + gsngRo * sngMax(1)) / (.sngParam1 + gsngRo * sngMax(1))
'                    End If
'                ElseIf gintSelCount = 2 Then
'                    If (.sngParam1 + gsngRo * sngMax(1)) = 0 Then
'                        .sngParam1 = 0
'                    Else
'                        .sngParam1 = (sngMin(1) + gsngRo * sngMax(1)) / (.sngParam1 + gsngRo * sngMax(1))
'                    End If
'                    If (.sngParam2 + gsngRo * sngMax(2)) = 0 Then
'                        .sngParam2 = 0
'                    Else
'                        .sngParam2 = (sngMin(2) + gsngRo * sngMax(2)) / (.sngParam2 + gsngRo * sngMax(2))
'                    End If
'                ElseIf gintSelCount = 3 Then
'                    If (.sngParam1 + gsngRo * sngMax(1)) = 0 Then
'                        .sngParam1 = 0
'                    Else
'                        .sngParam1 = (sngMin(1) + gsngRo * sngMax(1)) / (.sngParam1 + gsngRo * sngMax(1))
'                    End If
'                    If (.sngParam2 + gsngRo * sngMax(2)) = 0 Then
'                        .sngParam2 = 0
'                    Else
'                        .sngParam2 = (sngMin(2) + gsngRo * sngMax(2)) / (.sngParam2 + gsngRo * sngMax(2))
'                    End If
'                    If (.sngParam3 + gsngRo * sngMax(3)) = 0 Then
'                        .sngParam3 = 0
'                    Else
'                        .sngParam3 = (sngMin(3) + gsngRo * sngMax(3)) / (.sngParam3 + gsngRo * sngMax(3))
'                    End If
'                ElseIf gintSelCount = 4 Then
'                    If (.sngParam1 + gsngRo * sngMax(1)) = 0 Then
'                        .sngParam1 = 0
'                    Else
'                        .sngParam1 = (sngMin(1) + gsngRo * sngMax(1)) / (.sngParam1 + gsngRo * sngMax(1))
'                    End If
'                    If (.sngParam2 + gsngRo * sngMax(2)) = 0 Then
'                        .sngParam2 = 0
'                    Else
'                        .sngParam2 = (sngMin(2) + gsngRo * sngMax(2)) / (.sngParam2 + gsngRo * sngMax(2))
'                    End If
'                    If (.sngParam3 + gsngRo * sngMax(3)) = 0 Then
'                        .sngParam3 = 0
'                    Else
'                        .sngParam3 = (sngMin(3) + gsngRo * sngMax(3)) / (.sngParam3 + gsngRo * sngMax(3))
'                    End If
'                    If (.sngParam4 + gsngRo * sngMax(4)) = 0 Then
'                        .sngParam4 = 0
'                    Else
'                        .sngParam4 = (sngMin(4) + gsngRo * sngMax(4)) / (.sngParam4 + gsngRo * sngMax(4))
'                    End If
'                ElseIf gintSelCount = 5 Then
'                    If (.sngParam1 + gsngRo * sngMax(1)) = 0 Then
'                        .sngParam1 = 0
'                    Else
'                        .sngParam1 = (sngMin(1) + gsngRo * sngMax(1)) / (.sngParam1 + gsngRo * sngMax(1))
'                    End If
'                    If (.sngParam2 + gsngRo * sngMax(2)) = 0 Then
'                        .sngParam2 = 0
'                    Else
'                        .sngParam2 = (sngMin(2) + gsngRo * sngMax(2)) / (.sngParam2 + gsngRo * sngMax(2))
'                    End If
'                    If (.sngParam3 + gsngRo * sngMax(3)) = 0 Then
'                        .sngParam3 = 0
'                    Else
'                        .sngParam3 = (sngMin(3) + gsngRo * sngMax(3)) / (.sngParam3 + gsngRo * sngMax(3))
'                    End If
'                    If (.sngParam4 + gsngRo * sngMax(4)) = 0 Then
'                        .sngParam4 = 0
'                    Else
'                        .sngParam4 = (sngMin(4) + gsngRo * sngMax(4)) / (.sngParam4 + gsngRo * sngMax(4))
'                    End If
'                    If (.sngParam5 + gsngRo * sngMax(5)) = 0 Then
'                        .sngParam5 = 0
'                    Else
'                        .sngParam5 = (sngMin(5) + gsngRo * sngMax(5)) / (.sngParam5 + gsngRo * sngMax(5))
'                    End If
'                ElseIf gintSelCount = 6 Then
'                    If (.sngParam1 + gsngRo * sngMax(1)) = 0 Then
'                        .sngParam1 = 0
'                    Else
'                        .sngParam1 = (sngMin(1) + gsngRo * sngMax(1)) / (.sngParam1 + gsngRo * sngMax(1))
'                    End If
'                    If (.sngParam2 + gsngRo * sngMax(2)) = 0 Then
'                        .sngParam2 = 0
'                    Else
'                        .sngParam2 = (sngMin(2) + gsngRo * sngMax(2)) / (.sngParam2 + gsngRo * sngMax(2))
'                    End If
'                    If (.sngParam3 + gsngRo * sngMax(3)) = 0 Then
'                        .sngParam3 = 0
'                    Else
'                        .sngParam3 = (sngMin(3) + gsngRo * sngMax(3)) / (.sngParam3 + gsngRo * sngMax(3))
'                    End If
'                    If (.sngParam4 + gsngRo * sngMax(4)) = 0 Then
'                        .sngParam4 = 0
'                    Else
'                        .sngParam4 = (sngMin(4) + gsngRo * sngMax(4)) / (.sngParam4 + gsngRo * sngMax(4))
'                    End If
'                    If (.sngParam5 + gsngRo * sngMax(5)) = 0 Then
'                        .sngParam5 = 0
'                    Else
'                        .sngParam5 = (sngMin(5) + gsngRo * sngMax(5)) / (.sngParam5 + gsngRo * sngMax(5))
'                    End If
'                    If (.sngParam6 + gsngRo * sngMax(6)) = 0 Then
'                        .sngParam6 = 0
'                    Else
'                        .sngParam6 = (sngMin(6) + gsngRo * sngMax(6)) / (.sngParam6 + gsngRo * sngMax(6))
'                    End If
'                ElseIf gintSelCount = 7 Then
'                    If (.sngParam1 + gsngRo * sngMax(1)) = 0 Then
'                        .sngParam1 = 0
'                    Else
'                        .sngParam1 = (sngMin(1) + gsngRo * sngMax(1)) / (.sngParam1 + gsngRo * sngMax(1))
'                    End If
'                    If (.sngParam2 + gsngRo * sngMax(2)) = 0 Then
'                        .sngParam2 = 0
'                    Else
'                        .sngParam2 = (sngMin(2) + gsngRo * sngMax(2)) / (.sngParam2 + gsngRo * sngMax(2))
'                    End If
'                    If (.sngParam3 + gsngRo * sngMax(3)) = 0 Then
'                        .sngParam3 = 0
'                    Else
'                        .sngParam3 = (sngMin(3) + gsngRo * sngMax(3)) / (.sngParam3 + gsngRo * sngMax(3))
'                    End If
'                    If (.sngParam4 + gsngRo * sngMax(4)) = 0 Then
'                        .sngParam4 = 0
'                    Else
'                        .sngParam4 = (sngMin(4) + gsngRo * sngMax(4)) / (.sngParam4 + gsngRo * sngMax(4))
'                    End If
'                    If (.sngParam5 + gsngRo * sngMax(5)) = 0 Then
'                        .sngParam5 = 0
'                    Else
'                        .sngParam5 = (sngMin(5) + gsngRo * sngMax(5)) / (.sngParam5 + gsngRo * sngMax(5))
'                    End If
'                    If (.sngParam6 + gsngRo * sngMax(6)) = 0 Then
'                        .sngParam6 = 0
'                    Else
'                        .sngParam6 = (sngMin(6) + gsngRo * sngMax(6)) / (.sngParam6 + gsngRo * sngMax(6))
'                    End If
'                    If (.sngParam7 + gsngRo * sngMax(7)) = 0 Then
'                        .sngParam7 = 0
'                    Else
'                        .sngParam7 = (sngMin(7) + gsngRo * sngMax(7)) / (.sngParam7 + gsngRo * sngMax(7))
'                    End If
'                Else
'                    If (.sngParam1 + gsngRo * sngMax(1)) = 0 Then
'                        .sngParam1 = 0
'                    Else
'                        .sngParam1 = (sngMin(1) + gsngRo * sngMax(1)) / (.sngParam1 + gsngRo * sngMax(1))
'                    End If
'                    If (.sngParam2 + gsngRo * sngMax(2)) = 0 Then
'                        .sngParam2 = 0
'                    Else
'                        .sngParam2 = (sngMin(2) + gsngRo * sngMax(2)) / (.sngParam2 + gsngRo * sngMax(2))
'                    End If
'                    If (.sngParam3 + gsngRo * sngMax(3)) = 0 Then
'                        .sngParam3 = 0
'                    Else
'                        .sngParam3 = (sngMin(3) + gsngRo * sngMax(3)) / (.sngParam3 + gsngRo * sngMax(3))
'                    End If
'                    If (.sngParam4 + gsngRo * sngMax(4)) = 0 Then
'                        .sngParam4 = 0
'                    Else
'                        .sngParam4 = (sngMin(4) + gsngRo * sngMax(4)) / (.sngParam4 + gsngRo * sngMax(4))
'                    End If
'                    If (.sngParam5 + gsngRo * sngMax(5)) = 0 Then
'                        .sngParam5 = 0
'                    Else
'                        .sngParam5 = (sngMin(5) + gsngRo * sngMax(5)) / (.sngParam5 + gsngRo * sngMax(5))
'                    End If
'                    If (.sngParam6 + gsngRo * sngMax(6)) = 0 Then
'                        .sngParam6 = 0
'                    Else
'                        .sngParam6 = (sngMin(6) + gsngRo * sngMax(6)) / (.sngParam6 + gsngRo * sngMax(6))
'                    End If
'                    If (.sngParam7 + gsngRo * sngMax(7)) = 0 Then
'                        .sngParam7 = 0
'                    Else
'                        .sngParam7 = (sngMin(7) + gsngRo * sngMax(7)) / (.sngParam7 + gsngRo * sngMax(7))
'                    End If
'                    If (.sngParam8 + gsngRo * sngMax(8)) = 0 Then
'                        .sngParam8 = 0
'                    Else
'                        .sngParam8 = (sngMin(8) + gsngRo * sngMax(8)) / (.sngParam8 + gsngRo * sngMax(8))
'                    End If
'                End If
'            End With
'        Next
'        '--------------------------------------------------------------------
'        '* 計算GRG
'        '* 920705 --> 根據attributes處理
'        '--------------------------------------------------------------------
'        For i = 1 To intCount - 1
'            With udtGRGDiff(i)
'                '--- 採用均權的方式 ---
''                .sngGRG = (.sngParam1 + .sngParam2 + .sngParam3 + .sngParam4 + .sngParam5 + .sngParam6 + .sngParam7 + .sngParam8) / 8
'                '--------------------------------------------------------------------
'                '* 根據參數個數決定執行那個stmt
'                '--------------------------------------------------------------------
'                If gintSelCount = 1 Then
'                    .sngGRG = .sngParam1
'                ElseIf gintSelCount = 2 Then
'                    .sngGRG = (.sngParam1 + .sngParam2) / 2
'                ElseIf gintSelCount = 3 Then
'                    .sngGRG = (.sngParam1 + .sngParam2 + .sngParam3) / 3
'                ElseIf gintSelCount = 4 Then
'                    '--- 920706 : 採用均權方式，不適用，先改用加權方式 ---
'                    .sngGRG = (.sngParam1 + .sngParam2 + .sngParam3 + .sngParam4) / 4
''                    .sngGRG = .sngParam1 * 3 / 10 + .sngParam2 * 3 / 10 + .sngParam3 * 2 / 10 + .sngParam4 * 2 / 10
'                ElseIf gintSelCount = 5 Then
'                    '--- 920706 : 採用均權方式，不適用，先改用加權方式 ---
'                    .sngGRG = (.sngParam1 + .sngParam2 + .sngParam3 + .sngParam4 + .sngParam5) / 5
'                ElseIf gintSelCount = 6 Then
'                    '--- 920706 : 採用均權方式，不適用，先改用加權方式 ---
'                    .sngGRG = (.sngParam1 + .sngParam2 + .sngParam3 + .sngParam4 + .sngParam5 + .sngParam6) / 6
''                    .sngGRG = (.sngParam1 * 3 / 20 + .sngParam2 * 1 / 4 + .sngParam3 * 1 / 4 + .sngParam4 * 1 / 4 + .sngParam5 * 1 / 20 + .sngParam6 * 1 / 20)
'                ElseIf gintSelCount = 7 Then
'                    .sngGRG = (.sngParam1 + .sngParam2 + .sngParam3 + .sngParam4 + .sngParam5 + .sngParam6 + .sngParam7) / 7
'                Else: gintSelCount = 8
'                    .sngGRG = (.sngParam1 + .sngParam2 + .sngParam3 + .sngParam4 + .sngParam5 + .sngParam6 + .sngParam7 + .sngParam8) / 8
''                    .sngGRG = (.sngParam1 * 3 / 21 + .sngParam2 * 3 / 21 + .sngParam3 * 2 / 21 + .sngParam4 * 2 / 21 * 3 / 21 + .sngParam5 * 2 / 21 + .sngParam6 * 3 / 21 + .sngParam7 * 3 / 21 + .sngParam8 * 3 / 21)
'                End If
'            End With
'        Next
'
'        '--------------------------------------------------------------------
'        '* 聚類分析
'        '--------------------------------------------------------------------
'        sngSumRo = 0            ' 記錄大於閥值的筆數
'        blnChangeFlag = True    ' 若記錄之參考數列不等於此次的參考數列…
'        For i = 1 To intCount - 1
'            With udtGRGDiff(i)
'                If .sngGRG >= gsngFa Then
'                    ' intCount是指向母數列位置
''                    udtGRGDiff(intCount).sngParam1 = udtGRGDiff(intCount).sngParam1 + .sngParam1 * .sngGRG
''                    udtGRGDiff(intCount).sngParam2 = udtGRGDiff(intCount).sngParam2 + .sngParam2 * .sngGRG
''                    udtGRGDiff(intCount).sngParam3 = udtGRGDiff(intCount).sngParam3 + .sngParam3 * .sngGRG
''                    udtGRGDiff(intCount).sngParam4 = udtGRGDiff(intCount).sngParam4 + .sngParam4 * .sngGRG
''                    udtGRGDiff(intCount).sngParam5 = udtGRGDiff(intCount).sngParam5 + .sngParam5 * .sngGRG
''                    udtGRGDiff(intCount).sngParam6 = udtGRGDiff(intCount).sngParam6 + .sngParam6 * .sngGRG
''                    udtGRGDiff(intCount).sngParam7 = udtGRGDiff(intCount).sngParam7 + .sngParam7 * .sngGRG
''                    udtGRGDiff(intCount).sngParam8 = udtGRGDiff(intCount).sngParam8 + .sngParam8 * .sngGRG
'                    '--------------------------------------------------------------------
'                    '* 根據所選的Attributes數目來summarize GRG
'                    '--------------------------------------------------------------------
'                    If gintSelCount = 1 Then
'                        udtGRGDiff(intCount).sngParam1 = udtGRGDiff(intCount).sngParam1 + .sngParam1 * .sngGRG
'                    ElseIf gintSelCount = 2 Then
'                        udtGRGDiff(intCount).sngParam1 = udtGRGDiff(intCount).sngParam1 + .sngParam1 * .sngGRG
'                        udtGRGDiff(intCount).sngParam2 = udtGRGDiff(intCount).sngParam2 + .sngParam2 * .sngGRG
'                    ElseIf gintSelCount = 3 Then
'                        udtGRGDiff(intCount).sngParam1 = udtGRGDiff(intCount).sngParam1 + .sngParam1 * .sngGRG
'                        udtGRGDiff(intCount).sngParam2 = udtGRGDiff(intCount).sngParam2 + .sngParam2 * .sngGRG
'                        udtGRGDiff(intCount).sngParam3 = udtGRGDiff(intCount).sngParam3 + .sngParam3 * .sngGRG
'                    ElseIf gintSelCount = 4 Then
'                        udtGRGDiff(intCount).sngParam1 = udtGRGDiff(intCount).sngParam1 + .sngParam1 * .sngGRG
'                        udtGRGDiff(intCount).sngParam2 = udtGRGDiff(intCount).sngParam2 + .sngParam2 * .sngGRG
'                        udtGRGDiff(intCount).sngParam3 = udtGRGDiff(intCount).sngParam3 + .sngParam3 * .sngGRG
'                        udtGRGDiff(intCount).sngParam4 = udtGRGDiff(intCount).sngParam4 + .sngParam4 * .sngGRG
'                    ElseIf gintSelCount = 5 Then
'                        udtGRGDiff(intCount).sngParam1 = udtGRGDiff(intCount).sngParam1 + .sngParam1 * .sngGRG
'                        udtGRGDiff(intCount).sngParam2 = udtGRGDiff(intCount).sngParam2 + .sngParam2 * .sngGRG
'                        udtGRGDiff(intCount).sngParam3 = udtGRGDiff(intCount).sngParam3 + .sngParam3 * .sngGRG
'                        udtGRGDiff(intCount).sngParam4 = udtGRGDiff(intCount).sngParam4 + .sngParam4 * .sngGRG
'                        udtGRGDiff(intCount).sngParam5 = udtGRGDiff(intCount).sngParam5 + .sngParam5 * .sngGRG
'                    ElseIf gintSelCount = 6 Then
'                        udtGRGDiff(intCount).sngParam1 = udtGRGDiff(intCount).sngParam1 + .sngParam1 * .sngGRG
'                        udtGRGDiff(intCount).sngParam2 = udtGRGDiff(intCount).sngParam2 + .sngParam2 * .sngGRG
'                        udtGRGDiff(intCount).sngParam3 = udtGRGDiff(intCount).sngParam3 + .sngParam3 * .sngGRG
'                        udtGRGDiff(intCount).sngParam4 = udtGRGDiff(intCount).sngParam4 + .sngParam4 * .sngGRG
'                        udtGRGDiff(intCount).sngParam5 = udtGRGDiff(intCount).sngParam5 + .sngParam5 * .sngGRG
'                        udtGRGDiff(intCount).sngParam6 = udtGRGDiff(intCount).sngParam6 + .sngParam6 * .sngGRG
'                    ElseIf gintSelCount = 7 Then
'                        udtGRGDiff(intCount).sngParam1 = udtGRGDiff(intCount).sngParam1 + .sngParam1 * .sngGRG
'                        udtGRGDiff(intCount).sngParam2 = udtGRGDiff(intCount).sngParam2 + .sngParam2 * .sngGRG
'                        udtGRGDiff(intCount).sngParam3 = udtGRGDiff(intCount).sngParam3 + .sngParam3 * .sngGRG
'                        udtGRGDiff(intCount).sngParam4 = udtGRGDiff(intCount).sngParam4 + .sngParam4 * .sngGRG
'                        udtGRGDiff(intCount).sngParam5 = udtGRGDiff(intCount).sngParam5 + .sngParam5 * .sngGRG
'                        udtGRGDiff(intCount).sngParam6 = udtGRGDiff(intCount).sngParam6 + .sngParam6 * .sngGRG
'                        udtGRGDiff(intCount).sngParam7 = udtGRGDiff(intCount).sngParam7 + .sngParam7 * .sngGRG
'                    ElseIf gintSelCount = 8 Then
'                        udtGRGDiff(intCount).sngParam1 = udtGRGDiff(intCount).sngParam1 + .sngParam1 * .sngGRG
'                        udtGRGDiff(intCount).sngParam2 = udtGRGDiff(intCount).sngParam2 + .sngParam2 * .sngGRG
'                        udtGRGDiff(intCount).sngParam3 = udtGRGDiff(intCount).sngParam3 + .sngParam3 * .sngGRG
'                        udtGRGDiff(intCount).sngParam4 = udtGRGDiff(intCount).sngParam4 + .sngParam4 * .sngGRG
'                        udtGRGDiff(intCount).sngParam5 = udtGRGDiff(intCount).sngParam5 + .sngParam5 * .sngGRG
'                        udtGRGDiff(intCount).sngParam6 = udtGRGDiff(intCount).sngParam6 + .sngParam6 * .sngGRG
'                        udtGRGDiff(intCount).sngParam7 = udtGRGDiff(intCount).sngParam7 + .sngParam7 * .sngGRG
'                        udtGRGDiff(intCount).sngParam8 = udtGRGDiff(intCount).sngParam8 + .sngParam8 * .sngGRG
'                    End If
'                    sngSumRo = sngSumRo + 1
'                    '--------------------------------------------------------------------
'                    '* GRG > gsnFa 則udtGroup(i) = true --> ok
'                    '* else modified udtGroup(i) = true
'                    '--------------------------------------------------------------------
'                    If Not udtGroup(i) Then
'                        udtGroup(i) = True
'                        blnChangeFlag = False
'                    End If
'                Else
'                    If udtGroup(i) Then
'                        udtGroup(i) = False
'                        blnChangeFlag = False
'                    End If
'                End If
'            End With
'        Next
'
'        '--- 若blnChangeFlag = True則表示二次計算所聚類的資料皆相等，則可結束此次計算 ---
'        If blnChangeFlag Then
'            blnFlag = False
'        End If
'
'        '--- For Debug : 顯示的聚類的位置 ---
'        Debug.Print "---------------------------------------"
'        For i = 1 To intCount - 1
'            If udtGroup(i) = True Then
'                Debug.Print "i=" & i
'            End If
'            With udtGRG(i)
'                .sngGRG = udtGRGDiff(i).sngGRG
'            End With
'        Next
'        Debug.Print "---------------------------------------"
'
'        '--- 產生母數列資料 ---
'        With udtGRGDiff(intCount)
'            If sngSumRo = 0 Then
''                MsgBox "沒有符合的資料"
'                Err.Raise 10001, "CalGRG", "沒有符合的資料"
'            End If
'            .sngParam1 = .sngParam1 / sngSumRo
'            .sngParam2 = .sngParam2 / sngSumRo
'            .sngParam3 = .sngParam3 / sngSumRo
'            .sngParam4 = .sngParam4 / sngSumRo
'            .sngParam5 = .sngParam5 / sngSumRo
'            .sngParam6 = .sngParam6 / sngSumRo
'            .sngParam7 = .sngParam7 / sngSumRo
'            .sngParam8 = .sngParam8 / sngSumRo
'        End With
'
'        '--- 記錄下重心 ---
'        With udtGRG(intCount)
'            .sngParam1 = udtGRGDiff(intCount).sngParam1
'            .sngParam2 = udtGRGDiff(intCount).sngParam2
'            .sngParam3 = udtGRGDiff(intCount).sngParam3
'            .sngParam4 = udtGRGDiff(intCount).sngParam4
'            .sngParam5 = udtGRGDiff(intCount).sngParam5
'            .sngParam6 = udtGRGDiff(intCount).sngParam6
'            .sngParam7 = udtGRGDiff(intCount).sngParam7
'            .sngParam8 = udtGRGDiff(intCount).sngParam8
''            Debug.Print "the avergae heart parm1= " & .sngParam1 & ",parm2=" & .sngParam2 & ",parm3=" & .sngParam3 & ",parm4=" & .sngParam4 & ",parm5=" & .sngParam5 & ",parm6=" & .sngParam6 & ",parm7=" & .sngParam7 & ",parm8=" & .sngParam8
'        End With
'
''        ' For Debug : 顯示此次運算的次數
''        '             超過20次，先停下來看為什麼
''        intLoop = intLoop + 1
''        If intLoop > 100 Then
''            Debug.Print "intloop=" & intLoop
''            blnFlag = False
''        End If
'    Wend
'
'    Exit Sub
'
'
'ERR_HANDLE:
'    Select Case Err.Number
'        Case 10001
'            Debug.Print "[Method: CalGRG2()] -- Error Cluster"
''            Resume Next
'        Case Else
'            MsgBox "[Method: CalGRG2()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
'    End Select
'
'End Sub


'***********************************************************************************
'* 計算第intCount列資料的GRG
'* input param:
'*   udtStock      ：每日股價資料
'*   udtIndex      ：儲存每日股價資料的指數資料
'*   intStockNo    ：每日股價資料筆數
'*   gudtClusterDat：聚類結果dataset
'*   intClusterCnt ：聚類結果筆數
'***********************************************************************************
Public Sub CalculateHighLowPoints(ByRef udtStock() As StockData, _
                                ByRef udtIndex() As IndexData, _
                                ByVal intStockNo As Integer, _
                                ByRef gudtClusterDat() As udtGRGIndex, _
                                ByVal intClusterCnt As Integer)
    Dim i As Integer
    
    On Error GoTo ERR_HANDLE
    
    For i = 1 To intStockNo
        '--- 以90年以後的資料為traingng set ---
        If udtStock(i).sngDate >= 900109 Then
'            udtIndex(i).sngGRG = CalGRGValue2(udtIndex, gudtClusterDat, i, 1)     ' 算高點
            udtIndex(i).sngGRG = CalGRGValue2(udtIndex, gudtClusterDat, i, 2)      ' 算低點
            '--- 評量 -- 920924 ---
            If udtIndex(i).sngGRG > 0 Then
                '--- 評量該點是否正確 -- 920924 ---
'                Call EvaluateTheResult(udtStock, udtIndex, i, intStockNo, 1)     ' 評量高點
                Call EvaluateTheResult(udtStock, udtIndex, i, intStockNo, 2)      ' 評量低點
            End If
        End If
    Next
    
    Exit Sub
    
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.CalculateHighLowPoints()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
End Sub


'***********************************************************************************
'* 評估Method：計算第intCount列資料的GRG是否為高點或低點，若是則顯示出來
'* input param:
'*   udtStock       ：每日股價資料
'*   udtIndex       ：儲存每日股價資料的指數資料
'*   intStockNo     ：每日股價資料筆數
'*   gudtClusterDat ：聚類結果dataset
'*   intClusterCnt  ：聚類結果筆數
'*   intSelAttrCount：各聚類結果所選用來聚類的Attributes的個數
'***********************************************************************************
Public Sub CalculateHighLowPoints2(ByRef udtStock() As StockData, _
                                ByRef udtIndex() As IndexData, _
                                ByVal intStockNo As Integer, _
                                ByRef gudtClusterDat() As udtGRGIndex, _
                                ByVal intClusterCnt As Integer, _
                                ByRef intSelAttrCount() As Integer)
    Dim i As Integer
    
    On Error GoTo ERR_HANDLE
    
    For i = 1 To intStockNo
        '--- 以90年以後的資料為Traingng Set ---
        If udtStock(i).sngDate >= 900101 Then
            '--- 計算高點 ---
            If frmGeryExp.chkCalGRG.Value = 1 Then
                udtIndex(i).sngGRG = CalGRGValue2(udtIndex, gudtClusterDat, i, 1)
            End If
            '--- 計算低點 ---
            If frmGeryExp.chkCalGRG2.Value = 1 Then
                udtIndex(i).sngGRG = CalGRGValue2(udtIndex, gudtClusterDat, i, 2)
            End If
            
            '-----------------------------------
            '--- 評量該點是否正確 -- 920924 ---
            '-----------------------------------
            If udtIndex(i).sngGRG > 0 Then
                '--- 評量高點 ---
                If frmGeryExp.chkCalGRG.Value = 1 Then
                    Call EvaluateTheResult(udtStock, udtIndex, i, intStockNo, 1)
                End If
                '--- 評量低點 ---
                If frmGeryExp.chkCalGRG2.Value = 1 Then
                    Call EvaluateTheResult(udtStock, udtIndex, i, intStockNo, 2)
                End If
            End If
        End If
    Next
    
    Exit Sub
    
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.CalculateHighLowPoints2()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
End Sub


'***********************************************************************************
'* 還原聚類出來的值
'* Input Param :
'*   udtClusterDat：聚類結果的重心
'*   intPos       ：該資料的位置
'***********************************************************************************
Public Sub RestoreData(ByRef udtClusterDat() As udtGRGIndex, _
                        ByVal intPos As Integer)
    On Error GoTo ERR_HANDLE
    
'''    With udtClusterDat(intPos)
'''        If frmGeryExp.chkSelAttr(0).Value = "1" Then
'''            .sngParam1 = (sngMax(1) - sngMin(1)) * .sngParam1 + sngMin(1)
'''        End If
'''        If frmGeryExp.chkSelAttr(1).Value = "1" Then
'''            .sngParam2 = (sngMax(2) - sngMin(2)) * .sngParam2 + sngMin(2)
'''        End If
'''        If frmGeryExp.chkSelAttr(2).Value = "1" Then
'''            .sngParam3 = (sngMax(3) - sngMin(3)) * .sngParam3 + sngMin(3)
'''        End If
'''        If frmGeryExp.chkSelAttr(3).Value = "1" Then
'''            .sngParam4 = (sngMax(4) - sngMin(4)) * .sngParam4 + sngMin(4)
'''        End If
'''        If frmGeryExp.chkSelAttr(4).Value = "1" Then
'''            .sngParam5 = (sngMax(5) - sngMin(5)) * .sngParam5 + sngMin(5)
'''        End If
'''        If frmGeryExp.chkSelAttr(5).Value = "1" Then
'''            .sngParam6 = (sngMax(6) - sngMin(6)) * .sngParam6 + sngMin(6)
'''        End If
'''        If frmGeryExp.chkSelAttr(6).Value = "1" Then
'''            .sngParam7 = (sngMax(7) - sngMin(7)) * .sngParam7 + sngMin(7)
'''        End If
'''        If frmGeryExp.chkSelAttr(7).Value = "1" Then
'''            .sngParam8 = (sngMax(8) - sngMin(8)) * .sngParam8 + sngMin(8)
'''        End If
'''    End With

    Exit Sub
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.RestoreData()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
    
    Resume Next
End Sub





'***********************************************************************************
'* 對資料做前處理
'* Input Param.
'*   udtGRG  ：要處理的資料
'*   intCount：要處理的資料的筆數
'***********************************************************************************
Public Sub PreProcess2(ByRef udtGRG() As udtGRGIndex, ByVal intCount As Integer)
    Dim i As Integer
    Dim sngMax As Single, sngMin As Single

    On Error GoTo ERR_HANDLE
   ' 20080228 marked for complier...
'''    '--- Initialize Variables ---
'''    sngMax = -99999
'''    sngMin = 99999
'''
'''    '--------------------------------------------------------------------
'''    '* Find each Max value and Min value in each parameter
'''    '--------------------------------------------------------------------
'''    For i = 1 To intCount
'''        With udtGRG(i)
'''            If frmGeryExp.chkSelAttr(0).Value = "1" Then
'''                If sngMax < .sngParam1 Then
'''                    sngMax = .sngParam1
'''                End If
'''                If sngMin > .sngParam1 Then
'''                    sngMin = .sngParam1
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(1).Value = "1" Then
'''                If sngMax < .sngParam2 Then
'''                    sngMax = .sngParam2
'''                End If
'''                If sngMin > .sngParam2 Then
'''                    sngMin = .sngParam2
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(2).Value = "1" Then
'''                If sngMax < .sngParam3 Then
'''                    sngMax = .sngParam3
'''                End If
'''                If sngMin > .sngParam3 Then
'''                    sngMin = .sngParam3
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(3).Value = "1" Then
'''                If sngMax < .sngParam4 Then
'''                    sngMax = .sngParam4
'''                End If
'''                If sngMin > .sngParam4 Then
'''                    sngMin = .sngParam4
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(4).Value = "1" Then
'''                If sngMax < .sngParam5 Then
'''                    sngMax = .sngParam5
'''                End If
'''                If sngMin > .sngParam5 Then
'''                    sngMin = .sngParam5
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(5).Value = "1" Then
'''                If sngMax < .sngParam6 Then
'''                    sngMax = .sngParam6
'''                End If
'''                If sngMin > .sngParam6 Then
'''                    sngMin = .sngParam6
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(6).Value = "1" Then
'''                If sngMax < .sngParam7 Then
'''                    sngMax = .sngParam7
'''                End If
'''                If sngMin > .sngParam7 Then
'''                    sngMin = .sngParam7
'''                End If
'''            End If
'''            If frmGeryExp.chkSelAttr(7).Value = "1" Then
'''                If sngMax < .sngParam8 Then
'''                    sngMax = .sngParam8
'''                End If
'''                If sngMin > .sngParam8 Then
'''                    sngMin = .sngParam8
'''                End If
'''            End If
'''        End With
'''    Next
'''
'''    '--------------------------------------------------------------------
'''    '* 初值化
'''    '--------------------------------------------------------------------
'''    For i = 1 To intCount - 1
'''        With udtGRG(i)
''''            '--------------------------------------------------------------------
''''            '* gintSelFactorsMethod = 1 表「望大」
''''            '* gintSelFactorsMethod = 2 表「望小」
''''            '* gintSelFactorsMethod = 3 表「望目」
''''            '--------------------------------------------------------------------
''''            If frmGeryExp.chkSelAttr(0).Value = "1" Then
''''                If gintSelFactorsMethod(1) = 1 Then
''''                    .sngParam1 = (.sngParam1 - sngMin) / (sngMax - sngMin)
''''                ElseIf gintSelFactorsMethod(1) = 2 Then
''''                    .sngParam1 = (sngMax - .sngParam1) / (sngMax - sngMin)
''''                Else
''''                    .sngParam1 = 1 - Abs(.sngParam1 - udtGRG(intCount).sngParam1) / (IIf((sngMax - udtGRG(intCount).sngParam1 > udtGRG(intCount).sngParam1 - sngMin), sngMax - udtGRG(intCount).sngParam1, udtGRG(intCount).sngParam1 - sngMin))
''''                End If
''''            End If
''''
''''            If frmGeryExp.chkSelAttr(1).Value = "1" Then
''''                If gintSelFactorsMethod(2) = 1 Then
''''                    .sngParam2 = (.sngParam2 - sngMin) / (sngMax - sngMin)
''''                ElseIf gintSelFactorsMethod(2) = 2 Then
''''                    .sngParam2 = (sngMax - .sngParam2) / (sngMax - sngMin)
''''                Else
''''                    .sngParam2 = 1 - Abs(.sngParam2 - udtGRG(intCount).sngParam2) / (IIf((sngMax - udtGRG(intCount).sngParam2 > udtGRG(intCount).sngParam2 - sngMin), sngMax - udtGRG(intCount).sngParam2, udtGRG(intCount).sngParam2 - sngMin))
''''                End If
''''            End If
''''
''''            If frmGeryExp.chkSelAttr(2).Value = "1" Then
''''                If gintSelFactorsMethod(3) = 1 Then
''''                    .sngParam3 = (.sngParam3 - sngMin) / (sngMax - sngMin)
''''                ElseIf gintSelFactorsMethod(3) = 2 Then
''''                    .sngParam3 = (sngMax - .sngParam3) / (sngMax - sngMin)
''''                Else
''''                    .sngParam3 = 1 - Abs(.sngParam3 - udtGRG(intCount).sngParam3) / (IIf((sngMax - udtGRG(intCount).sngParam3 > udtGRG(intCount).sngParam3 - sngMin), sngMax - udtGRG(intCount).sngParam3, udtGRG(intCount).sngParam3 - sngMin))
''''                End If
''''            End If
''''
''''            If frmGeryExp.chkSelAttr(3).Value = "1" Then
''''                If gintSelFactorsMethod(4) = 1 Then
''''                    .sngParam4 = (.sngParam4 - sngMin) / (sngMax - sngMin)
''''                ElseIf gintSelFactorsMethod(4) = 2 Then
''''                    .sngParam4 = (sngMax - .sngParam4) / (sngMax - sngMin)
''''                Else
''''                    .sngParam4 = 1 - Abs(.sngParam4 - udtGRG(intCount).sngParam4) / (IIf((sngMax - udtGRG(intCount).sngParam4 > udtGRG(intCount).sngParam4 - sngMin), sngMax - udtGRG(intCount).sngParam4, udtGRG(intCount).sngParam4 - sngMin))
''''                End If
''''            End If
''''
''''            If frmGeryExp.chkSelAttr(4).Value = "1" Then
''''                If gintSelFactorsMethod(5) = 1 Then
''''                    .sngParam5 = (.sngParam5 - sngMin) / (sngMax - sngMin)
''''                ElseIf gintSelFactorsMethod(5) = 2 Then
''''                    .sngParam5 = (sngMax - .sngParam5) / (sngMax - sngMin)
''''                Else
''''                    .sngParam5 = 1 - Abs(.sngParam5 - udtGRG(intCount).sngParam5) / (IIf((sngMax - udtGRG(intCount).sngParam5 > udtGRG(intCount).sngParam5 - sngMin), sngMax - udtGRG(intCount).sngParam5, udtGRG(intCount).sngParam5 - sngMin))
''''                End If
''''            End If
''''
''''            If frmGeryExp.chkSelAttr(5).Value = "1" Then
''''                If gintSelFactorsMethod(6) = 1 Then
''''                    .sngParam6 = (.sngParam6 - sngMin) / (sngMax - sngMin)
''''                ElseIf gintSelFactorsMethod(6) = 2 Then
''''                    .sngParam6 = (sngMax - .sngParam6) / (sngMax - sngMin)
''''                Else
''''                    .sngParam6 = 1 - Abs(.sngParam6 - udtGRG(intCount).sngParam6) / (IIf((sngMax - udtGRG(intCount).sngParam6 > udtGRG(intCount).sngParam6 - sngMin), sngMax - udtGRG(intCount).sngParam6, udtGRG(intCount).sngParam6 - sngMin))
''''                End If
''''            End If
''''
''''            If frmGeryExp.chkSelAttr(6).Value = "1" Then
''''                If gintSelFactorsMethod(7) = 1 Then
''''                    .sngParam7 = (.sngParam7 - sngMin) / (sngMax - sngMin)
''''                ElseIf gintSelFactorsMethod(7) = 2 Then
''''                    .sngParam7 = (sngMax - .sngParam7) / (sngMax - sngMin)
''''                Else
''''                    .sngParam7 = 1 - Abs(.sngParam7 - udtGRG(intCount).sngParam7) / (IIf((sngMax - udtGRG(intCount).sngParam7 > udtGRG(intCount).sngParam7 - sngMin), sngMax - udtGRG(intCount).sngParam7, udtGRG(intCount).sngParam7 - sngMin))
''''                End If
''''            End If
''''
''''            If frmGeryExp.chkSelAttr(7).Value = "1" Then
''''                If gintSelFactorsMethod(8) = 1 Then
''''                    .sngParam8 = (.sngParam8 - sngMin) / (sngMax - sngMin)
''''                ElseIf gintSelFactorsMethod(8) = 2 Then
''''                    .sngParam8 = (sngMax - .sngParam8) / (sngMax - sngMin)
''''                Else
''''                    .sngParam8 = 1 - Abs(.sngParam8 - udtGRG(intCount).sngParam8) / (IIf((sngMax - udtGRG(intCount).sngParam8 > udtGRG(intCount).sngParam8 - sngMin), sngMax - udtGRG(intCount).sngParam8, udtGRG(intCount).sngParam8 - sngMin))
''''                End If
''''            End If
'''            ' 全部改為望目處理
'''            .sngParam1 = 1 - Abs(.sngParam1 - udtGRG(intCount).sngParam1) / (IIf((sngMax - udtGRG(intCount).sngParam1 > udtGRG(intCount).sngParam1 - sngMin), sngMax - udtGRG(intCount).sngParam1, udtGRG(intCount).sngParam1 - sngMin))
'''            .sngParam2 = 1 - Abs(.sngParam2 - udtGRG(intCount).sngParam2) / (IIf((sngMax - udtGRG(intCount).sngParam2 > udtGRG(intCount).sngParam2 - sngMin), sngMax - udtGRG(intCount).sngParam2, udtGRG(intCount).sngParam2 - sngMin))
'''            .sngParam3 = 1 - Abs(.sngParam3 - udtGRG(intCount).sngParam3) / (IIf((sngMax - udtGRG(intCount).sngParam3 > udtGRG(intCount).sngParam3 - sngMin), sngMax - udtGRG(intCount).sngParam3, udtGRG(intCount).sngParam3 - sngMin))
'''            .sngParam4 = 1 - Abs(.sngParam4 - udtGRG(intCount).sngParam4) / (IIf((sngMax - udtGRG(intCount).sngParam4 > udtGRG(intCount).sngParam4 - sngMin), sngMax - udtGRG(intCount).sngParam4, udtGRG(intCount).sngParam4 - sngMin))
'''            .sngParam5 = 1 - Abs(.sngParam5 - udtGRG(intCount).sngParam5) / (IIf((sngMax - udtGRG(intCount).sngParam5 > udtGRG(intCount).sngParam5 - sngMin), sngMax - udtGRG(intCount).sngParam5, udtGRG(intCount).sngParam5 - sngMin))
'''            .sngParam6 = 1 - Abs(.sngParam6 - udtGRG(intCount).sngParam6) / (IIf((sngMax - udtGRG(intCount).sngParam6 > udtGRG(intCount).sngParam6 - sngMin), sngMax - udtGRG(intCount).sngParam6, udtGRG(intCount).sngParam6 - sngMin))
'''            .sngParam7 = 1 - Abs(.sngParam7 - udtGRG(intCount).sngParam7) / (IIf((sngMax - udtGRG(intCount).sngParam7 > udtGRG(intCount).sngParam7 - sngMin), sngMax - udtGRG(intCount).sngParam7, udtGRG(intCount).sngParam7 - sngMin))
'''            .sngParam8 = 1 - Abs(.sngParam8 - udtGRG(intCount).sngParam8) / (IIf((sngMax - udtGRG(intCount).sngParam8 > udtGRG(intCount).sngParam8 - sngMin), sngMax - udtGRG(intCount).sngParam8, udtGRG(intCount).sngParam8 - sngMin))
'''        End With
'''    Next
    
    Exit Sub
    
    
ERR_HANDLE:
    Select Case Err.Number
        Case 6
            Debug.Print "[GCFR_General_Module.PreProcess2()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
        Case Else
            MsgBox "[GCFR_General_Module.PreProcess2()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
    
    Err.Clear
    Resume Next
End Sub


'***********************************************************************************
'* 計算90年01月01日以後的資料 --
'* Input Param :
'*   udtStock  ：儲存每日的股票資料
'*   udtIndex  ：儲存每日的指數資料
'*   intPos    ：要讀取udtIndex的那一筆資料(即所在位置)
'*   intStockNo：udtStock的總筆數
'*   intHighLow：評量高點或是低點 (高點=1、低點=0.5)
'* Return      ：回傳與所有聚類DataSet的GRG結果中最高者
'*
'***********************************************************************************
Public Sub EvaluateTheResult(ByRef udtStock() As StockData, _
                            ByRef udtIndex() As IndexData, _
                            ByVal intPos As Integer, _
                            ByVal intStockNo As Integer, _
                            ByVal intHighLow As Integer)
    Dim intDayLimit As Integer
    Dim sngHighDate As Single
    Dim sngHighEndP As Single
    Dim sngLowDate As Single
    Dim sngLowEndP As Single
    Dim sngGetDate As Single
    Dim sngGetEndP As Single
    Dim sngDiff As Single
    Dim i As Integer
    
    On Error GoTo ERR_HANDLE
    
    '---------------------------------------------------
    '* Initialize Variables
    '---------------------------------------------------
    intDayLimit = frmGeryExp.txtDayLimit.Text
    sngHighDate = 0
    sngHighEndP = -99999
    sngLowDate = 0
    sngLowEndP = 99999
    sngGetDate = udtStock(intPos).sngDate
    sngGetEndP = udtStock(intPos).sngEndprice
    
    For i = 1 To intDayLimit
        If sngHighEndP < udtStock(intPos).sngEndprice Then
            sngHighDate = udtStock(intPos).sngDate
            sngHighEndP = udtStock(intPos).sngEndprice
        End If
        If sngLowEndP > udtStock(intPos).sngEndprice Then
            sngLowDate = udtStock(intPos).sngDate
            sngLowEndP = udtStock(intPos).sngEndprice
        End If
        
        If intPos < intStockNo Then
            intPos = intPos + 1
        End If
    Next
    
    sngDiff = Abs(sngHighEndP - sngLowEndP) * 10 / 100
    
    If sngGetEndP <= (sngHighEndP + sngDiff) Then
        Debug.Print "nice-date=," & sngGetDate & ", is ok" & "src end-price=," & sngGetEndP & ", high end-price=," & sngHighEndP & ",low end-price=," & sngLowEndP & ", sngDiff=," & sngDiff
    Else
        Debug.Print "fail-date=," & sngGetDate & ", is ok" & "src end-price=," & sngGetEndP & ", high end-price,=" & sngHighEndP & ",low end-price=," & sngLowEndP & ", sngDiff=," & sngDiff & ",fa-value=," & (sngHighEndP - sngDiff) & ", need-value=," & (sngGetEndP - (sngHighEndP - sngDiff))
    End If
 
    Exit Sub
    
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.EvaluateTheResult()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
End Sub


'***********************************************************************************
'* 計算90年01月01日以後的資料 -- 使用GRG方法
'* Input Param :
'*   udtIndex        : 儲存每日的指數資料
'*   gudtClusterDat  : 儲存聚類結果的資料
'*   intPos          : 要讀取udtIndex的那一筆資料(即所在位置)
'*  Return :回傳與所有聚類dataset的GRG結果中最高者
'*
'***********************************************************************************
Public Function CalGRGValue2(ByRef udtIndex() As IndexData, _
                            ByRef gudtClusterDat() As udtGRGIndex, _
                            ByVal intPos As Integer, _
                            ByVal intHighLow As Integer) As Single
    Dim udtGRGDiff As udtGRGIndex                   ' 儲放差序列的變數
    Dim M(8) As Single                              ' 母數列的attribute的值
    Dim S(8) As Single                              ' 子數列的attribute的值
    Dim sngDiff(8) As Single                        ' 差序列的值
    Dim sngTestGRG(2) As udtGRGIndex                ' 用做前處理的暫存變數
    Dim sngGRG As Single                            ' 此次GRG的結果
    Dim sngPrevGRG As Single                        ' 前一次GRG的結果
    Dim sngSumGRG As Single                         ' 各個變變數的GRC合計
    Dim intGRGMethod As Integer                     ' 1 表示用GRGa, 2 表示用GRGd
    Dim sngMax As Single, sngMin As Single          '
    Dim i As Integer, j As Integer, k As Integer    '
    ReDim gstrSelFactors(8)                         ' Redefine 選擇儲存的attributes的size
          
    On Error GoTo ERR_HANDLE
    
'''    '--------------------------------------------------------------------
'''    '* Initialize Variables
'''    '--------------------------------------------------------------------
'''    sngMax = -99999
'''    sngMin = 99999
'''    sngGRG = 0
'''    sngPrevGRG = 0
'''    i = 0
'''
'''    '--- 選擇聚類的方法 ---
'''    If frmGeryExp.cboGRGMethod.Text = "GRGa" Then
'''        intGRGMethod = 1
'''    ElseIf frmGeryExp.cboGRGMethod.Text = "GRGd" Then
'''        intGRGMethod = 2
'''    Else
'''        Err.Raise 11000, "[CalGRGValue]", "本次測試，尚未選擇聚類的方法"
'''        Exit Function
'''    End If
'''
'''    '--- 若gintClusterCnt = 0 表示沒有聚類成功 ---
'''    If gintClusterCnt = 0 Then
'''        Err.Raise 11000, "[CalGRGValue]", "本次測試，沒有聚類成功"
'''        Exit Function
'''    End If
'''    '--- 判斷是抓高點或低點 ---
'''    If intHighLow = 1 Then          ' 高點
'''        If udtIndex(intPos).sngP24 < udtIndex(intPos).sngP72 Then
'''            Exit Function
'''        End If
'''    ElseIf intHighLow = 2 Then      ' 低點
'''        If udtIndex(intPos).sngP24 > udtIndex(intPos).sngP72 Then
'''            Exit Function
'''        End If
'''    End If
'''
'''    '--- 取得要比較的日期的資料 ---
''''    S(1) = GetFactorValue(udtIndex, intPos, 1)
''''    S(2) = GetFactorValue(udtIndex, intPos, 2)
''''    S(3) = GetFactorValue(udtIndex, intPos, 3)
''''    S(4) = GetFactorValue(udtIndex, intPos, 4)
''''    S(5) = GetFactorValue(udtIndex, intPos, 5)
''''    S(6) = GetFactorValue(udtIndex, intPos, 6)
''''    S(7) = GetFactorValue(udtIndex, intPos, 7)
''''    S(8) = GetFactorValue(udtIndex, intPos, 8)
'''    For i = 1 To gintClusterCnt     ' gintClusterCnt 表示聚類結果的數目
'''        '--- Initialize Variables ---
'''        sngMax = -99999
'''        sngMin = 99999
'''        sngSumGRG = 0
'''        gintSelCount = gintSelAttrCount(i)
'''        With gudtSelAttrs(i)
'''            '-------------------------------------------------------
'''            '* 將第i項的聚類結悲 assign 至gstrSelFactors變數中，
'''            '* 供Method:GetFactorValue()判斷用
'''            '-------------------------------------------------------
'''            gstrSelFactors(1) = .strParam1
'''            gstrSelFactors(2) = .strParam2
'''            gstrSelFactors(3) = .strParam3
'''            gstrSelFactors(4) = .strParam4
'''            gstrSelFactors(5) = .strParam5
'''            gstrSelFactors(6) = .strParam6
'''            gstrSelFactors(7) = .strParam7
'''            gstrSelFactors(8) = .strParam8
'''            '--- 取得要比較的日期的資料 ---
'''            S(1) = GetFactorValue(udtIndex, intPos, 1)
'''            S(2) = GetFactorValue(udtIndex, intPos, 2)
'''            S(3) = GetFactorValue(udtIndex, intPos, 3)
'''            S(4) = GetFactorValue(udtIndex, intPos, 4)
'''            S(5) = GetFactorValue(udtIndex, intPos, 5)
'''            S(6) = GetFactorValue(udtIndex, intPos, 6)
'''            S(7) = GetFactorValue(udtIndex, intPos, 7)
'''            S(8) = GetFactorValue(udtIndex, intPos, 8)
'''        End With
'''        '--- Assign母數列的值至變數M中 ---
'''        M(1) = gudtClusterDat(i).sngParam1
'''        M(2) = gudtClusterDat(i).sngParam2
'''        M(3) = gudtClusterDat(i).sngParam3
'''        M(4) = gudtClusterDat(i).sngParam4
'''        M(5) = gudtClusterDat(i).sngParam5
'''        M(6) = gudtClusterDat(i).sngParam6
'''        M(7) = gudtClusterDat(i).sngParam7
'''        M(8) = gudtClusterDat(i).sngParam8
'''
'''        '-------------------------------------------------------------------
'''        '* 取得差序列
'''        '-------------------------------------------------------------------
'''        If gintSelCount < 2 Then
'''            Err.Raise 11001, "[CalGRGValue]", _
'''                        "選擇之attributes個數不得小於２，否則無法GRG"
'''        Else
'''            If intGRGMethod = 1 Then
'''                '--- 計算各個attributes間的面積 -- For GRGa ---
'''                For j = 1 To gintSelCount - 1
'''                    If (M(j + 1) - S(j + 1)) * (M(j) - S(j)) >= 0 Then
'''                        sngDiff(j) = (Abs(M(j + 1) - S(j + 1)) + Abs(M(j) - S(j))) / 2
'''                    Else
'''                        sngDiff(j) = (Abs(M(j + 1) - S(j + 1)) * Abs(M(j + 1) - S(j + 1)) + Abs(M(j) - S(j)) * Abs(M(j) - S(j))) / (2 * (Abs(M(j + 1) - S(j + 1)) + Abs(M(j) - S(j))))
'''                    End If
'''                Next
'''            ElseIf intGRGMethod = 2 Then
'''                '--- 計算各個attributes間的距離 -- For GRGd ---
'''                For j = 1 To gintSelCount
'''                    sngDiff(j) = Abs(M(j) - S(j))
'''                Next
'''            End If
'''        End If
'''        '--------------------------------------------------------------------
'''        '* 找出最大值及最小值
'''        '--------------------------------------------------------------------
'''        If intGRGMethod = 1 Then
'''            For j = 1 To gintSelCount - 1
'''                If sngDiff(j) > sngMax Then
'''                    sngMax = sngDiff(j)
'''                End If
'''                If sngDiff(j) < sngMin Then
'''                    sngMin = sngDiff(j)
'''                End If
'''            Next
'''        ElseIf intGRGMethod = 2 Then
'''            '--- 找出差序列中的最大值及最小值 ---
'''            For j = 1 To gintSelCount
'''                If sngDiff(j) > sngMax Then
'''                    sngMax = sngDiff(j)
'''                End If
'''                If sngDiff(j) < sngMin Then
'''                    sngMin = sngDiff(j)
'''                End If
'''            Next
'''        End If
'''        '--------------------------------------------------------------------
'''        '* 計算GRC
'''        '--------------------------------------------------------------------
'''        If intGRGMethod = 1 Then
'''            '--- Mthod 1 : Use GRGa 計算各個attributes間的面積 ---
'''            For j = 1 To gintSelCount - 1
'''                If (sngMax + sngDiff(j)) = 0 Then
'''                    sngDiff(j) = 1
'''                Else
'''                    sngDiff(j) = (sngMax * gsngRo) / (sngDiff(j) + sngMax * gsngRo)
'''                End If
'''            Next
'''        ElseIf intGRGMethod = 2 Then
'''            '--- Method 2 : Use GRGd 計算各個attributes間的距離 ---
'''            With gudtSelAttrs(i)
'''                If Len(.strParam1) > 0 Then
'''                    If (sngDiff(1) + gsngRo * sngMax) = 0 Then
'''                        sngDiff(1) = 0
'''                    Else
'''                        sngDiff(1) = (sngMin + gsngRo * sngMax) / (sngDiff(1) + gsngRo * sngMax)
'''                    End If
'''                End If
'''                If Len(.strParam2) > 0 Then
'''                    If (sngDiff(2) + gsngRo * sngMax) = 0 Then
'''                        sngDiff(2) = 0
'''                    Else
'''                        sngDiff(2) = (sngMin + gsngRo * sngMax) / (sngDiff(2) + gsngRo * sngMax)
'''                    End If
'''                End If
'''                If Len(.strParam3) > 0 Then
'''                    If (sngDiff(3) + gsngRo * sngMax) = 0 Then
'''                        sngDiff(3) = 0
'''                    Else
'''                        sngDiff(3) = (sngMin + gsngRo * sngMax) / (sngDiff(3) + gsngRo * sngMax)
'''                    End If
'''                End If
'''                If Len(.strParam4) > 0 Then
'''                    If (sngDiff(4) + gsngRo * sngMax) = 0 Then
'''                        sngDiff(4) = 0
'''                    Else
'''                        sngDiff(4) = (sngMin + gsngRo * sngMax) / (sngDiff(4) + gsngRo * sngMax)
'''                    End If
'''                End If
'''                If Len(.strParam5) > 0 Then
'''                    If (sngDiff(5) + gsngRo * sngMax) = 0 Then
'''                        sngDiff(5) = 0
'''                    Else
'''                        sngDiff(5) = (sngMin + gsngRo * sngMax) / (sngDiff(5) + gsngRo * sngMax)
'''                    End If
'''                End If
'''                If Len(.strParam6) > 0 Then
'''                    If (sngDiff(6) + gsngRo * sngMax) = 0 Then
'''                        sngDiff(6) = 0
'''                    Else
'''                        sngDiff(6) = (sngMin + gsngRo * sngMax) / (sngDiff(6) + gsngRo * sngMax)
'''                    End If
'''                End If
'''                If Len(.strParam7) > 0 Then
'''                    If (sngDiff(7) + gsngRo * sngMax) = 0 Then
'''                        sngDiff(7) = 0
'''                    Else
'''                        sngDiff(7) = (sngMin + gsngRo * sngMax) / (sngDiff(7) + gsngRo * sngMax)
'''                    End If
'''                End If
'''                If Len(.strParam8) > 0 Then
'''                    If (sngDiff(8) + gsngRo * sngMax) = 0 Then
'''                        sngDiff(8) = 0
'''                    Else
'''                        sngDiff(8) = (sngMin + gsngRo * sngMax) / (sngDiff(8) + gsngRo * sngMax)
'''                    End If
'''                End If
'''            End With
'''            '--------------------------------------------------------------------
'''            '* 計算GRG
'''            '--------------------------------------------------------------------
'''            If intGRGMethod = 1 Then
'''                '--- For GRGa ---
'''                For j = 1 To gintSelCount - 1
'''                    sngSumGRG = sngSumGRG + sngDiff(j)
'''                Next
'''                sngGRG = sngSumGRG / (gintSelCount - 1)
'''            ElseIf intGRGMethod = 2 Then
'''                '--- For GRGd ---
'''                For j = 1 To gintSelCount
'''                    sngSumGRG = sngSumGRG + sngDiff(j)
'''                Next
'''                sngGRG = sngSumGRG / gintSelCount
'''            End If
'''
'''            '--- 找出與所有母數列運算中最大的值 ---
'''            If sngGRG > sngPrevGRG Then
'''                sngPrevGRG = sngGRG
'''            End If
'''        End If
'''    Next
'''
'''    '--------------------------------------------------------------------
'''    '* 傳回運算後的GRG結果
'''    '--------------------------------------------------------------------
'''    If sngGRG >= gsngCompFa Then
'''        '--- 傳回結果 ---
'''        CalGRGValue2 = sngGRG
'''    Else
'''        CalGRGValue2 = 0
'''    End If
           
    Exit Function
   
    
ERR_HANDLE:
    Select Case Err.Number
        Case 11001
            MsgBox "[GCFR_General_Module.CalGRGValue2()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
            Err.Raise11001
            Exit Function
        Case Else
            MsgBox "[GCFR_General_Module.CalGRGValue2()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
            Err.Raise 11001
            Exit Function
    End Select
End Function


' 921221 -- Modified -- Not Finished
'***********************************************************************************
'*
'***********************************************************************************
Sub CaluateGRGs(ByRef udtStock() As StockData, ByRef udtIndex() As IndexData, _
                ByVal intDayIndex As Integer, ByVal sngBegDate As Single, _
                ByVal sngEndDate As Single, ByVal intDay As Integer, _
                ByVal strIndex As String)
    Dim sngGRGDiff() As Single
    Dim sngGRGRo() As Single
    Dim sngRoSum As Single
    Dim sngGRG As Single
    Dim i As Integer
    Dim j As Integer
    Dim intSize As Integer
    Dim intBegPos As Integer    ' 計算起始位置
    Dim intEndPos As Integer    ' 計算結束位置
    Dim sngMax As Single
    Dim sngMin As Single
    
    On Error GoTo ERR_HANDLE
    
    '--- Initialize Variables ---
    intSize = 0
    i = 1
    sngMax = 0
    sngMin = 9999
    
    While ((udtStock(i).sngDate <= sngEndDate) And (i < intDayIndex))
        '--- 記錄起始位置 ---
        If udtStock(i).sngDate = sngBegDate Then
            intBegPos = i
        End If
        '--- 記錄結束位置 ---
        If udtStock(i).sngDate = sngEndDate Then
            intEndPos = i
        End If
        '--- 記錄Size ---
        If (udtStock(i).sngDate >= sngBegDate) And (udtStock(i).sngDate <= sngEndDate) Then
            intSize = intSize + 1
        End If

        i = i + 1
    Wend
    
    ReDim sngGRGDiff(intSize)
    ReDim sngGRGRo(intSize)
    j = 1
    For i = intBegPos To intEndPos - 1
        '--- 計算GRG Difference ---
        If strIndex = "PSY" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngPSY - udtIndex(i).sngPSY))
        ElseIf strIndex = "WMS" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngWMS - udtIndex(i).sngWMS))
        ElseIf strIndex = "K" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngK - udtIndex(i).sngK))
        ElseIf strIndex = "D" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngD - udtIndex(i).sngD))
        ElseIf strIndex = "EMA_S" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngDIF - udtIndex(i).sngDIF))
        ElseIf strIndex = "EMA_L" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngDIF - udtIndex(i).sngDIF))
        ElseIf strIndex = "MACD" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngMACD - udtIndex(i).sngMACD))
        ElseIf strIndex = "RSI_S" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngRSI_S - udtIndex(i).sngRSI_S))
        ElseIf strIndex = "RSI_L" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngRSI_L - udtIndex(i).sngRSI_L))
        ElseIf strIndex = "BIAS" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngBias - udtIndex(i).sngBias))
'''''        ElseIf strIndex = "AR" Then
'''''            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngAR - udtIndex(i).sngAR))
'''''        ElseIf strIndex = "BR" Then
'''''            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngBR - udtIndex(i).sngBR))
'''''        ElseIf strIndex = "OBV" Then
'''''            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngOBV - udtIndex(i).sngOBV))
'''''        ElseIf strIndex = "TAPI" Then
'''''            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngTAPI - udtIndex(i).sngTAPI))
        ElseIf strIndex = "VR" Then
            sngGRGDiff(j) = Abs((udtStock(i + 1).sngEndprice - udtStock(i).sngEndprice) - (udtIndex(i + 1).sngVR - udtIndex(i).sngVR))
        Else
            MsgBox "沒有符合的指標", vbOKOnly
        End If
        j = j + 1
    Next
    
    For i = 1 To intSize
        If sngMax < sngGRGDiff(i) Then
            sngMax = sngGRGDiff(i)
        End If
        If sngMin > sngGRGDiff(i) Then
            sngMin = sngGRGDiff(i)
        End If
    Next
    
    For i = 1 To intSize - 1
        sngGRGRo(i) = (sngMin + sngMax) / (sngGRGDiff(i) + sngMax)
    Next
    
    For i = 1 To intSize - 1
        sngRoSum = sngRoSum + sngGRGRo(i)
    Next
    
    sngGRG = sngRoSum / (intSize - 1)
    
    If gsngTestGRG < sngGRG Then
        gsngTestGRG = sngGRG
        frmGeryExp.cboIndDays.Text = intDay
    End If
       
    frmGeryExp.txtGRGResults.Text = "Day-index= " & intDay & ", snggrg=" & sngGRG & vbCrLf & frmGeryExp.txtGRGResults.Text
    
    Exit Sub
    
ERR_HANDLE:
'    MsgBox "[Method: CaluateGRGs()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    Debug.Print "[GCFR_General_Module.CaluateGRGs()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
End Sub



'***************************************************************************************************
'* 說    明: 利用灰聚類方法，找出每個Sector(Use CY split)間的高低點
'* 輸入參數: udtStock  : 每日股價資料
'*           udtIndex  ：儲存的指數資料
'*           intStockNo：資料筆數
'*           strStartDT：傳入的資料起始日期
'*           strEndDT  ：傳入的資料結束日期
'* 輸出參數: 無
'* 版    本: 1.00: 20041214 新增
'*           1.10: 20050604 Modified
'***************************************************************************************************
Public Sub GetHighLowPoints4CY(ByRef udtStock() As StockData, _
                            ByRef udtIndex() As IndexData, _
                            ByVal intStockNo As Integer, _
                            ByVal strStartDT As String, _
                            ByVal strEndDT As String)
    Dim intCurrentPos As Integer    ' 記錄目前指標的位置
    Dim intBegPos As Integer        ' 記錄此Sector的起始位置
    Dim intEndPos As Integer        ' 記錄下一個Sector的起始位置
    Dim sngFlag As Single           ' 記錄此次要聚類的是高點或低點
    Dim udtGRG() As udtGRGIndex     ' 儲存要做聚類的資料
    Dim udtGroup() As Boolean       ' 記錄是那幾筆資料聚成高(低)點
    Dim intMPos As Integer          ' 記錄此聚類的的高(低)點的位置
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ERR_HANDLE
    
    '--- Initialize Variables ---
    ReDim gudtClusterDat(100)   ' 記錄每個sector最後聚類結果的值 (最多可記100個)
    ReDim gudtCYSHDat(50)
    ReDim gudtCYSLDat(50)
    intCurrentPos = 1
    gintCYSHCount = 1
    gintCYSLCount = 1
    '--- Reset gudtCYHDat and gudtCYSLDat ---
    For i = 1 To 50
        With gudtCYSHDat(i)
            .sngDate = 0
            .sngGRG = 0
            .sngMax = 0
            .sngMin = 0
            .sngParam1 = 0
            .sngParam2 = 0
            .sngParam3 = 0
            .sngParam4 = 0
            .sngParam5 = 0
            .sngParam6 = 0
            .sngParam7 = 0
            .sngParam8 = 0
        End With
        With gudtCYSLDat(i)
            .sngDate = 0
            .sngGRG = 0
            .sngMax = 0
            .sngMin = 0
            .sngParam1 = 0
            .sngParam2 = 0
            .sngParam3 = 0
            .sngParam4 = 0
            .sngParam5 = 0
            .sngParam6 = 0
            .sngParam7 = 0
            .sngParam8 = 0
        End With
    Next
    
    '===================================================================
    '* 只處理Learning Set部份
    '===================================================================
    While intCurrentPos < intStockNo
        If udtStock(intCurrentPos).sngDate >= LEARN_START_DATE And _
            udtStock(intCurrentPos).sngDate <= LEARN_END_DATE Then
            intBegPos = intCurrentPos
            intEndPos = intCurrentPos
            sngFlag = udtIndex(intBegPos).sngSector
            
            While (sngFlag = udtIndex(intEndPos).sngSector) And (intEndPos < intStockNo)
                If sngFlag = udtIndex(intEndPos).sngSector Then
                    intEndPos = intEndPos + 1
                End If
            Wend
            intCurrentPos = intEndPos - 1           ' 記錄下一個要處理的Sector
            ReDim udtGroup(intEndPos - intBegPos)   ' 記錄該Sector那幾筆聚成一類 (+1的原因是因為最後1筆要儲在母數列)
            ReDim udtGRG(intEndPos - intBegPos)     ' 儲存該Sector的聚類資料
            j = 1
            
            '===================================================================
            '* 在做分析前，要先是所有指標重新計算
            '* 目前高點與低點各自有一套自已的指標天數的記錄
            '===================================================================
            If sngFlag = 1 Then
                PSY_No = 13
                WMS_No = 9
                KD_No = 9
                MACD_No = 24
                EMA_S = 12
                EMA_L = 26
                RSI_S = 6
                RSI_L = 12
                Bias_No = 10
                VR_NO = 12
                OBV_No = 12
                TAPI_No = 12
                AR_No = 12
            Else
                PSY_No = 13
                WMS_No = 9
                KD_No = 9
                MACD_No = 24
                EMA_S = 12
                EMA_L = 26
                RSI_S = 6
                RSI_L = 6
                Bias_No = 24
                VR_NO = 12
                OBV_No = 12
                TAPI_No = 12
                AR_No = 12
            End If
            Call subCalculateIndex(gudtStockDay, gudtIndexDay, gintDayIndex)
            Call subCalculateIndex(gudtStockWeek, gudtIndexWeek, gintWeekIndex)
            Call subCalculateIndex(gudtStockMonth, gudtIndexMonth, gintMonthIndex)
            
            '--- 將此次要分析的udtIndex()資料assign to udtGRG() ---
            For i = intBegPos To intEndPos - 1
                udtGRG(j).sngDate = udtStock(i).sngDate
                udtGRG(j).sngParam1 = GetAttributesValue4CY(udtIndex, i, 1, sngFlag)
                udtGRG(j).sngParam2 = GetAttributesValue4CY(udtIndex, i, 2, sngFlag)
                udtGRG(j).sngParam3 = GetAttributesValue4CY(udtIndex, i, 3, sngFlag)
                udtGRG(j).sngParam4 = GetAttributesValue4CY(udtIndex, i, 4, sngFlag)
                udtGRG(j).sngParam5 = GetAttributesValue4CY(udtIndex, i, 5, sngFlag)
                udtGRG(j).sngParam6 = GetAttributesValue4CY(udtIndex, i, 6, sngFlag)
                udtGRG(j).sngParam7 = GetAttributesValue4CY(udtIndex, i, 7, sngFlag)
                udtGRG(j).sngParam8 = GetAttributesValue4CY(udtIndex, i, 8, sngFlag)
                j = j + 1
            Next
            '--- 找出欲分析Sector's母數列並將之assign至array的最後一列 ---
            intMPos = FindHighLowDate(udtStock, intBegPos, intEndPos - 1, sngFlag)
            udtGRG(intEndPos - intBegPos).sngDate = udtStock(intMPos).sngDate
            udtGRG(intEndPos - intBegPos).sngParam1 = GetAttributesValue4CY(udtIndex, intMPos, 1, sngFlag)
            udtGRG(intEndPos - intBegPos).sngParam2 = GetAttributesValue4CY(udtIndex, intMPos, 2, sngFlag)
            udtGRG(intEndPos - intBegPos).sngParam3 = GetAttributesValue4CY(udtIndex, intMPos, 3, sngFlag)
            udtGRG(intEndPos - intBegPos).sngParam4 = GetAttributesValue4CY(udtIndex, intMPos, 4, sngFlag)
            udtGRG(intEndPos - intBegPos).sngParam5 = GetAttributesValue4CY(udtIndex, intMPos, 5, sngFlag)
            udtGRG(intEndPos - intBegPos).sngParam6 = GetAttributesValue4CY(udtIndex, intMPos, 6, sngFlag)
            udtGRG(intEndPos - intBegPos).sngParam7 = GetAttributesValue4CY(udtIndex, intMPos, 7, sngFlag)
            udtGRG(intEndPos - intBegPos).sngParam8 = GetAttributesValue4CY(udtIndex, intMPos, 8, sngFlag)
                       
            '--- 資料前處理 ---
            Call PreProcess4CY(udtGRG, intEndPos - intBegPos, sngFlag)
            '--- 聚類 ---
            Call CalculateGRG4CY(udtGRG, intEndPos - intBegPos, udtStock(intMPos).sngDate, sngFlag)
            
            '--- 聚類後處理 ---
            '===================================================================
            '* 記錄下這次聚類的結果 --
            '===================================================================
            If udtGRG(UBound(udtGRG)).sngGRG <> -1 Then
                If sngFlag = 1 Then
                    gudtCYSHDat(gintCYSHCount) = udtGRG(UBound(udtGRG))
                    ' 將聚類結果的attribute值還原
                    Call RestoreData4CY(gudtCYSHDat, gintCYSHCount, sngFlag)
                    gintCYSHCount = gintCYSHCount + 1
                Else
                    gudtCYSLDat(gintCYSLCount) = udtGRG(UBound(udtGRG))
                    ' 將聚類結果的attribute值還原
                    Call RestoreData4CY(gudtCYSLDat, gintCYSLCount, sngFlag)
                    gintCYSLCount = gintCYSLCount + 1
                End If
            
                '===================================================================
                '* 將該sector的資料的GRG值Assign回udtindex()供程式display the value
                '===================================================================
                i = 1   ' 陣列起始位置為0
                gsngCompFa = GRG_HIGH
                gsngCompFa2 = GRG_LOW
            
                For j = intBegPos To intEndPos - 1
                    If udtGRG(i).sngGRG >= gsngCompFa Then
                        udtIndex(j).sngGRG = udtGRG(i).sngGRG
                    ElseIf udtGRG(i).sngGRG <= gsngCompFa2 Then
                        udtIndex(j).sngGRG = udtGRG(i).sngGRG
                    Else
                        udtIndex(j).sngGRG = 0
                    End If
                    i = i + 1
                Next
            End If
        End If
        intCurrentPos = intCurrentPos + 1
    Wend
    
    '-----------------------------------------------------------
    '* Print High-Pioints and Low-Points' Pattern STA
    '-----------------------------------------------------------
    For i = 1 To gintCYSHCount
        Debug.Print gudtCYSHDat(i).sngDate & "," & _
            gudtCYSHDat(i).sngParam1 & "," & _
            gudtCYSHDat(i).sngParam2 & "," & _
            gudtCYSHDat(i).sngParam3 & "," & _
            gudtCYSHDat(i).sngParam4 & "," & _
            gudtCYSHDat(i).sngParam5 & "," & _
            gudtCYSHDat(i).sngParam6 & "," & _
            gudtCYSHDat(i).sngParam7 & "," & _
            gudtCYSHDat(i).sngParam8
    Next
    
    For i = 1 To gintCYSLCount
        Debug.Print gudtCYSLDat(i).sngDate & "," & _
            gudtCYSLDat(i).sngParam1 & "," & _
            gudtCYSLDat(i).sngParam2 & "," & _
            gudtCYSLDat(i).sngParam3 & "," & _
            gudtCYSLDat(i).sngParam4 & "," & _
            gudtCYSLDat(i).sngParam5 & "," & _
            gudtCYSLDat(i).sngParam6 & "," & _
            gudtCYSLDat(i).sngParam7 & "," & _
            gudtCYSLDat(i).sngParam8
    Next
    
    '-----------------------------------------------------------
    '* Print High-Pioints and Low-Points' Pattern END
    '-----------------------------------------------------------
    
    '--- 取得目前資料中最新的一筆的日期 ---
    Dim strTestEndDate As String
    strTestEndDate = GetTheLatestDate(udtStock, intStockNo)
    '--- 產生Testing-Set的GRG值 ---
    Call GenearteTestingSetGRG(udtStock, udtIndex, intStockNo, TEST_START_DATE, strTestEndDate)
    intCurrentPos = 2
    While intCurrentPos < intStockNo
        If udtStock(intCurrentPos).sngDate >= LEARN_END_DATE Then
            If udtIndex(intCurrentPos).sngGRG < GRG_HIGH _
                And udtIndex(intCurrentPos).sngGRG > GRG_LOW Then
                udtIndex(intCurrentPos).sngGRG = 0
            End If
        End If
        intCurrentPos = intCurrentPos + 1
    Wend
   
    Exit Sub
    
ERR_HANDLE:
    Select Case Err.Number
        Case 6
            MsgBox "[GCFR_General_Module.GetHighLowPoints4CY()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
            Resume Next
        Case 10001
            MsgBox "[GCFR_General_Module.GetHighLowPoints4CY()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
        Case Else
            MsgBox "[GCFR_General_Module.GetHighLowPoints4CY()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
   
    Err.Clear
    Resume Next
End Sub


'***************************************************************************************************
'* 說    明: 顯示買賣訊號 (sngGRG值大於或小於設定顯示訊號者，才計算；計算結束後，如
'*           果高於設定)
'* 輸入參數: udtStock  : 每日股價資料
'*           udtIndex  ：儲存的指數資料
'*           intStockNo：資料筆數
'* 輸出參數: 無
'* 版    本: 1.00: 20040411 新增
'*           1.10: 20050604 Modified
'***************************************************************************************************
Public Sub DisplaySignals4CY(ByRef udtStock() As StockData, _
                            ByRef udtIndex() As IndexData, _
                            ByVal intStockNo As Integer)
    Dim i As Integer
    Dim bln1stAHSAppear As Boolean  ' CY>0區間中，若第一次出現AHS則此旗標設為true
    Dim bln1stALSAppear As Boolean  ' CY<0區間中，若第一次出現ALS則此旗標設為true
    Dim bln2ndAHSAppear As Boolean  ' CY>0區間中，若第二次出現AHS則此旗標設為true
    Dim bln2ndALSAppear As Boolean  ' CY<0區間中，若第二次出現ALS則此旗標設為true
    Dim blnHSAppear As Boolean
    Dim blnHS2Appear As Boolean
    Dim intHSAppearDays As Integer
    Dim intLSAppearDays As Integer
    Dim blnLSAppear As Boolean
    Dim blnLS2Appear As Boolean
    Dim blnAHSCross As Boolean      ' 記錄是否出現當日收盤價 cross-down 60MAP
    Dim blnALSCross As Boolean      ' 記錄是否出現當日收盤價 cross-upon 60MAP
    Dim sngLSBiasValue As Single    ' 記錄Fuzzy化後的 LSBias 值
    Dim sngMASValue As Single       ' 記錄Fuzzy化後的 MAS 值
    Dim blnLSBiasH3 As Boolean

    On Error GoTo ERR_HANDLE

    bln1stAHSAppear = False
    bln1stALSAppear = False
    bln2ndAHSAppear = False
    bln2ndALSAppear = False
    blnHSAppear = False
    blnHS2Appear = False
    blnLSAppear = False
    blnLS2Appear = False
    intHSAppearDays = 0
    intLSAppearDays = 0
    blnAHSCross = False
    blnALSCross = False
    blnLSBiasH3 = False

    For i = 3 To intStockNo
        '*** Initialize Variables ***
        udtIndex(i).sngSignal = 0
        
        ' 851018以後才有CY值可做分析
        If udtStock(i).sngDate > LEARN_START_DATE Then
            If udtIndex(i).sngCyOp > 0 Then
                ' Initialize Variables
                bln1stALSAppear = False
                bln2ndALSAppear = False
                blnLSAppear = False
                blnLS2Appear = False
                intLSAppearDays = 0
                blnALSCross = False

                '*** 產生高點訊號 ***
                If Not bln1stALSAppear Then
                    If udtIndex(i).sngGRG >= GRG_HIGH Then
                        ' Bias Fuzzy Rules
                        If udtIndex(i).sngLSBias >= LSBIASH3 Then
                            sngLSBiasValue = 1
                        ElseIf udtIndex(i).sngLSBias >= LSBIASH2 Then
                            sngLSBiasValue = (udtIndex(i).sngLSBias - LSBIASH2) / (LSBIASH3 - LSBIASH2) * (1 - 0.9) + 0.9
                        ElseIf udtIndex(i).sngLSBias >= LSBIASH1 Then
                            sngLSBiasValue = (udtIndex(i).sngLSBias - LSBIASH1) / (LSBIASH2 - LSBIASH1) * (0.9 - 0.8) + 0.8
                        ElseIf udtIndex(i).sngLSBias >= 5 Then
                            sngLSBiasValue = (udtIndex(i).sngLSBias - 5) / (LSBIASH1 - 5) * (0.8 - 0.7) + 0.7
                        Else
                            sngLSBiasValue = 0
                        End If
                        ' MAPS Fuzzy Rules
                        If udtIndex(i).sngMASlope >= MASH5 Then
                            sngMASValue = 1
                        ElseIf udtIndex(i).sngMASlope >= MASH4 Then
                            sngMASValue = (udtIndex(i).sngMASlope - MASH4) / (MASH5 - MASH4) * (1 - 0.9) + 0.9
                        ElseIf udtIndex(i).sngMASlope >= MASH3 Then
                            sngMASValue = (udtIndex(i).sngMASlope - MASH3) / (MASH4 - MASH3) * (0.9 - 0.8) + 0.8
                        ElseIf udtIndex(i).sngMASlope >= MASH2 Then
                            sngMASValue = (udtIndex(i).sngMASlope - MASH2) / (MASH3 - MASH2) * (0.8 - 0.7) + 0.7
                        ElseIf udtIndex(i).sngMASlope >= MASH1 Then
                            sngMASValue = (udtIndex(i).sngMASlope - MASH1) / (MASH2 - MASH1) * (0.7 - 0.6) + 0.6
                        Else
                            sngMASValue = 0
                        End If
                        ' Calculate The Signal Value
                        If sngLSBiasValue * sngMASValue > 0 Then
                            udtIndex(i).sngSignal = udtIndex(i).sngGRG * 150 * sngLSBiasValue * sngMASValue
                            If Not (udtIndex(i).sngSignal >= HIGHSIGNAL) Then
                                udtIndex(i).sngSignal = 0
                            Else
                                blnHSAppear = True
                                intHSAppearDays = 1
                            End If
                        End If
                    End If
                    ' 如果高點出現，則開始計算經過天數，若超過5天則重新歸 0
                    If intHSAppearDays > 0 And intHSAppearDays < 5 Then
                        intHSAppearDays = intHSAppearDays + 1
                    Else
                        intHSAppearDays = 0
                    End If
                    '*** 判斷是否出現第一個AHS ***
                    ' 出現高點訊號5日內，若當日收盤價貫5MA and 10MA
                    If intHSAppearDays > 0 _
                        And udtStock(i).sngEndprice < udtIndex(i).sngP5 _
                        And udtStock(i).sngEndprice < udtIndex(i).sngP10 Then
                        udtIndex(i).sngSignal = 145
                        udtIndex(i).sngSignal = 140
                        bln1stAHSAppear = True
                        intHSAppearDays = 0
                    End If

                    If blnHSAppear And Not bln1stAHSAppear Then
                        ' 當日收盤價貫5MA and 10MA and 出現死亡交叉
                        If (udtStock(i).sngEndprice < udtIndex(i).sngP5 _
                            And udtStock(i).sngEndprice < udtIndex(i).sngP10) _
                            And (udtIndex(i - 1).sngP5 >= udtIndex(i - 1).sngP10 _
                            And udtIndex(i).sngP5 <= udtIndex(i).sngP10) Then
                            udtIndex(i).sngSignal = 145
                            udtIndex(i).sngSignal = 140
                            bln1stAHSAppear = True
                        End If
                        ' 當日收盤價貫5MA and 10MA
                        If udtStock(i).sngEndprice < udtIndex(i).sngP5 _
                            And udtStock(i).sngEndprice < udtIndex(i).sngP10 Then
                            udtIndex(i).sngSignal = 145
                            udtIndex(i).sngSignal = 140
                            bln1stAHSAppear = True
                        End If
                        ' 出現死亡交叉
                        If udtIndex(i - 1).sngP5 >= udtIndex(i - 1).sngP10 _
                            And udtIndex(i).sngP5 <= udtIndex(i).sngP10 Then
                            udtIndex(i).sngSignal = 145
                            udtIndex(i).sngSignal = 140
                            bln1stAHSAppear = True
                        End If
                    End If
                Else
                    ' 找第二個AHS
                    If bln1stAHSAppear And Not bln2ndAHSAppear Then   ' 第一個AHS出現後
                        If udtIndex(i).sngGRG >= (GRG_HIGH - 0.1) Then
                            ' Bias Fuzzy Rules
                            If udtIndex(i).sngLSBias >= 4 Then
                                sngLSBiasValue = 1
                            ElseIf udtIndex(i).sngLSBias >= 3.5 Then
                                sngLSBiasValue = (udtIndex(i).sngLSBias - 3.5) / (4 - 3.5) * (1 - 0.9) + 0.9
                            ElseIf udtIndex(i).sngLSBias >= 3 Then
                                sngLSBiasValue = (udtIndex(i).sngLSBias - 3) / (3.5 - 3) * (0.9 - 0.8) + 0.8
                            ElseIf udtIndex(i).sngLSBias >= 2 Then
                                sngLSBiasValue = (udtIndex(i).sngLSBias - 2) / (3 - 2) * (0.8 - 0.7) + 0.7
                            ElseIf udtIndex(i).sngLSBias >= 1 Then
                                sngLSBiasValue = (udtIndex(i).sngLSBias - 1) / (2 - 1) * (0.7 - 0.6) + 0.6
                            Else
                                sngLSBiasValue = 0
                            End If
                            ' MAPS Fuzzy Rules
                            If udtIndex(i).sngMASlope >= MASH5 Then
                                sngMASValue = 1
                            ElseIf udtIndex(i).sngMASlope >= MASH4 Then
                                sngMASValue = (udtIndex(i).sngMASlope - MASH4) / (MASH5 - MASH4) * (1 - 0.9) + 0.9
                            ElseIf udtIndex(i).sngMASlope >= MASH3 Then
                                sngMASValue = (udtIndex(i).sngMASlope - MASH3) / (MASH4 - MASH3) * (0.9 - 0.8) + 0.8
                            ElseIf udtIndex(i).sngMASlope >= MASH2 Then
                                sngMASValue = (udtIndex(i).sngMASlope - MASH2) / (MASH3 - MASH2) * (0.8 - 0.7) + 0.7
                            ElseIf udtIndex(i).sngMASlope >= MASH1 Then
                                sngMASValue = (udtIndex(i).sngMASlope - MASH1) / (MASH2 - MASH1) * (0.7 - 0.6) + 0.6
                            Else
                                sngMASValue = 0
                            End If
                            ' Calcualte Signal Value
                            If sngLSBiasValue * sngMASValue > 0 Then
                                udtIndex(i).sngSignal = udtIndex(i).sngGRG * 150 * sngLSBiasValue * sngMASValue
                                If Not (udtIndex(i).sngSignal >= HIGHSIGNAL) Then
                                    udtIndex(i).sngSignal = 0
                                Else
                                    blnHS2Appear = True
                                End If
                            End If
                        End If
                        
                        If blnHS2Appear And Not bln2ndAHSAppear Then
                            ' 當日收盤價貫5MA and 10MA and 出現死亡交叉
                            If (udtStock(i).sngEndprice < udtIndex(i).sngP5 _
                                And udtStock(i).sngEndprice < udtIndex(i).sngP10) _
                                And (udtIndex(i - 1).sngP5 >= udtIndex(i - 1).sngP10 _
                                And udtIndex(i).sngP5 <= udtIndex(i).sngP10) Then
                                udtIndex(i).sngSignal = 140
                                bln2ndAHSAppear = True
                            End If
                            ' 當日收盤價貫5MA and 10MA
                            If udtStock(i).sngEndprice < udtIndex(i).sngP5 _
                                And udtStock(i).sngEndprice < udtIndex(i).sngP10 Then
                                udtIndex(i).sngSignal = 140
                                bln2ndAHSAppear = True
                            End If
                            ' 出現死亡交叉
                            If udtIndex(i - 1).sngP5 >= udtIndex(i - 1).sngP10 _
                                And udtIndex(i).sngP5 <= udtIndex(i).sngP10 Then
                                udtIndex(i).sngSignal = 140
                                bln2ndAHSAppear = True
                            End If
                        End If
                    End If
                End If

                ' 收盤價貫穿60MAP停利、停損點 且 連3天收黑K
                If (udtIndex(i).sngP60 > udtStock(i).sngEndprice _
                    And udtIndex(i - 1).sngP60 < udtStock(i - 1).sngEndprice) _
                    Or (udtIndex(i).sngP60 > udtStock(i).sngEndprice _
                    And udtIndex(i - 2).sngP60 < udtStock(i - 2).sngEndprice) Then
                    If udtStock(i - 2).sngStartprice > udtStock(i - 2).sngEndprice _
                        And udtStock(i - 1).sngStartprice > udtStock(i - 1).sngEndprice _
                        And udtStock(i).sngStartprice > udtStock(i).sngEndprice Then
                        If udtIndex(i).sngUpDownDays = -1 Then
                            udtIndex(i).sngSignal = 150
                            bln1stAHSAppear = True
                        End If
                    End If
                End If
                
                ' 當日收盤價貫穿60MAP時，如果接下來的日子收盤價低於60MAP達100點，則顯示AHS
                If udtStock(i).sngEndprice < udtIndex(i).sngP60 _
                    And udtStock(i - 1).sngEndprice > udtIndex(i - 1).sngP60 Then
                    blnAHSCross = True
                End If
                If blnAHSCross Then
                    If udtStock(i).sngEndprice > udtIndex(i).sngP60 Then
                        blnAHSCross = False
                    End If
                End If
                If blnAHSCross And Abs(udtIndex(i).sngMAPDis) > 100 Then
                    udtIndex(i).sngSignal = 150
                    bln1stAHSAppear = True
                    blnAHSCross = False
                End If
            ElseIf udtIndex(i).sngCyOp < 0 Then
                '*** Initialize Variables ***
                bln1stAHSAppear = False
                bln2ndAHSAppear = False
                blnHSAppear = False
                blnHS2Appear = False
                intHSAppearDays = 0
                blnAHSCross = False

                '*** 產生低點訊號 ***
                If udtIndex(i).sngGRG <= GRG_LOW And udtIndex(i).sngGRG > 0 Then
                    ' Bias Fuzzy Rules
                    If udtIndex(i).sngLSBias <= LSBIASL3 Then
                        sngLSBiasValue = 1
                    ElseIf udtIndex(i).sngLSBias <= LSBIASL2 Then
                        sngLSBiasValue = (udtIndex(i).sngLSBias - LSBIASL2) / (LSBIASL3 - LSBIASL2) * (1 - 0.9) + 0.9
                    ElseIf udtIndex(i).sngLSBias <= LSBIASL1 Then
                        sngLSBiasValue = (udtIndex(i).sngLSBias - LSBIASL1) / (LSBIASL3 - LSBIASL1) * (0.9 - 0.8) + 0.8
                    Else
                        sngLSBiasValue = 0
                    End If
                    ' MAPS Fuzzy Rules
                    If udtIndex(i).sngMASlope <= MASL5 Then
                        sngMASValue = 1
                    ElseIf udtIndex(i).sngMASlope <= MASL4 Then
                        sngMASValue = (udtIndex(i).sngMASlope - MASL4) / (MASL5 - MASL4) * (1 - 0.9) + 0.9
                    ElseIf udtIndex(i).sngMASlope <= MASL3 Then
                        sngMASValue = (udtIndex(i).sngMASlope - MASL3) / (MASH4 - MASL3) * (0.9 - 0.8) + 0.8
                    ElseIf udtIndex(i).sngMASlope <= MASL2 Then
                        sngMASValue = (udtIndex(i).sngMASlope - MASL2) / (MASH3 - MASL2) * (0.8 - 0.7) + 0.7
                    ElseIf udtIndex(i).sngMASlope <= MASL1 Then
                        sngMASValue = (udtIndex(i).sngMASlope - MASL1) / (MASH2 - MASL1) * (0.7 - 0.6) + 0.6
                    Else
                        sngMASValue = 0
                    End If
                    ' Calculate The Signal Value
                    If sngLSBiasValue * sngMASValue > 0 Then
                        udtIndex(i).sngSignal = udtIndex(i).sngGRG * 150 * (1 - (sngLSBiasValue * sngMASValue)) + 10
                        ' Signal符合條件，則表示本區間可找到低點
                        If Not (udtIndex(i).sngSignal >= LOWSIGNAL) Then
                            udtIndex(i).sngSignal = 10
                        End If
                    End If
                End If

                ' 當日收盤價上穿60MAP時，如果接下來的日子收盤價高於60MAP達200點或是上揚超過5天，則顯示LHS
                If udtStock(i).sngEndprice > udtIndex(i).sngP60 _
                    And udtStock(i - 1).sngEndprice < udtIndex(i - 1).sngP60 Then
                    blnALSCross = True
                End If
                If blnALSCross Then
                    If udtStock(i).sngEndprice < udtIndex(i).sngP60 Then
                        blnALSCross = False
                    End If
                End If
                If blnALSCross And (Abs(udtIndex(i).sngMAPDis) > 200 Or (udtIndex(i).sngUpDownDays > 5 And Abs(udtIndex(i).sngMAPDis) > 150)) Then
                    udtIndex(i).sngSignal = 5
                    blnALSCross = False
                End If
            End If
        End If
    Next


    bln1stAHSAppear = False
    Dim blnFind2ndAHS As Boolean
    blnFind2ndAHS = False
    Dim intHSDay As Integer
    For i = 1 To intStockNo
        If udtIndex(i).sngSignal > 10 And udtIndex(i).sngSignal < 140 Then
            udtIndex(i).sngSignal = 0
        End If
        ' 如果訊號=150 但有盤整狀況，則取消該訊號
        If udtIndex(i).sngSignal = 150 And udtIndex(i).sngInRange > 0 Then
            udtIndex(i).sngSignal = 0
        End If
        '出現ALS訊號，至少前一天也要是GRG低點前算
        If udtIndex(i).sngSignal = 10 Then
'            If Not (udtIndex(i - 1).sngGRG > 0 And udtIndex(i - 1).sngGRG <= frmGeryExp.txtGRGThreshold2) Then
            If Not (udtIndex(i - 1).sngGRG > 0 And udtIndex(i - 1).sngGRG <= GRG_LOW) Then
                udtIndex(i).sngSignal = 0
            End If
        End If

        ' *** 60-10Bias乖離大於LSBIASH3時，第2個AHS尋找方式 ***
        ' 當60-10Bias值>=LSBIASH3，由觀察知趨勢將強勢向上，此時第2個AHS的產生方式如下:
        ' Cond1: 60-10Bias>=LSBIASH3 且出現第1個AHS
        ' Cond2: 60MAS<MASH1
        ' Cond3: HS appear and 5日內收盤價要貫穿5MAP, 10MAP 且當日收黑K且開盤價與收盤價的差>=10點
        If udtIndex(i).sngLSBias >= LSBIASH3 Then
            blnLSBiasH3 = True
        End If
        If udtIndex(i).sngCyOp < 0 Then
            blnLSBiasH3 = False
            blnFind2ndAHS = False
        End If
        ' 60-10 Bias乖離率大過設定範圍，未達第2個AHS的標準時，若再有聚出AHS，一律清除
        If blnLSBiasH3 And bln1stAHSAppear Then
            If udtIndex(i).sngSignal > 0 And udtIndex(i).sngSignal <> 150 Then
                udtIndex(i).sngSignal = 0
            End If
        End If
        ' 第1次出現AHS
        If udtIndex(i).sngSignal >= 140 Then
            bln1stAHSAppear = True
        End If
        If udtIndex(i).sngCyOp < 0 Then
            bln1stAHSAppear = False
        End If
        ' 符合60-10Bias 及第1個AHS出現後，開始判斷是否符合Cond2
        If blnLSBiasH3 And bln1stAHSAppear Then
            If udtIndex(i).sngMASlope < MASH1 Then
                blnFind2ndAHS = True
            End If
        Else
            blnFind2ndAHS = False
        End If
        ' 符合Cond1 and Cond2 開始判斷是否符合 Cond3，若符合即表示找到第2個AHS
        If blnFind2ndAHS Then
'            If udtIndex(i).sngGRG >= frmGeryExp.txtGRGThreshold Then
            If udtIndex(i).sngGRG >= GRG_HIGH Then
                blnHSAppear = True
                intHSDay = 5
            End If
        End If
        If intHSDay > 0 Then
            intHSDay = intHSDay - 1
            If udtStock(i).sngEndprice < udtIndex(i).sngP5 And udtStock(i).sngEndprice < udtIndex(i).sngP10 _
                And udtStock(i).sngStartprice > udtStock(i).sngEndprice And Abs(udtStock(i).sngStartprice - udtStock(i).sngEndprice) >= 10 Then
                udtIndex(i).sngSignal = 140
                blnFind2ndAHS = False
                intHSDay = 0
            End If
        Else
            blnHSAppear = False
        End If
    Next


    Dim int140AHS As Integer
    Dim int150AHS As Integer
    For i = 1 To intStockNo
        If udtIndex(i).sngCyOp > 0 Then
            If udtIndex(i).sngSignal = 140 Then
                int140AHS = int140AHS + 1
            End If
            If udtIndex(i).sngSignal = 150 Then
                int150AHS = int150AHS + 1
            End If
            If udtIndex(i).sngSignal = 150 Then
                If (int140AHS + int150AHS) > 2 Then
                    udtIndex(i).sngSignal = 0
                End If
            End If
        ElseIf udtIndex(i).sngCyOp < 0 Then
            int140AHS = 0
            int150AHS = 0
        End If
    Next

    Exit Sub

ERR_HANDLE:
    MsgBox "[GCFR_General_Module.DisplaySignals4CY()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    Resume Next
End Sub


'***************************************************************************************************
'* 說    明:
'*    分割出每個Sector -- use CY
'* 輸入參數:
'*    udtStock 每日股價資料(原始資料)
'*    udtIndex 每日股價資料(技術指標)
'*    intStockNo 股價資料筆數
'*    strDayNo 區間內最少天數(小於此值者需合併)
'*    strStkDis 區間內最高與最低點差距值(小於此值者需合併)
'*    indexName1 傳入的MAP字串--起日
'*    indexName2 傳入的MAP字串--迄日
'* 輸出參數: 無
'* 版    本:
'*    2.00: 20080911 新增
'* 備    註:
'*    sngSector = 1 表示為高點區段；sngSector = 0.5 表示為低點區段
'***************************************************************************************************
Public Sub SplitSector4CY(ByRef udtStock() As StockData, _
                        ByRef udtIndex() As IndexData, _
                        ByVal intStockNo As Integer)
    Dim sngFlag As Single       ' 值=1表高點、值=0.5表低點
    Dim i As Integer
      
    On Error GoTo ERR_HANDLE
    
    '*** Initialize Variables ***
    sngFlag = 1
    gintSector = 0
    
    For i = 2 To intStockNo
        With udtIndex(i)
            ' 851019 以後才有CY index的資料
            If udtStock(i).sngDate >= LEARN_START_DATE Then
                '--- 記錄該點是屬於高點或低點 ---
                If .sngCyOp >= 0 Then
                    sngFlag = 1     ' 高點
                Else
                    sngFlag = 0.5   ' 低點
                End If
                .sngSector = sngFlag
            End If
        End With
    Next
        
    Exit Sub
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.SplitSector4CY()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
End Sub


'***************************************************************************************************
'* 說    明:
'*    對資料做前處理 取得每一個各別指標的最大值及最小值
'* 輸入參數:
'*    udtGRG 要處理的資料
'*    intCount 要處理的資料的筆數
'*    sngFlag 此次前處理為高點或是低點
'* 輸出參數: 無
'* 版    本:
'*    2.00  20080911 新增
'***************************************************************************************************
Public Sub PreProcess4CY(ByRef udtGRG() As udtGRGIndex, _
                        ByVal intCount As Integer, _
                        ByVal sngFlag As Single)
    Dim i As Integer
    Dim j As Integer

    On Error GoTo ERR_HANDLE
    
    '--- Initialize Variables ---
    For i = 1 To 8
        sngMax(i) = -99999
        sngMin(i) = 99999
    Next
    
    '--------------------------------------------------------------------
    '* Find each Max value and Min value in each parameter
    '--------------------------------------------------------------------
    For i = 1 To intCount
        If sngFlag = 1 Then     ' 高點
            With udtGRG(i)
                If sngMax(1) < .sngParam1 Then
                    sngMax(1) = .sngParam1
                End If
                If sngMin(1) > .sngParam1 Then
                    sngMin(1) = .sngParam1
                End If
                If sngMax(2) < .sngParam2 Then
                    sngMax(2) = .sngParam2
                End If
                If sngMin(2) > .sngParam2 Then
                    sngMin(2) = .sngParam2
                End If
                If sngMax(3) < .sngParam3 Then
                    sngMax(3) = .sngParam3
                End If
                If sngMin(3) > .sngParam3 Then
                    sngMin(3) = .sngParam3
                End If
                If sngMax(4) < .sngParam4 Then
                    sngMax(4) = .sngParam4
                End If
                If sngMin(4) > .sngParam4 Then
                    sngMin(4) = .sngParam4
                End If
                If sngMax(5) < .sngParam5 Then
                    sngMax(5) = .sngParam5
                End If
                If sngMin(5) > .sngParam5 Then
                    sngMin(5) = .sngParam5
                End If
                If sngMax(6) < .sngParam6 Then
                    sngMax(6) = .sngParam6
                End If
                If sngMin(6) > .sngParam6 Then
                    sngMin(6) = .sngParam6
                End If
            End With
        Else        ' 低點
            With udtGRG(i)
                If sngMax(1) < .sngParam1 Then
                    sngMax(1) = .sngParam1
                End If
                If sngMin(1) > .sngParam1 Then
                    sngMin(1) = .sngParam1
                End If
                If sngMax(2) < .sngParam2 Then
                    sngMax(2) = .sngParam2
                End If
                If sngMin(2) > .sngParam2 Then
                    sngMin(2) = .sngParam2
                End If
                If sngMax(3) < .sngParam3 Then
                    sngMax(3) = .sngParam3
                End If
                If sngMin(3) > .sngParam3 Then
                    sngMin(3) = .sngParam3
                End If
                If sngMax(4) < .sngParam4 Then
                    sngMax(4) = .sngParam4
                End If
                If sngMin(4) > .sngParam4 Then
                    sngMin(4) = .sngParam4
                End If
                If sngMax(5) < .sngParam5 Then
                    sngMax(5) = .sngParam5
                End If
                If sngMin(5) > .sngParam5 Then
                    sngMin(5) = .sngParam5
                End If
                If sngMax(6) < .sngParam6 Then
                    sngMax(6) = .sngParam6
                End If
                If sngMin(6) > .sngParam6 Then
                    sngMin(6) = .sngParam6
                End If
            End With
        End If
    Next
    
    '--------------------------------------------------------------------
    '* 初值化
    '*
    '* Note: gintSelFactorsMethod()    陣列 start with 1
    '*       gintSelLowFactorsMethod() 陣列 start with 1
    '*
    '* gintSelAttributesMethod = 1 表「望大」
    '* gintSelAttributesMethod = 2 表「望小」
    '* gintSelAttributesMethod = 3 表「望目」
    '--------------------------------------------------------------------
    For i = 1 To intCount
        If sngFlag = 1 Then     ' 高點
            With udtGRG(i)
                If UBound(gintSelFactorsMethod) >= 1 Then
                    If gintSelFactorsMethod(1) = 1 And UBound(gintSelFactorsMethod) >= 1 Then
                        If (sngMax(1) - sngMin(1)) <> 0 Then
                            .sngParam1 = (.sngParam1 - sngMin(1)) / (sngMax(1) - sngMin(1))
                        End If
                    ElseIf gintSelFactorsMethod(1) = 2 Then
                        If (sngMax(1) - sngMin(1)) <> 0 Then
                            .sngParam1 = (sngMax(1) - .sngParam1) / (sngMax(1) - sngMin(1))
                        End If
                    ElseIf gintSelFactorsMethod(1) = 3 Then
                        If (sngMax(1) - udtGRG(intCount).sngParam1) <> 0 Or _
                            (udtGRG(intCount).sngParam1 - sngMin(1)) <> 0 Then
                            .sngParam1 = 1 - Abs(.sngParam1 - udtGRG(intCount).sngParam1) / (IIf((sngMax(1) - udtGRG(intCount).sngParam1 > udtGRG(intCount).sngParam1 - sngMin(1)), sngMax(1) - udtGRG(intCount).sngParam1, udtGRG(intCount).sngParam1 - sngMin(1)))
                        End If
                    End If
                End If
                
                If UBound(gintSelFactorsMethod) >= 2 Then
                    If gintSelFactorsMethod(2) = 1 And UBound(gintSelFactorsMethod) >= 2 Then
                        If (sngMax(2) - sngMin(2)) <> 0 Then
                            .sngParam2 = (.sngParam2 - sngMin(2)) / (sngMax(2) - sngMin(2))
                        End If
                    ElseIf gintSelFactorsMethod(2) = 2 Then
                        If (sngMax(2) - sngMin(2)) <> 0 Then
                            .sngParam2 = (sngMax(2) - .sngParam2) / (sngMax(2) - sngMin(2))
                        End If
                    ElseIf gintSelFactorsMethod(2) = 3 Then
                        If (sngMax(2) - udtGRG(intCount).sngParam2) <> 0 Or _
                            (udtGRG(intCount).sngParam2 - sngMin(2)) <> 0 Then
                            .sngParam2 = 1 - Abs(.sngParam2 - udtGRG(intCount).sngParam2) / (IIf((sngMax(2) - udtGRG(intCount).sngParam2 > udtGRG(intCount).sngParam2 - sngMin(2)), sngMax(2) - udtGRG(intCount).sngParam2, udtGRG(intCount).sngParam2 - sngMin(2)))
                        End If
                    End If
                End If
                
                If UBound(gintSelFactorsMethod) >= 3 Then
                    If gintSelFactorsMethod(3) = 1 And UBound(gintSelFactorsMethod) >= 3 Then
                        If (sngMax(3) - sngMin(3)) <> 0 Then
                            .sngParam3 = (.sngParam3 - sngMin(3)) / (sngMax(3) - sngMin(3))
                        End If
                    ElseIf gintSelFactorsMethod(3) = 2 Then
                        If (sngMax(3) - sngMin(3)) <> 0 Then
                            .sngParam3 = (sngMax(3) - .sngParam3) / (sngMax(3) - sngMin(3))
                        End If
                    ElseIf gintSelFactorsMethod(2) = 3 Then
                        If (sngMax(3) - udtGRG(intCount).sngParam3) <> 0 Or _
                            (udtGRG(intCount).sngParam3 - sngMin(3)) <> 0 Then
                            .sngParam3 = 1 - Abs(.sngParam3 - udtGRG(intCount).sngParam3) / (IIf((sngMax(3) - udtGRG(intCount).sngParam3 > udtGRG(intCount).sngParam3 - sngMin(3)), sngMax(3) - udtGRG(intCount).sngParam3, udtGRG(intCount).sngParam3 - sngMin(3)))
                        End If
                    End If
                End If
            
                If UBound(gintSelFactorsMethod) >= 4 Then
                    If gintSelFactorsMethod(4) = 1 And UBound(gintSelFactorsMethod) >= 4 Then
                        If (sngMax(4) - sngMin(4)) <> 0 Then
                            .sngParam4 = (.sngParam4 - sngMin(4)) / (sngMax(4) - sngMin(4))
                        End If
                    ElseIf gintSelFactorsMethod(4) = 2 Then
                        If (sngMax(4) - sngMin(4)) <> 0 Then
                            .sngParam4 = (sngMax(4) - .sngParam4) / (sngMax(4) - sngMin(4))
                        End If
                    ElseIf gintSelFactorsMethod(4) = 3 Then
                        If (sngMax(4) - udtGRG(intCount).sngParam4) <> 0 Or _
                            (udtGRG(intCount).sngParam4 - sngMin(4)) <> 0 Then
                            .sngParam4 = 1 - Abs(.sngParam4 - udtGRG(intCount).sngParam4) / (IIf((sngMax(4) - udtGRG(intCount).sngParam4 > udtGRG(intCount).sngParam4 - sngMin(4)), sngMax(4) - udtGRG(intCount).sngParam4, udtGRG(intCount).sngParam4 - sngMin(4)))
                        End If
                    End If
                End If
            
                If UBound(gintSelFactorsMethod) >= 5 Then
                    If gintSelFactorsMethod(5) = 1 And UBound(gintSelFactorsMethod) >= 5 Then
                        If (sngMax(5) - sngMin(5)) <> 0 Then
                            .sngParam5 = (.sngParam5 - sngMin(5)) / (sngMax(5) - sngMin(5))
                        End If
                    ElseIf gintSelFactorsMethod(5) = 2 Then
                        If (sngMax(5) - sngMin(5)) <> 0 Then
                            .sngParam5 = (sngMax(5) - .sngParam5) / (sngMax(5) - sngMin(5))
                        End If
                    ElseIf gintSelFactorsMethod(5) = 3 Then
                        If (sngMax(5) - udtGRG(intCount).sngParam5) <> 0 Or _
                            (udtGRG(intCount).sngParam5 - sngMin(5)) <> 0 Then
                            .sngParam5 = 1 - Abs(.sngParam5 - udtGRG(intCount).sngParam5) / (IIf((sngMax(5) - udtGRG(intCount).sngParam5 > udtGRG(intCount).sngParam5 - sngMin(5)), sngMax(5) - udtGRG(intCount).sngParam5, udtGRG(intCount).sngParam5 - sngMin(5)))
                        End If
                    End If
                End If
            
                If UBound(gintSelFactorsMethod) >= 6 Then
                    If gintSelFactorsMethod(6) = 1 Then
                        If (sngMax(6) - sngMin(6)) <> 0 Then
                            .sngParam6 = (.sngParam6 - sngMin(6)) / (sngMax(6) - sngMin(6))
                        End If
                    ElseIf gintSelFactorsMethod(6) = 2 Then
                        If (sngMax(6) - sngMin(6)) <> 0 Then
                            .sngParam6 = (sngMax(6) - .sngParam6) / (sngMax(6) - sngMin(6))
                        End If
                    ElseIf gintSelFactorsMethod(6) = 3 Then
                        If (sngMax(6) - udtGRG(intCount).sngParam6) <> 0 Or _
                            (udtGRG(intCount).sngParam6 - sngMin(6)) <> 0 Then
                            .sngParam6 = 1 - Abs(.sngParam6 - udtGRG(intCount).sngParam6) / (IIf((sngMax(6) - udtGRG(intCount).sngParam6 > udtGRG(intCount).sngParam6 - sngMin(6)), sngMax(6) - udtGRG(intCount).sngParam6, udtGRG(intCount).sngParam6 - sngMin(6)))
                        End If
                    End If
                End If
            End With
        Else        ' 低點
            With udtGRG(i)
                If UBound(gintSelLowFactorsMethod) >= 1 Then
                    If gintSelLowFactorsMethod(1) = 1 Then
                        If (sngMax(1) - sngMin(1)) <> 0 Then
                            .sngParam1 = (.sngParam1 - sngMin(1)) / (sngMax(1) - sngMin(1))
                        End If
                    ElseIf gintSelLowFactorsMethod(1) = 2 Then
                        If (sngMax(1) - sngMin(1)) <> 0 Then
                            .sngParam1 = (sngMax(1) - .sngParam1) / (sngMax(1) - sngMin(1))
                        End If
                    ElseIf gintSelLowFactorsMethod(1) = 3 Then
                        If (sngMax(1) - udtGRG(intCount).sngParam1) <> 0 Or _
                            (udtGRG(intCount).sngParam1 - sngMin(1)) <> 0 Then
                            .sngParam1 = 1 - Abs(.sngParam1 - udtGRG(intCount).sngParam1) / (IIf((sngMax(1) - udtGRG(intCount).sngParam1 > udtGRG(intCount).sngParam1 - sngMin(1)), sngMax(1) - udtGRG(intCount).sngParam1, udtGRG(intCount).sngParam1 - sngMin(1)))
                        End If
                    End If
                End If
            
                If UBound(gintSelLowFactorsMethod) >= 2 Then
                    If gintSelLowFactorsMethod(2) = 1 Then
                        If (sngMax(2) - sngMin(2)) <> 0 Then
                            .sngParam2 = (.sngParam2 - sngMin(2)) / (sngMax(2) - sngMin(2))
                        End If
                    ElseIf gintSelLowFactorsMethod(2) = 2 Then
                        If (sngMax(2) - sngMin(2)) <> 0 Then
                            .sngParam2 = (sngMax(2) - .sngParam2) / (sngMax(2) - sngMin(2))
                        End If
                    ElseIf gintSelLowFactorsMethod(2) = 3 Then
                        If (sngMax(2) - udtGRG(intCount).sngParam2) <> 0 Or _
                            (udtGRG(intCount).sngParam2 - sngMin(2)) <> 0 Then
                            .sngParam2 = 1 - Abs(.sngParam2 - udtGRG(intCount).sngParam2) / (IIf((sngMax(2) - udtGRG(intCount).sngParam2 > udtGRG(intCount).sngParam2 - sngMin(2)), sngMax(2) - udtGRG(intCount).sngParam2, udtGRG(intCount).sngParam2 - sngMin(2)))
                        End If
                    End If
                End If
            
                If UBound(gintSelLowFactorsMethod) >= 3 Then
                    If gintSelLowFactorsMethod(3) = 1 Then
                        If (sngMax(3) - sngMin(3)) <> 0 Then
                            .sngParam3 = (.sngParam3 - sngMin(3)) / (sngMax(3) - sngMin(3))
                        End If
                    ElseIf gintSelLowFactorsMethod(3) = 2 Then
                        If (sngMax(3) - sngMin(3)) <> 0 Then
                            .sngParam3 = (sngMax(3) - .sngParam3) / (sngMax(3) - sngMin(3))
                        End If
                    ElseIf gintSelLowFactorsMethod(3) = 3 Then
                        If (sngMax(3) - udtGRG(intCount).sngParam3) <> 0 Or _
                            (udtGRG(intCount).sngParam3 - sngMin(3)) <> 0 Then
                            .sngParam3 = 1 - Abs(.sngParam3 - udtGRG(intCount).sngParam3) / (IIf((sngMax(3) - udtGRG(intCount).sngParam3 > udtGRG(intCount).sngParam3 - sngMin(3)), sngMax(3) - udtGRG(intCount).sngParam3, udtGRG(intCount).sngParam3 - sngMin(3)))
                        End If
                    End If
                End If
            
                If UBound(gintSelLowFactorsMethod) >= 4 Then
                    If gintSelLowFactorsMethod(4) = 1 Then
                        If (sngMax(4) - sngMin(4)) <> 0 Then
                            .sngParam4 = (.sngParam4 - sngMin(4)) / (sngMax(4) - sngMin(4))
                        End If
                    ElseIf gintSelLowFactorsMethod(4) = 2 Then
                        If (sngMax(4) - sngMin(4)) <> 0 Then
                            .sngParam4 = (sngMax(4) - .sngParam4) / (sngMax(4) - sngMin(4))
                        End If
                    ElseIf gintSelLowFactorsMethod(4) = 3 Then
                        If (sngMax(4) - udtGRG(intCount).sngParam4) <> 0 Or _
                            (udtGRG(intCount).sngParam4 - sngMin(4)) <> 0 Then
                            .sngParam4 = 1 - Abs(.sngParam4 - udtGRG(intCount).sngParam4) / (IIf((sngMax(4) - udtGRG(intCount).sngParam4 > udtGRG(intCount).sngParam4 - sngMin(4)), sngMax(4) - udtGRG(intCount).sngParam4, udtGRG(intCount).sngParam4 - sngMin(4)))
                        End If
                    End If
                End If
            
                If UBound(gintSelLowFactorsMethod) >= 5 Then
                    If gintSelLowFactorsMethod(5) = 1 Then
                        If (sngMax(5) - sngMin(5)) <> 0 Then
                            .sngParam5 = (.sngParam5 - sngMin(5)) / (sngMax(5) - sngMin(5))
                        End If
                    ElseIf gintSelLowFactorsMethod(5) = 2 Then
                        If (sngMax(5) - sngMin(5)) <> 0 Then
                            .sngParam5 = (sngMax(5) - .sngParam5) / (sngMax(5) - sngMin(5))
                        End If
                    ElseIf gintSelLowFactorsMethod(5) = 3 Then
                        If (sngMax(5) - udtGRG(intCount).sngParam5) <> 0 Or _
                            (udtGRG(intCount).sngParam5 - sngMin(5)) <> 0 Then
                            .sngParam5 = 1 - Abs(.sngParam5 - udtGRG(intCount).sngParam5) / (IIf((sngMax(5) - udtGRG(intCount).sngParam5 > udtGRG(intCount).sngParam5 - sngMin(5)), sngMax(5) - udtGRG(intCount).sngParam5, udtGRG(intCount).sngParam5 - sngMin(5)))
                        End If
                    End If
                End If
            
                If UBound(gintSelLowFactorsMethod) >= 6 Then
                    If gintSelLowFactorsMethod(6) = 1 Then
                        If (sngMax(6) - sngMin(6)) <> 0 Then
                            .sngParam6 = (.sngParam6 - sngMin(6)) / (sngMax(6) - sngMin(6))
                        End If
                    ElseIf gintSelLowFactorsMethod(6) = 2 Then
                        If (sngMax(6) - sngMin(6)) <> 0 Then
                            .sngParam6 = (sngMax(6) - .sngParam6) / (sngMax(6) - sngMin(6))
                        End If
                    ElseIf gintSelLowFactorsMethod(6) = 3 Then
                        If (sngMax(6) - udtGRG(intCount).sngParam6) <> 0 Or _
                            (udtGRG(intCount).sngParam6 - sngMin(6)) <> 0 Then
                            .sngParam6 = 1 - Abs(.sngParam6 - udtGRG(intCount).sngParam6) / (IIf((sngMax(6) - udtGRG(intCount).sngParam6 > udtGRG(intCount).sngParam6 - sngMin(6)), sngMax(6) - udtGRG(intCount).sngParam6, udtGRG(intCount).sngParam6 - sngMin(6)))
                        End If
                    End If
                End If
            End With
        End If
    Next
    
    Exit Sub
    
    
ERR_HANDLE:
    Select Case Err.Number
        Case 6
            Debug.Print "[GCFR_General_Module.PreProcess4CY()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
        Case Else
            Debug.Print "[GCFR_General_Module.PreProcess4CY()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
    
    Err.Clear
    Resume Next
End Sub



'***************************************************************************************************
'* 說    明:
'*    Get attributes value
'* 輸入參數:
'*    udtIndex 儲存的指數資料
'*    intPos 要讀取資料的位置
'*    intSelAttribute 選擇的attribute
'* 輸出參數:
'*    Single 所選擇之資料之attribute的值
'* 版    本:
'*    2.00  20080911 Modified
'***************************************************************************************************
Public Function GetAttributesValue4CY(ByRef udtIndex() As IndexData, _
                                ByVal intPos As Integer, _
                                ByVal intSelAttribute As Integer, _
                                ByVal sngFlag As Single) As Single
    
    On Error GoTo ERR_HANDLE
    
    If sngFlag = 1 Then
        '--- 大於選擇的attribute數目，均回傳99999 ---
        If intSelAttribute > gintSelCount Then
            GetAttributesValue4CY = 99999
            Exit Function
        End If
    
        If gstrSelFactors(intSelAttribute) = "BIAS" Then
            GetAttributesValue4CY = udtIndex(intPos).sngBias
        ElseIf gstrSelFactors(intSelAttribute) = "MACD" Then
            GetAttributesValue4CY = udtIndex(intPos).sngMACD
        ElseIf gstrSelFactors(intSelAttribute) = "PSY" Then
            GetAttributesValue4CY = udtIndex(intPos).sngPSY
        ElseIf gstrSelFactors(intSelAttribute) = "RSI_L" Then
            GetAttributesValue4CY = udtIndex(intPos).sngRSI_L
        ElseIf gstrSelFactors(intSelAttribute) = "RSI_S" Then
            GetAttributesValue4CY = udtIndex(intPos).sngRSI_S
        ElseIf gstrSelFactors(intSelAttribute) = "WMS" Then
            GetAttributesValue4CY = udtIndex(intPos).sngWMS
        ElseIf gstrSelFactors(intSelAttribute) = "K" Then
            GetAttributesValue4CY = udtIndex(intPos).sngK
        ElseIf gstrSelFactors(intSelAttribute) = "D" Then
            GetAttributesValue4CY = udtIndex(intPos).sngD
'        ElseIf gstrSelFactors(intSelAttribute) = "KDDis" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngKDDis
        ElseIf gstrSelFactors(intSelAttribute) = "MAPDis" Then
            GetAttributesValue4CY = udtIndex(intPos).sngMAPDis
'        ElseIf gstrSelFactors(intSelAttribute) = "RSIDis" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngRSIDis
        ElseIf gstrSelFactors(intSelAttribute) = "MAP3" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP3
        ElseIf gstrSelFactors(intSelAttribute) = "MAP4" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP4
        ElseIf gstrSelFactors(intSelAttribute) = "MAP5" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP5
        ElseIf gstrSelFactors(intSelAttribute) = "MAP6" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP6
        ElseIf gstrSelFactors(intSelAttribute) = "MAP8" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP8
        ElseIf gstrSelFactors(intSelAttribute) = "MAP10" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP10
        ElseIf gstrSelFactors(intSelAttribute) = "MAP12" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP12
        ElseIf gstrSelFactors(intSelAttribute) = "MAP24" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP24
        ElseIf gstrSelFactors(intSelAttribute) = "MAP30" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP30
        ElseIf gstrSelFactors(intSelAttribute) = "MAP72" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP72
        ElseIf gstrSelFactors(intSelAttribute) = "MAP144" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP144
        ElseIf gstrSelFactors(intSelAttribute) = "MAP288" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP288
        ElseIf gstrSelFactors(intSelAttribute) = "DIF" Then
            GetAttributesValue4CY = udtIndex(intPos).sngDIF
        ElseIf gstrSelFactors(intSelAttribute) = "DIF_MACD" Then
            GetAttributesValue4CY = udtIndex(intPos).sngDIF_MACD
'        ElseIf gstrSelFactors(intSelAttribute) = "TAPI" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngTAPI
'        ElseIf gstrSelFactors(intSelAttribute) = "OBV" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngOBV
        ElseIf gstrSelFactors(intSelAttribute) = "VR" Then
            GetAttributesValue4CY = udtIndex(intPos).sngVR
        ElseIf gstrSelFactors(intSelAttribute) = "CY" Then
            GetAttributesValue4CY = Str(Round(udtIndex(intPos).sngDIF_MACD, 2))
'        ElseIf gstrSelFactors(intSelAttribute) = "AR" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngAR
'        ElseIf gstrSelFactors(intSelAttribute) = "BR" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngBR
        ElseIf gstrSelFactors(intSelAttribute) = "VOL24" Then
            GetAttributesValue4CY = udtIndex(intPos).sngVol24
        ElseIf gstrSelFactors(intSelAttribute) = "CY" Then
            GetAttributesValue4CY = udtIndex(intPos).sngCyOp
'        ElseIf gstrSelFactors(intSelAttribute) = "CYS" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngCYS
        Else
            Err.Raise 10002
        End If
    Else
        '--- 大於選擇的attribute數目，均回傳99999 ---
        If intSelAttribute > gintSelLowCount Then
            GetAttributesValue4CY = 99999
            Exit Function
        End If
    
        If gstrSelLowFactors(intSelAttribute) = "BIAS" Then
            GetAttributesValue4CY = udtIndex(intPos).sngBias
        ElseIf gstrSelLowFactors(intSelAttribute) = "MACD" Then
            GetAttributesValue4CY = udtIndex(intPos).sngMACD
        ElseIf gstrSelLowFactors(intSelAttribute) = "PSY" Then
            GetAttributesValue4CY = udtIndex(intPos).sngPSY
        ElseIf gstrSelLowFactors(intSelAttribute) = "RSI_L" Then
            GetAttributesValue4CY = udtIndex(intPos).sngRSI_L
        ElseIf gstrSelLowFactors(intSelAttribute) = "RSI_S" Then
            GetAttributesValue4CY = udtIndex(intPos).sngRSI_S
        ElseIf gstrSelLowFactors(intSelAttribute) = "WMS" Then
            GetAttributesValue4CY = udtIndex(intPos).sngWMS
        ElseIf gstrSelLowFactors(intSelAttribute) = "K" Then
            GetAttributesValue4CY = udtIndex(intPos).sngK
        ElseIf gstrSelLowFactors(intSelAttribute) = "D" Then
            GetAttributesValue4CY = udtIndex(intPos).sngD
'        ElseIf gstrSelLowFactors(intSelAttribute) = "KDDis" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngKDDis
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAPDis" Then
            GetAttributesValue4CY = udtIndex(intPos).sngMAPDis
'        ElseIf gstrSelLowFactors(intSelAttribute) = "RSIDis" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngRSIDis
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP3" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP3
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP4" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP4
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP5" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP5
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP6" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP6
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP8" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP8
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP10" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP10
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP12" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP12
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP24" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP24
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP30" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP30
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP72" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP72
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP144" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP144
        ElseIf gstrSelLowFactors(intSelAttribute) = "MAP288" Then
            GetAttributesValue4CY = udtIndex(intPos).sngP288
        ElseIf gstrSelLowFactors(intSelAttribute) = "DIF" Then
            GetAttributesValue4CY = udtIndex(intPos).sngDIF
        ElseIf gstrSelLowFactors(intSelAttribute) = "DIF_MACD" Then
            GetAttributesValue4CY = udtIndex(intPos).sngDIF_MACD
'        ElseIf gstrSelLowFactors(intSelAttribute) = "TAPI" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngTAPI
'        ElseIf gstrSelLowFactors(intSelAttribute) = "OBV" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngOBV
        ElseIf gstrSelLowFactors(intSelAttribute) = "VR" Then
            GetAttributesValue4CY = udtIndex(intPos).sngVR
        ElseIf gstrSelLowFactors(intSelAttribute) = "CY" Then
            GetAttributesValue4CY = Str(Round(udtIndex(intPos).sngDIF_MACD, 2))
'        ElseIf gstrSelLowFactors(intSelAttribute) = "AR" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngAR
'        ElseIf gstrSelLowFactors(intSelAttribute) = "BR" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngBR
        ElseIf gstrSelLowFactors(intSelAttribute) = "VOL24" Then
            GetAttributesValue4CY = udtIndex(intPos).sngVol24
        ElseIf gstrSelLowFactors(intSelAttribute) = "CY" Then
            GetAttributesValue4CY = udtIndex(intPos).sngCyOp
'        ElseIf gstrSelLowFactors(intSelAttribute) = "CYS" Then
'            GetAttributesValue4CY = udtIndex(intPos).sngCYS
        Else
            Err.Raise 10003
        End If
    End If
    
    Exit Function
    
    
ERR_HANDLE:
    MsgBox "[GCFR_General_Module.GetAttributesValue4CY()] -- " & Err.Number & ":" & Err.Description & " -- 找不到對應不到的指標", vbOKOnly
'    Err.Clear
    Resume Next
'    Exit Function
End Function


'***************************************************************************************************
'* 說    明:
'*    還原聚類出來的值
'* 輸入參數:
'*    udtClusterDat 聚類結果的重心
'*    intPos 該資料的位置
'*    intPos 該資料的位置
'* 輸出參數: 無
'* 版    本:
'*    2.00  20080911 Modified
'***************************************************************************************************
Public Sub RestoreData4CY(ByRef udtClusterDat() As udtGRGIndex, _
                            ByVal intPos As Integer, _
                            ByVal sngFlag)
    On Error GoTo ERR_HANDLE
    
    If sngFlag = 1 Then
        With udtClusterDat(intPos)
            .sngParam1 = (sngMax(1) - sngMin(1)) * .sngParam1 + sngMin(1)
            .sngParam2 = (sngMax(2) - sngMin(2)) * .sngParam2 + sngMin(2)
            .sngParam3 = (sngMax(3) - sngMin(3)) * .sngParam3 + sngMin(3)
            .sngParam4 = (sngMax(4) - sngMin(4)) * .sngParam4 + sngMin(4)
            .sngParam5 = (sngMax(5) - sngMin(5)) * .sngParam5 + sngMin(5)
            .sngParam6 = (sngMax(6) - sngMin(6)) * .sngParam6 + sngMin(6)
        End With
    Else
        With udtClusterDat(intPos)
            .sngParam1 = (sngMax(1) - sngMin(1)) * .sngParam1 + sngMin(1)
            .sngParam2 = (sngMax(2) - sngMin(2)) * .sngParam2 + sngMin(2)
            .sngParam3 = (sngMax(3) - sngMin(3)) * .sngParam3 + sngMin(3)
            .sngParam4 = (sngMax(4) - sngMin(4)) * .sngParam4 + sngMin(4)
            .sngParam5 = (sngMax(5) - sngMin(5)) * .sngParam5 + sngMin(5)
            .sngParam6 = (sngMax(6) - sngMin(6)) * .sngParam6 + sngMin(6)
        End With
    End If

    Exit Sub
    
    
ERR_HANDLE:
    Select Case Err.Number
        Case Else
            MsgBox "[GCFR_General_Module.RestoreData4CY()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    End Select
    
    Resume Next
End Sub


'***************************************************************************************************
'* 說    明:
'*    產生Testing Sets' GRG values
'* 輸入參數:
'*    udtStock 每日股價資料
'*    udtIndex 儲存的指數資料
'*    intStockNo 資料筆數
'*    strStartDT 傳入的資料起始日期
'*    strEndDT 傳入的資料結束日期
'* 輸出參數: 無
'* 版    本:
'*    2.00: 20080913 Earvin   New
'***************************************************************************************************
Public Sub GenearteTestingSetGRG(ByRef udtStock() As StockData, _
                            ByRef udtIndex() As IndexData, _
                            ByVal intStockNo As Integer, _
                            ByVal strStartDT As String, _
                            ByVal strEndDT As String)
    Dim intCurrentPos As Integer
                            
    On Error GoTo ERR_HANDLE
    
    '---------------------------------------------------
    '* 分2次做，以結省重複運算的時間 -- CY-Index > 0
    '---------------------------------------------------
    PSY_No = 13
    WMS_No = 9
    KD_No = 9
    MACD_No = 24
    EMA_S = 12
    EMA_L = 26
    RSI_S = 6
    RSI_L = 12
    Bias_No = 10
    VR_NO = 12
    OBV_No = 12
    TAPI_No = 12
    AR_No = 12
    
    Call subCalculateIndex(gudtStockDay, gudtIndexDay, gintDayIndex)
    Call subCalculateIndex(gudtStockWeek, gudtIndexWeek, gintWeekIndex)
    Call subCalculateIndex(gudtStockMonth, gudtIndexMonth, gintMonthIndex)
    Call frmEarvinStocks.DrawStockForm(gsngEndIndex, GetStockName(frmEarvinStocks.cboStocks.Text))
    
    intCurrentPos = 2
    While intCurrentPos < intStockNo
        If udtStock(intCurrentPos).sngDate >= strStartDT And udtStock(intCurrentPos).sngDate <= strEndDT Then
            If udtIndex(intCurrentPos).sngCyOp >= 0 Then
                udtIndex(intCurrentPos).sngGRG = CalculateTestingSetGRG4CY(udtStock, udtIndex, intCurrentPos, intStockNo, 1)
            End If
        End If
        intCurrentPos = intCurrentPos + 1
    Wend

    '---------------------------------------------------
    '* 分2次做，以結省重複運算的時間 -- CY-Index < 0
    '---------------------------------------------------
    PSY_No = 13
    WMS_No = 9
    KD_No = 9
    MACD_No = 24
    EMA_S = 12
    EMA_L = 26
    RSI_S = 6
    RSI_L = 6
    Bias_No = 24
    VR_NO = 12
    OBV_No = 12
    TAPI_No = 12
    AR_No = 12
                
    Call subCalculateIndex(gudtStockDay, gudtIndexDay, gintDayIndex)
    Call subCalculateIndex(gudtStockWeek, gudtIndexWeek, gintWeekIndex)
    Call subCalculateIndex(gudtStockMonth, gudtIndexMonth, gintMonthIndex)
    Call frmEarvinStocks.DrawStockForm(gsngEndIndex, GetStockName(frmEarvinStocks.cboStocks.Text))
    
    intCurrentPos = 2
    While intCurrentPos < intStockNo
        If udtStock(intCurrentPos).sngDate >= strStartDT And udtStock(intCurrentPos).sngDate <= strEndDT Then
            If udtIndex(intCurrentPos).sngCyOp < 0 Then
                udtIndex(intCurrentPos).sngGRG = CalculateTestingSetGRG4CY(udtStock, udtIndex, intCurrentPos, intStockNo, 2)
            End If
        End If
        intCurrentPos = intCurrentPos + 1
    Wend

    Exit Sub
    
    
ERR_HANDLE:
    Debug.Print "[GCFR_General_Module.GenearteTestingSetGRG()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
End Sub



''''***************************************************************************************************
''''* 說    明:
''''*    程式一啟動自動聚類並執行POS策略
''''* 輸入參數:
''''*    udtStock 每日股價資料
''''*    udtIndex 儲存的指數資料
''''*    intIndexCnt 資料筆數
''''*    strSTA 欲使用之策略(POS, SSOS, EOS, SOS)
''''*    intKind 1表示產生日資料；2表示產生週資料；3表示產生月資料
''''* 輸出參數: 無
''''* 版    本:
''''*    2.00: 20080913 Earvin   New
''''***************************************************************************************************
'''Public Sub GenSignals(ByRef udtStock() As Stockdata, _
'''                    ByRef udtIndex() As IndexData, _
'''                    ByVal intIndexCnt As Integer, _
'''                    ByVal strSTA As String, _
'''                    ByVal intKind As Integer)
'''    Dim i As Integer, j As Integer
'''
'''    On Error GoTo ERR_HANDLE
'''
'''    '*** Initialize Variables ***
'''    gintSelCount = 0                ' 記錄高點一共用了幾個attributes做聚類
'''    gintSelLowCount = 0             ' 記錄低點一共用了幾個attributes做聚類
'''    gintClusterCnt = 0              ' 記錄成功了多少的聚類pattern
'''    gintSum = 0
'''
'''    If intKind = 1 Then     ' 日
'''        If Not blnClusterDayOK Then
'''            Call ExecGRG(gudtStockDay, gudtIndexDay, gintDayIndex)  ' 聚類
'''            blnClusterDayOK = True
'''        End If
''''''''        If strSTA = "POS" Then
''''''''            Call CalculateProfits4CY1(gudtStockDay, gudtIndexDay, gintDayIndex) ' POS
''''''''            Call GenSTA1Report2
''''''''        ElseIf strSTA = "SSOS" Then
''''''''            Call CalculateProfits4CY2(gudtStockDay, gudtIndexDay, gintDayIndex) ' SSOS
''''''''            Call GenSTA2Report2
''''''''        ElseIf strSTA = "EOS" Then
''''''''            Call CalculateProfits4CY3NGLA(gudtStockDay, gudtIndexDay, gintDayIndex) ' EOS
''''''''            Call GenSTA3Report2
''''''''        ElseIf strSTA = "SOS" Then
''''''''            Call CalculateProfits4CY4(gudtStockDay, gudtIndexDay, gintDayIndex) ' SOS
''''''''            Call GenSTA4Report2
''''''''        End If
'''    ElseIf intKind = 2 Then     ' 週
'''        If Not blnClusterWeekOK Then
'''            Call ExecGRG(gudtStockWeek, gudtIndexWeek, gintWeekIndex)
'''            blnClusterWeekOK = True
'''        End If
''''''''        Call CalculateProfits4CY1(gudtStockWeek, gudtIndexWeek, gintWeekIndex)
'''    ElseIf intKind = 3 Then     ' 月
'''        If Not blnClusterMonthOK Then
'''            Call ExecGRG(gudtStockMonth, gudtIndexMonth, gintMonthIndex)
'''            blnClusterMonthOK = True
'''        End If
''''''''        Call CalculateProfits4CY1(gudtStockMonth, gudtIndexMonth, gintMonthIndex)
'''    End If
'''
'''    Exit Sub
'''
'''
'''ERR_HANDLE:
'''    MsgBox "[Method: GenSignals()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
'''End Sub

                                                                                    


'***************************************************************************************************
'* 說    明:
'*    取得最新一筆資料的日期
'* 輸入參數:
'*    udtStock 股市資料
'*    intStockNo 股市資料總筆數
'* 輸出參數: 無
'* 版    本:
'*    2.00  20050604 Earvin 新增
'***************************************************************************************************
Public Function GetTheLatestDate(ByRef udtStock() As StockData, _
                                    ByVal intStockNo As Integer) As String
    GetTheLatestDate = udtStock(intStockNo).sngDate
End Function



'***************************************************************************************************
'* 說    明:
'*    將6項指標寫入記錄技術指標的陣列中
'* 輸入參數:
'*    udt6IndexData:
'*    intCount     :
'*    udtStock     :
'*    udtIndex     :
'*    intStockNo   :
'* 輸出參數: 無
'* 版    本:
'*    2.00 20050604  Earvin   New
'***************************************************************************************************
Public Sub Write6Index(ByRef udt6IndexData() As SixStockData, _
                        ByVal intCount As Integer, _
                        ByRef udtStock() As StockData, _
                        ByRef udtIndex() As IndexData, _
                        ByVal intStockNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    '--- 平滑化CY index使用變數 ---
    Dim sngValue As Single      ' 儲存要遞增或遞減的值
    Dim dtTheDate As Date       ' 日期 : 起始日期
    Dim sngDate As Single       ' 日期 : 結束日期
    Dim intWeekDay As Integer   ' 記錄該日期為星期幾
    
    On Error GoTo ERR_HANDLE

    '--- Initialize Variables ---
    j = 1
    
    For i = 1 To intStockNo
        If j <= intCount Then
            If udt6IndexData(j).sngDate = udtStock(i).sngDate Then
                With udt6IndexData(j)
                    udtIndex(i).sngDiffOp1 = .sngDiffOp1
                    udtIndex(i).sngDiffOp2 = .sngDiffOp2
                    udtIndex(i).sngQMOp = .sngQMOp
                    udtIndex(i).sngTrendOp = .sngTrendOp
                    udtIndex(i).sngCyOp = .sngCyOp
                    udtIndex(i).sngLstOp = .sngLstOp
                    udtIndex(i).sngLwOp = .sngLwOp
                    
                    j = j + 1
                End With
            ElseIf udt6IndexData(j).sngDate > udtStock(i).sngDate Then
                ' Do nothing
            Else
                j = j + 1
            End If
        Else
            With udtIndex(i)
                .sngDiffOp1 = 0
                .sngDiffOp2 = 0
                .sngQMOp = 0
                .sngTrendOp = 0
                .sngCyOp = 0
                .sngLstOp = 0
                .sngLwOp = 0
            End With
        End If
    Next
    
    Exit Sub
    
ERR_HANDLE:
    Debug.Print "[GCFR_General_Module.Write6Index()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    Resume Next
End Sub

'***************************************************************************************************
'* 說    明: 取得股市高點日期
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'* 輸出參數: 無
'* 版    本: 1.00 20100610 新增
'***************************************************************************************************
Public Sub subGenHighPointDate(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim intUp As Integer
    Dim sngPSY As Single
    
    On Error GoTo ERR_HANDLE
    
    j = 2
    
    While j <= intStockNo
        ' 分開算
'        If KDGreaterTheValueAndCrossDown(udtIndex(j - 1).sngK, udtIndex(j - 1).sngD, udtIndex(j).sngK, udtIndex(j).sngD, theValue) Then
'            udtIndex(j).sngSignal = udtIndex(j).sngSignal + 30
'        End If
'        If MACDGreaterTheValueAndCrossDown(udtIndex(j - 1).sngMACD, udtIndex(j - 1).sngDIF, udtIndex(j).sngMACD, udtIndex(j).sngDIF, 0) Then
'            udtIndex(j).sngSignal = udtIndex(j).sngSignal + 30
'        End If
'        If BIASGreaterTheValue(udtIndex(j).sngBias, -10) Then
'            udtIndex(j).sngSignal = udtIndex(j).sngSignal + 30
'        End If
        
        ' 合在一起算
        If (KDGreaterTheValueAndCrossDown(udtIndex(j - 1).sngK, udtIndex(j - 1).sngD, udtIndex(j).sngK, udtIndex(j).sngD, 30) Or _
            MACDGreaterTheValueAndCrossDown(udtIndex(j - 1).sngMACD, udtIndex(j - 1).sngDIF, udtIndex(j).sngMACD, udtIndex(j).sngDIF, 0)) And _
            BIASGreaterTheValue(udtIndex(j).sngBias, 0) Then
            udtIndex(j).sngSignal = udtIndex(j).sngSignal + 130
        End If
   
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    Debug.Print "[GCFR_General_Module.subGenHighPointDate()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: 取得股市低點日期
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'* 輸出參數: 無
'* 版    本: 1.00 20100610 新增
'***************************************************************************************************
Public Sub subGenLowPointDate(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer)
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ERR_HANDLE
    
    j = 2
   
    ' 高、低點訊號值請參考 Form : frmGeryExp頁面設定
    ' KD值 20以下 且 K 交叉向上 D
    ' MACD < 0 且 MACD 交叉向上 DIF
    ' BIAS < 0
    While j <= intStockNo
        ' 分開算
'        If KDUnderTheValueAndCrossUp(udtIndex(j - 1).sngK, udtIndex(j - 1).sngD, udtIndex(j).sngK, udtIndex(j).sngD, theValue) Then
'            udtIndex(j).sngSignal = udtIndex(j).sngSignal + 30
'        End If
'        If MACDUnderTheValueAndCrossUp(udtIndex(j - 1).sngMACD, udtIndex(j - 1).sngDIF, udtIndex(j).sngMACD, udtIndex(j).sngDIF, 0) Then
'            udtIndex(j).sngSignal = udtIndex(j).sngSignal + 30
'        End If
'        If BIASUnderTheValue(udtIndex(j).sngBias, -10) Then
'            udtIndex(j).sngSignal = udtIndex(j).sngSignal + 30
'        End If
        
        ' 合在一起算
        If (KDUnderTheValueAndCrossUp(udtIndex(j - 1).sngK, udtIndex(j - 1).sngD, udtIndex(j).sngK, udtIndex(j).sngD, 30) Or _
            MACDUnderTheValueAndCrossUp(udtIndex(j - 1).sngMACD, udtIndex(j - 1).sngDIF, udtIndex(j).sngMACD, udtIndex(j).sngDIF, 0)) And _
            BIASUnderTheValue(udtIndex(j).sngBias, 0) Then
            udtIndex(j).sngSignal = udtIndex(j).sngSignal + 30
        End If
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    Debug.Print "[GCFR_General_Module.subGenLowPointDate()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
End Sub


'''' 取得要處理的股票清單
'''Public Function GetProcessStockList() As String
'''    Dim openFile As String
'''    Dim lineData As String
'''    Dim splitData As Variant
'''    Dim path As String
'''    Dim stockList As String
'''    Dim i As Integer
'''
'''    path = GetAppPath(1)
'''    ' Read StockNo file form "taiwan_stkno.bat" file
'''    openFile = path + "stocks-files\台灣股市\個股\操作個股.txt"
'''    i = 0
'''
'''    Open openFile For Input As #1
'''    While Not EOF(1)
'''        Line Input #1, lineData
'''        splitData = Split(lineData, " ")
'''        stockList = stockList & splitData(1) & ","
'''        i = i + 1
'''    Wend
'''    Close #1
'''
'''    GetProcessStockList = stockList
'''End Function

'''' 尋找高、低點訊號
'''Public Sub FindHighAndLowSignals()
'''    Dim stockString As String
'''    Dim stockList
'''    Dim i As Integer
'''    Dim sp As SplitParameter
'''
'''    sp = GetSplitParameter
'''    i = 1
'''    stockString = GetProcessStockList
'''    stockList = Split(stockString, ",")
'''
'''    While (UBound(stockList) >= i)
'''        Call ChangeStockData(CStr(stockList(i)), sp)
'''        i = i + 1
'''    Wend
'''
'''End Sub




Private Function KDGreaterTheValueAndCrossDown(ByVal prevK As Single, ByVal prevD As Single, ByVal currK As Single, ByVal currD As Single, ByVal theValue As Single) As Boolean
    If prevK > prevD And _
        currK < currD And _
        prevK > theValue And _
        prevD > theValue And _
        currK > theValue And _
        currD > theValue Then
        KDGreaterTheValueAndCrossDown = True
    Else
        KDGreaterTheValueAndCrossDown = False
    End If
End Function


Private Function MACDUnderTheValueAndCrossUp(ByVal prevMACD As Single, ByVal prevDIF As Single, ByVal currMACD As Single, ByVal currDIF As Single, ByVal theValue As Single) As Boolean
    If prevMACD < prevDIF And currMACD > currDIF And _
        prevMACD < theValue And _
        prevDIF < theValue And _
        currMACD < theValue And _
        currDIF < theValue Then
        MACDUnderTheValueAndCrossUp = True
    Else
        MACDUnderTheValueAndCrossUp = False
    End If
End Function

Private Function MACDGreaterTheValueAndCrossDown(ByVal prevMACD As Single, ByVal prevDIF As Single, ByVal currMACD As Single, ByVal currDIF As Single, ByVal theValue As Single) As Boolean
    If prevMACD > prevDIF And currMACD < currDIF And _
        prevMACD > theValue And _
        prevDIF > theValue And _
        currMACD > theValue And _
        currDIF > theValue Then
        MACDGreaterTheValueAndCrossDown = True
    Else
        MACDGreaterTheValueAndCrossDown = False
    End If
End Function


Private Function BIASUnderTheValue(ByVal currBIAS As Single, ByVal theValue As Single) As Boolean
    If currBIAS < theValue Then
        BIASUnderTheValue = True
    Else
        BIASUnderTheValue = False
    End If
End Function


Private Function BIASGreaterTheValue(ByVal currBIAS As Single, ByVal theValue As Single) As Boolean
    If currBIAS > theValue Then
        BIASGreaterTheValue = True
    Else
        BIASGreaterTheValue = False
    End If
End Function



'''' 設定存放個股資料的路徑(取得程式存放路徑的上1(n)層(ex: D:\A\B --> D:\A\))
'''Public Function GetAppPath(ByVal upLevel As Integer) As String
'''    Dim blnFlag As Boolean
'''    Dim path As String
'''    Dim strChar As String
'''    Dim intUpDir As Integer
'''    Dim i As Integer
'''    blnFlag = False
'''    intUpDir = 0
'''    path = App.path
'''    For i = Len(path) To 1 Step -1
'''        strChar = Mid(path, i, 1)
'''        If strChar = "\" Then
'''            intUpDir = intUpDir + 1
'''        End If
'''        If intUpDir = upLevel Then
'''            path = Mid(path, 1, i)
'''            Exit For
'''        End If
'''    Next
'''    GetAppPath = path
'''End Function
'''


