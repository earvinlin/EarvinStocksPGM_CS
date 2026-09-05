Option Explicit
Option Base 1


Public VR_NO As Integer
Public OBV_No As Integer
Public TAPI_No As Integer
Public AR_No As Integer
Public MASlope_No As Integer
Public MADistance_No As Integer
Public LBias_No As Integer
Public SBias_No As Integer

'*******************************************************************************************
'* 老師給的六項指標的值
'*******************************************************************************************
Public gudt6IndexDataType As SixStockData
Public gudt6IndexData() As SixStockData

'*******************************************************************************************
'*  6-index (FileName=6index.dat)
'*******************************************************************************************
Public Type SixStockData
    sngDate As Single
    sngDiffOp1 As Single
    sngDiffOp2 As Single
    sngQMOp As Single
    sngTrendOp As Single
    sngCyOp As Single
    sngLstOp As Single
    sngLwOp As Single
End Type

'*******************************************************************************************
'* 儲存使用聚類參數及聚類結果(GRG)
'*******************************************************************************************
Public Type udtGRGIndex
    sngDate As Single
    sngParam1 As Single
    sngParam2 As Single
    sngParam3 As Single
    sngParam4 As Single
    sngParam5 As Single
    sngParam6 As Single
    sngParam7 As Single
    sngParam8 As Single
    sngGRG As Single
    sngMax As Single    ' 最大值
    sngMin As Single    ' 最小值
End Type

'*******************************************************************************************
' Save the cluster attributes
'*******************************************************************************************
Public Type SelectAttributes
   ' 20080124 modified
    sngParam() As Single
'    strParam1 As String
'    strParam2 As String
'    strParam3 As String
'    strParam4 As String
'    strParam5 As String
'    strParam6 As String
'    strParam7 As String
'    strParam8 As String
End Type

'***************************************************************************************************
'* GCFR Module 相關參數設定
'***************************************************************************************************
' 20060808
'Public Const gsngFa = 0.6                   ' 閥值
Public Const gsngRo = 0.3                    ' 關聯係數
Public gsngFa As Single                      ' 閥值
'Public gsngRo As Single                       ' 關聯係數
                                             
Public gsngRoMax As Single                   ' 關聯係數
Public gsngRoMin As Single                   ' 關聯係數
Public gintSector As Integer                 ' 記錄共分了多少個sector
Public gintSelCount As Integer               ' 記錄選擇了試驗factor的數目
Public gstrSelFactors() As String            ' 記錄選擇了試驗factor的名稱
Public gintSelFactorsMethod() As Integer     ' 記錄選擇了試驗factor的處理方式 (1:望大，2:望小，3:望目)
Public gintSelLowCount As Integer            ' 記錄選擇低點的試驗factor的數目
Public gstrSelLowFactors() As String         ' 記錄選擇低點的試驗factor的名稱
Public gintSelLowFactorsMethod() As Integer  ' 記錄選擇低點的試驗factor的處理方式 (1:望大，2:望小，3:望目)
Public gintLowSum As Integer                 ' 記錄選取低點的attribute的個數
Public gsngHighLow As Single                 ' 儲存高低點
Public gudtClusterDat() As udtGRGIndex       ' 儲存聚類的資料
Public gintClusterCnt As Integer             ' 儲存聚類的資料
Public gintSum As Integer                    ' 記錄選取的Attribute的個數
Public gsngCompFa As Single                  ' 記錄要比較的閥值(高值)
Public gsngCompFa2 As Single                 ' 記錄要比較的閥值(低值)
Public gudtSelAttrs() As SelectAttributes    ' 有關選取的聚類的指標: 記錄所選取的各個Attribute名稱
Public gintSelAttrCount() As Integer         ' 有關選取的聚類的指標: 記錄此聚類結果使用了多少個Attributes來聚類
Public gsngTestGRG As Single                 ' 測試結果的GRG

'*******************************************************************************************
'* 有關聚類的變數
'*******************************************************************************************
Public gudtCYSHDat() As udtGRGIndex
Public gudtCYSLDat() As udtGRGIndex
Public gintCYSHCount As Integer       ' 儲存cluster位置
Public gintCYSLCount As Integer       ' 儲存cluster位置

'***************************************************************************************************
'* Defin L-S Bias Value for CGFR to Generate AHS, LHS
'***************************************************************************************************
Public Const LSBIASH1 = 8
Public Const LSBIASH2 = 10
Public Const LSBIASH3 = 12
Public Const LSBIASL1 = -10
Public Const LSBIASL2 = -12
Public Const LSBIASL3 = -14

'***************************************************************************************************
'* Defin MAS for CGFR to Generate AHS, LHS
'***************************************************************************************************
Public Const MASH1 = 17.5
Public Const MASH2 = 21
Public Const MASH3 = 24
Public Const MASH4 = 27
Public Const MASH5 = 30
Public Const MASH6 = 35
Public Const MASL1 = -25
Public Const MASL2 = -26
Public Const MASL3 = -27
Public Const MASL4 = -28
Public Const MASL5 = -29
Public Const MASL6 = -30

'***************************************************************************************************
'* 訊號定義
'***************************************************************************************************
Public Const ABSOLUTEHIGH = 150     ' Rule-Define的最高點(AHS)
Public Const ABSOLUTELOW = 5        ' Rule-Define的最低點(LHS)
Public Const ABSOLUTEHIGH2 = 140    ' CGFR產生的最高點(AHS)
Public Const ABSOLUTELOW2 = 10      ' CGFR產生的最低點(LHS)
'--- 高低點訊號的值 ---
Public Const HIGHSIGNAL = 80        ' 高點訊號
Public Const LOWSIGNAL = 40         ' 低點訊號

'***************************************************************************************************
'* Testing Set, Learning Set 的起迄日期(Testing Set的結束日期=資料的最新的一筆的日期)
'***************************************************************************************************
Public Const LEARN_START_DATE = 851019  ' Learning Set Start Date
Public Const LEARN_END_DATE = 900108    ' Learning Set End Date
Public Const TEST_START_DATE = 900109   ' Testing Set Start Date

'***************************************************************************************************
'* 各項策略的設定值
'***************************************************************************************************
'--- 儲存POS 各項設定值 ---
Public intSTA1TP As Integer
Public sngSTA1TSG As Single
Public sngSTA1TSL As Single
Public intSTA1Steps As Integer
Public intSTA1Buy() As Integer
Public sngSTA1Loss() As Single
Public sngSTA1Gain() As Single
'--- 儲存SSOS 各項設定值 ---
Public intSTA2TP As Integer
Public sngSTA2TSG As Single
Public sngSTA2TSL As Single
Public intSTA2Steps As Integer
Public intSTA2Buy() As Integer
Public sngSTA2Loss() As Single
Public sngSTA2Gain() As Single
'--- 儲存EOS 各項設定值 ---
Public intSTA3TP As Integer
Public sngSTA3TSG As Single
Public sngSTA3TSL As Single
Public intSTA3BuySteps As Integer
Public intSTA3SellSteps As Integer
Public intSTA3Buy() As Integer
Public intSTA3Sell() As Integer
Public sngSTA3Loss() As Single
Public sngSTA3Gain() As Single
'--- 記錄 GRG高點、低點的值 ---
Public Const GRG_HIGH = 0.7
Public Const GRG_LOW = 0.4
'--- 趨勢的定義 ---
Public Const ISUP = 1               ' 做多
Public Const ISDOWN = -1            ' 放空
Public Const ISWAIT = 0             ' 等待
'--- Others ---
Public Const SELECTED = 0           ' 未使用
Public Const UNSELECTED = 1         ' 未使用
Public Const MAXHOLD = 5            ' 最大持有口數
Public Const MINHOLD = 1            ' 最小持有口數
'--- 用來判斷是否已經完成聚類的動作 ---
Public blnClusterDayOK As Boolean           ' 日資料已經完成聚類=true 否則=false
Public blnClusterWeekOK As Boolean          ' 週資料已經完成聚類=true 否則=false
Public blnClusterMonthOK As Boolean         ' 月資料已經完成聚類=true 否則=false


Dim sngMax(8) As Single, sngMin(8) As Single    ' 定義聚類過程中的最大值、最小值
Public gsngEndPrice(5) As Single                ' 記錄每口購買時的價格 (收盤價)
Public gblnAHSAppear As Boolean                 ' 記錄在做多的Sector中尋找放空點的過程是否有出現高點訊號
'===================================================================================
' 下面變數用來判斷在CY>0下放空點
'===================================================================================
Public mintHP As Integer
Public mintKDCross As Integer
Public mintRSICross As Integer
Public mintWMS As Integer



'***************************************************************************************************
'* 說    明: Calculate the Willams value.
'* 輸入參數: Stockdata  股價資料
'*           IndexData  技術指標資料
'*           intStockNo 資料筆數
'*           intWMSNo   指標天數
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subRWMS(ByRef udtStock() As StockData, _
                   ByRef udtIndex() As IndexData, _
                   ByVal intStockNo As Integer, _
                   ByVal intWMSNo As Integer)
    Dim sngWMS As Single
    Dim i As Integer
    Dim j As Integer
    Dim sngMax As Single
    Dim sngMin As Single
    
    On Error GoTo ERR_HANDLE
    
    j = 1
    While j <= intStockNo
        sngMax = -1:   sngMin = 99999999
        If j <= intWMSNo Then
            For i = 1 To j
                If sngMax < udtStock(i).sngHighPrice Then
                    sngMax = udtStock(i).sngHighPrice
                End If
                If sngMin > udtStock(i).sngLowPrice Then
                    sngMin = udtStock(i).sngLowPrice
                End If
            Next
        Else
            For i = j To j - intWMSNo + 1 Step -1
                If sngMax < udtStock(i).sngHighPrice Then
                    sngMax = udtStock(i).sngHighPrice
                End If
                If sngMin > udtStock(i).sngLowPrice Then
                    sngMin = udtStock(i).sngLowPrice
                End If
            Next
        End If
        If sngMax <> sngMin Then
            sngWMS = (sngMax - udtStock(j).sngEndprice) / (sngMax - sngMin) * 100
        Else
            sngWMS = 50
        End If
        udtIndex(j).sngRWMS = 100 - sngWMS
        
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subRWMS()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the StochRSI value.
'* 輸入參數: Stockdata  股價資料
'*           IndexData  技術指標資料
'*           intStockNo 資料筆數
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subStochRSI(ByRef udtStock() As StockData, _
                        ByRef udtIndex() As IndexData, _
                        ByVal intStockNo As Integer)
    Dim sngMax As Single
    Dim sngMin As Single
    Dim Temp As Single
    Dim i As Integer
    Dim j As Integer
    sngMax = 50
    sngMin = 50
    Temp = 0
    j = 1
    
    On Error GoTo ERR_HANDLE
    
    While j <= intStockNo
        If j <= StochRSI_No Then
            sngMax = 0
            sngMin = 100
            For i = 1 To j
                If sngMax < udtIndex(i).sngRSI_S Then
                    sngMax = udtIndex(i).sngRSI_S
                End If
                If sngMin > udtIndex(i).sngRSI_S Then
                    sngMin = udtIndex(i).sngRSI_S
                End If
            Next i
        Else
            sngMax = 0
            sngMin = 100
            For i = j To j - StochRSI_No + 1 Step -1
                If sngMax < udtIndex(i).sngRSI_S Then
                    sngMax = udtIndex(i).sngRSI_S
                End If
                If sngMin > udtIndex(i).sngRSI_S Then
                    sngMin = udtIndex(i).sngRSI_S
                End If
            Next i
        End If
        If (sngMax - sngMin <> 0) Then
            udtIndex(j).sngStochRSI = (udtIndex(j).sngRSI_S - sngMin) * 100 / (sngMax - sngMin)
        Else
            udtIndex(j).sngStochRSI = 0.5
        End If
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subStochRSI()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the MACD and Cy value.
'* 輸入參數: Stockdata  股價資料
'*           IndexData  技術指標資料
'*           intStockNo 資料筆數
'*           intMACDNo  ?????
'*           intSEMANo  ?????
'*           intLEMANo  ?????
'*           IsDaily    ?????
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subMACD(ByRef udtStock() As StockData, _
                   ByRef udtIndex() As IndexData, _
                    ByVal intStockNo As Integer, _
                    ByVal intMACDNo As Integer, _
                    ByVal intSEMANo As Integer, _
                    ByVal intLEMANo As Integer, _
                    ByVal IsDaily As Boolean)
'   Dim sngTemp As Single
    Dim sngEMA_S As Single
    Dim sngEMA_L As Single
    Dim sngPreEMA_S As Single
    Dim sngPreEMA_L As Single
    Dim sngPreMACD As Single
    Dim sngMACD As Single
    Dim sngDIF As Single
    Dim sngDIF_MACD As Single
    Dim sngDI As Single
    Dim j As Integer
    Dim i As Integer
    
    On Error GoTo ERR_HANDLE
    
'    If IsDaily Then
'        sngPreEMA_S = 6500
'        sngPreEMA_L = 6500
'    Else
'        sngPreEMA_S = 6750
'        sngPreEMA_L = 6750
'    End If
'    sngTemp = 0
'    For j = 1 To 5
'      With udtStock(j)
'        sngTemp = sngTemp + .sngEndprice
'      End With
'    Next
    sngPreEMA_S = udtStock(1).sngEndprice
    sngPreEMA_L = udtStock(1).sngEndprice
    sngPreMACD = 0
    j = 1
    While j <= intStockNo
        sngDI = (udtStock(j).sngHighPrice + udtStock(j).sngLowPrice + udtStock(j).sngEndprice * 2) / 4
        sngEMA_S = sngPreEMA_S + (2 * (sngDI - sngPreEMA_S) / (1 + intSEMANo))
        sngEMA_L = sngPreEMA_L + (2 * (sngDI - sngPreEMA_L) / (1 + intLEMANo))
        sngDIF = sngEMA_S - sngEMA_L
        sngMACD = sngPreMACD + (2 * (sngDIF - sngPreMACD) / (1 + intMACDNo))
        sngDIF_MACD = sngDIF - sngMACD
        
        udtIndex(j).sngDIF = sngDIF
        udtIndex(j).sngMACD = sngMACD
        udtIndex(j).sngDIF_MACD = sngDIF_MACD
        
        sngPreEMA_S = sngEMA_S
        sngPreEMA_L = sngEMA_L
        sngPreMACD = sngMACD
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subMACD()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the average value of volume and main value.
'* 輸入參數: udtStock   股價資料
'*           udtIndex   技術指標資料
'*           intStockNo 資料筆數
'*           intDayNo   指標天數
'*           intPos     運算結果的值要儲存到(udtIndex)那一個位置
'*           blnIsP     是計算成交價(=true) 或 成交量(=false)
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subAverage(ByRef udtStock() As StockData, _
                      ByRef udtIndex() As IndexData, _
                      ByVal intStockNo As Integer, _
                      ByVal intDayNo As Integer, _
                      ByVal intPos As Integer, _
                      ByVal blnIsP As Boolean)
    Dim i As Integer
    Dim j As Integer
    Dim sngAverage As Single
    
    On Error GoTo ERR_HANDLE
    
    j = 1
    While j <= intStockNo
        sngAverage = 0
        If j < intDayNo Then
            For i = 1 To j
                If blnIsP Then
                    sngAverage = sngAverage + udtStock(i).sngEndprice
                Else
                    sngAverage = sngAverage + udtStock(i).sngVol
                End If
            Next
            sngAverage = sngAverage / j
        ElseIf intDayNo <> 0 Then
            For i = j To j - intDayNo + 1 Step -1
                If blnIsP Then
                    sngAverage = sngAverage + udtStock(i).sngEndprice
                Else
                    sngAverage = sngAverage + udtStock(i).sngVol
                End If
            Next
            sngAverage = sngAverage / intDayNo
        Else
            sngAverage = 0
        End If
        
        If blnIsP Then
            If intPos = 1 Then
                udtIndex(j).sngMAP1 = sngAverage
            ElseIf intPos = 2 Then
                udtIndex(j).sngMAP2 = sngAverage
            ElseIf intPos = 3 Then
                udtIndex(j).sngMAP3 = sngAverage
            ElseIf intPos = 4 Then
                udtIndex(j).sngMAP4 = sngAverage
            ElseIf intPos = 5 Then
                udtIndex(j).sngMAP5 = sngAverage
            ElseIf intPos = 6 Then
                udtIndex(j).sngMAP6 = sngAverage
            ElseIf intPos = 7 Then
                udtIndex(j).sngMAP7 = sngAverage
            ElseIf intPos = 0 Then
                udtIndex(j).sngMAP0 = sngAverage
            End If
        Else
            If intPos = 1 Then
                udtIndex(j).sngMAV1 = sngAverage
            ElseIf intPos = 2 Then
                udtIndex(j).sngMAV2 = sngAverage
            ElseIf intPos = 3 Then
                udtIndex(j).sngMAV3 = sngAverage
            ElseIf intPos = 4 Then
                udtIndex(j).sngMAV4 = sngAverage
            ElseIf intPos = 5 Then
                udtIndex(j).sngMAV5 = sngAverage
            End If
        End If
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subAverage()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


Public Sub subDMI(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer, _
                  ByVal intDayNo As Integer)
                  
    Dim sngPreDM_P As Single
    Dim sngPreDM_N As Single
    Dim sngPreTR As Single
    Dim sngDM_P As Single
    Dim sngDM_N As Single
    Dim sngTR As Single
    Dim sngDI_P As Single
    Dim sngDI_N As Single
    Dim sngDX As Single
    Dim sngADX As Single
    Dim sngTemp As Single
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ERR_HANDLE
    
    For j = 1 To intStockNo
        sngDM_P = 0: sngDM_N = 0: sngTR = 0
        If j > intDayNo Then
            sngDM_P = udtStock(j).sngHighPrice - udtStock(j - 1).sngLowPrice
            sngDM_N = udtStock(j).sngLowPrice - udtStock(j - 1).sngLowPrice
            sngTemp = 0
            sngTemp = udtStock(j).sngHighPrice - udtStock(j).sngLowPrice
            sngTemp = IIf((udtStock(j).sngHighPrice - udtStock(j - 1).sngEndprice) > sngTemp, (udtStock(j).sngHighPrice - udtStock(j - 1).sngEndprice), sngTemp)
            sngTemp = IIf((udtStock(j).sngLowPrice - udtStock(j - 1).sngEndprice) > sngTemp, (udtStock(j).sngLowPrice - udtStock(j - 1).sngEndprice), sngTemp)
        
            sngDM_P = sngPreDM_P * (intDayNo - 1) / intDayNo + sngDM_P
            sngDM_N = sngPreDM_N * (intDayNo - 1) / intDayNo + sngDM_N
            sngTR = sngPreTR * (intDayNo - 1) / intDayNo + sngTemp
                          
            udtIndex(j).sngDI_P = sngDM_P / sngTR * 100
            udtIndex(j).sngDI_N = sngDM_N / sngTR * 100
        
            sngDX = Abs(udtIndex(j).sngDI_P - udtIndex(j).sngDI_N) / (udtIndex(j).sngDI_P + udtIndex(j).sngDI_N) * 100
            udtIndex(j).sngADX = udtIndex(j - 1).sngADX * (intDayNo - 1) / intDayNo + (sngDX / intDayNo)
            sngPreDM_P = sngDM_P
            sngPreDM_N = sngDM_N
            sngPreTR = sngTR
        End If
    Next
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subDMI()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the Wilder’s RSI value.
'* 輸入參數: Stockdata  股價資料
'*           IndexData  技術指標資料
'*           intStockNo 資料筆數
'*           intWRSINo  指標天數
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subWRSI(ByRef udtStock() As StockData, _
                     ByRef udtIndex() As IndexData, _
                     ByVal intStockNo As Integer, _
                     ByVal intWRSINo As Integer)
    Dim sngWRSI As Single
    Dim sngUp As Single
    Dim sngDown As Single
    Dim sngDiff As Single
    Dim i As Integer
    Dim j As Integer
    Dim sngPreUp As Single
    Dim sngPreDown As Single
    
    On Error GoTo ERR_HANDLE
    
    sngPreUp = 0
    sngPreDown = 0
    
    j = 2
    While j <= intStockNo
        sngUp = 0: sngDown = 0
        sngDiff = udtStock(j).sngEndprice - udtStock(j - 1).sngEndprice
        If sngDiff > 0 Then
            sngUp = sngDiff
        Else
            sngDown = Abs(sngDiff)
        End If
        
        If j <= intWRSINo Then
            sngUp = ((j - 1) / j) * sngPreUp + (1 / j) * sngUp
            sngDown = ((j - 1) / j) * sngPreDown + (1 / j) * sngDown
        Else
            sngUp = ((intWRSINo - 1) / intWRSINo) * sngPreUp + (1 / intWRSINo) * sngUp
            sngDown = ((intWRSINo - 1) / intWRSINo) * sngPreDown + (1 / intWRSINo) * sngDown
        End If
        
        If j <> 1 And (sngUp + sngDown) <> 0 Then
            sngWRSI = sngUp / (sngUp + sngDown) * 100
        Else
            sngWRSI = 50
        End If
        sngPreUp = sngUp
        sngPreDown = sngDown
    
        udtIndex(j).sngWRSI = sngWRSI

        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subWRSI()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the Wilder’s MARSI value.
'* 輸入參數: Stockdata   股價資料
'*           IndexData   技術指標資料
'*           intStockNo  資料筆數
'*           intMAWRSINo 指標天數
'*           blnIsEMA    ?????
'* 輸出參數: 無
'* 版    本: 2.00 20080910 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subMAWRSI(ByRef udtStock() As StockData, _
                     ByRef udtIndex() As IndexData, _
                     ByVal intStockNo As Integer, _
                     ByVal intMAWRSINo As Integer, _
                     ByVal blnIsEMA As Boolean)
    Dim sngMAWRSI As Single
    Dim i As Integer
    Dim j As Integer
    Dim sngPreMAWRSI As Single
    
    On Error GoTo ERR_HANDLE
    
    j = 1
    
    If blnIsEMA Then
        While j <= intStockNo
            sngMAWRSI = 0
            If j < intMAWRSINo Then
                sngMAWRSI = sngPreMAWRSI + (2 * (udtIndex(j).sngWRSI - sngPreMAWRSI) / (1 + j))
            Else
                sngMAWRSI = sngPreMAWRSI + (2 * (udtIndex(j).sngWRSI - sngPreMAWRSI) / (1 + intMAWRSINo))
            End If
            udtIndex(j).sngMAWRSI = sngMAWRSI
            sngPreMAWRSI = sngMAWRSI
            
            j = j + 1
        Wend
        
    Else
        While j <= intStockNo
            sngMAWRSI = 0
            If j < intMAWRSINo Then
                For i = 1 To j
                    sngMAWRSI = sngMAWRSI + udtIndex(i).sngWRSI
                Next
                sngMAWRSI = sngMAWRSI / j
            Else
                For i = j To j - intMAWRSINo + 1 Step -1
                    sngMAWRSI = sngMAWRSI + udtIndex(i).sngWRSI
                Next
                sngMAWRSI = sngMAWRSI / intMAWRSINo
            End If
            udtIndex(j).sngMAWRSI = sngMAWRSI
            j = j + 1
        Wend
    End If
   
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subMAWRSI()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub

Private Function Max(ByRef MaxRule() As Single, ByVal count As Integer, ByRef index As Integer) As Single
    Dim i As Integer
    Dim Mval As Single
    
    On Error GoTo ERR_HANDLE
    
    Mval = MaxRule(1)
    index = 1
    
    For i = 1 To count - 1
        If MaxRule(i + 1) > Mval Then
            Mval = MaxRule(i + 1)
            index = i + 1
        End If
    Next
    
    Max = Mval
    
    Exit Function
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.Max()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Function


'***************************************************************************************************
'* 說    明:
'*    調整CY Sector
'* 輸入參數:
'*    udtStock 每日股價資料
'*    udtIndex 儲存的指數資料
'*    intStockNo 資料筆數
'* 輸出參數: 無
'* 版    本:
'*    2.00: 20080913 Earvin   New
'***************************************************************************************************
Public Sub subAdjustCYSector(ByRef udtStock() As StockData, ByRef udtIndex() As IndexData, _
                            ByVal intStockNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim sngAverage As Single
    
    On Error GoTo ERR_HANDLE
        
    j = 1
    
    While j <= intStockNo
        If udtIndex(j).sngP60 > udtStock(j).sngEndprice _
            And udtIndex(j).sngCyOp < 1 And udtIndex(j).sngCyOp > 0 Then
            udtIndex(j).sngCyOp = -0.1
        End If
        If udtIndex(j).sngP60 < udtStock(j).sngEndprice _
            And udtIndex(j).sngCyOp > -3 And udtIndex(j).sngCyOp < 0 Then
            udtIndex(j).sngCyOp = -0.1
        End If
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subAdjustCYSector()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: 顯示當日是否為盤整日期
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'* 輸出參數: 無
'* 版    本: 2.00 20080913 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subIsInRange(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer)
    Dim i As Integer
    Dim intAA As Integer
    Dim intBB As Integer
    Dim intDays As Integer
    
    On Error GoTo ERR_HANDLE
    
    '--- Initialize Variables ---
    i = 2
    intAA = 0
    intBB = 0
    intDays = 0
    
    While i <= intStockNo
        udtIndex(i).sngInRange = 0

        ' 若均線陷入盤整，則ASSIGN VALUE = 60
        If Abs(udtStock(i).sngEndprice - udtIndex(i).sngP5) < 100 _
            And Abs(udtIndex(i).sngP5 - udtIndex(i).sngP10) < 100 _
            And Abs(udtIndex(i).sngP10 - udtIndex(i).sngp20) < 100 _
            And Abs(udtStock(i).sngEndprice - udtIndex(i).sngP10) < 100 _
            And Abs(udtStock(i).sngEndprice - udtIndex(i).sngp20) < 150 Then
            intAA = 0
            intDays = intDays + 1
            udtIndex(i).sngInRange = 5
            If Abs(udtStock(i).sngEndprice - udtIndex(i).sngP60) < 100 Then
                udtIndex(i).sngInRange = 10
            End If
            udtIndex(i).sngInRange = intDays
            intBB = 0
            ' 前一天若為盤整，則本次檢查放寬
        ElseIf udtIndex(i - 1).sngInRange <= 10 _
            And udtIndex(i - 1).sngInRange > 0 _
            And intAA < 3 _
            And Abs(udtStock(i).sngEndprice - udtIndex(i).sngP5) < 150 _
            And Abs(udtIndex(i).sngP5 - udtIndex(i).sngP10) < 130 _
            And Abs(udtIndex(i).sngP10 - udtIndex(i).sngp20) < 130 _
            And Abs(udtStock(i).sngEndprice - udtIndex(i).sngP10) < 100 _
            And Abs(udtStock(i).sngEndprice - udtIndex(i).sngp20) < 150 Then
            intAA = intAA + 1
            intDays = intDays + 1
            udtIndex(i).sngInRange = 15
            If Abs(udtStock(i).sngEndprice - udtIndex(i).sngP60) < 100 Then
                udtIndex(i).sngInRange = 10
            End If
            udtIndex(i).sngInRange = intDays
            intBB = 0
        ' 前二天若為盤整，則本次檢查放寬
        ElseIf udtIndex(i - 1).sngInRange <= 15 _
            And udtIndex(i - 1).sngInRange > 10 _
            And intAA < 3 _
            And Abs(udtStock(i).sngEndprice - udtIndex(i).sngP5) < 150 _
            And Abs(udtIndex(i).sngP5 - udtIndex(i).sngP10) < 150 _
            And Abs(udtIndex(i).sngP10 - udtIndex(i).sngp20) < 150 _
            And Abs(udtStock(i).sngEndprice - udtIndex(i).sngP10) < 120 _
            And Abs(udtStock(i).sngEndprice - udtIndex(i).sngp20) < 150 Then
            intAA = intAA + 1
            intDays = intDays + 1
            udtIndex(i).sngInRange = 15
            If Abs(udtStock(i).sngEndprice - udtIndex(i).sngP60) < 120 Then
                udtIndex(i).sngInRange = 15
            End If
            udtIndex(i).sngInRange = intDays
            intBB = 0
        Else
            If intBB < 1 Then
                intBB = intBB + 1
                udtIndex(i).sngInRange = intDays
            Else
                intDays = 0
                intBB = 0
            End If
        End If
        
        i = i + 1
    Wend
        
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subIsInRange()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明:
'*    調整CY Sector
'* 輸入參數:
'*    udtStock 每日股價資料
'*    udtIndex 儲存的指數資料
'*    intStockNo 資料筆數
'* 輸出參數: 無
'* 版    本:
'*    2.00: 20080913 Earvin   New
'***************************************************************************************************
Public Sub subAdjustCYSector2(ByRef udtStock() As StockData, ByRef udtIndex() As IndexData, _
                                ByVal intStockNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim sngAverage As Single
    
    On Error GoTo ERR_HANDLE
        
    j = 2
    
    While j <= intStockNo
        If udtIndex(j - 1).sngCyOp > 0 And udtIndex(j).sngCyOp < 0 Then
            If udtIndex(j).sngCyOp > -10 Then
                udtIndex(j).sngCyOp = -udtIndex(j).sngCyOp
            End If
        End If
        If udtIndex(j - 1).sngCyOp < 0 And udtIndex(j).sngCyOp > 0 Then
            If udtIndex(j).sngCyOp < 10 Then
                udtIndex(j).sngCyOp = -udtIndex(j).sngCyOp
            End If
        End If
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subAdjustCYSector2()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    Resume Next
End Sub


'***************************************************************************************************
'* 說    明: Calculate the Long-Term - Short-Term BIAS value
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intBiasNo  欲計算指標之天數
'* 輸出參數: 無
'* 版    本: 1.00 20041214 新增
'*           1.10 20050604 Modified
'*           2.01 20081222 調整格式及增加註解
'***************************************************************************************************
Public Sub subLSBias(ByRef udtStock() As StockData, _
                    ByRef udtIndex() As IndexData, _
                    ByVal intStockNo As Integer)
    Dim i As Integer
    
    On Error GoTo ERR_HANDLE
        
    i = 1
    
    While i <= intStockNo
        udtIndex(i).sngLSBias = udtIndex(i).sngLBias - udtIndex(i).sngSBias
        i = i + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subLSBias()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: 計算形成多頭排列後持續天數
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'* 輸出參數: 無
'* 版    本: 2.00 20080915 新增
'*           2.01 20081222 調整格式及增加註解
'***************************************************************************************************
Public Sub subBullList(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer)
    Dim i As Integer
    Dim intBullListDay As Integer
    
    On Error GoTo ERR_HANDLE
    
    i = 1
    intBullListDay = 0
    
    While i <= intStockNo
        With udtIndex(i)
            If .sngP5 > .sngP10 And .sngP10 > .sngp20 And .sngp20 > .sngP60 Then
                intBullListDay = intBullListDay + 1
            Else
                intBullListDay = 0
            End If
            .sngBullList = intBullListDay
        End With
        i = i + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subBullList()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: 計算形成空頭排列後持續天數
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subBearList(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer)
    Dim i As Integer
    Dim intBearListDay As Integer
   
    On Error GoTo ERR_HANDLE
    
    i = 1
    intBearListDay = 0
    
    While i <= intStockNo
        With udtIndex(i)
            If .sngP5 < .sngP10 And .sngP10 < .sngp20 Then
                intBearListDay = intBearListDay + 1
            Else
                intBearListDay = 0
            End If
            .sngBearList = intBearListDay
        End With
        i = i + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subBearList()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明:
'*    顯示當日是否為波段高點或波段低點
'* 輸入參數:
'*    udtStoc 股市資料
'*    udtIndex 股市指標資料
'*    intStockNo 股市資料總筆數
'* 輸出參數: 無
'* 版    本:
'*    2.00: 20080916 Earvin   New
'***************************************************************************************************
Public Sub subDisplayHighAndLowArea(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer)
    Dim intCurrentPos As Integer
    Dim intStartPos As Integer
    Dim sngHighPrice As Single
    Dim sngLowPrice As Single
    Dim i As Integer
    
    On Error GoTo ERR_HANDLE
    
    intStartPos = 0
    intCurrentPos = 2
    
    While (intCurrentPos < intStockNo)
        If (udtIndex(intCurrentPos - 1).sngCyOp > 0 And udtIndex(intCurrentPos).sngCyOp < 0) _
            Or (udtIndex(intCurrentPos - 1).sngCyOp < 0 And udtIndex(intCurrentPos).sngCyOp > 0) _
            Or (udtIndex(intCurrentPos - 1).sngCyOp = 0 And udtIndex(intCurrentPos).sngCyOp > 0) Then
            If intStartPos > 0 Then
                sngHighPrice = sngHighPrice - (sngHighPrice - sngLowPrice) / 10 * 2
                sngLowPrice = sngLowPrice + (sngHighPrice - sngLowPrice) / 10 * 2
                For i = intStartPos To intCurrentPos - 1
                    If udtStock(i).sngEndprice >= sngHighPrice Then
                        If udtIndex(i).sngCyOp > 0 Then
                            udtIndex(i).sngHighLowArea = 2
                        Else
                            udtIndex(i).sngHighLowArea = 0
                        End If
                    ElseIf udtStock(i).sngEndprice <= sngLowPrice Then
                        If udtIndex(i).sngCyOp < 0 Then
                            udtIndex(i).sngHighLowArea = -2
                        Else
                            udtIndex(i).sngHighLowArea = 0
                        End If
                    Else
                        udtIndex(i).sngHighLowArea = 0
                    End If
                Next
            End If
            intStartPos = intCurrentPos
            sngHighPrice = udtStock(intCurrentPos).sngEndprice
            sngLowPrice = udtStock(intCurrentPos).sngEndprice
        Else
            If udtStock(intCurrentPos).sngEndprice > sngHighPrice Then
                sngHighPrice = udtStock(intCurrentPos).sngEndprice
            End If
            If udtStock(intCurrentPos).sngEndprice < sngLowPrice Then
                sngLowPrice = udtStock(intCurrentPos).sngEndprice
            End If
        End If
        
        intCurrentPos = intCurrentPos + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    Debug.Print "[GCFR_Index_Module.subDisplayHighAndLowArea()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    MsgBox "[GCFR_Index_Module.subDisplayHighAndLowArea()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    Resume Next
End Sub


'***************************************************************************************************
'* 說    明: Calculate the MA-Slope value.
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intMAPNo   MA的斜率
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081222 調整格式及增加註解
'***************************************************************************************************
Public Sub subMASlope(ByRef udtStock() As StockData, _
                        ByRef udtIndex() As IndexData, _
                        ByVal intStockNo As Integer, _
                        ByVal intMAPNo As Integer)
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ERR_HANDLE
        
    j = 2
    
    While j <= intStockNo
        Select Case intMAPNo
            Case 3
                udtIndex(j).sngMASlope = udtIndex(j).sngP3 - udtIndex(j - 1).sngP3
            Case 5
                udtIndex(j).sngMASlope = udtIndex(j).sngP5 - udtIndex(j - 1).sngP5
            Case 6
                udtIndex(j).sngMASlope = udtIndex(j).sngP6 - udtIndex(j - 1).sngP6
            Case 10
                udtIndex(j).sngMASlope = udtIndex(j).sngP10 - udtIndex(j - 1).sngP10
            Case 12
                udtIndex(j).sngMASlope = udtIndex(j).sngP12 - udtIndex(j - 1).sngP12
            Case 20
                udtIndex(j).sngMASlope = udtIndex(j).sngp20 - udtIndex(j - 1).sngp20
            Case 24
                udtIndex(j).sngMASlope = udtIndex(j).sngP24 - udtIndex(j - 1).sngP24
            Case 30
                udtIndex(j).sngMASlope = (udtIndex(j).sngP30 - udtIndex(j - 1).sngP30) / udtIndex(j).sngP30 * 10000
            Case 60
                ' 只計算差值
'                udtIndex(j).sngMASlope = udtIndex(j).sngP60 - udtIndex(j - 1).sngP60
                ' 取斜率
'                udtIndex(j).sngMASlope = (udtIndex(j).sngP60 - udtIndex(j - 1).sngP60) / udtIndex(j).sngP60
                ' 取斜率 * 10000
                udtIndex(j).sngMASlope = (udtIndex(j).sngP60 - udtIndex(j - 1).sngP60) / udtIndex(j).sngP60 * 10000
            Case 72
                udtIndex(j).sngMASlope = udtIndex(j).sngP72 - udtIndex(j - 1).sngP72
'            Case 120
'                udtIndex(j).sngMASlope = udtIndex(j).sngP120 - udtIndex(j - 1).sngP120
            Case 144
                udtIndex(j).sngMASlope = udtIndex(j).sngP144 - udtIndex(j - 1).sngP144
'            Case 240
'                udtIndex(j).sngMASlope = udtIndex(j).sngP240 - udtIndex(j - 1).sngP240
            Case 288
                udtIndex(j).sngMASlope = udtIndex(j).sngP288 - udtIndex(j - 1).sngP288
        End Select
        j = j + 1
    Wend
    
   Exit Sub
    
ERR_HANDLE:
    Debug.Print "[GCFR_Index_Module.subMASlope()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    MsgBox "[GCFR_Index_Module.subMASlope()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    Resume Next
End Sub


'***************************************************************************************************
'* 說    明: 60MAP vs EndPrice
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subUpDownDays(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer)
    Dim intCurrentPos As Integer
    Dim intStartPos As Integer
    Dim sngHighPrice As Single
    Dim sngLowPrice As Single
    Dim i As Integer
    
    On Error GoTo ERR_HANDLE
    
    intStartPos = 0
    intCurrentPos = 2
    
    While (intCurrentPos < intStockNo)
        If udtIndex(intCurrentPos).sngP60 > udtStock(intCurrentPos).sngEndprice Then
            If udtIndex(intCurrentPos - 1).sngUpDownDays > 0 Then
                udtIndex(intCurrentPos).sngUpDownDays = -1
            Else
                udtIndex(intCurrentPos).sngUpDownDays = udtIndex(intCurrentPos - 1).sngUpDownDays - 1
            End If
        ElseIf udtIndex(intCurrentPos).sngP60 < udtStock(intCurrentPos).sngEndprice Then
            If udtIndex(intCurrentPos - 1).sngUpDownDays < 0 Then
                udtIndex(intCurrentPos).sngUpDownDays = 1
            Else
                udtIndex(intCurrentPos).sngUpDownDays = udtIndex(intCurrentPos - 1).sngUpDownDays + 1
            End If
        Else
            udtIndex(intCurrentPos).sngUpDownDays = 0
        End If
        intCurrentPos = intCurrentPos + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    Debug.Print "[GCFR_Index_Module.subUpDownDays()] -- " & Err.Number & ":" & Err.Description, vbOKOnly
    MsgBox "[GCFR_Index_Module.subUpDownDays()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    Resume Next
End Sub


'***************************************************************************************************
'* 說    明: VR Index (成交量比率)
'*    依據「量先於價」及「價量同步同向」之理論，計算一段期間內上漲日交易金
'*    額與下跌日交易金額之比率關係，以為研依據
'*    Formula: VR(n) = [UpTotVol(n)+(1/2×EquTotVol(n)] /
'*                     [DownTotVol(n)+(1/2×EquTotVol(n)] × 100
'*                     UpTotVol(n)  ：表示過去n日股價上漲日之成交量總數
'*                     DownTotVol(n)：表示過去n日股價下跌日之成交量總數
'*                     EquTotVol(n) ：表示過去n日股價不變日之成交量總數
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intVRNo    欲計算指標之天數
'* 輸出參數: 無
'* 版    本: 2.00 20050604 新增
'*           2.01 20081222 調整格式及增加註解
'***************************************************************************************************
Public Sub subVR(ByRef udtStock() As StockData, ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer, ByVal intVRNo As Integer)
    Dim i As Integer, j As Integer
    Dim sngVR As Single
    Dim sngVolUP As Single
    Dim sngVolDown As Single
    Dim sngVolEqu As Single
       
    On Error GoTo ERR_HANDLE
    
    j = 1
    
    While j < intStockNo
        If j <= intVRNo Then
            sngVR = 0
        Else
            sngVolUP = 0
            sngVolDown = 0
            sngVolEqu = 0
            
            For i = j To j - intVRNo + 1 Step -1
                If (udtStock(i).sngEndprice > udtStock(i).sngStartprice) Then
                    sngVolUP = sngVolUP + udtStock(i).sngVol
                ElseIf (udtStock(i).sngEndprice < udtStock(i).sngStartprice) Then
                    sngVolDown = sngVolDown + udtStock(i).sngVol
                Else
                    sngVolEqu = sngVolEqu + udtStock(i).sngVol
                End If
            Next
            '--- 若分母為0，則將值設為0 ---
            If (sngVolDown + (sngVolEqu / 2)) > 0 Then
                sngVR = ((sngVolUP + (sngVolEqu / 2)) / (sngVolUP + sngVolDown + (sngVolEqu / 2))) * 100
            Else
                sngVR = 0
            End If
        End If
        udtIndex(j).sngVR = sngVR
        j = j + 1
    Wend
       
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subVR()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明:
'*    VCalculate the MA-Slope value.
'* 輸入參數:
'*    udtStock  :
'*    udtIndex  :
'*    intStockNo:
'*    intMAPNo  : MA的斜率
'* 輸出參數: 無
'* 版    本:
'*    2.00  20050604 Earvin   New
'***************************************************************************************************
Public Sub subMADistance(ByRef udtStock() As StockData, _
                            ByRef udtIndex() As IndexData, _
                            ByVal intStockNo As Integer, _
                            ByVal intMAPNo As Integer)
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ERR_HANDLE
        
    j = 2
    
    While j <= intStockNo
        Select Case intMAPNo
            Case 3
                udtIndex(j).sngMADistance = udtIndex(j).sngP3 - udtIndex(j - 1).sngP3
            Case 5
                udtIndex(j).sngMADistance = udtIndex(j).sngP5 - udtIndex(j - 1).sngP5
            Case 6
                udtIndex(j).sngMADistance = udtIndex(j).sngP6 - udtIndex(j - 1).sngP6
            Case 10
                udtIndex(j).sngMADistance = udtIndex(j).sngP10 - udtIndex(j - 1).sngP10
            Case 12
                udtIndex(j).sngMADistance = udtIndex(j).sngP12 - udtIndex(j - 1).sngP12
            Case 20
                udtIndex(j).sngMADistance = udtIndex(j).sngp20 - udtIndex(j - 1).sngp20
            Case 24
                udtIndex(j).sngMADistance = udtIndex(j).sngP24 - udtIndex(j - 1).sngP24
            Case 30
                udtIndex(j).sngMADistance = udtIndex(j).sngP30 - udtIndex(j - 1).sngP30
            Case 60
                udtIndex(j).sngMADistance = udtStock(j).sngEndprice - udtIndex(j).sngP60
            Case 72
                udtIndex(j).sngMADistance = udtIndex(j).sngP72 - udtIndex(j - 1).sngP72
'            Case 120
'                udtIndex(j).sngMADistance = udtIndex(j).sngP120 - udtIndex(j - 1).sngP120
            Case 144
                udtIndex(j).sngMADistance = udtIndex(j).sngP144 - udtIndex(j - 1).sngP144
'            Case 240
'                udtIndex(j).sngMADistance = udtIndex(j).sngP240 - udtIndex(j - 1).sngP240
            Case 288
                udtIndex(j).sngMADistance = udtIndex(j).sngP288 - udtIndex(j - 1).sngP288
        End Select
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subMADistance()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the Long-Term BIAS value
'*           Formula : LBias(N) = (Cn - MA(N)) / MA(N) * 100
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intBiasNo  欲計算指標之天數
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081222 調整格式及增加註解
'***************************************************************************************************
Public Sub subLBias(ByRef udtStock() As StockData, _
                    ByRef udtIndex() As IndexData, _
                    ByVal intStockNo As Integer, _
                    ByVal intBiasNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim sngAverage As Single
    
    On Error GoTo ERR_HANDLE
        
    j = 1
    
    While j <= intStockNo
        sngAverage = 0
        If j < intBiasNo Then
            For i = 1 To j
                sngAverage = sngAverage + udtStock(i).sngEndprice
            Next
            sngAverage = sngAverage / j
        Else
            For i = j To j - intBiasNo + 1 Step -1
                sngAverage = sngAverage + udtStock(i).sngEndprice
            Next
            sngAverage = sngAverage / intBiasNo
        End If
    
        udtIndex(j).sngLBias = (udtStock(j).sngEndprice - sngAverage) / sngAverage * 100
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subLBias()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the Short-Term BIAS value
'*           Formula : SBias(N) = (Cn - MA(N)) / MA(N) * 100
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intBiasNo  欲計算指標之天數
'* 輸出參數: 無
'* 版    本: 2.00 20050604 新增
'*           2.01 20081222 調整格式及增加註解
'***************************************************************************************************
Public Sub subSBias(ByRef udtStock() As StockData, _
                    ByRef udtIndex() As IndexData, _
                    ByVal intStockNo As Integer, _
                    ByVal intBiasNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim sngAverage As Single
    
    On Error GoTo ERR_HANDLE
        
    j = 1
    
    While j <= intStockNo
        sngAverage = 0
        If j < intBiasNo Then
            For i = 1 To j
                sngAverage = sngAverage + udtStock(i).sngEndprice
            Next
            sngAverage = sngAverage / j
        Else
            For i = j To j - intBiasNo + 1 Step -1
                sngAverage = sngAverage + udtStock(i).sngEndprice
            Next
            sngAverage = sngAverage / intBiasNo
        End If
    
        udtIndex(j).sngSBias = (udtStock(j).sngEndprice - sngAverage) / sngAverage * 100
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
'    MsgBox "[GCFR_Index_Module.subSBias()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    Debug.Print "[GCFR_Index_Module.subSBias()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
    Resume Next
End Sub


'***************************************************************************************************
'* 說    明:
'*    Calculate the average value of volume and main value.
'* 輸入參數:
'*    udtStock  : 股市資料
'*    udtIndex  : 股市指標資料
'*    intStockNo: 股市資料的筆數
'*    intDayNo  :
'*    blnIsP    : 要計算的天數
'* 輸出參數: 無
'* 版    本:
'* 2.00  20050604 Earvin   New
'***************************************************************************************************
Public Sub subAverage2(ByRef udtStock() As StockData, _
                      ByRef udtIndex() As IndexData, _
                      ByVal intStockNo As Integer, _
                      ByVal intDayNo As Integer, _
                      ByVal blnIsP As Boolean)
    Dim i As Integer
    Dim j As Integer
    Dim sngAverage As Single
    
    On Error GoTo ERR_HANDLE
    
    j = 1
    While j <= intStockNo
        sngAverage = 0
        If j < intDayNo Then
            For i = 1 To j
                If blnIsP Then
                    sngAverage = sngAverage + udtStock(i).sngEndprice
                Else
                    sngAverage = sngAverage + udtStock(i).sngVol
                End If
            Next
            sngAverage = sngAverage / j
        Else
            For i = j To j - intDayNo + 1 Step -1
                If blnIsP Then
                    sngAverage = sngAverage + udtStock(i).sngEndprice
                Else
                    sngAverage = sngAverage + udtStock(i).sngVol
                End If
            Next
            sngAverage = sngAverage / intDayNo
        End If
        
        If blnIsP Then
            If intDayNo = 3 Then
                udtIndex(j).sngP3 = sngAverage
            ElseIf intDayNo = 4 Then
                udtIndex(j).sngP4 = sngAverage
            ElseIf intDayNo = 5 Then
                udtIndex(j).sngP5 = sngAverage
            ElseIf intDayNo = 6 Then
                udtIndex(j).sngP6 = sngAverage
            ElseIf intDayNo = 8 Then
                udtIndex(j).sngP8 = sngAverage
            ElseIf intDayNo = 10 Then
                udtIndex(j).sngP10 = sngAverage
            ElseIf intDayNo = 12 Then
                udtIndex(j).sngP12 = sngAverage
            ElseIf intDayNo = 20 Then
                udtIndex(j).sngp20 = sngAverage
            ElseIf intDayNo = 24 Then
                udtIndex(j).sngP24 = sngAverage
            ElseIf intDayNo = 30 Then
                udtIndex(j).sngP30 = sngAverage
            ElseIf intDayNo = 60 Then
                udtIndex(j).sngP60 = sngAverage
            ElseIf intDayNo = 72 Then
                udtIndex(j).sngP72 = sngAverage
            ElseIf intDayNo = 120 Then
                udtIndex(j).sngP120 = sngAverage
            ElseIf intDayNo = 144 Then
                udtIndex(j).sngP144 = sngAverage
            ElseIf intDayNo = 240 Then
                udtIndex(j).sngP240 = sngAverage
            ElseIf intDayNo = 288 Then
                udtIndex(j).sngP288 = sngAverage
            End If
        Else
            If intDayNo = 3 Then
                udtIndex(j).sngVol3 = sngAverage
            ElseIf intDayNo = 5 Then
                udtIndex(j).sngVol5 = sngAverage
            ElseIf intDayNo = 6 Then
                udtIndex(j).sngVol6 = sngAverage
            ElseIf intDayNo = 10 Then
                udtIndex(j).sngVol10 = sngAverage
            ElseIf intDayNo = 12 Then
                udtIndex(j).sngVol12 = sngAverage
            ElseIf intDayNo = 20 Then
                udtIndex(j).sngVol20 = sngAverage
            ElseIf intDayNo = 24 Then
                udtIndex(j).sngVol24 = sngAverage
            End If
        End If
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subAverage2()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the MAP-distance value
'*           Formula: MAPDis = P(N1) - P(N2)
'*                    N1, N2：表示天數，與選擇之MAP值天數相同
'*           目前是用MAP6 - MAP24來測試
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081222 調整格式及增加註解
'***************************************************************************************************
Public Sub subMAPDis(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim intUp As Integer
    
    On Error GoTo ERR_HANDLE
    
    j = 1
    
    While j <= intStockNo
        With udtIndex(j)
            .sngMAPDis = .sngP60 - udtStock(j).sngEndprice
        End With
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subMAPDis()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the BIAS value.
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intBiasNo  欲計算指標之天數
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subBias(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                   ByVal intStockNo As Integer, _
                   ByVal intBiasNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim sngAverage As Single
    
    On Error GoTo ERR_HANDLE
        
    j = 1
    
    While j <= intStockNo
        sngAverage = 0
        If j < intBiasNo Then
            For i = 1 To j
                sngAverage = sngAverage + udtStock(i).sngEndprice
            Next
            sngAverage = sngAverage / j
        Else
            For i = j To j - intBiasNo + 1 Step -1
                sngAverage = sngAverage + udtStock(i).sngEndprice
            Next
            sngAverage = sngAverage / intBiasNo
        End If
    
        udtIndex(j).sngBias = (udtStock(j).sngEndprice - sngAverage) / sngAverage * 100
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
'    MsgBox "[GCFR_Index_Module.subBias()], Err-Number= " & Err.Number & _
'    ", Err-Desc= " & _
'    ", Err-Data= " & udtStock(i).sngDate & _
'    Err.Description, vbOKOnly
    
    Debug.Print "[GCFR_Index_Module.subBias()], Err-Number= " & Err.Number & _
    ", Err-Desc= " & _
    ", Err-Data= " & udtStock(i).sngDate & _
    Err.Description, vbOKOnly
    
    Resume Next
    
End Sub


'***************************************************************************************************
'* 說    明: Calculate the K and D value.
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intKDNo    欲計算之KD值
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subKD(ByRef udtStock() As StockData, _
                 ByRef udtIndex() As IndexData, _
                 ByVal intStockNo As Integer, _
                 ByVal intKDNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim sngPreK As Single
    Dim sngPreD As Single
    Dim sngRSV As Single
    Dim sngK As Single
    Dim sngD As Single
    Dim sngMax As Single
    Dim sngMin As Single
    
    On Error GoTo ERR_HANDLE
    
    sngPreK = 50: sngPreD = 50
    j = 1
    
    While j <= intStockNo
        sngMax = -1: sngMin = 999999
        If j <= intKDNo Then
            For i = 1 To j
                If sngMax < udtStock(i).sngHighPrice Then
                    sngMax = udtStock(i).sngHighPrice
                End If
                If sngMin > udtStock(i).sngLowPrice Then
                    sngMin = udtStock(i).sngLowPrice
                End If
            Next
        Else
            For i = j To j - intKDNo + 1 Step -1
                If sngMax < udtStock(i).sngHighPrice Then
                    sngMax = udtStock(i).sngHighPrice
                End If
                If sngMin > udtStock(i).sngLowPrice Then
                    sngMin = udtStock(i).sngLowPrice
                End If
            Next
        End If
        If sngMax <> sngMin Then
            sngRSV = (udtStock(j).sngEndprice - sngMin) / (sngMax - sngMin) * 100
        Else
            sngRSV = 50
        End If
        sngK = sngPreK * 2 / 3 + sngRSV / 3
        sngD = sngPreD * 2 / 3 + sngK / 3
        sngPreK = sngK
        sngPreD = sngD
        udtIndex(j).sngK = sngK
        udtIndex(j).sngD = sngD
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subKD()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the PSY value
'*           Formula: PSY(N) = A / N * 100%
'*                    N: 表示天數
'*                    A: 表示N天中股價上漲的天數
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intPSYNo   欲計算的心理線天數
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subPSY(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer, _
                  ByVal intPSYNo As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim intUp As Integer
    Dim sngPSY As Single
    
    On Error GoTo ERR_HANDLE
    
    j = 1
    
    While j <= intStockNo
        intUp = 0
        If j <= intPSYNo Then
            sngPSY = 50
        Else
            For i = j To j - intPSYNo + 1 Step -1
                If (udtStock(i).sngEndprice - udtStock(i - 1).sngEndprice) > 0 Then
                    intUp = intUp + 1
                End If
            Next
            sngPSY = intUp / intPSYNo * 100
        End If
        udtIndex(j).sngPSY = sngPSY
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subPSY()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the Willams value.
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intWMSNo   欲計算之WMS天數
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subWMS(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                   ByVal intStockNo As Integer, _
                   ByVal intWMSNo As Integer)
    Dim sngWMS As Single
    Dim i As Integer
    Dim j As Integer
    Dim sngMax As Single
    Dim sngMin As Single
    
    On Error GoTo ERR_HANDLE
    
    j = 1
    While j <= intStockNo
        sngMax = -1:   sngMin = 99999999
        If j <= intWMSNo Then
            For i = 1 To j
                If sngMax < udtStock(i).sngHighPrice Then
                    sngMax = udtStock(i).sngHighPrice
                End If
                If sngMin > udtStock(i).sngLowPrice Then
                    sngMin = udtStock(i).sngLowPrice
                End If
            Next
        Else
            For i = j To j - intWMSNo + 1 Step -1
                If sngMax < udtStock(i).sngHighPrice Then
                    sngMax = udtStock(i).sngHighPrice
                End If
                If sngMin > udtStock(i).sngLowPrice Then
                    sngMin = udtStock(i).sngLowPrice
                End If
            Next
        End If
        If sngMax <> sngMin Then
            ' 920719 marked
'            sngWMS = (sngMax - udtstock(j).sngEndprice) / (sngMax - sngMin) * 100
            sngWMS = (udtStock(j).sngEndprice - sngMin) / (sngMax - sngMin) * 100
'            If sngWMS < 0 Or sngWMS > 100 Then
'                Debug.Print udtStock(j).sngDate & ":" & sngWMS
'            End If
        Else
            sngWMS = 50
        End If
        udtIndex(j).sngWMS = sngWMS
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subWMS()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


'***************************************************************************************************
'* 說    明: Calculate the RSI value.
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intRSINo   欲計算的RSI天數
'*           blnIsShort 表示該RSI 是否為短期RSI （因為在顯示RSI 圖形時會同時顯示長短期之RSI 線，
'*                      故計算時需知此次計算是算長期或短期之RSI 以存入相對欄位
'* 輸出參數: 無
'* 版    本: 2.00 20080916 新增
'*           2.01 20081209 調整格式及增加註解
'***************************************************************************************************
Public Sub subRSI(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                   ByVal intStockNo As Integer, _
                   ByVal intRSINo As Integer, _
                   ByVal blnIsShort As Boolean)
    Dim strRSIfile As String
    Dim sngRSI As Single
    Dim sngUp As Single
    Dim sngDown As Single
    Dim sngDiff As Single
    Dim intRSILow As Integer
    Dim intRSIHigh As Integer
    Dim i As Integer
    Dim j As Integer

    On Error GoTo ERR_HANDLE
    
    j = 1
    
    While j <= intStockNo
        sngUp = 0: sngDown = 0
        If j <= intRSINo Then
            For i = 2 To j
                sngDiff = udtStock(i).sngEndprice - udtStock(i - 1).sngEndprice
                If sngDiff > 0 Then
                    sngUp = sngUp + sngDiff
                Else
                    sngDown = sngDown + Abs(sngDiff)
                End If
            Next
            sngUp = sngUp / j
            sngDown = sngDown / j
        Else
            For i = j To j - intRSINo + 1 Step -1
                sngDiff = udtStock(i).sngEndprice - udtStock(i - 1).sngEndprice
                If sngDiff > 0 Then
                    sngUp = sngUp + sngDiff
                Else
                    sngDown = sngDown + Abs(sngDiff)
                End If
            Next
            sngUp = sngUp / intRSINo
            sngDown = sngDown / intRSINo
        End If
        
        If j <> 1 And (sngUp + sngDown) <> 0 Then
            sngRSI = sngUp / (sngUp + sngDown) * 100
        Else
            sngRSI = 50
        End If
        
        If blnIsShort Then
            udtIndex(j).sngRSI_S = sngRSI
        Else
            udtIndex(j).sngRSI_L = sngRSI
        End If
        j = j + 1
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subRSI()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


Public Function KDUnderTheValueAndCrossUp(ByVal prevK As Single, ByVal prevD As Single, ByVal currK As Single, ByVal currD As Single, ByVal theValue As Single) As Boolean
    If prevK < prevD And _
        currK > currD And _
        prevK < theValue And _
        prevD < theValue And _
        currK < theValue And _
        currD < theValue Then
        KDUnderTheValueAndCrossUp = True
    Else
        KDUnderTheValueAndCrossUp = False
    End If
End Function


'***************************************************************************************************
'* 說    明: Calculate the EMA value (平滑移動平均線)
'*           Formula: EMAt = EMAt-1 *[(N-1)/N] + Pt * (1/N)
'*                    N : 表示天數
'*                    Pt: 表示當天收盤價
'*
'*                    EMAt = EMAt-1 + A * (Pt -EMAt-1)
'*                    A : 表示平滑因子 ==> 常用 2/(N+1)
'*
'* 輸入參數: udtStock   股市資料
'*           udtIndex   股市指標資料
'*           intStockNo 股市資料的筆數
'*           intEMANo   欲計算的EMA天數
'* 輸出參數: 無
'* 版    本: 1.00 20100819 新增
'***************************************************************************************************
Public Sub subEMA(ByRef udtStock() As StockData, _
                  ByRef udtIndex() As IndexData, _
                  ByVal intStockNo As Integer, _
                  ByVal intEMANo As Integer)
    Dim i As Integer
    Dim sngNowEMA As Single
    Dim sngPrevEMA As Single
    
    On Error GoTo ERR_HANDLE
    
    ' 第1筆資料的EMA值直接設定為當日收盤價
    i = 1
    sngNowEMA = udtStock(i).sngEndprice
    sngPrevEMA = udtStock(i).sngEndprice
    udtIndex(i).sngEMA = sngNowEMA
    
    While i < intStockNo
        i = i + 1
        sngNowEMA = (sngPrevEMA * (intEMANo - 1) / intEMANo) + (udtStock(i).sngEndprice / intEMANo)
        udtIndex(i).sngEMA = sngNowEMA
    Wend
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[GCFR_Index_Module.subEMA()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub






'--- LoadDFTS ------------------------------------------------------------------------------------------------------
'''Public Sub Load_DFTS()
'''    Dim InputFileTrend_Day As String
'''    Dim InputFileQM_Day As String
'''    Dim InputFileDIFF_Day As String
'''    Dim InputFileDC_Day As String
'''
'''    Dim intInputFileTrendNo_Day As Integer
'''    Dim intInputFileQMNo_Day As Integer
'''    Dim intInputFileDIFFNo_Day As Integer
'''    Dim intInputFileDCNo_Day As Integer
'''
'''    Dim InputFileWLW_Week As String
'''    Dim InputFileWCY_Week As String
'''    Dim InputFileWLST_Week As String
'''
'''    Dim intInputFileWLWNo_Week As Integer
'''    Dim intInputFileWCYNo_Week As Integer
'''    Dim intInputFileWLSTNo_Week As Integer
'''    Dim ReadLine As String
'''    Dim i, j, intCount As Long
'''
'''
''''==== DFTS_Day.dat loading ====
'''    InputFileTrend_Day = "data\OPData\trend\weight.opd"
'''    InputFileQM_Day = "data\OPData\qm\weight.opd"
'''    InputFileDIFF_Day = "data\OPData\diff\weight.opd"
'''    InputFileDC_Day = "data\OPData\dc\weight.opd"
'''
'''    intInputFileTrendNo_Day = 1
'''    intInputFileQMNo_Day = 2
'''    intInputFileDIFFNo_Day = 3
'''    intInputFileDCNo_Day = 4
'''
'''    Open InputFileTrend_Day For Binary As #intInputFileTrendNo_Day    ' Opening the DFTS_Day.dat file
'''    Open InputFileQM_Day For Binary As #intInputFileQMNo_Day    ' Opening the DFTS_Day.dat file
'''    Open InputFileDIFF_Day For Binary As #intInputFileDIFFNo_Day    ' Opening the DFTS_Day.dat file
'''    Open InputFileDC_Day For Binary As #intInputFileDCNo_Day    ' Opening the DFTS_Day.dat file
'''        '--------
'''    '    intCount = LOF(intInputFileNo_Day) / Len(gudtDFTSTrend)
'''
'''    ReDim gudtTrend(gintDayIndex)                    'DFTS模組會使用到
'''    ReDim gudtQM(gintDayIndex)
'''    ReDim gudtDIFF(gintDayIndex)
'''    ReDim gudtDC(gintDayIndex)
'''    Get #intInputFileTrendNo_Day, , gudtTrend ' read the data of DFTS_Day.dat
'''    Get #intInputFileQMNo_Day, , gudtQM ' read the data of DFTS_Day.dat
'''    Get #intInputFileDIFFNo_Day, , gudtDIFF ' read the data of DFTS_Day.dat
'''    Get #intInputFileDCNo_Day, , gudtDC ' read the data of DFTS_Day.dat
'''    Close #intInputFileTrendNo_Day
'''    Close #intInputFileQMNo_Day
'''    Close #intInputFileDIFFNo_Day
'''    Close #intInputFileDCNo_Day
'''
'''    For i = 1 To gintDayIndex
'''        gudtIndexDay(i).sngTrend = gudtTrend(i).sngTrend
'''        gudtIndexDay(i).sngQM = gudtQM(i).sngQM
'''        gudtIndexDay(i).sngDIFF1 = gudtDIFF(i).sngDIFF1
'''        gudtIndexDay(i).sngDIFF2 = gudtDIFF(i).sngDIFF2
'''        gudtIndexDay(i).sngDC = gudtDC(i).sngDC
'''    Next
'''
'''
'''    '==== DFTS_Week.dat loading ====
'''    InputFileWLW_Week = "data\OPData\wlw\weight.opd"
'''    InputFileWCY_Week = "data\OPData\wcy\weight.opd"
'''    InputFileWLST_Week = "data\OPData\wlst\weight.opd"
'''
'''    intInputFileWLWNo_Week = 5
'''    intInputFileWCYNo_Week = 6
'''    intInputFileWLSTNo_Week = 7
'''
'''    Open InputFileWLW_Week For Binary As #intInputFileWLWNo_Week    ' Opening the DFTS_Week.dat file
'''    Open InputFileWCY_Week For Binary As #intInputFileWCYNo_Week    ' Opening the DFTS_Week.dat file
'''    Open InputFileWLST_Week For Binary As #intInputFileWLSTNo_Week    ' Opening the DFTS_Week.dat file
'''
'''    '----此處會重設 gsngEndIndex ----
'''    intCount = LOF(intInputFileWLWNo_Week) / Len(gudtDFTSWLW)
'''
'''    ReDim gudtWLW(intCount + 1)                  'DFTS模組會使用到
'''    ReDim gudtWCY(intCount + 1)                  'DFTS模組會使用到
'''    ReDim gudtWLST(intCount + 1)                  'DFTS模組會使用到
'''
'''
'''    Get #intInputFileWLWNo_Week, , gudtWLW ' read the data of DFTS_Week.dat
'''    Get #intInputFileWCYNo_Week, , gudtWCY ' read the data of DFTS_Week.dat
'''    Get #intInputFileWLSTNo_Week, , gudtWLST ' read the data of DFTS_Week.dat
'''    Close #intInputFileWLWNo_Week
'''    Close #intInputFileWCYNo_Week
'''    Close #intInputFileWLSTNo_Week
'''    '------- 將相對應之日資料向前補滿
'''    j = 1
'''    i = 1
'''
'''
'''    While j <= intCount And i <= gintDayIndex
'''        If (gudtWLW(j + 1).sngDate > gudtStockDay(i).sngDate) Then
'''            gudtIndexDay(i).sngWLST = gudtWLST(j).sngWLST
'''            gudtIndexDay(i).sngWCY = gudtWCY(j).sngWCY
'''            gudtIndexDay(i).sngWLW = gudtWLW(j).sngWLW
'''            i = i + 1
'''        Else
'''            j = j + 1
'''            gudtIndexDay(i).sngWLST = gudtWLST(j).sngWLST
'''            gudtIndexDay(i).sngWCY = gudtWCY(j).sngWCY
'''            gudtIndexDay(i).sngWLW = gudtWLW(j).sngWLW
'''            i = i + 1
'''        End If
'''    Wend
'''
'''    j = 1
'''
'''    While (j <= intCount + 1)
'''        gudtIndexWeek(j).sngWLST = gudtWLST(j).sngWLST
'''        gudtIndexWeek(j).sngWCY = gudtWCY(j).sngWCY
'''        gudtIndexWeek(j).sngWLW = gudtWLW(j).sngWLW
'''        j = j + 1
'''    Wend
'''   ' MsgBox "作業完成"
'''
'''    Exit Sub
'''
'''End Sub



