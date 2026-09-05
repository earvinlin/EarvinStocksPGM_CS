'***************************************************************************************************************************
'* 說    明:
'*    利用灰聚類方法，預測高低點
'*    為基於研究所論文程式來改良，使更具彈性、移植性及易於維護、除錯與擴充
'* 版    本:
'*    2.00  20080917 Earvin   新增
'***************************************************************************************************************************
Option Explicit

Public Const PARAM_MAX_CNT = 8         ' 最多可選用的聚類的attributes數目
Public CLUSTER_MAX As Integer          ' 望大
Public CLUSTER_MIN As Integer          ' 望小
Public CLUSTER_OBJ As Integer          ' 望目
Public CLUSTER_FA As Single            ' 聚類的閥值
Public CLUSTER_RO As Single            ' 聚類的關係係數
Public Const HIGH_SIGNAL = 1           ' 聚類高點
Public Const LOW_SIGNAL = -1           ' 聚類低點


Public Const CHKBOX_SELECTED = 1       ' the form's checkbox is selected

''Public totalHighAttrCount As Integer   ' 記錄聚類之高點選用attribute的數目
''Public totalHighWightValue As Integer  ' 記錄聚類之高點選用attribute的總權重
''Public totalLowAttrCount As Integer    ' 記錄聚類之低點選用attribute的數目
''Public totalLowWightValue As Integer   ' 記錄聚類之低點選用attribute的總權重

Public paramMax(PARAM_MAX_CNT) As Single  ' 記錄聚類各參數之資料前處理前的最大值，以供回復原值之用
Public paramMin(PARAM_MAX_CNT) As Single


' 定義呼叫切割區間函式所需的傳入參數
Public Type SplitParameter
    indexDay1 As String ' 目前只實做MAP、KD指標
    indexDay2 As String ' 目前只實做MAP、KD指標
    splitType As String ' 目前只可選擇MAP；KD不確定
    stockType As String ' 股票資料：選擇 日/週/月線資料
End Type

' 20151107 儲存最高值與最低值
Public Type HighAndLowValues
    HighValue As Single
    LowValue As Single
End Type




'***************************************************************************************************************************
'* 定義繪製各項指標的名稱、天數及繪製方式
'*    strIndexName 指標名稱
'*    intIndexDay 指標天數
'*    intDrawType 繪製方式
'***************************************************************************************************************************
Public Type DrawIndexInfo
    strIndexName As String
    intIndexDay As Integer
    intDrawType As String
End Type


'***************************************************************************************************************************
'* 儲存要顯示的指標資料
'*    sngIndex
'***************************************************************************************************************************
Public Type DrawIndexData
    sngIndex(MAX_FRAME) As Single
End Type


'***************************************************************************************************************************
'* Define the type of individual data, include date, price, quantity
'*    sngDate 交易日期
'*    sngStartprice 開盤價
'*    sngHighPrice 最高價
'*    sngLowPrice 最低價
'*    sngEndprice 收盤價
'*    sngVol 成交量
'*    sngAcc 融資
'*    sngTome 融券
'*--- 20180527 新增下列4個指標 ---------------
'*    sngForeignStock 外資庫存
'*    sngSitAndCbStock 投信庫存
'*    sngSelfEmployedStock 自營商庫存
'*    sngLegalPersonStock 法人庫存
'***************************************************************************************************************************
Public Type StockData
    sngDate As Single
    sngStartprice As Single
    sngHighPrice As Single
    sngLowPrice As Single
    sngEndprice As Single
    sngVol As Single
    sngAcc As Single
    sngTome As Single
    ' 20180526
    sngForeignStock As Single
    sngSitAndCbStock As Single
    sngSelfEmployedStock As Single
    sngLegalPersonStock As Single
End Type


'***************************************************************************************************************************
'* Define the type of stock index
'***************************************************************************************************************************
Public Type IndexData
    sngMAP0 As Single
    sngMAP1 As Single
    sngMAP2 As Single
    sngMAP3 As Single
    sngMAP4 As Single
    sngMAP5 As Single
    sngMAP6 As Single
    sngMAP7 As Single
         
    sngMAV1 As Single
    sngMAV2 As Single
    sngMAV3 As Single
    sngMAV4 As Single
    sngMAV5 As Single
         
    sngDate As Single
    sngK As Single
    sngD As Single
    sngPSY As Single
    sngWMS As Single
    sngRWMS As Single          ' WMS的反轉
    sngRSI_S As Single         ' 台灣RSI
    sngRSI_L As Single         ' 台灣RSI
    sngStochRSI As Single      ' StochRSI
    sngWRSI As Single          ' Wilder's RSI
    sngMAWRSI As Single
    sngMACD As Single
    sngBias As Single
    sngDIF As Single
    sngDIF_MACD As Single
    sngDI_P As Single
    sngDI_N As Single
    sngADX As Single
    sngDIF_MACD_AREA As Single
         
    sngTrend   As Single
    sngQM      As Single
    sngDIFF1   As Single
    sngDIFF2   As Single
    sngDC      As Single
    sngWLST    As Single
    sngWCY     As Single
    sngWLW     As Single
         
    '=======================================================
    '* GCFR Model 新增使用的變數    --- 20050620 STA ---
    '=======================================================
    sngP3 As Single
    sngP4 As Single
    sngP5 As Single
    sngP6 As Single
    sngP8 As Single
    sngP10 As Single
    sngP12 As Single
    sngp20 As Single
    sngP24 As Single
    sngP30 As Single
    sngP60 As Single
    sngP72 As Single
    sngP120 As Single
    sngP144 As Single
    sngP240 As Single
    sngP288 As Single
    sngVol3 As Single
    sngVol5 As Single
    sngVol6 As Single
    sngVol10 As Single
    sngVol12 As Single
    sngVol20 As Single
    sngVol24 As Single
    sngGRG As Single            ' 儲存計算出來的GRG
    sngSector As Single         ' 聚類sector
    sngMAPDis As Single         ' 920728: Use distance to calculate
    sngVR As Single             ' 921217: Add VR index
    sngDiffOp1 As Single        ' 930320: Add the 6 index
    sngDiffOp2 As Single        ' 930320: Add the 6 index
    sngQMOp As Single           ' 930320: Add the 6 index
    sngTrendOp As Single        ' 930320: Add the 6 index
    sngCyOp As Single           ' 930320: Add the 6 index
    sngLstOp As Single          ' 930320: Add the 6 index
    sngLwOp As Single           ' 930320: Add the 6 index
    sngSignal As Single         ' 930421: Add Signal index
    sngHold As Single           ' 930428: Add Hold and Profit
    sngProfit As Single
    sngBiasS As Single
    sngBiasM As Single
    sngBiasL As Single
    sngBullBear As Single           ' 記錄目前是操作「做多(1)」、「做空(-1)」或「等待(0)」
    sngMASlope As Single            ' MA的斜率
    sngMADistance As Single         ' 記錄今日的MA 減 昨日MA的值
    sngLSBias As Single             ' 長天數Bias - 短天數Bias的差值
    sngLBias As Single              ' 長天數Bias
    sngSBias As Single              ' 短天數Bias
    sngBullList As Single           ' 形成多頭排列後持續天數
    sngBearList As Single           ' 形成空頭排列後持續天數
    sngInRange As Single            ' 盤整
    sngHighLowArea As Single
    sngUpDownDays As Single         ' 當日收盤價>60MAP +1天；當日收盤價<60MAP -1天；當日收盤價=60MAP value=0
    sngSTAandSTPMarked As Single    ' 標記每個波段操作的起迄點(起=1, 迄=2, 起+迄=3, others=0)
    sngSTARealProfit(5) As Single   ' 第1~5口實際損益
    sngStep As Single               ' 操作(買賣)次數
    '--- For Strategy3 使用 -------------
    sngBuyStep As Single
    sngSellStep As Single
    '--- 增加毛利率、毛損率 -------------
    sngGrossGainRate(5) As Single
    sngGrossLossRate(5) As Single
    '--- 損益率 -------------------------
    sngProfitRate As Single
    '--- 顯示畫面用 ---------------------
'    sngIndexValue As Single
'    sngIndexValue1 As Single
'    sngIndexValue2 As Single
    sngLineValue(5) As Single
    sngBarValue(5) As Single
    
    '=======================================================
    '* GCFR Model 新增使用的變數    --- 20050620 END ---
    '=======================================================
       
    ' New Indexes
    sngEMA As Single
End Type


'***************************************************************************************************************************
'* 儲存使用聚類參數及聚類結果(GRG)
'*    sngDate   指數資料日期
'*    sngParam  選擇的技術指標值
'*    sngGRG    GRG值
'*    sngMax    分析區間sngParam資料前處理前的最大值(主要用來將來將技術指標值還原之用)
'*    sngMin    分析區間sngParam資料前處理前的最小值(主要用來將來將技術指標值還原之用)
'*    isCluster 此次聚類是否使用
'***************************************************************************************************************************
Public Type udtGRGIndex2
    sngDate As Single
    sngParam() As Single
    sngGRG As Single
    sngMax As Single
    sngMin As Single
    isCluster As Boolean
End Type

'***************************************************************************************************************************
'* 記錄聚類的選用之個別參數資訊
'*    isInUse        此資料是否使用中(true: 使用; false: 未使用)
'*    paramName      選擇的技術指標
'*    clusterMethod  聚類方式(望大、望小及望目)
'*    paramDay       天數
'*    paramTarget    望目值
'*    paramWeight    權重
'***************************************************************************************************************************
Public Type ClusterParam
'    isInUse As Boolean
    paramName As String
    clusterMethod As Single
    paramDay As Integer
    paramTarget As Double
    paramWeight As Single
End Type

'***************************************************************************************************************************
'* 記錄聚類的選用參數資訊
'*    selectParamCnt     選擇的參數總個數
'*    selectParamsWeight 選擇的參數總權重
'*    paramsMax
'*    paramsMin
'*    clustParams        選擇的個別參數資訊
'***************************************************************************************************************************
Public Type ClusterParamGroup
    selectParamsCnt As Integer
    selectParamsWeight As Single
    paramsMax() As Single
    paramsMin() As Single
    clustParams() As ClusterParam
End Type

'***************************************************************************************************************************
'* 記錄聚類的的結果
'*    clstIsInUse       此資料是否使用中(true: 使用; false: 未使用)
'*    sectorValue       此次聚類分析是高點 或 低點
'*    sectorBegDate     此次聚類分析之起始日期
'*    sectorEndDate     此次聚類分析之結束日期
'*    ResultParamsValue 選用的參數的聚類結果最後的值
'***************************************************************************************************************************
Public Type ClusterResult
    isInUse As Boolean
    sectorValue As Single
    sectorBegDate As Single
    sectorEndDate As Single
    returnGRGIndex As udtGRGIndex2
End Type

'***************************************************************************************************************************
'* 記錄預測區間最高點日期、最低點日期、最高點指數、最低點指數
'*    isInUse 此資料是否使用中(true: 使用; false: 未使用)
'*    sectStartDate  區間起始日期
'*    sectEndDate    區間結束日期
'*    sectHighDate   區間最高點日期
'*    sectLowDate    區間最低點日期
'*    sectHighPoints 區間最高點指數
'*    sectLowPoints  區間最低點指數
'***************************************************************************************************************************
Public Type SectorInfo
    isInUse As Boolean
    sectStartDate As Single
    sectEndDate As Single
    sectHighDate As Single
    sectLowDate As Single
    sectHighPoints As Single
    sectLowPoints As Single
End Type

'***************************************************************************************************************************
'* 記錄預測區間的股票及技術指標資料
'*    sectStockData  區間股票資料
'*    sectStockIndex 區間技術指標資料
'*    sectStockCount 區間資料筆數
'***************************************************************************************************************************
Public Type SectorStockIndex
    sectStockData() As StockData
    sectStockIndex() As IndexData
    sectStockCount As Integer
End Type



'***************************************************************************************************************************
'* 儲存使用聚類參數及聚類結果(GRG)
'*    sngDate 指數資料日期
'*    sngOrgValue 選擇的技術指標值
'*    sngvalue 前處理後的技術指標值
'***************************************************************************************************************************
Public Type udtPreProcessData
    sngDate As Single
    sngOrgValue() As Single
    sngValue() As Single
End Type



'***************************************************************************************************************************
'* 記錄聚類的選用之個別參數資訊
'*    paramName 選擇的技術指標
'*    useMethod 聚類方式(望大、望小及望目)
'*    paramDay 天數
'*    paramTarget 望目值
'*    paramWeight 權重
'***************************************************************************************************************************
Public Type GeryClusterParam
    paramName As String
    useMethod As Single
    paramDay As Integer
    paramTarget As Double
    paramWeight As Single
End Type



'***************************************************************************************************************************
'* 記錄聚類的選用參數資訊
'*    selectParamCnt 選擇的參數總個數
'*    selectParamsWeight 選擇的參數總權重
'*    paramsMax
'*    paramsMin
'*    clustParams 選擇的個別參數資訊
'***************************************************************************************************************************
Public Type GeryClusterParamGRP
    selectParamsCnt As Integer
    selectParamsWeight As Single
    paramsMax() As Single
    paramsMin() As Single
    clustParams() As GeryClusterParam
End Type



'***************************************************************************************************************************
'* 記錄聚類的的結果
'*    clstIsInUse       : 此資料是否使用中(true: 使用; false: 未使用)
'*    ResultParamsValue : 選用的參數的聚類結果最後的值
'***************************************************************************************************************************
Public Type GeryClusterResult
    isInUse As Boolean
    returnGRGIndex As udtGRGIndex2
End Type




'--------------------------------------------------------------------------------------------------
' 以下可能不需要用 -- 2010.08.11 --
'--------------------------------------------------------------------------------------------------

Public Type DFTSType
    sngDate As Single
    sngTrend As Single
    sngQM As Single
    sngDIFF1 As Single
    sngDIFF2 As Single
    sngWLST As Single
    sngWCY As Single
    sngWLW As Single
End Type


'========DFTS==========
Public Type RawTrend
    sngDate As Single
    sngStart As Single
    sngHigh As Single
    sngLow As Single
    sngEnd As Single
    sngVol As Single
    sngTrend As Single
End Type

Public Type QM
    sngDate As Single
    sngStart As Single
    sngHigh As Single
    sngLow As Single
    sngEnd As Single
    sngVol As Single
    sngQM As Single
End Type

Public Type DIFF
    sngDate As Single
    sngStart As Single
    sngHigh As Single
    sngLow As Single
    sngEnd As Single
    sngVol As Single
    sngDIFF1 As Single
    sngDIFF2 As Single
End Type

Public Type WLST
    sngDate As Single
    sngStart As Single
    sngHigh As Single
    sngLow As Single
    sngEnd As Single
    sngVol As Single
    sngWLST As Single
End Type

Public Type WCY
    sngDate As Single
    sngStart As Single
    sngHigh As Single
    sngLow As Single
    sngEnd As Single
    sngVol As Single
    sngWCY As Single
End Type

Public Type WLW
    sngDate As Single
    sngStart As Single
    sngHigh As Single
    sngLow As Single
    sngEnd As Single
    sngVol As Single
    sngWLW As Single
End Type

Public Type DC
    sngDate As Single
    sngStart As Single
    sngHigh As Single
    sngLow As Single
    sngEnd As Single
    sngVol As Single
    sngDC As Single
End Type




'Rcd = Record, W = weekly, Cy = Cylinder
Public Type RcdWCyType
    sngDate As Single
    intSignal As Integer    '1 = Buy, -1 = Sell, 0 = nothing
End Type

Public Type RcdWDType
    sngStartDate As Single
    sngEndDate As Single
    intSignal As Integer
End Type

Public Type RcdQType
    sngDate As Single
    sngVol As Single
End Type


'Oct. 19
Public Type BSType
    sngBDegree As Single
    sngSDegree As Single
End Type
Public Type GreySpaceType
    sngDate As Single
    sngStartprice As Single
    sngHighPrice As Single
    sngLowPrice As Single
    sngEndprice As Single
    sngVol As Single
    sngOperation As Single
End Type

Public Type GreySpaceType_2
    sngDate As Single
    sngStartprice As Single
    sngHighPrice As Single
    sngLowPrice As Single
    sngEndprice As Single
    sngVol As Single
    sngOpe1 As Single
    sngOpe2 As Single
End Type


Public Type Trend
    sngDate As Single
    sngStartprice As Single
    sngHighPrice As Single
    sngLowPrice As Single
    sngEndprice As Single
    sngVol As Single
    intTrendValue As Integer
End Type

Public Type FrameData
    sngHeight As Single
    bytAttribute As Byte
End Type

