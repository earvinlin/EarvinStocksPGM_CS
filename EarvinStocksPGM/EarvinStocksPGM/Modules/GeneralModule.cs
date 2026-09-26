using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace EarvinStocksPGM.Modules
{
    public static class GeneralModule
    {
        // 灰聚類模型之聚類方式
        enum GRG_TYPE
        {
            GRG_LTB = 0,    // 望大Larger-the-Better (LTB)
            GRG_STB = 1,    // 望小Smaller-the-Better (STB)
            GRG_NTB = 2     // 適中Nominal-the-Best (NTB)
        }

        // GRG模型中相關屬性資訊
        public struct SelectAttributesSettings {
            public String IndexName;    // 指標名稱
            public int IndexDay;        // 指標天數
            public int GRGType;         // 聚類方式
            public int GRGTarget;       // 聚類方式若為NTB，需指定目標值
            public float weight;        // 指標權重
        }

        // 保存Frame的座標
        public struct FramePoints
        {
            public float frameX;
            public float frameY;
        }

        public const int MAP_UNSELECTED = 0;
        public const int MAP_KBAR = 1;
        public const int MAP_VOLUME = 2;
        public const int MAP_BIAS = 3;
        public const int MAP_WMS = 4;
        public const int MAP_PSY = 5;
        public const int MAP_SRSI = 6;
        public const int MAP_LRSI = 7;
        public const int MAP_RSI = 8;
        public const int MAP_K = 9;
        public const int MAP_D = 10;
        public const int MAP_KD = 11;
        public const int MAP_DIF = 12;
        public const int MAP_DIF_MACD = 13;
        public const int MAP_MACD = 14;
        public const int MAP_SECTORS = 15;

        //-- Keep the Index's Days --//
        public static int BIASDay = 10;
        public static int WMSDay = 10;
        public static int PSYDay = 10;
        public static int SRSIDay = 6;
        public static int LRSIDay = 20;
        public static int KDay = 9;
        public static int DDay = 9;
        public static int DIFDay = 12;
        public static int MACDDay = 26;
        public static int OSCDay = 9;

        // 記錄每個FRAME選擇顯示的資料(最多只能選9個；第1個一定是MAP_K)
        public static int[] SelectShowMapOnFrames = new int[9];

        public class HighLowValues
        {
            public double highValue { get; set; }
            public double lowValue { get; set; }
        }

        public static HighLowValues GetHighLowValue(StockData[] sd, IndexData[] idx, int startIndex, int displayCount, int type)
        {
            double highValue = double.MinValue;
            double lowValue = double.MaxValue;

            switch (type)
            {
                case MAP_KBAR :
                    for (int i = startIndex; i < (startIndex + displayCount); i++)
                    {
                        if (highValue < sd[i].HighPrice)
                            highValue = sd[i].HighPrice;
                        if (lowValue > sd[i].LowPrice)
                            lowValue = sd[i].LowPrice;
                    }
                    break;

                case MAP_VOLUME :
                    for (int i = startIndex; i < (startIndex + displayCount); i++)
                    {
                        if (highValue < sd[i].Volume)
                            highValue = sd[i].Volume;
                        if (lowValue > sd[i].Volume)
                            lowValue = sd[i].Volume;
                    }
                    break;

                case MAP_BIAS :
                    for (int i = startIndex; i < (startIndex + displayCount); i++)
                    {
                        if (highValue < idx[i].BIAS)
                            highValue = idx[i].BIAS;
                        if (lowValue > idx[i].BIAS)
                            lowValue = idx[i].BIAS;
                    }
                    highValue = Math.Round(highValue, 2);
                    lowValue = Math.Round(lowValue, 2);

                    if (highValue >= 0 && lowValue >= 0)
                    {
                        highValue = Math.Ceiling(Math.Abs(highValue));
                        lowValue = -highValue;
                    } 
                    else if (highValue >= 0 && lowValue < 0)
                    {
                        if (Math.Abs(highValue) >= Math.Abs(lowValue))
                        {
                            highValue = Math.Ceiling(Math.Abs(highValue));
                            lowValue = -highValue;
                        }
                        else
                        {
                            highValue = Math.Ceiling(Math.Abs(lowValue));
                            lowValue = -highValue;
                        }
                    }
                    else 
                    {
                        highValue = Math.Ceiling(Math.Abs(lowValue));
                        lowValue = -highValue;
                    }
                    break;

                case MAP_MACD:
                    double[] difValues = idx[startIndex..(startIndex + displayCount - 1)].Select(d => d.DIF).ToArray();
                    double[] macdValues = idx[startIndex..(startIndex + displayCount - 1)].Select(d => d.MACD).ToArray();
                    double[] oscValues = idx[startIndex..(startIndex + displayCount - 1)].Select(d => d.OSC).ToArray();
                    highValue = Math.Max(difValues.Max(), Math.Max(macdValues.Max(),oscValues.Max()));
                    lowValue = Math.Min(difValues.Min(), Math.Min(macdValues.Min(),oscValues.Min()));
                    break;

                case MAP_WMS:
                case MAP_PSY:
                case MAP_SRSI:
                case MAP_LRSI:
                case MAP_K:
                case MAP_D:
                case MAP_KD:
                    highValue = 100;
                    lowValue = 0;
                    break;

                case MAP_SECTORS:
                    highValue = 1;
                    lowValue = -1;
                    break;
            }
            HighLowValues values = new HighLowValues();
            values.highValue = highValue;
            values.lowValue = lowValue;

            return values;
        }

        public static int GetSelectFrame(FramePoints[] frmLeft, FramePoints[] frmRight, Point cursorPos, int frameNum)
        {
            int selectFrame = 0;

            for (int i = 0; i < frameNum; i++)
            {
                if ((cursorPos.Y) < frmLeft[i].frameY)
                {
                    break;
                }
                else if ((cursorPos.Y) >= frmLeft[i].frameY && (cursorPos.Y) <= frmLeft[i + 1].frameY)
                {
                    selectFrame = (i + 1);
                    break;
                }
                else if ((cursorPos.Y) > frmLeft[frameNum].frameY)
                {
                    break;
                }
            }
            return selectFrame;
        }
    }
}


/*
 *  -- Sector --
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





 */