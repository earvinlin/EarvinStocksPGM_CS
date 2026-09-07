Private Sub DrawFrameData(ByRef gudtStock() As StockData, ByRef gudtIndex() As IndexData)
    Dim processFrame As Integer
    Dim frameCount As Integer
    
    On Error GoTo ERR_HANDLE
    
    frameCount = frmEarvinStocks.cboFrameNum
    gbytFrameNum = frameCount   ' 還不能拿掉，後面會用到…
    
    For processFrame = 1 To frameCount
        Select Case mudtFrame(processFrame).bytAttribute
            Case mKDmap
                Call Chalk_Line_Map(gudtStock, gudtIndex, mKDmap, 3, processFrame)
            Case mRSImap
                Call Chalk_Line_Map(gudtStock, gudtIndex, mRSImap, 3, processFrame)
            Case mPSYmap
                Call Chalk_Line_Map(gudtStock, gudtIndex, mPSYmap, 2, processFrame)
            Case mEMAmap
                Call Chalk_Line_Map(gudtStock, gudtIndex, mEMAmap, 2, processFrame)
            Case mBiasmap
                Call Chalk_Line_Map(gudtStock, gudtIndex, mBiasmap, 1, processFrame)
            Case mSectorMap
                Call Chalk_Bar_MAP(gudtStock, gudtIndex, mSectorMap, 2, processFrame)
            Case mMACDmap
                Call Chalk_MACD_Map(gudtStock, gudtIndex, processFrame)
            Case mQuantityMap ' 20100415 , Has problem ...
                Call Chalk_Quantity_Map(gudtStock, gudtIndex, processFrame)
         
            '--- 下面有空再改~~ 應該著重在預測分析模型才對!! 又不是要練寫程式說~~ (20100416) ----
            Case mKmap
                Call Chalk_K_Map(gudtStock, gudtIndex, processFrame)
            Case mStochRSImap
                Call Chalk_StochRSI_Map(gudtStock, gudtIndex, processFrame)
            Case mWRSImap
                Call Chalk_WRSI_Map(gudtStock, gudtIndex, processFrame)
            Case mWMSmap
                Call Chalk_William_Map(gudtStock, gudtIndex, processFrame)
            Case mRWMSmap
                Call Chalk_RWilliam_Map(gudtStock, gudtIndex, processFrame)
            Case mAccMap
                Call Chalk_Acc_Map(gudtStock, gudtIndex, processFrame)
            Case mTomeMap
                Call Chalk_Tome_Map(gudtStock, gudtIndex, processFrame)
            Case mGRGMap
                Call Chalk_GRG_MAP(gudtStock, gudtIndex, processFrame)
            Case mSignalMap
                Call Chalk_Signal_Map(gudtStock, gudtIndex, processFrame)
            ' 20180825
            Case mForeignStockMap
                Call Chalk_ForeignStock_Map(gudtStock, gudtIndex, processFrame)
            Case mSitAndCbStockMap
                Call Chalk_SitAndCbStock_Map(gudtStock, gudtIndex, processFrame)
            Case mSelfEmployedStockMap
                Call Chalk_SelfEmployedStock_Map(gudtStock, gudtIndex, processFrame)
            Case mLegalPersonStockMap
                Call Chalk_LegalPersonStock_Map(gudtStock, gudtIndex, processFrame)
                
            Case Else
        End Select
    Next
   
    Exit Sub
ERR_HANDLE:
    MsgBox "[Method: frmEarvinStocks.DrawFrameData()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub



Private Sub Chalk_Line_Map(ByRef gudtStock() As StockData, _
                           ByRef gudtIndex() As IndexData, _
                           ByVal dispMap As Integer, _
                           ByVal intDrawType As Integer, _
                           ByVal bytCurrentFrame As Byte)
    Dim sngLowPrice As Single, sngHighPrice As Single
    Dim sngCount As Single
    Dim sngIndex As Single
    Dim strIndexName As String ' 顯示 [指標天數+指標名稱] 於畫面上
    Dim bytRcolor As Byte, bytGcolor As Byte, bytBcolor As Byte
    Dim highAndLowValue As HighAndLowValues
    
    On Error GoTo ERR_HANDLE
   
    ' 取得欲顯示於畫面上該指標值之最大值、最小值 -- 20151107 --
    Call GetHighLow(sngLowPrice, sngHighPrice, dispMap, gudtStock, gudtIndex)
'    highAndLowValue = GetHighAndLowValues(dispMap, gudtStock, gudtIndex)
   
    ' 單線、對稱於0軸
    If intDrawType = 1 Then
        ' 對稱於0軸
        If Abs(sngLowPrice) < Abs(sngHighPrice) Then
            sngLowPrice = -sngHighPrice
        ElseIf Abs(sngLowPrice) > Abs(sngHighPrice) Then
            sngHighPrice = -sngLowPrice
        End If
    
        ' set the x axial and y axial scale
        gsngXinterval = (ScaleWidth - gsngLeftLevel - gsngRightLevel) / (frmEarvinStocks.Width / gsngBarWidth + 2)
        If (sngHighPrice - sngLowPrice) <> 0 Then
            gsngYinterval = mudtFrame(bytCurrentFrame).sngHeight / (sngHighPrice - sngLowPrice)
        End If
        ' caculate the underline value of the current frame
        gsngYshift = gsngBottomFrame - mudtFrame(bytCurrentFrame).sngHeight
        For sngCount = bytCurrentFrame To gbytFrameNum
            gsngYshift = gsngYshift + mudtFrame(sngCount).sngHeight
        Next
        ' plot the horizontal dash-line for roughly showing the value of price and mark the lable
        DrawStyle = vbDot
        ForeColor = QBColor(12)
        For sngCount = 1 To 3
            Line (gsngLeftLevel, sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift)-(frmEarvinStocks.ScaleWidth - gsngRightLevel, sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift), QBColor(6)
            frmEarvinStocks.CurrentX = 0.5 * gsngLeftLevel
            frmEarvinStocks.CurrentY = sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift + 100
            Print Format(sngLowPrice + sngCount * (sngHighPrice - sngLowPrice) / 4, "0")
        Next
        frmEarvinStocks.CurrentX = 0
        frmEarvinStocks.CurrentY = 0.1 * (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift + 100
        ForeColor = QBColor(7)
      
        Call GetIndexInfomation(strIndexName, dispMap, gudtStock, gudtIndex)
        Print strIndexName
   
        DrawStyle = vbSolid
        sngIndex = 0
   
        For sngCount = gsngStartIndex To gsngEndIndex
            If sngCount > 1 Then
                With gudtIndex(sngCount)
                    If sngCount = gsngStartIndex Then
                        Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift) _
                            -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(14)
                    Else
                        Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift) _
                            -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(14)
                    End If
                End With
            End If
            sngIndex = sngIndex + gsngXinterval
        Next
   
    ' 單線、上下限範圍 0 ~ 100
    ElseIf intDrawType = 2 Then
        ' set the x axial and y axial scale
        gsngXinterval = (ScaleWidth - gsngLeftLevel - gsngRightLevel) / (frmEarvinStocks.Width / gsngBarWidth + 2)
        If (sngHighPrice - sngLowPrice) <> 0 Then
            gsngYinterval = mudtFrame(bytCurrentFrame).sngHeight / (sngHighPrice - sngLowPrice)
        End If
        ' caculate the underline value of the current frame
        gsngYshift = gsngBottomFrame - mudtFrame(bytCurrentFrame).sngHeight
        For sngCount = bytCurrentFrame To gbytFrameNum
            gsngYshift = gsngYshift + mudtFrame(sngCount).sngHeight
        Next
        ' plot the horizontal dash-line for roughly showing the value of price and mark the lable
        DrawStyle = vbDot
        ForeColor = QBColor(12)
      
        For sngCount = 0 To 2
            Line (gsngLeftLevel, (20 + sngCount * 30) * gsngYinterval + gsngYshift)-(frmEarvinStocks.ScaleWidth - gsngRightLevel, 20 + (20 + sngCount * 30) * gsngYinterval + gsngYshift), QBColor(6)
            frmEarvinStocks.CurrentX = 0.6 * gsngLeftLevel
            frmEarvinStocks.CurrentY = (20 + sngCount * 30) * gsngYinterval + gsngYshift + 100
            Print Format(sngLowPrice + 20 + sngCount * 30, "#")
        Next
        frmEarvinStocks.CurrentX = 0
        frmEarvinStocks.CurrentY = 8 * gsngYinterval + gsngYshift + 100
        ForeColor = QBColor(7)
      
        Call GetIndexInfomation(strIndexName, dispMap, gudtStock, gudtIndex)
        Print strIndexName
      
        DrawStyle = vbSolid
        sngIndex = 0
      
        For sngCount = gsngStartIndex To gsngEndIndex
            With gudtIndex(sngCount)
                If sngCount > 1 Then
                    If sngCount = gsngStartIndex Then
                        Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift) _
                            -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(12)
                    Else
                        Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift) _
                            -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(12)
                    End If
                End If
                sngIndex = sngIndex + gsngXinterval
            End With
        Next
   
    ' 雙線、對稱於0軸
    ElseIf intDrawType = 3 Then
        gsngXinterval = (ScaleWidth - gsngLeftLevel - gsngRightLevel) / (frmEarvinStocks.Width / gsngBarWidth + 2)
        If (sngHighPrice - sngLowPrice) <> 0 Then
            gsngYinterval = mudtFrame(bytCurrentFrame).sngHeight / (sngHighPrice - sngLowPrice)
        End If
        ' caculate the underline value of the current frame
        gsngYshift = gsngBottomFrame - mudtFrame(bytCurrentFrame).sngHeight
        For sngCount = bytCurrentFrame To gbytFrameNum
            gsngYshift = gsngYshift + mudtFrame(sngCount).sngHeight
        Next
        'plot the horizontal dash-line for roughly showing the value of price and mark the lable
        DrawStyle = vbDot
        ForeColor = QBColor(12)
        For sngCount = 0 To 2
            Line (gsngLeftLevel, (20 + sngCount * 30) * gsngYinterval + gsngYshift)-(frmEarvinStocks.ScaleWidth - gsngRightLevel, 20 + (20 + sngCount * 30) * gsngYinterval + gsngYshift), QBColor(6)
            frmEarvinStocks.CurrentX = 0.6 * gsngLeftLevel
            frmEarvinStocks.CurrentY = (20 + sngCount * 30) * gsngYinterval + gsngYshift + 100
            Print Format(sngLowPrice + 20 + sngCount * 30, "#")
        Next
        frmEarvinStocks.CurrentX = 0
        frmEarvinStocks.CurrentY = 8 * gsngYinterval + gsngYshift + 100
        ForeColor = QBColor(7)
      
        Call GetIndexInfomation(strIndexName, dispMap, gudtStock, gudtIndex)
        Print strIndexName
      
        DrawStyle = vbSolid
        sngIndex = 0
      
        ' plot the KD line
        For sngCount = gsngStartIndex To gsngEndIndex
            With gudtIndex(sngCount)
                If sngCount > 1 Then
                    If sngCount = gsngStartIndex Then
                        DrawStyle = vbSolid
                        Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift) _
                            -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(12)
                        'plot the D line
                        DrawStyle = vbDot
                        Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift) _
                            -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(10)
                    Else
                        'plot the K line / MACD line
                        DrawStyle = vbSolid
                        Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift) _
                            -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(12)
                        'plot the D line / DIF line
                        DrawStyle = vbDot
                        Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift) _
                            -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(10)
                    End If
                End If
                sngIndex = sngIndex + gsngXinterval
            End With
        Next
        DrawStyle = vbSolid
   
    ' 3 Lines
    ElseIf intDrawType = 4 Then
        For sngCount = gsngStartIndex To gsngEndIndex
            With gudtIndex(sngCount)
                If sngCount > 1 Then
                    If sngCount = gsngStartIndex Then
                        ' Plot the 6 average line
                        If MAV_1 <> 0 And .sngLineValue(0) > sngLowPrice And .sngLineValue(0) < sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(14)
                        End If
                        ' Plot the 13 average line
                        If MAV_2 <> 0 And .sngLineValue(1) > sngLowPrice And .sngLineValue(1) < sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                        End If
                        ' Plot the 26 average line
                        If MAV_3 <> 0 And .sngLineValue(2) > sngLowPrice And .sngLineValue(2) < sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(2) - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(2) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                        End If
                        If MAV_4 <> 0 And .sngMAV4 > sngLowPrice And .sngMAV4 < sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAV4 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAV4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                        End If
                        If MAV_5 <> 0 And .sngMAV5 > sngLowPrice And .sngMAV5 < sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAV5 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAV5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                        End If
                    Else
                        ' Plot the 6 average line
                        If MAV_1 <> 0 Then
                            If .sngLineValue(0) >= sngLowPrice And .sngLineValue(0) <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(14)
                            ElseIf gudtIndex(sngCount - 1).sngLineValue(0) < sngLowPrice And .sngLineValue(0) >= sngLowPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(14)
                            ElseIf gudtIndex(sngCount - 1).sngLineValue(0) > sngHighPrice And .sngLineValue(0) <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(0) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(14)
                            End If
                        End If
                        ' Plot the 13 average line
                        If MAV_2 <> 0 Then
                            If .sngLineValue(1) >= sngLowPrice And .sngLineValue(1) <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                            ElseIf gudtIndex(sngCount - 1).sngLineValue(1) < sngLowPrice And .sngLineValue(1) >= sngLowPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                            ElseIf gudtIndex(sngCount - 1).sngLineValue(1) > sngHighPrice And .sngLineValue(1) <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(1) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                            End If
                        End If
                        ' Plot the 26 average line
                        If MAV_3 <> 0 Then
                            If .sngLineValue(2) >= sngLowPrice And .sngLineValue(2) <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngLineValue(2) - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(2) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                            ElseIf gudtIndex(sngCount - 1).sngLineValue(2) < sngLowPrice And .sngLineValue(2) >= sngLowPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(2) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                            ElseIf gudtIndex(sngCount - 1).sngLineValue(2) > sngHighPrice And .sngLineValue(2) <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLineValue(2) - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                            End If
                        End If
                        If MAV_4 <> 0 Then
                            If .sngMAV4 >= sngLowPrice And .sngMAV4 <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAV4 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAV4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                            ElseIf gudtIndex(sngCount - 1).sngMAV4 < sngLowPrice And .sngMAV4 >= sngLowPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAV4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                            ElseIf gudtIndex(sngCount - 1).sngMAV4 > sngHighPrice And .sngMAV4 <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAV4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                            End If
                        End If
                        If MAV_5 <> 0 Then
                            If .sngMAV5 >= sngLowPrice And .sngMAV5 <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAV5 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAV5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                            ElseIf gudtIndex(sngCount - 1).sngMAV5 < sngLowPrice And .sngMAV5 >= sngLowPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAV5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                            ElseIf gudtIndex(sngCount - 1).sngMAV5 > sngHighPrice And .sngMAV5 <= sngHighPrice Then
                                Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAV5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                            End If
                        End If
                    End If
                End If
            End With
            sngIndex = sngIndex + gsngXinterval
        Next
    End If
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[frmEarvinStocks.Chalk_Line_Map()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub



Private Sub Chalk_Bar_MAP(ByRef gudtStock() As StockData, _
                           ByRef gudtIndex() As IndexData, _
                           ByVal dispMap As Integer, _
                           ByVal intDrawType As Integer, _
                           ByVal bytCurrentFrame As Byte)
    Dim sngLowPrice As Single, sngHighPrice As Single
    Dim sngCount As Single
    Dim sngIndex As Single
    Dim strIndexName As String ' 顯示 [指標天數+指標名稱] 於畫面上
    Dim bytRcolor As Byte, bytGcolor As Byte, bytBcolor As Byte
    Dim sngChangeColor As Single

    On Error GoTo ERR_HANDLE
   
    ' 取得欲顯示於畫面上該指標值之最大值、最小值
    Call GetHighLow(sngLowPrice, sngHighPrice, dispMap, gudtStock, gudtIndex)
   
    ' 對稱0軸
    If intDrawType = 1 Then
   
    ' 0~100
    ElseIf intDrawType = 2 Then
        ' set the x axial and y axial scale
        gsngXinterval = (ScaleWidth - gsngLeftLevel - gsngRightLevel) / (frmEarvinStocks.Width / gsngBarWidth + 2)
        If (sngHighPrice - sngLowPrice) <> 0 Then
            gsngYinterval = mudtFrame(bytCurrentFrame).sngHeight / (sngHighPrice - sngLowPrice)
        End If
   
        ' caculate the underline value of the current frame
        gsngYshift = gsngBottomFrame - mudtFrame(bytCurrentFrame).sngHeight
        For sngCount = bytCurrentFrame To gbytFrameNum
            gsngYshift = gsngYshift + mudtFrame(sngCount).sngHeight
        Next
   
        ' plot the horizontal dash-line for roughly showing the value of price and mark the lable
        DrawStyle = vbDot
        ForeColor = QBColor(12)
        For sngCount = 1 To 3
            Line (gsngLeftLevel, sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift)-(frmEarvinStocks.ScaleWidth - gsngRightLevel, sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift), QBColor(6)
            frmEarvinStocks.CurrentX = 0.5 * gsngLeftLevel
            frmEarvinStocks.CurrentY = sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift + 100
            Print Format(sngLowPrice + sngCount * (sngHighPrice - sngLowPrice) / 4, "#")
        Next
        frmEarvinStocks.CurrentX = 0
        frmEarvinStocks.CurrentY = 0.1 * (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift + 100
        ForeColor = QBColor(7)
      
        Call GetIndexInfomation(strIndexName, dispMap, gudtStock, gudtIndex)
        Print strIndexName
      
        DrawStyle = vbSolid
        sngIndex = 0
      
        ' plot the SECTOR line
        For sngCount = gsngStartIndex To gsngEndIndex
            With gudtIndex(sngCount)
                If .sngBarValue(0) > 0 Then
                    bytRcolor = 255      ' Red Bar
                    bytGcolor = 0
                    bytBcolor = 0
                    Line (sngIndex + gsngLeftLevel, gsngYshift + (mudtFrame(bytCurrentFrame).sngHeight) / 2) _
                        -(sngIndex + 0.7 * gsngXinterval + gsngLeftLevel, (.sngBarValue(0)) * gsngYinterval + gsngYshift + (mudtFrame(bytCurrentFrame).sngHeight) / 2), RGB(bytRcolor, bytGcolor, bytBcolor), BF
                Else
                    bytRcolor = 50    ' Green Bar
                    bytGcolor = 200
                    bytBcolor = 50
                    Line (sngIndex + gsngLeftLevel, gsngYshift + (mudtFrame(bytCurrentFrame).sngHeight) / 2) _
                        -(sngIndex + 0.7 * gsngXinterval + gsngLeftLevel, (.sngBarValue(0)) * gsngYinterval + gsngYshift + (mudtFrame(bytCurrentFrame).sngHeight) / 2), RGB(bytRcolor, bytGcolor, bytBcolor), BF
                End If
                sngIndex = sngIndex + gsngXinterval
            End With
        Next
    ElseIf intDrawType = 3 Then   ' 20100415
        sngHighPrice = 1.1 * sngHighPrice
        ' Set the x axial and y axial scale
        gsngXinterval = (ScaleWidth - gsngLeftLevel - gsngRightLevel) / (frmEarvinStocks.Width / gsngBarWidth + 2)
        If (sngHighPrice - sngLowPrice) <> 0 Then
            gsngYinterval = mudtFrame(bytCurrentFrame).sngHeight / (sngHighPrice - sngLowPrice)
        End If
        ' Caculate the underline value of the current frame
        gsngYshift = gsngBottomFrame - mudtFrame(bytCurrentFrame).sngHeight
        For sngCount = bytCurrentFrame To gbytFrameNum
            gsngYshift = gsngYshift + mudtFrame(sngCount).sngHeight
        Next
        ' Plot the horizontal dash-line for roughly showing the value of price and mark the lable
        DrawStyle = 2
        ForeColor = QBColor(12)
        For sngCount = 1 To 3
            Line (gsngLeftLevel, sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift)-(frmEarvinStocks.ScaleWidth - gsngRightLevel, sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift), QBColor(6)
            frmEarvinStocks.CurrentX = 0.1 * gsngLeftLevel
            frmEarvinStocks.CurrentY = sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift + 100
            Print Format(sngLowPrice + sngCount * (sngHighPrice - sngLowPrice) / 4, "#")
        Next
        frmEarvinStocks.CurrentX = 0
        frmEarvinStocks.CurrentY = 50 * gsngYinterval + gsngYshift + 100
        ForeColor = QBColor(7)
        Print "成交量"
        DrawStyle = 0
        sngIndex = 0
    End If
    
    Exit Sub
    
ERR_HANDLE:
    MsgBox "[frmEarvinStocks.Chalk_Bar_MAP()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub