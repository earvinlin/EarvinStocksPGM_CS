Public Sub DrawStockForm(ByVal displayEndIndex As Integer, ByVal stockName As String)
    Dim frameCount As Integer
    
    On Error GoTo ERR_HANDLE
    
    Cls
    ' 若開檔成功，顯示股票名稱
    If blnOpenFileSuccess Then
        lblStockName.ForeColor = QBColor(14)
        lblStockName.Caption = stockName
    End If
        
    frameCount = cboFrameNum.Text ' 記錄目前的frame數目
    Call SetEachFrameHigh(frameCount)
    Call DrawOutlineOfFrames(frameCount)
    Call SetDisplayStartIndex
    ' 傳入選擇處理資料是「日/週/月線」；要顯示的最後一筆資料位置
    Call DrawStockFormByStockType(cboStocksType.Text, displayEndIndex)
    
    Exit Sub
ERR_HANDLE:
    MsgBox "[Method: frmEarvinStocks.DrawStockForm()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub


Private Sub SetEachFrameHigh(ByVal frameCount As Integer)
    Dim processFrame As Integer ' 目前準備處理的frame
    
    If frameCount = 1 Then
        mudtFrame(1).sngHeight = Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand - gsngBottomFrame
    Else
        mudtFrame(1).sngHeight = 3 * ((Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand - gsngBottomFrame)) / (frameCount + 2)
        For processFrame = 2 To frameCount
            mudtFrame(processFrame).sngHeight = (Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand - gsngBottomFrame) / (frameCount + 2)
        Next
    End If
End Sub




Private Sub Chalk_K_Map(ByRef gudtStock() As StockData, _
                        ByRef gudtIndex() As IndexData, _
                        ByVal bytCurrentFrame As Byte)
    Dim sngLowPrice As Single, sngHighPrice As Single
    Dim sngCount As Single
    Dim sngIndex As Single
    Dim intCounter As Single
    Dim bytRcolor As Byte, bytGcolor As Byte, bytBcolor As Byte
    Dim Temp
    
    On Error GoTo ERR_HANDLE
    
    Call GetHighLow(sngLowPrice, sngHighPrice, mKmap, gudtStock, gudtIndex) ' get the higher and lower value
    ' Set the x axial and y axial scale
    gsngXinterval = (ScaleWidth - gsngLeftLevel - gsngRightLevel) / (frmEarvinStocks.Width / gsngBarWidth + 2)
    gsngYinterval = mudtFrame(bytCurrentFrame).sngHeight / (sngHighPrice - sngLowPrice)
    ' Caculate the underline value of the current frame
    gsngYshift = gsngBottomFrame - mudtFrame(bytCurrentFrame).sngHeight
    For sngCount = bytCurrentFrame To gbytFrameNum
        gsngYshift = gsngYshift + mudtFrame(sngCount).sngHeight
    Next
    Temp = gsngYshift
    ' Plot the horizontal dash-line for roughly showing the value of price and mark the lable
    DrawStyle = 2
    ForeColor = QBColor(11)
    Line (gsngLeftLevel, 0.1 * gsngYinterval * (sngHighPrice - sngLowPrice) / 5 + gsngYshift)-(frmEarvinStocks.ScaleWidth - gsngRightLevel, 0.1 * gsngYinterval * (sngHighPrice - sngLowPrice) / 5 + gsngYshift), QBColor(2)
    frmEarvinStocks.CurrentX = 0.1 * gsngLeftLevel
    frmEarvinStocks.CurrentY = 0.1 * gsngYinterval * (sngHighPrice - sngLowPrice) / 5 + gsngYshift + 100
    Print Format(sngLowPrice + 0.1 * (sngHighPrice - sngLowPrice) / 5, "0.00")
    For sngCount = 1 To 3
        Line (gsngLeftLevel, sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift)-(frmEarvinStocks.ScaleWidth - gsngRightLevel, sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift), QBColor(6)
        frmEarvinStocks.CurrentX = 0.1 * gsngLeftLevel
        frmEarvinStocks.CurrentY = sngCount * gsngYinterval * (sngHighPrice - sngLowPrice) / 4 + gsngYshift + 100
        Print Format(sngLowPrice + sngCount * (sngHighPrice - sngLowPrice) / 4, "0.00")
    Next
    frmEarvinStocks.CurrentX = 0.1 * gsngLeftLevel
    frmEarvinStocks.CurrentY = 4.9 * gsngYinterval * (sngHighPrice - sngLowPrice) / 5 + gsngYshift + 100
    Print Format(sngLowPrice + 4.9 * (sngHighPrice - sngLowPrice) / 5, "0.00")
    Line (gsngLeftLevel, 4.9 * gsngYinterval * (sngHighPrice - sngLowPrice) / 5 + gsngYshift)-(frmEarvinStocks.ScaleWidth - gsngRightLevel, 4.9 * gsngYinterval * (sngHighPrice - sngLowPrice) / 5 + gsngYshift), QBColor(2)
    ' plot the K map
    sngIndex = 0
     
    For sngCount = gsngStartIndex To gsngEndIndex
        With gudtStock(sngCount)
            ' select the color of markup, sell-off and parity
            If .sngStartprice = .sngEndprice And .sngEndprice = .sngHighPrice And .sngHighPrice = .sngLowPrice Then
                bytRcolor = 255   ' White Bar
                bytGcolor = 255
                bytBcolor = 255
            ElseIf .sngEndprice > .sngStartprice Then
                bytRcolor = 255   ' Red Bar
                bytGcolor = 0
                bytBcolor = 0
            Else
                bytRcolor = 50    ' Green Bar
                bytGcolor = 200
                bytBcolor = 50
            End If
            If cboStocksType.Text = "日線" Then
                frmEarvinStocks.CurrentX = sngIndex + gsngLeftLevel
                frmEarvinStocks.CurrentY = gsngBottomFrame
                DrawStyle = 2
                ForeColor = QBColor(11)
                If sngCount = gsngStartIndex Then   ' For starting case
                    Print Int(gudtStock(sngCount).sngDate / 10000)
                    frmEarvinStocks.CurrentX = sngIndex + gsngLeftLevel
                    frmEarvinStocks.CurrentY = gsngBottomFrame / 2
                    Print " " & Format((gudtStock(sngCount).sngDate / 100) Mod 100, "00")
                ElseIf Int(gudtStock(sngCount).sngDate / 100) <> Int(gudtStock(sngCount - 1).sngDate / 100) Then
                    Print Int(gudtStock(sngCount).sngDate / 10000)
                    frmEarvinStocks.CurrentX = sngIndex + gsngLeftLevel
                    frmEarvinStocks.CurrentY = gsngBottomFrame / 2
                    Print " " & Format((gudtStock(sngCount).sngDate / 100) Mod 100, "00")
                    Line (sngIndex + gsngLeftLevel - 0.15 * gsngXinterval, gsngBottomFrame)-(sngIndex + gsngLeftLevel - 0.15 * gsngXinterval, Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand), QBColor(6)
                End If
            ElseIf cboStocksType.Text = "週線" Then
                frmEarvinStocks.CurrentX = sngIndex + gsngLeftLevel
                frmEarvinStocks.CurrentY = gsngBottomFrame
                DrawStyle = 2
                ForeColor = QBColor(11)
                If sngCount = gsngStartIndex Then   ' For starting case
                    Print Int(gudtStock(sngCount).sngDate / 10000)
                    frmEarvinStocks.CurrentX = sngIndex + gsngLeftLevel
                    frmEarvinStocks.CurrentY = gsngBottomFrame / 2
                    Print " " & Format((gudtStock(sngCount).sngDate / 100) Mod 100, "00")
                    intCounter = Int(gudtStock(sngCount).sngDate / 100) Mod 100
                ElseIf Int(gudtStock(sngCount).sngDate / 100) <> Int(gudtStock(sngCount - 1).sngDate / 100) Then
                    intCounter = intCounter + 1
                    If intCounter = 12 Then
                        intCounter = 0
                    End If
                    If intCounter = 7 Or intCounter = 1 Then
                        Print Int(gudtStock(sngCount).sngDate / 10000)
                        frmEarvinStocks.CurrentX = sngIndex + gsngLeftLevel
                        frmEarvinStocks.CurrentY = gsngBottomFrame / 2
                        Print " " & Format((gudtStock(sngCount).sngDate / 100) Mod 100, "00")
                        Line (sngIndex + gsngLeftLevel - 0.15 * gsngXinterval, gsngBottomFrame)-(sngIndex + gsngLeftLevel - 0.15 * gsngXinterval, Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand), QBColor(6)
                    End If
                End If
            ElseIf cboStocksType.Text = "月線" Then
                frmEarvinStocks.CurrentX = sngIndex + gsngLeftLevel
                frmEarvinStocks.CurrentY = gsngBottomFrame
                DrawStyle = 2
                ForeColor = QBColor(11)
                If sngCount = gsngStartIndex Then   ' For starting case
                    Print Int(gudtStock(sngCount).sngDate / 100)
                ElseIf Int(gudtStock(sngCount).sngDate / 100) <> Int(gudtStock(sngCount - 1).sngDate / 100) Then
                    Print Int(gudtStock(sngCount).sngDate / 100)
                    Line (sngIndex + gsngLeftLevel - 0.15 * gsngXinterval, gsngBottomFrame)-(sngIndex + gsngLeftLevel - 0.15 * gsngXinterval, Abs(frmEarvinStocks.ScaleHeight) - gsngTopFrame - gsngTopCommand), QBColor(6)
                End If
            End If
            ' Plot the block of bar
            DrawStyle = 0
            Line (sngIndex + gsngLeftLevel, (.sngStartprice - sngLowPrice) * gsngYinterval + gsngYshift) _
                -(sngIndex + 0.7 * gsngXinterval + gsngLeftLevel, (.sngEndprice - sngLowPrice) * gsngYinterval + gsngYshift), RGB(bytRcolor, bytGcolor, bytBcolor), BF
            ' Plot the line of bar
            Line (sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngLowPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift), RGB(bytRcolor, bytGcolor, bytBcolor)
            ' 標示出畫面中的最高與最低價
            If .sngHighPrice = sngHighPrice Then
                ' 顯示出最高價
                LabHighprice.Left = frmEarvinStocks.CurrentX - gsngXinterval - gsngLeftLevel
                LabHighprice.Top = (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift - 10       '減10只是為了橋位置
                LabHighprice.Caption = sngHighPrice
                LabHighprice.Visible = True
            ElseIf .sngLowPrice = sngLowPrice Then
                ' 顯示出最低價
                LabLowprice.Left = frmEarvinStocks.CurrentX - gsngXinterval - gsngLeftLevel
                LabLowprice.Top = gsngYshift + 180                                  '加上180只是為了調整位置
                LabLowprice.Caption = sngLowPrice
                LabLowprice.Visible = True
            End If
        End With
        
        '***********當Kmap被選取時，畫出KMAP ********************
        If mnuKmap.HelpContextID = 0 Then
            ' Plot the average line
            With gudtIndex(sngCount)
                If sngCount > 1 Then
                    If sngCount = gsngStartIndex Then
                        ' Plot the 6 average line
                        If MAP_0 <> 0 Then
                            If gudtIndex(sngCount - 1).sngMAP0 >= sngLowPrice And gudtIndex(sngCount - 1).sngMAP0 <= sngHighPrice Then
                                Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP0 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP0 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(5)
                            ElseIf gudtIndex(sngCount - 1).sngMAP0 < sngLowPrice And .sngMAP0 >= sngLowPrice Then
                                Line (0.98 * gsngLeftLevel, gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP0 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(5)
                            ElseIf gudtIndex(sngCount - 1).sngMAP0 > sngHighPrice And .sngMAP0 <= sngHighPrice Then
                                Line (0.98 * gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                    -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP0 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(5)
                        End If
                    End If
                    
                    If MAP_1 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP1 >= sngLowPrice And gudtIndex(sngCount - 1).sngMAP1 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP1 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP1 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(14)
                        ElseIf gudtIndex(sngCount - 1).sngMAP1 < sngLowPrice And .sngMAP1 >= sngLowPrice Then
                            Line (0.98 * gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP1 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(14)
                        ElseIf gudtIndex(sngCount - 1).sngMAP1 > sngHighPrice And .sngMAP1 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP1 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(14)
                        End If
                    End If
                    ' Plot the 12 average line
                    If MAP_2 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP2 >= sngLowPrice And gudtIndex(sngCount - 1).sngMAP2 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP2 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP2 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                        ElseIf gudtIndex(sngCount - 1).sngMAP2 < sngLowPrice And .sngMAP2 >= sngLowPrice Then
                            Line (0.98 * gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP2 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                        ElseIf gudtIndex(sngCount - 1).sngMAP2 > sngHighPrice And .sngMAP2 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP2 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                        End If
                    End If
                    ' Plot the 24 average line
                    If MAP_3 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP3 >= sngLowPrice And gudtIndex(sngCount - 1).sngMAP3 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP3 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP3 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                        ElseIf gudtIndex(sngCount - 1).sngMAP3 < sngLowPrice And .sngMAP3 >= sngLowPrice Then
                            Line (0.98 * gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP3 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                        ElseIf gudtIndex(sngCount - 1).sngMAP3 > sngHighPrice And .sngMAP3 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP3 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                        End If
                    End If
                    ' Plot the 72 average line (72 average line = season line)
                    If MAP_4 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP4 >= sngLowPrice And gudtIndex(sngCount - 1).sngMAP4 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP4 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                        ElseIf gudtIndex(sngCount - 1).sngMAP4 < sngLowPrice And .sngMAP4 >= sngLowPrice Then
                            Line (0.98 * gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                        ElseIf gudtIndex(sngCount - 1).sngMAP4 > sngHighPrice And .sngMAP4 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                        End If
                    End If
                    ' Plot the 144 average line
                    If MAP_5 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP5 >= sngLowPrice And gudtIndex(sngCount - 1).sngMAP5 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP5 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                        ElseIf gudtIndex(sngCount - 1).sngMAP5 < sngLowPrice And .sngMAP5 >= sngLowPrice Then
                            Line (0.98 * gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                        ElseIf gudtIndex(sngCount - 1).sngMAP5 > sngHighPrice And .sngMAP5 <= sngHighPrice Then
                            Line (0.98 * gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                        End If
                    End If
                Else
                    ' Plot the 6 average line
                    If MAP_0 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP0 >= sngLowPrice And gudtIndex(sngCount - 1).sngMAP0 <= sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP0 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP0 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(5)
                        ElseIf gudtIndex(sngCount - 1).sngMAP0 < sngLowPrice And .sngMAP0 >= sngLowPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP0 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(5)
                        ElseIf gudtIndex(sngCount - 1).sngMAP0 > sngHighPrice And .sngMAP0 <= sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP0 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(5)
                        End If
                    End If
                  
                    If MAP_1 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP1 >= sngLowPrice And gudtIndex(sngCount - 1).sngMAP1 <= sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP1 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP1 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(3)
                        ElseIf gudtIndex(sngCount - 1).sngMAP1 < sngLowPrice And .sngMAP1 >= sngLowPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP1 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(3)
                        ElseIf gudtIndex(sngCount - 1).sngMAP1 > sngHighPrice And .sngMAP1 <= sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP1 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(3)
                        End If
                    End If
                    ' Plot the 12 average line
                    If MAP_2 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP2 >= sngLowPrice And gudtIndex(sngCount - 1).sngMAP2 <= sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP2 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP2 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                        ElseIf gudtIndex(sngCount - 1).sngMAP2 < sngLowPrice And .sngMAP2 >= sngLowPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP2 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                        ElseIf gudtIndex(sngCount - 1).sngMAP2 > sngHighPrice And .sngMAP2 <= sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP2 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(9)
                        End If
                    End If
                    ' Plot the 24 average line
                    If MAP_3 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP3 > sngLowPrice And gudtIndex(sngCount - 1).sngMAP3 < sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP3 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP3 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                        ElseIf gudtIndex(sngCount - 1).sngMAP3 < sngLowPrice And .sngMAP3 >= sngLowPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP3 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                        ElseIf gudtIndex(sngCount - 1).sngMAP3 > sngHighPrice And .sngMAP3 <= sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP3 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(13)
                        End If
                    End If
                    ' Plot the 72 average line (72 average line = season line)
                    If MAP_4 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP4 > sngLowPrice And gudtIndex(sngCount - 1).sngMAP4 < sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP4 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                        ElseIf gudtIndex(sngCount - 1).sngMAP4 < sngLowPrice And .sngMAP4 >= sngLowPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                        ElseIf gudtIndex(sngCount - 1).sngMAP4 > sngHighPrice And .sngMAP4 <= sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP4 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(11)
                        End If
                    End If
                    ' Plot the 144 average line
                    If MAP_5 <> 0 Then
                        If gudtIndex(sngCount - 1).sngMAP5 > sngLowPrice And gudtIndex(sngCount - 1).sngMAP5 < sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (gudtIndex(sngCount - 1).sngMAP5 - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                        ElseIf gudtIndex(sngCount - 1).sngMAP5 < sngLowPrice And .sngMAP5 >= sngLowPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                        ElseIf gudtIndex(sngCount - 1).sngMAP5 > sngHighPrice And .sngMAP5 <= sngHighPrice Then
                            Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftLevel, (sngHighPrice - sngLowPrice) * gsngYinterval + gsngYshift) _
                                -(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftLevel, (.sngMAP5 - sngLowPrice) * gsngYinterval + gsngYshift), QBColor(7)
                        End If
                    End If
                End If
                End If
                '------畫在K線下之三態線------
                'frmEarvinStocksWidth = 4
                'Line (sngIndex + (0.7 * gsngXinterval / 2 - gsngXinterval) + gsngLeftlevel, temp)-(sngIndex + (0.7 * gsngXinterval / 2) + gsngLeftlevel, temp), QBColor(Tr_arr(sngcount) + 1)
                'frmEarvinStocksWidth = 1
            End With
        End If
        sngIndex = sngIndex + gsngXinterval
    Next
  
    Exit Sub
  
ERR_HANDLE:
    MsgBox "[Method: frmEarvinStocks.Chalk_K_Map()], Err-Number= " & Err.Number & ", Err-Desc= " & Err.Description, vbOKOnly
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''
''' For IndexModule.cs 版本 (20260928 Moved) '''
''''''''''''''''''''''''''''''''''''''''''''''''
//////// 唸書時的 MACD 版本 (未完成)
//////public static (double[] DIF, double[] MACD, double[] DIF_MACD) CalculateMACD(StockData[] sd,
//////    int period = 12, int ema_s = 12, int ema_l = 26, int ema = 5)
//////{
//////    double sngEMA_S;
//////    double sngEMA_L;
//////    double sngPreEMA_S;
//////    double sngPreEMA_L;
//////    double sngPreMACD;
//////    double sngMACD;
//////    double sngDIF;
//////    double sngDIF_MACD;
//////    double sngDI;
//////    int i;
//////    int j;

//////    double[] difValues = New double[sd.Length];
//////    double[] macdValues = New double[sd.Length];
//////    double[] dif_macdValues = New double[sd.Length];

//////    sngPreEMA_S = sd[0].EndPrice;
//////    sngPreEMA_L = sd[0].EndPrice;
//////    sngPreMACD = 0;
//////    j = 1;

//////    while (j <= sd.Length) 
//////    {
//////        sngDI = (sd[j].HighPrice + sd[j].LowPrice + sd[j].EndPrice * 2) / 4;
//////        sngEMA_S = sngPreEMA_S + (2 * (sngDI - sngPreEMA_S) / (1 + ema_s));
//////        sngEMA_L = sngPreEMA_L + (2 * (sngDI - sngPreEMA_L) / (1 + ema_l));
//////        sngDIF = sngEMA_S - sngEMA_L;
//////        //sngMACD = sngPreMACD + (2 * (sngDIF - sngPreMACD) / (1 + period));
//////        sngMACD = sngPreMACD + (2 * (sngDIF - sngPreMACD) / (1 + period));
//////        sngDIF_MACD = sngDIF - sngMACD;

//////        difValues[j] = sngDIF;
//////        macdValues[j] = sngMACD;
//////        dif_macdValues[j] = sngDIF_MACD;

//////        sngPreEMA_S = sngEMA_S;
//////        sngPreEMA_L = sngEMA_L;
//////        sngPreMACD = sngMACD;
//////        j++;
//////    }

//////    return (difValues, macdValues, dif_macdValues);
//////    /*

//////        sngPreEMA_S = udtStock(1).sngEndprice
//////        sngPreEMA_L = udtStock(1).sngEndprice
//////        sngPreMACD = 0
//////        j = 1
//////        While j <= intStockNo
//////            sngDI = (udtStock(j).sngHighPrice + udtStock(j).sngLowPrice + udtStock(j).sngEndprice * 2) / 4
//////            sngEMA_S = sngPreEMA_S + (2 * (sngDI - sngPreEMA_S) / (1 + intSEMANo))
//////            sngEMA_L = sngPreEMA_L + (2 * (sngDI - sngPreEMA_L) / (1 + intLEMANo))
//////            sngDIF = sngEMA_S - sngEMA_L
//////            sngMACD = sngPreMACD + (2 * (sngDIF - sngPreMACD) / (1 + intMACDNo))
//////            sngDIF_MACD = sngDIF - sngMACD

//////            udtIndex(j).sngDIF = sngDIF
//////            udtIndex(j).sngMACD = sngMACD
//////            udtIndex(j).sngDIF_MACD = sngDIF_MACD

//////            sngPreEMA_S = sngEMA_S
//////            sngPreEMA_L = sngEMA_L
//////            sngPreMACD = sngMACD
//////            j = j + 1
//////        Wend


//////     * 標準 MACD 的寫法應該類似
//////        for (int i = 1; i < sd.Length; i++)
//////        {
//////            double close = sd[i].EndPrice;

//////            ema12 = ema12 + (close - ema12) * 2.0 / 13.0;
//////            ema26 = ema26 + (close - ema26) * 2.0 / 27.0;

//////            dif = ema12 - ema26;

//////            dea = dea + (dif - dea) * 2.0 / 10.0;

//////            hist = dif - dea;

//////            difValues[i] = dif;
//////            macdValues[i] = dea;
//////            dif_macdValues[i] = hist;
//////        }
//////    *-------------------------------


//////     */
//////}






//////        public static (double[] K, double[] D) CalculateKD(StockData[] sd, int intDayNo)
//////        {
//////            int i = 0, j = 0;
//////            double dblPrevK = 50, dblPrevD = 50, dblRsv = 0;
//////            double dblK = 0, dblD = 0;
//////            double dblMax = 0, dblMin = 0;
//////            double[] dblKValues = New double[sd.Length];
//////            double[] dblDValues = New double[sd.Length];

//////            while (i < sd.Length)
//////            {
//////                dblMax = double.MinValue;
//////                dblMin = double.MaxValue;
//////                if (i < (intDayNo - 1))
//////                {
//////                    dblKValues[i] = double.NaN;
//////                    dblDValues[i] = double.NaN;
//////                    i++;
//////                    continue;
//////                }
//////                else
//////                {
////////                    for (j = i; j > (i - intDayNo); j--)
//////                    for (j = i - intDayNo + 1; j <= i; j++)
//////                    {
//////                        if (dblMax < sd[j].HighPrice)
//////                            dblMax = sd[j].HighPrice;
//////                        if (dblMin > sd[j].LowPrice)
//////                            dblMin = sd[j].LowPrice;
//////                    }
//////                }

//////                if (dblMax != dblMin)
//////                    dblRsv = (sd[i].EndPrice - dblMin) / (dblMax - dblMin) * 100;
//////                else
//////                    dblRsv = 50;

//////                Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", dblMax= " + Math.Round(dblMax, 2) +
//////                    ", dblMin= " + Math.Round(dblMin, 2) + ", EndPrice= " + sd[i].EndPrice + ", dblRsv= " + dblRsv);
//////                dblK = dblPrevK * 2 / 3 + dblRsv / 3;
//////                dblD = dblPrevD * 2 / 3 + dblK / 3;
//////                Debug.WriteLine("dblPrevK= " + dblPrevK + ", dblPrevD= " + dblPrevD + ", dblK= " + dblK + ", dblD= " + dblD);

//////                if (i < (intDayNo - 1))
//////                {
//////                    dblKValues[i] = double.NaN;
//////                    dblDValues[i] = double.NaN;
//////                } 
//////                else
//////                {
//////                    dblKValues[i] = dblK;
//////                    dblDValues[i] = dblD;
//////                }
//////                i++;
//////            }
//////            // DEBUG : Display the PSY values For verification
//////            for (i = 0; i < sd.Length; i++)
//////            {
//////                Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", K= " + Math.Round(dblKValues[i], 2) + ", D= " + Math.Round(dblDValues[i], 2));
//////            }

//////            return (dblKValues, dblDValues);
//////        }