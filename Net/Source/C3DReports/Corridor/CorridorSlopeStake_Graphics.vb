''//
''// (C) Copyright 2009 by Autodesk, Inc.
''//
''// Permission to use, copy, modify, and distribute this software in
''// object code form for any purpose and without fee is hereby granted,
''// provided that the above copyright notice appears in all copies and
''// that both that copyright notice and the limited warranty and
''// restricted rights notice below appear in all supporting
''// documentation.
''//
''// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
''// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
''// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
''// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
''// UNINTERRUPTED OR ERROR FREE.
''//
''// Use, duplication, or disclosure by the U.S. Government is subject to
''// restrictions set forth in FAR 52.227-19 (Commercial Computer
''// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
''// (Rights in Technical Data and Computer Software), as applicable.
'
'
'
Friend Class SSRGraphics
    Public mCodePointList As New List(Of Slope_PointData)
    Public mDictLinkPoints As SortedDictionary(Of Integer, List(Of Slope_PointData))
    Public mDictEGDatas As SortedDictionary(Of Integer, List(Of Slope_PointData))
    Public mWidth, mHeight As New Integer
    Public mHorizontalMargin As Integer = 20
    Public mVerticalMargin As Integer = 20
    Public mPtMin, mPtMax As New Slope_PointData
    Public mXMax, mYMax As New Double
    Public mVertlineRatio_Y, mLinkRatio_Y As New Double
    Public mMaxLadderLevel As Integer = 10

    Public Sub AddCodePoint(ByVal ptData As Slope_PointData)
        mCodePointList.Add(ptData)
    End Sub

    Public Sub SetLinks(ByRef links As SortedDictionary(Of Integer, List(Of Slope_PointData)))
        mDictLinkPoints = links
    End Sub

    Public Sub SetEGDatas(ByRef EGs As SortedDictionary(Of Integer, List(Of Slope_PointData)))
        mDictEGDatas = EGs
    End Sub

    Public Sub GetGeomExtent(ByRef ptMin As Slope_PointData, _
                             ByRef ptMax As Slope_PointData)
        ' Get max X/Y geometry extent of whole graph(all points in range of ext of section)
        Dim bFirstTime As Boolean = True
        For Each ptData As Slope_PointData In mCodePointList
            If bFirstTime = True Then
                bFirstTime = False
                ptMin.mOffsetValue = ptData.mOffsetValue
                ptMax.mOffsetValue = ptData.mOffsetValue
                ptMin.mElevationValue = ptData.mElevationValue
                ptMax.mElevationValue = ptData.mElevationValue
            Else
                ptMin.mOffsetValue = Math.Min(ptMin.mOffsetValue, ptData.mOffsetValue)
                ptMin.mElevationValue = Math.Min(ptMin.mElevationValue, ptData.mElevationValue)
                ptMax.mOffsetValue = Math.Max(ptMax.mOffsetValue, ptData.mOffsetValue)
                ptMax.mElevationValue = Math.Max(ptMax.mElevationValue, ptData.mElevationValue)
            End If
        Next

        For Each idxOfLinkOnGraph As Integer In mDictLinkPoints.Keys()
            For Each ptData As Slope_PointData In mDictLinkPoints(idxOfLinkOnGraph)
                ptMin.mOffsetValue = Math.Min(ptMin.mOffsetValue, ptData.mOffsetValue)
                ptMin.mElevationValue = Math.Min(ptMin.mElevationValue, ptData.mElevationValue)
                ptMax.mOffsetValue = Math.Max(ptMax.mOffsetValue, ptData.mOffsetValue)
                ptMax.mElevationValue = Math.Max(ptMax.mElevationValue, ptData.mElevationValue)
            Next
        Next

        ' Trim the EGs' points if its offset out-of-range of section link point, except for the nearest one only
        Dim ptDataTemp As New Slope_PointData
        ' Get left nearest point of surface
        ptDataTemp.mOffsetValue = ptMin.mOffsetValue
        ptDataTemp.mElevationValue = ptMin.mElevationValue
        Dim ptDataList As List(Of Slope_PointData)
        For Each nIdxOfEG As Integer In mDictEGDatas.Keys()
            ptDataList = mDictEGDatas(nIdxOfEG)

            For Each ptData As Slope_PointData In ptDataList
                If ptData.mOffsetValue <= ptMin.mOffsetValue Then
                    ptDataTemp.mOffsetValue = ptData.mOffsetValue
                    ptDataTemp.mElevationValue = ptData.mElevationValue
                Else
                    ptMin.mOffsetValue = ptDataTemp.mOffsetValue
                    ptMin.mElevationValue = Math.Min(ptMin.mElevationValue, ptDataTemp.mElevationValue)
                    ptMax.mElevationValue = Math.Max(ptMax.mElevationValue, ptDataTemp.mElevationValue)
                    Exit For
                End If
            Next
        Next

        ' Get right furthest point of surface
        For Each nIdxOfEG As Integer In mDictEGDatas.Keys()
            ptDataList = mDictEGDatas(nIdxOfEG)

            For Each ptData As Slope_PointData In ptDataList
                If ptData.mOffsetValue < ptMax.mOffsetValue Then
                    Continue For
                Else
                    ptMax.mOffsetValue = ptData.mOffsetValue
                    ptMax.mElevationValue = Math.Max(ptMax.mElevationValue, ptData.mElevationValue)
                    ptMin.mElevationValue = Math.Min(ptMin.mElevationValue, ptData.mElevationValue)
                    Exit For
                End If
            Next
        Next
    End Sub

    Public Sub GetScreenResolution(ByRef Width As Integer, _
                                ByRef Height As Integer)
        Width = Screen.PrimaryScreen.Bounds.Width
        Height = Screen.PrimaryScreen.Bounds.Height
    End Sub

    Public Function GetDrawableWidth()
        Return mWidth - 2 * mHorizontalMargin
    End Function
    Public Function GetDrawableHeight()
        Return mHeight - 2 * mVerticalMargin
    End Function

    Public Function GetPointOnGraph(ByRef ptData As Slope_PointData) As Point


        ' ratio for mapping from geometry to pixel
        Dim ratioX, ratioY As Double
        ratioX = GetDrawableWidth() / mXMax
        ratioY = GetDrawableHeight() * (1 - mVertlineRatio_Y) / mYMax
        GetPointOnGraph.X = mHorizontalMargin + GetDrawableWidth() * Math.Abs(mPtMin.mOffsetValue) / (mPtMax.mOffsetValue - mPtMin.mOffsetValue) + ratioX * ptData.mOffsetValue

        If ptData.mCodes <> "CL" Then
            GetPointOnGraph.Y = mVerticalMargin + GetDrawableHeight() * (1.0 - mLinkRatio_Y / 2) - ratioY * (ptData.mElevationValue - (mPtMax.mElevationValue + mPtMin.mElevationValue) / 2)
        Else
            ' if center line, find the elevation of sub-assembly's nearest link code
            Dim dNearestLinkCodeOffsetToCL, dNearestLinkCodeElevToCL As Double
            Dim bFirstTime As Boolean = True
            For Each idxOfLinkOnGraph As Integer In mDictLinkPoints.Keys()
                Dim ptDataList As List(Of Slope_PointData)
                ptDataList = mDictLinkPoints(idxOfLinkOnGraph)
                For Each ptDataOfLink As Slope_PointData In ptDataList
                    If bFirstTime = True Then
                        bFirstTime = False
                        dNearestLinkCodeOffsetToCL = ptDataOfLink.mOffsetValue
                        dNearestLinkCodeElevToCL = ptDataOfLink.mElevationValue
                    Else
                        If Math.Abs(dNearestLinkCodeOffsetToCL) > Math.Abs(ptDataOfLink.mOffsetValue) Then
                            dNearestLinkCodeOffsetToCL = ptDataOfLink.mOffsetValue
                            dNearestLinkCodeElevToCL = ptDataOfLink.mElevationValue
                        End If
                    End If
                Next
            Next
            GetPointOnGraph.Y = mVerticalMargin + GetDrawableHeight() * (1.0 - mLinkRatio_Y / 2) - ratioY * (dNearestLinkCodeElevToCL - (mPtMax.mElevationValue + mPtMin.mElevationValue) / 2.0)
        End If

    End Function

    Public Class CodeWeightPos
        Public mOffsetWeight As New Double
        Public mIndexOfCode As New Integer
        Public mVertLevel As New Integer
        Public mIsLeftAlign As New Boolean ' True: right-to-left text layout, False: left-to-right
        Public Sub New(ByVal weight As Double, _
                       ByVal nIndexOfCode As Integer, _
                       Optional ByVal vertLevel As Integer = 0, _
                       Optional ByVal isLeftAlign As Boolean = False)
            mOffsetWeight = weight
            mIndexOfCode = nIndexOfCode
            mVertLevel = vertLevel
            mIsLeftAlign = isLeftAlign
        End Sub
    End Class

    Public Function FindIndexAtCodeIndex(ByVal nIndexOfCode As Integer, _
                                         ByRef ladderWeights As List(Of CodeWeightPos)) As Integer
        FindIndexAtCodeIndex = -1

        If nIndexOfCode < 0 Or nIndexOfCode >= mCodePointList.Count() Then
            Exit Function
        End If

        For nIdxOfWeight As Integer = 0 To ladderWeights.Count() - 1
            Dim weightPos As CodeWeightPos = ladderWeights.Item(nIdxOfWeight)
            If weightPos.mIndexOfCode = nIndexOfCode Then
                FindIndexAtCodeIndex = nIdxOfWeight
                Exit For
            End If
        Next
    End Function

    Public Class CodeWeightPos_Sorter
        Implements IComparer(Of CodeWeightPos)
        Function Compare(ByVal x As CodeWeightPos, _
                         ByVal y As CodeWeightPos) As Integer _
            Implements IComparer(Of CodeWeightPos).Compare
            If x.mOffsetWeight < y.mOffsetWeight Then
                Return -1
            ElseIf x.mOffsetWeight = y.mOffsetWeight Then
                If x.mIndexOfCode < y.mIndexOfCode Then
                    Return -1
                ElseIf x.mIndexOfCode = y.mIndexOfCode Then
                    Return 0
                Else
                    Return 1
                End If
            Else
                Return 1
            End If
        End Function 'IComparer.Compare
    End Class 'CodeWeightPos_Sorter

    Public Function GetladderWeights() As List(Of CodeWeightPos)
        GetladderWeights = New List(Of CodeWeightPos)
        Dim dLowestWeight As Double = 1000000.0
        Dim dWeight As Double = 0
        For nIdx As Integer = 0 To mCodePointList.Count() - 1
            If (nIdx = 0) Or (nIdx = mCodePointList.Count() - 1) Then
                GetladderWeights.Add(New CodeWeightPos(dLowestWeight, nIdx))
            Else
                dWeight = mCodePointList.Item(nIdx + 1).mOffsetValue - mCodePointList.Item(nIdx - 1).mOffsetValue ' left&right offset distance as weight
                GetladderWeights.Add(New CodeWeightPos(dWeight, nIdx))
            End If
        Next

        ' sort weight before return
        Dim weightSorter As New CodeWeightPos_Sorter
        GetladderWeights.Sort(weightSorter)

    End Function

    Public Function GetDefaultCodePosList(ByVal nVertLevel As Integer, _
                                          ByRef gr As Graphics, _
                                          ByRef codeFont As Font, _
                                          ByRef offsetFont As Font) As List(Of CodeWeightPos)
        ' get ladder weights
        Dim ladderWeights As List(Of CodeWeightPos)
        ladderWeights = GetladderWeights()

        ' number of groups each of which holding nVertLevel codes
        Dim nGroupNum As Integer = ladderWeights.Count() \ nVertLevel
        If (ladderWeights.Count() / nVertLevel) > (ladderWeights.Count() \ nVertLevel) Then
            nGroupNum += 1
        End If

        ' init ladder level
        Dim nLadderLevel As Integer = -1
        For nIdx As Integer = 0 To ladderWeights.Count() - 1
            nLadderLevel = (nVertLevel - 1) - (nIdx \ nGroupNum)
            ladderWeights.Item(nIdx).mVertLevel = nLadderLevel
        Next

        ' init text align
        Dim dFurthest As Double = 1000000.0
        Dim nIdxLeft As Integer = 0
        Dim nIdxRight As Integer = 0
        Dim offsetLeft As Double = 0.0
        Dim offsetRight As Double = 0.0
        Dim nLadderIndex As Integer = 0
        For nIdx As Integer = 0 To ladderWeights.Count() - 1

            nLadderIndex = ladderWeights.Item(nIdx).mIndexOfCode
            nLadderLevel = ladderWeights.Item(nIdx).mVertLevel
            nIdxLeft = FindSameVertLevel(nLadderLevel, _
                                         nLadderIndex - 1, _
                                         False, _
                                         ladderWeights)
            nIdxRight = FindSameVertLevel(nLadderLevel, _
                                          nLadderIndex + 1, _
                                          True, _
                                          ladderWeights)
            If nIdxLeft < 0 Then
                offsetLeft = dFurthest
            Else
                offsetLeft = mCodePointList.Item(nLadderIndex).mOffsetValue _
                             - mCodePointList.Item(nIdxLeft).mOffsetValue
            End If
            If nIdxRight < 0 Then
                offsetRight = dFurthest
            Else
                offsetRight = mCodePointList.Item(nIdxRight).mOffsetValue _
                             - mCodePointList.Item(nLadderIndex).mOffsetValue
            End If

            If offsetLeft > offsetRight Then
                ladderWeights.Item(nIdx).mIsLeftAlign = True
            Else
                ladderWeights.Item(nIdx).mIsLeftAlign = False
            End If

            If IsLayoutClipped(ladderWeights.Item(nIdx).mIndexOfCode, ladderWeights, gr, codeFont, offsetFont) Then
                ' reset it back
                ladderWeights.Item(nIdx).mIsLeftAlign = Not ladderWeights.Item(nIdx).mIsLeftAlign
            End If
        Next

        Return ladderWeights
    End Function

    Public Function FindSameVertLevel(ByVal nLadderLevel As Integer, _
                                      ByVal startIdx As Integer, _
                                      ByVal isForward As Boolean, _
                                      ByRef codeWeightPosList As List(Of CodeWeightPos)) As Integer
        FindSameVertLevel = -1

        If startIdx < 0 Or startIdx >= codeWeightPosList.Count() Then
            Exit Function
        End If

        Dim nStep As Integer
        Dim endIdx As Integer
        If isForward = True Then
            endIdx = codeWeightPosList.Count() - 1
            nStep = 1
        Else
            endIdx = 0
            nStep = -1
        End If
        Dim nIdxOfWeightList As Integer = -1
        For idx As Integer = startIdx To endIdx Step nStep
            nIdxOfWeightList = FindIndexAtCodeIndex(idx, codeWeightPosList)
            If nIdxOfWeightList < 0 Then
                Exit Function
            End If
            If codeWeightPosList.Item(nIdxOfWeightList).mVertLevel = nLadderLevel Then
                FindSameVertLevel = idx
                Exit Function
            End If
        Next
    End Function

    Public Function IsLayoutOverlapped(ByVal nIdxOfCode As Integer, _
                                       ByRef codeWeights As List(Of CodeWeightPos), _
                                       ByRef gr As Graphics, _
                                       ByRef codeFont As Font, _
                                       ByRef offsetFont As Font) As Boolean
        IsLayoutOverlapped = False

        Dim nIdxOfWeight As Integer = FindIndexAtCodeIndex(nIdxOfCode, codeWeights)
        If nIdxOfWeight < 0 Then
            Exit Function ' should never happen
        End If

        Dim codePos As CodeWeightPos = codeWeights.Item(nIdxOfWeight)
        Dim nLadderLevel As Integer = codePos.mVertLevel

        Dim startIdx As Integer
        Dim isForward As Boolean
        If codePos.mIsLeftAlign = True Then
            startIdx = nIdxOfCode - 1
            isForward = False
        Else
            startIdx = nIdxOfCode + 1
            isForward = True
        End If

        Dim sameLevelIdxCode As Integer = FindSameVertLevel(nLadderLevel, startIdx, isForward, codeWeights)
        Dim sameLevelIdxWeight As Integer = FindIndexAtCodeIndex(sameLevelIdxCode, codeWeights)
        Dim ratioX As Double = GetDrawableWidth() / mXMax

        If sameLevelIdxCode >= 0 And sameLevelIdxWeight >= 0 Then
            Dim ptData, ptDataNextTo As Slope_PointData
            Dim szCodeText, szCodeTextNextTo As SizeF
            Dim szOffsetText, szOffsetTextNextTo As SizeF

            ptData = mCodePointList.Item(nIdxOfCode)
            ptDataNextTo = mCodePointList.Item(sameLevelIdxCode)

            szCodeText = gr.MeasureString(ptData.mCodes, codeFont)
            szCodeTextNextTo = gr.MeasureString(ptDataNextTo.mCodes, codeFont)

            szOffsetText = gr.MeasureString(ptData.mOffsetString, offsetFont)
            szOffsetTextNextTo = gr.MeasureString(ptDataNextTo.mOffsetString, offsetFont)

            Dim dDistToSameLevelVLine As Double = ratioX * Math.Abs(ptData.mOffsetValue - ptDataNextTo.mOffsetValue)
            Dim dDistOccupyBySameLevelCode, dDistOccupyBySameLevelOffset As Double
            If codeWeights.Item(nIdxOfWeight).mIsLeftAlign <> codeWeights.Item(sameLevelIdxWeight).mIsLeftAlign Then
                dDistOccupyBySameLevelCode = szCodeTextNextTo.Width
                dDistOccupyBySameLevelOffset = szOffsetTextNextTo.Width
            Else
                dDistOccupyBySameLevelCode = 0
                dDistOccupyBySameLevelOffset = 0
            End If

            ' see if code or offset text overlapped
            If (szCodeText.Width + dDistOccupyBySameLevelCode) > dDistToSameLevelVLine _
               Or (szOffsetText.Width + dDistOccupyBySameLevelOffset) > dDistToSameLevelVLine Then
                IsLayoutOverlapped = True
            Else
                IsLayoutOverlapped = False
            End If
        End If
    End Function

    Public Function IsLayoutClipped(ByVal nIdxOfCode As Integer, _
                                    ByRef codeWeights As List(Of CodeWeightPos), _
                                    ByRef gr As Graphics, _
                                    ByRef codeFont As Font, _
                                    ByRef offsetFont As Font) As Boolean
        IsLayoutClipped = False

        Dim nIdxOfWeight As Integer = FindIndexAtCodeIndex(nIdxOfCode, codeWeights)
        If nIdxOfWeight < 0 Then
            Exit Function ' should never happen
        End If

        Dim codePos As CodeWeightPos = codeWeights.Item(nIdxOfWeight)

        Dim ptData As Slope_PointData
        Dim szCodeText As SizeF
        Dim szOffsetText As SizeF
        Dim ptOnGraph, ptClip As Point

        ptData = mCodePointList.Item(nIdxOfCode)
        szCodeText = gr.MeasureString(ptData.mCodes, codeFont)
        szOffsetText = gr.MeasureString(ptData.mOffsetString, offsetFont)
        ptOnGraph = GetPointOnGraph(ptData)

        ' see if code or offset text clipped(fall out of graph)
        If codePos.mIsLeftAlign = True Then
            ptClip.X = Math.Min(ptOnGraph.X - szCodeText.Width, ptOnGraph.X - szOffsetText.Width)
            IsLayoutClipped = ptClip.X < 0
        Else
            ptClip.X = Math.Max(ptOnGraph.X + szCodeText.Width, ptOnGraph.X + szOffsetText.Width)
            IsLayoutClipped = ptClip.X > (mHorizontalMargin + GetDrawableWidth())
        End If
    End Function

    Public Function IsLayoutSolved(ByRef codeWeights As List(Of CodeWeightPos), _
                                   ByVal nVertLevel As Integer, _
                                   ByRef gr As Graphics, _
                                   ByRef codeFont As Font, _
                                   ByRef offsetFont As Font) As Boolean

        Dim nIndexOfCode As Integer = -1
        For nIndexOfWeight As Integer = 0 To codeWeights.Count() - 1
            nIndexOfCode = codeWeights.Item(nIndexOfWeight).mIndexOfCode
            If True = IsLayoutOverlapped(nIndexOfCode, codeWeights, gr, codeFont, offsetFont) Then
                Return False
            End If
        Next
        Return True
    End Function

    Public Function Relayout(ByVal nIndexOfCode As Integer, _
                             ByRef gr As Graphics, _
                             ByRef codeFont As Font, _
                             ByRef offsetFont As Font, _
                             ByVal nVertLevel As Integer, _
                             ByRef codeWeights As List(Of CodeWeightPos)) As Boolean
        Dim bRet As Boolean = False

        Dim nIndexOfWeight As Integer = FindIndexAtCodeIndex(nIndexOfCode, codeWeights)
        Dim codePos As CodeWeightPos = codeWeights(nIndexOfWeight)

        ' try relayout
        Dim nOldUpper As Integer = codePos.mVertLevel
        Dim nTempUpper As Integer = nVertLevel - 1
        While nTempUpper >= 0
            ' Try to change text layout direction
            codePos.mIsLeftAlign = Not codePos.mIsLeftAlign
            If IsLayoutClipped(nIndexOfCode, codeWeights, gr, codeFont, offsetFont) Then
                ' reset it back
                codePos.mIsLeftAlign = Not codePos.mIsLeftAlign
            Else
                If False = IsLayoutOverlapped(nIndexOfCode, codeWeights, gr, codeFont, offsetFont) Then
                    ' succeed to layout
                    bRet = True
                    Exit While
                Else
                    ' reset it back
                    codePos.mIsLeftAlign = Not codePos.mIsLeftAlign
                End If
            End If

            ' Try to change vertical level for layout
            If nTempUpper >= 0 Then
                codePos.mVertLevel = nTempUpper

                If False = IsLayoutOverlapped(nIndexOfCode, codeWeights, gr, codeFont, offsetFont) Then
                    ' succeed to layout
                    bRet = True
                    Exit While
                Else
                    ' reset it back
                    codePos.mVertLevel = nOldUpper
                End If
            End If

            nTempUpper -= 1

        End While

        Return bRet
    End Function

    Public Function TryRelayoutNear(ByVal bFirstRun As Boolean, _
                                    ByVal nIndexOfCode As Integer, _
                                    ByVal nVertLevel As Integer, _
                                    ByRef codeWeights As List(Of CodeWeightPos), _
                                    ByRef gr As Graphics, _
                                    ByRef codeFont As Font, _
                                    ByRef offsetFont As Font) As Boolean
        Dim bRet As Boolean = False

        If nIndexOfCode < 0 Then
            Return bRet
        End If

        If False = IsLayoutOverlapped(nIndexOfCode, codeWeights, gr, codeFont, offsetFont) Then
            ' succeed to layout as no overlap now
            bRet = True
            Return bRet
        End If

        Dim nIndexOfWeight, nIndexOfCodeNear, nIndexOfWeightNear As Integer
        nIndexOfWeight = FindIndexAtCodeIndex(nIndexOfCode, codeWeights)
        If nIndexOfWeight < 0 Then
            Return bRet
        End If

        Dim codePos As CodeWeightPos = codeWeights.Item(nIndexOfWeight)
        Dim nLadderLevel As Integer = codePos.mVertLevel

        Dim startIdx As Integer
        Dim isForward As Boolean
        If codePos.mIsLeftAlign = True Then
            startIdx = codePos.mIndexOfCode - 1
            isForward = False
        Else
            startIdx = codePos.mIndexOfCode + 1
            isForward = True
        End If

        nIndexOfCodeNear = FindSameVertLevel(nLadderLevel, startIdx, isForward, codeWeights)
        If nIndexOfCodeNear < 0 Then
            Return bRet
        End If
        nIndexOfWeightNear = FindIndexAtCodeIndex(nIndexOfCodeNear, codeWeights)
        If nIndexOfWeightNear < 0 Then
            Return bRet
        End If

        Dim codePosNear As CodeWeightPos = codeWeights.Item(nIndexOfWeightNear)

        ' near one is more important
        If codePosNear.mOffsetWeight <= codePos.mOffsetWeight Then
            Return bRet ' stop relayout further
        ElseIf codePos.mIsLeftAlign = codePosNear.mIsLeftAlign Then
            Return bRet ' same text direction, stop
        Else
            ' try relayout near one
            If False = Relayout(nIndexOfCodeNear, gr, codeFont, offsetFont, nVertLevel, codeWeights) Then
                ' if failed to layout the near one, try the further near one
                If True = TryRelayoutNear(False, nIndexOfCodeNear, nVertLevel, codeWeights, gr, codeFont, offsetFont) Then
                    ' if can resolve near one
                    If True = Relayout(nIndexOfCodeNear, gr, codeFont, offsetFont, nVertLevel, codeWeights) Then
                        ' and currrent one
                        If True = Relayout(nIndexOfCode, gr, codeFont, offsetFont, nVertLevel, codeWeights) Then
                            bRet = True ' all resolved
                            Return bRet
                        End If
                    End If
                    Return bRet
                Else
                    Return bRet
                End If
            Else
                bRet = Relayout(nIndexOfCode, gr, codeFont, offsetFont, nVertLevel, codeWeights)
                Return bRet
            End If

        End If

    End Function

    Public Function ResolveLayout(ByRef gr As Graphics, _
                                  ByRef codeFont As Font, _
                                  ByRef offsetFont As Font, _
                                  ByVal nVertLevel As Integer, _
                                  ByRef codeWeights As List(Of CodeWeightPos)) As Boolean

        Dim bSameOffset As New Boolean
        Dim ptData As Slope_PointData
        Dim nIndexOfCode As Integer
        For nIndexOfWeight As Integer = 0 To codeWeights.Count() - 1
            nIndexOfCode = codeWeights.Item(nIndexOfWeight).mIndexOfCode
            ptData = mCodePointList.Item(nIndexOfCode)

            ' skip empty codes
            If ptData.mCodes = "" Then
                Continue For
            End If

            If True = IsLayoutOverlapped(nIndexOfCode, codeWeights, gr, codeFont, offsetFont) Then
                ' If near one is less important that current one, try change the near one
                If True = TryRelayoutNear(True, nIndexOfCode, nVertLevel, codeWeights, gr, codeFont, offsetFont) Then
                    Continue For
                Else
                    ' try relayout current
                    Relayout(nIndexOfCode, gr, codeFont, offsetFont, nVertLevel, codeWeights)
                End If
            End If
        Next

        Return IsLayoutSolved(codeWeights, nVertLevel, gr, codeFont, offsetFont)

    End Function

    Public Function CreateGraph(ByRef szFileName As String, _
                                ByRef szExt As String) As Boolean
        CreateGraph = False
        Try
            ' For sort
            Dim OffsetSorter As New Slope_PointData_Sorter

            ' Sort before drawing
            mCodePointList.Sort(OffsetSorter)
            For Each idxOfEG As Integer In mDictEGDatas.Keys()
                Dim ptDataList As List(Of Slope_PointData)
                ptDataList = mDictEGDatas(idxOfEG)
                ptDataList.Sort(OffsetSorter)
            Next

            Dim scrWidth, scrHeight As New Integer
            GetScreenResolution(scrWidth, scrHeight)

            GetGeomExtent(mPtMin, mPtMax)

            mXMax = (mPtMax.mOffsetValue - mPtMin.mOffsetValue)
            mYMax = (mPtMax.mElevationValue - mPtMin.mElevationValue)

            Dim codeFont As New Font("Arial", 12, FontStyle.Bold)
            Dim offsetFont As New Font("Arial", 12, FontStyle.Regular)

            ' set height to be consist in geom ratio of X,Y Axis
            mWidth = scrWidth - 40 ' -40 to prevent IE compressing graph even though IE window already maximized

            Dim nSectionHeight As Integer = GetDrawableWidth() * mYMax / mXMax ' 1:1 ratio of X/Y Axises
            Dim nladderHeight As Integer

            Dim codeWeights As List(Of CodeWeightPos)
            Dim nVertLevel As Integer = 1 ' initial vertical layout level
            Dim bResolved As Boolean = False

            Do
                nladderHeight = (nVertLevel + 2) * (codeFont.Height + offsetFont.Height) ' nVertLevel for code/offset, 2 for spaces between them
                mHeight = nSectionHeight + nladderHeight + 2 * mVerticalMargin

                mLinkRatio_Y = nSectionHeight / GetDrawableHeight()
                mVertlineRatio_Y = nladderHeight / GetDrawableHeight() ' ratio for vertical lines from point to code label

                Dim tempBitmap As Bitmap
                tempBitmap = New Bitmap(mWidth, mHeight)
                ' Create a graphics object for drawing.
                Dim tempGR As Graphics
                tempGR = Graphics.FromImage(tempBitmap)

                ' defalut ladder by nVertLevel
                codeWeights = GetDefaultCodePosList(nVertLevel, tempGR, codeFont, offsetFont)
                For nTryTimes As Integer = 0 To codeWeights.Count - 1
                    ' try resolve code layout
                    If True = ResolveLayout(tempGR, codeFont, offsetFont, nVertLevel, codeWeights) Then
                        bResolved = True
                        Exit For
                    End If
                Next

                tempGR.Dispose()
                tempBitmap.Dispose()

                nVertLevel += 1

            Loop While (Not bResolved) And (nVertLevel <= mMaxLadderLevel)

            nVertLevel -= 1

            Dim bitmap As Bitmap
            bitmap = New Bitmap(mWidth, mHeight)
            ' Create a graphics object for drawing.
            Dim gr As Graphics
            gr = Graphics.FromImage(bitmap)
            gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality

            ' Set bkground
            Dim rect As Rectangle
            rect = New Rectangle(0, 0, mWidth, mHeight)

            Dim bgBrush As System.Drawing.SolidBrush
            bgBrush = New System.Drawing.SolidBrush(Color.White)
            gr.FillRectangle(bgBrush, rect)

            ' Draw point codes
            Dim solidBrush As SolidBrush = New SolidBrush(Color.Black)
            Dim pen As Pen = New Pen(solidBrush)
            Dim textBrush As New Drawing2D.LinearGradientBrush(rect, _
            Color.Black, Color.Black, Drawing2D.LinearGradientMode.Horizontal)
            Dim textFormat As New StringFormat(StringFormat.GenericTypographic())

            ' draw links
            Dim linkPt1, linkPt2 As Point
            For Each idxOfLinkOnGraph As Integer In mDictLinkPoints.Keys()
                Dim ptDataList As List(Of Slope_PointData)
                ptDataList = mDictLinkPoints(idxOfLinkOnGraph)
                For nLoop As Integer = 0 To ptDataList.Count - 2
                    linkPt1 = GetPointOnGraph(ptDataList.Item(nLoop))
                    linkPt2 = GetPointOnGraph(ptDataList.Item(nLoop + 1))
                    gr.DrawLine(pen, linkPt1, linkPt2)
                Next
            Next

            ' draw codes
            Dim pt As New Point(-1, -1)
            Dim nLadderLevel As Integer = -1
            Dim bLeftAlign As Boolean = True

            Dim bSameOffset As New Boolean
            Dim ptData As Slope_PointData
            Dim nIndexOfWeights As Integer
            For nIndexOfCode As Integer = 0 To mCodePointList.Count() - 1
                ptData = mCodePointList.Item(nIndexOfCode)

                nIndexOfWeights = FindIndexAtCodeIndex(nIndexOfCode, codeWeights)
                nLadderLevel = codeWeights.Item(nIndexOfWeights).mVertLevel
                bLeftAlign = codeWeights.Item(nIndexOfWeights).mIsLeftAlign

                ' skip empty codes
                If ptData.mCodes = "" Then
                    Continue For
                End If

                pt = GetPointOnGraph(ptData)

                ' draw vertical line
                Dim xLineTop As Integer = pt.X
                Dim yLineTop As Integer = mVerticalMargin + GetDrawableHeight() * (mVertlineRatio_Y * nLadderLevel / nVertLevel)
                gr.DrawLine(pen, pt.X, pt.Y, xLineTop, yLineTop)

                ' Set text alignment, left at left side, right at right side
                If bLeftAlign = True Then
                    textFormat.Alignment = StringAlignment.Far ' default to left direction
                Else
                    textFormat.Alignment = StringAlignment.Near ' default to right direction
                End If

                ' draw code
                Dim xCode As Integer = pt.X
                Dim yCode As Integer
                yCode = mVerticalMargin + GetDrawableHeight() * (mVertlineRatio_Y * nLadderLevel / nVertLevel)
                gr.DrawString(ptData.mCodes, codeFont, textBrush, xCode, yCode, textFormat)

                ' draw offset
                ' If same with last offset, don't display offset again
                If nIndexOfCode <> 0 Then
                    If mCodePointList.Item(nIndexOfCode - 1).mOffsetValue = ptData.mOffsetValue Then
                        ' don't draw same offset
                        bSameOffset = True
                    Else
                        bSameOffset = False
                    End If
                Else
                    bSameOffset = False
                End If

                If bSameOffset = False Then
                    Dim xOffset As Integer = pt.X
                    Dim yOffset As Integer
                    yOffset = mVerticalMargin + GetDrawableHeight() * (mVertlineRatio_Y * nLadderLevel / nVertLevel) + codeFont.Height
                    gr.DrawString(ptData.mOffsetString, offsetFont, textBrush, xOffset, yOffset, textFormat)
                End If
            Next

            pen.DashCap = System.Drawing.Drawing2D.LineCap.Round
            pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash
            ' Connect surfaces' point
            Dim pt1 As New Point(-1, -1)
            Dim pt2 As New Point(-1, -1)

            For Each idxOfEG As Integer In mDictEGDatas.Keys()
                Dim ptDataList As List(Of Slope_PointData) = mDictEGDatas(idxOfEG)
                For Each ptEG As Slope_PointData In ptDataList
                    ' Trim EG point if out-of-range of offset
                    If ptEG.mOffsetValue < mPtMin.mOffsetValue _
                       Or ptEG.mOffsetValue > mPtMax.mOffsetValue Then
                        Continue For
                    End If

                    If pt1.X < 0 Then
                        pt1 = GetPointOnGraph(ptEG)
                    Else
                        pt2 = GetPointOnGraph(ptEG)
                        gr.DrawLine(pen, pt1, pt2)
                        pt1 = pt2
                    End If
                Next
            Next

            ' Clean up
            bgBrush.Dispose()
            solidBrush.Dispose()
            pen.Dispose()
            codeFont.Dispose()
            offsetFont.Dispose()
            textBrush.Dispose()

            ' Save the picture as a bitmap, JPEG, and GIF.
            If "bmp" = szExt Then
                bitmap.Save(szFileName & ".bmp", _
                            System.Drawing.Imaging.ImageFormat.Bmp)
                CreateGraph = True
            ElseIf "jpeg" = szExt Then
                bitmap.Save(szFileName & ".jpg", _
                            System.Drawing.Imaging.ImageFormat.Jpeg)
                CreateGraph = True
            ElseIf "png" = szExt Then
                bitmap.Save(szFileName & ".png", _
                            System.Drawing.Imaging.ImageFormat.Png)
                CreateGraph = True
            Else
                ' Do nothing
            End If

            ' Clean up
            gr.Dispose()
            bitmap.Dispose()

        Catch ex As Exception
            System.Diagnostics.Debug.Assert(False, ex.Message)
        End Try

    End Function
End Class