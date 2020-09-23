Attribute VB_Name = "modListviewSort"
Option Explicit
'// A module for sorting a listview which considers numbers, dates and size.
'// Heavily based on code written Copyright ©1997, Karl E. Peterson [http://www.mvps.org/vb]

Private Const M_COLL_KEY = "BT"

'// Private vars
Private m_lvIColumn   As Long
Private m_columnData  As Collection


'// Listview Column Sort Callbacks
'// Compare returns:
'//  -1 = Less Than
'//  0 = Equal
'//  1 = Greater Than
Private Function LvwCompareCallback(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As Long) As Long
    With m_columnData
        LvwCompareCallback = Sgn(.Item(M_COLL_KEY & lParam1) - .Item(M_COLL_KEY & lParam2))
    End With
    
    If SortOrder = lvwDescending Then LvwCompareCallback = -LvwCompareCallback
End Function

Private Sub PopulateCollection(ByVal hWnd As Long, ByVal columnIndex As Long, ByVal sortMode As ColumnSortConstants)
Dim i As Long, lvItemCount As Long, lRet As Long
Dim LVI As LV_ITEM, tmpStr As String
Dim fSizeNum As Double  '// For the cscSortFormattedSize mode
    
    Set m_columnData = New Collection
    
    '// Get the numer of items in the listview
    lvItemCount = SendMessage(hWnd, LVM_GETITEMCOUNT, 0&, 0&)
    
    If lvItemCount > 0 Then
        With LVI
            .mask = LVIF_TEXT Or LVIF_PARAM
            .iSubItem = columnIndex
            .cchTextMax = MAX_PATH
            .pszText = Space$(MAX_PATH)
        
            For i = 0 To lvItemCount - 1
                .iItem = i
                Call SendMessage(hWnd, LVM_GETITEM, 0&, LVI)
                tmpStr = TrimNulls(.pszText)
                
                If sortMode = cscSortDate Then
                    If tmpStr <> vbNullString Then
                        m_columnData.Add CDate(tmpStr), M_COLL_KEY & .lParam
                    Else
                        m_columnData.Add DateSerial(0, 1, 1), M_COLL_KEY & .lParam
                    End If
                ElseIf sortMode = cscSortNumber Then
                    If tmpStr <> vbNullString Then
                        m_columnData.Add Val(tmpStr), M_COLL_KEY & .lParam
                    Else
                        m_columnData.Add 0, M_COLL_KEY & .lParam
                    End If
                ElseIf sortMode = cscSortFormattedSize Then
                    If tmpStr <> vbNullString Then
                        fSizeNum = Val(tmpStr)

                        If InStr(1, tmpStr, "KB", vbTextCompare) > 0 Then
                            fSizeNum = fSizeNum * 1024
                        ElseIf InStr(1, tmpStr, "MB", vbTextCompare) > 0 Then
                            fSizeNum = fSizeNum * 1048576
                        ElseIf InStr(1, tmpStr, "GB", vbTextCompare) > 0 Then
                            fSizeNum = fSizeNum * 1073741824
                        End If
                        
                        m_columnData.Add fSizeNum, M_COLL_KEY & .lParam
                    Else
                        m_columnData.Add 0, M_COLL_KEY & .lParam
                    End If
                End If
            Next
        End With
    End If
End Sub

Public Sub SortColumn(ByVal hWnd As Long, ByVal columnIndex As Long, ByVal sortMode As ColumnSortConstants, ByVal SortOrder As ListSortOrderConstants)
    '// Populate the collection with the items we wish to sort
    Call PopulateCollection(hWnd, columnIndex, sortMode)
    
    '// Instruct the listview to use a callback function for a sort
    Call SendMessageLong(hWnd, LVM_SORTITEMS, SortOrder, FnPtr(AddressOf LvwCompareCallback))
    
    '// Cleanup
    Set m_columnData = Nothing
End Sub
