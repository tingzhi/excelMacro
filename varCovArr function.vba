Function VarCovArr(RNG As Range)
Dim Arr(), OutputArr(), i As Long, j As Long, k As Long, ReturnArr(), WF As Object, TempArr(), TempArr2()
Set WF = Application.WorksheetFunction

Arr = RNG

'build our return array using logrithmic method
    ' (1) size our return array
    ReDim ReturnArr(LBound(Arr, 1) To UBound(Arr, 1) - 1, LBound(Arr, 2) To UBound(Arr, 2))
    For i = LBound(ReturnArr, 1) To UBound(ReturnArr, 1)
        For j = LBound(ReturnArr, 2) To UBound(ReturnArr, 2)
            ReturnArr(i, j) = WF.Ln(Arr((i + 1), j) / Arr(i, j))
        Next j
    Next i

ReDim OutputArr(LBound(Arr, 2) To UBound(Arr, 2), LBound(Arr, 2) To UBound(Arr, 2))
ReDim TempArr(LBound(ReturnArr, 1) To UBound(ReturnArr, 1))
ReDim TempArr2(LBound(ReturnArr, 1) To UBound(ReturnArr, 1))

    For i = LBound(OutputArr, 1) To UBound(OutputArr, 1)
        For j = LBound(OutputArr, 2) To UBound(OutputArr, 2)
            ' (1) fill temp arrays
            For k = LBound(ReturnArr, 1) To UBound(ReturnArr, 1)
                TempArr(k) = ReturnArr(k, i)
                TempArr2(k) = ReturnArr(k, j)
            Next k
            ' (2) store the covariance
            OutputArr(i, j) = WF.Covar(TempArr, TempArr2)
            'For excel 2011, there is no Covariance_S function, it only has Covar function. However, in excel 2010 and 2013, they do.
        Next j
    Next i

VarCovArr = OutputArr

End Function

