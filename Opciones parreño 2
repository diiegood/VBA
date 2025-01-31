
'' ====================================
Function BSCall(Stock As Double, Exercise As Double, Rate As Double, Sigma As Double, Time As Double) As Double
Dim d1 As Double, d2 As Double
    With Application
        d1 = (.Ln(Stock / Exercise) + (Rate + (Sigma ^ 2) / 2) * Time) / (Sigma * Sqr(Time))
        d2 = (.Ln(Stock / Exercise) + (Rate - (Sigma ^ 2) / 2) * Time) / (Sigma * Sqr(Time))
        BSCall = Stock * .Norm_S_Dist(d1, True) - Exercise * Exp(-Rate * Time) * .Norm_S_Dist(d2, True)
    End With
End Function

'' ====================================
Function BSPut(Stock As Double, Exercise As Double, Rate As Double, Sigma As Double, Time As Double) As Double
Dim d1 As Double, d2 As Double
    With Application
        d1 = (.Ln(Stock / Exercise) + (Rate + (Sigma ^ 2) / 2) * Time) / (Sigma * Sqr(Time))
        d2 = (.Ln(Stock / Exercise) + (Rate - (Sigma ^ 2) / 2) * Time) / (Sigma * Sqr(Time))
        BSPut = Exercise * Exp(-Rate * Time) * .Norm_S_Dist(-d2, True) - Stock * .Norm_S_Dist(-d1, True)
    End With
End Function

'' ====================================
Sub TestBSModel()
Dim CallP As Double, PutP As Double
    CallP = BSCall(42, 40, 0.05, 0.2, 0.5)
    PutP = BSPut(42, 40, 0.05, 0.2, 0.5)
        Debug.Print "Time: " & Time
        Debug.Print "===================================="
        Debug.Print "BSCall(42, 40, 0.05, 0.2, 0.5): returns " & Format(CallP, "Currency")
        Debug.Print "BSPut(42, 40, 0.05, 0.2, 0.5): returns " & Format(PutP, "Currency")
End Sub

'' ====================================
Function BSOption(Stock As Double, _
                  Exercise As Double, _
                  Rate As Double, _
                  Sigma As Double, _
                  Time As Double, _
                  Optional OptType As Variant) As Variant
                  ' OptType TRUE (default) for Call, FALSE for Put
                  
Dim d1 As Double, d2 As Double
Dim BSCall As Double, BSPut As Double

    If IsMissing(OptType) Then OptType = True
    ' Check that Variant has sub type Boolean
    If VBA.TypeName(OptType) <> "Boolean" Then GoTo ErrHandler
    On Error GoTo ErrHandler
    
    With Application
        d1 = (.Ln(Stock / Exercise) + (Rate + (Sigma ^ 2) / 2) * Time) / (Sigma * Sqr(Time))
        d2 = (.Ln(Stock / Exercise) + (Rate - (Sigma ^ 2) / 2) * Time) / (Sigma * Sqr(Time))
        BSCall = Stock * .Norm_S_Dist(d1, True) - Exercise * Exp(-Rate * Time) * .Norm_S_Dist(d2, True)
        BSPut = Exercise * Exp(-Rate * Time) * .Norm_S_Dist(-d2, True) - Stock * .Norm_S_Dist(-d1, True)
    End With
    
    If OptType Then
        BSOption = BSCall
    Else
        BSOption = BSPut
    End If
    
Exit Function
ErrHandler:
    BSOption = CVErr(xlErrValue)    ' Return #VALUE! error
End Function

'' ====================================
Sub BSOption_Test()
Dim Price As Double
Dim S As Double, E As Double, R As Double, V As Double, T As Double
Dim OT As Boolean

S = 42: E = 40: R = 5 / 100: V = 20 / 100: T = 0.5
Price = BSOption(S, E, R, V, T)
    Debug.Print " ================================"
    Debug.Print " Time: " & Format(Time, "hh:mm:ss") & vbNewLine
    Debug.Print " S = 42: E = 40: R = 5 / 100: V = 20 / 100: T = 0.5"
    Debug.Print " BSPrice = " & Format(Price, "$0.0000") & vbNewLine
    
S = 42: E = 40: R = 5 / 100: V = 20 / 100: T = 0.5: OT = True
Price = BSOption(S, E, R, V, T, OT)
    Debug.Print " S = 42: E = 40: R = 5 / 100: V = 20 / 100: T = 0.5: OT = True"
    Debug.Print " [OT = True = Call] "
    Debug.Print " BSPrice = " & Format(Price, "$0.0000") & vbNewLine

S = 42: E = 40: R = 5 / 100: V = 20 / 100: T = 0.5: OT = False
Price = BSOption(S, E, R, V, T, OT)
    Debug.Print " S = 42: E = 40: R = 5 / 100: V = 20 / 100: T = 0.5: OT = False"
    Debug.Print " [OT = False = Put] "
    Debug.Print " BSPrice = " & Format(Price, "$0.0000") & vbNewLine
    Debug.Print " ================================" & vbNewLine

End Sub

