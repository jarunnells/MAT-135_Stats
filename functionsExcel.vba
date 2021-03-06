'
'    Developer: J.A. Runnells
'      Updated: 2020-07-22 00:45
'    Last Push: 2020-07-22 00:45
'       Branch: master
'      License: MIT
'

' Alt+F11 <=> Insert > Module
' Excel (VBA): https://support.microsoft.com/en-us/office/create-custom-functions-in-excel-2f06c10b-3622-40d6-a1b2-b6748ae8231f
' Application.SQRT()

' =============================================================================
' AVAILABLE FUNCTIONS (VBA):
'   [01] PHAT() <=> [x] TESTING  [x] DEV COMPLETE
'   [02] STDEV_PHAT() <=> [x] TESTING  [x] DEV COMPLETE
'   [03] MARGINERROR_P() <=> [x] TESTING  [x] DEV COMPLETE
'   [04] MARGINERROR_M() <=> [x] TESTING  [x] DEV COMPLETE
'   [05] CONFIDENCE_INTERVAL() <=> [x] TESTING  [x] DEV COMPLETE
'   [06] SAMPLE_MIN_P() <=> [x] TESTING  [ ] DEV COMPLETE
'   [07] SAMPLE_MIN() <=> [x] TESTING  [ ] DEV COMPLETE
'   [08] CONFIDENCE_INTERVAL_DEFAULT_TEST() <=> [x] TESTING  [ ] DEV COMPLETE
'   [09] PHAT_I() <=> [x] TESTING  [ ] DEV COMPLETE
'   [10] MARGINERROR_I() <=> [x] TESTING  [ ] DEV COMPLETE
'   [11] SAMPLE_MIN_MEAN() <=> [x] TESTING  [ ] DEV COMPLETE
' =============================================================================

' Calculates p̂ (population proportion) from given X and n.
Public Function PHAT(x As Double, n As Double) As Double
    PHAT = x / n
End Function
' Calculates standard deviation of the sampling distribution of p̂ (population proportion).
Public Function STDEV_PHAT(phat As Double, n As Double) As Double
    STDEV_PHAT = Sqr(phat * (1 - phat) / n)
End Function
' Calculates the margin of error [E or ME] for a proportion.
Public Function MARGINERROR_P(z_alpha2 As Double, phat As Double, n As Double) As Double
    MARGINERROR_P = z_alpha2 * Sqr(phat * (1 - phat) / n)
End Function
' Calculates the margin of error [E or ME] for the mean.
Public Function MARGINERROR_M(t_alpha2 As Double, s As Double, n As Double) As Double
    MARGINERROR_M = t_alpha2 * (s / Sqr(n))
End Function
' Calculates the specified upper or lower bound value.
Public Function CONFIDENCE_INTERVAL(z_alpha2 As Double, phat As Double, n As Double, lower As Variant) As Double
    Select Case lower        
        Case True, 1: CONFIDENCE_INTERVAL = phat - MARGINERROR_P(z_alpha2, phat, n)
        Case False, 0: CONFIDENCE_INTERVAL = phat + MARGINERROR_P(z_alpha2, phat, n)
        Case Else: MsgBox "[ERROR] Oops, try again!"
    End Select
End Function

' =============================================================================
' --------------------------------  UNDER DEV  --------------------------------
' =============================================================================
' Estimates population proportion using prior estimate (p̂).
Public Function SAMPLE_MIN_P(phat As Double, z_alpha2 As Double, m_err As Double) As Double
    SAMPLE_MIN_P = phat*(1-phat)*Math.Pow((z_alpha2/m_err),2)
End Function
' Estimates population proportion without a prior estimate.
Public Function SAMPLE_MIN(z_alpha2 As Double, m_err As Double) As Double
    SAMPLE_MIN = 0.25*Math.Pow((z_alpha2/me),2)
End Function
' Calculates the specified upper or lower bound value <=> REFACTOR: DEFAULT VALUE ADDED
Public Function CONFIDENCE_INTERVAL_DEFAULT_TEST(z_alpha2 As Double, phat As Double, n As Double, Optional lower As Variant = TRUE) As Double
    Select Case lower        
        Case True, 1: CONFIDENCE_INTERVAL = phat - MARGINERROR_P(z_alpha2, phat, n)
        Case False, 0: CONFIDENCE_INTERVAL = phat + MARGINERROR_P(z_alpha2, phat, n)
        Case Else: MsgBox "[ERROR] Oops, try again!"
    End Select
End Function
' Calculates p̂ (sample proportion) given upper and lower bounds.
Public Function PHAT_I(lower As Double, upper As Double) As Double
    PHAT_I = (lower+upper)/2
End Function
' Calculates the margin of error given upper and lower bounds.
Public Function MARGINERROR_I(lower As Double, upper As Double) As Double
    MARGINERROR_I = Math.Abs(lower-upper)/2
End Function
' Estimates population mean.
Public Function SAMPLE_MIN_MEAN(z_alpha2 As Double, s As Double, m_err As Double) As Double
    SAMPLE_MIN_MEAN = Math.Pow((z_alpha2*s/m_err),2)
End Function
