'
'    Developer: J.A. Runnells
'      Updated: 2020-07-15 17:30
'    Last Push: 2020-07-15 17:30
'       Branch: master
'      License: MIT
'

'Alt+F11 <=> Insert > Module
'Excel (VBA): https://support.microsoft.com/en-us/office/create-custom-functions-in-excel-2f06c10b-3622-40d6-a1b2-b6748ae8231f


Public Function PHAT(x As Double, n As Double) As Double
    PHAT = x / n
End Function

Public Function STDEV_PHAT(p As Double, n As Double) As Double
    STDEV_PHAT = Sqr((phat * (1 - phat)) / n)
    'STDEV_PHAT = Application.SQRT((phat * (1 - phat)) / n)
End Function

Public Function MARGINERROR_P(z_alpha2 As Double, phat As Double, n As Double) As Double
    MARGINERROR_P = z_alpha2 * Sqr((phat * (1 - phat)) / n)
    'MARGINERROR_P = z_alpha2 * Application.SQRT((phat * (1 - phat)) / n)
End Function

Public Function MARGINERROR_M(t_alpha2 As Double, s As Double, n As Double) As Double
    MARGINERROR_M = t_alpha2 * (s / Sqr(n))
    'MARGINERROR_M = t_alpha2 * (s / Application.SQRT(n))
End Function

Public Function CONFIDENCE_INTERVAL(z_alpha2 As Double, phat As Double, n As Double, lower As Variant) As Double
'Public Function CONFIDENCE_INTERVAL(z_alpha2 As Double, phat As Double, n As Double, Optional lower=FALSE As Boolean) As Double
    'Dim ci As Double
    Select Case lower        
        Case True, 1: CONFIDENCE_INTERVAL = phat - MARGINERROR_P(z_alpha2, phat, n)
        Case False, 0: CONFIDENCE_INTERVAL = phat + MARGINERROR_P(z_alpha2, phat, n)
        Case Else: MsgBox "[ERROR] Oops, try again!"
    End Select
    'CONFIDENCE_INTERVAL = ci
End Function

' =============================================================================
' UNDER DEV - NOT TESTED!!
' =============================================================================
Public Function SAMPLE_MIN_P(phat As Double, z_alpha2 As Double, me As Double) As Double
    SAMPLE_MIN_P = phat*(1-phat)*Math.Pow((z_alpha2/me),2)
End Function

Public Function SAMPLE_MIN(z_alpha2 As Double, me As Double) As Double
    SAMPLE_MIN = 0.25*Math.Pow((z_alpha2/me),2)
End Function