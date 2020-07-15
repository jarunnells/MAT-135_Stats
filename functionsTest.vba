'Alt+F11 <=> Insert > Module
'Excel (VBA): https://support.microsoft.com/en-us/office/create-custom-functions-in-excel-2f06c10b-3622-40d6-a1b2-b6748ae8231f


Public Function STDEV_PHAT(p As Double, n As Double) As Variant

    Sqr((phat * (1 - phat)) / n)
    'Application.SQRT((phat * (1 - phat)) / n)

End Function

Public Function PHAT(x As Double, n As Double) As Variant

    PHAT = x / n

End Function


Public Function MARGINERROR_P(z_alpha2 As Double, phat As Double, n As Double) As Variant

    z_alpha2 * Sqr((phat * (1 - phat)) / n)
    'z_alpha2 * Application.SQRT((phat * (1 - phat)) / n)

End Function

Public Function MARGINERROR_M(t_alpha2 As Double, s As Double, n As Double) As Variant

    t_alpha2 * (s / Sqr(n))
    't_alpha2 * (s / Application.SQRT(n))

End Function

Public Function MARGINERROR_TEST(z_alpha2 As Double, phat As Double, n As Double) As Variant

    z_alpha2 * STDEV_PHAT(phat,n)

End Function
