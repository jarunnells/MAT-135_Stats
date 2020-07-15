'Alt+F11 <=> Insert > Module
'Excel (VBA): https://support.microsoft.com/en-us/office/create-custom-functions-in-excel-2f06c10b-3622-40d6-a1b2-b6748ae8231f


Public Function PHAT(x,n)

    PHAT = x / n

End Function


Public Function MARGINERROR(z_alpha2,phat,n)

    z_alpha2 * Sqr((phat * (1 - phat)) / n)
    'z_alpha2 * Application.SQRT((phat * (1 - phat)) / n)

End Function
