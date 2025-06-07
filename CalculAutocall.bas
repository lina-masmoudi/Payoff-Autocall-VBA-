Attribute VB_Name = "Module1"

Sub CalculAutocall()

    MsgBox "Debut du calcul"

    Dim prixInitial As Double
    Dim barriereRappel As Double
    Dim barriereProtection As Double
    Dim coupon As Double
    Dim prixFinal(1 To 5) As Double
    Dim i As Integer
    Dim rappel As Boolean
    Dim remboursement As Double
    Dim totalCoupon As Double

    With Sheets("Inputs")
        prixInitial = .Range("B1").Value
        For i = 1 To 5
            prixFinal(i) = .Range("B" & (i + 1)).Value
        Next i
        barriereRappel = .Range("B7").Value / 100
        barriereProtection = .Range("B8").Value / 100
        coupon = .Range("B9").Value / 100
    End With

    rappel = False
    totalCoupon = 0

    For i = 1 To 5
        If prixFinal(i) >= prixInitial * barriereRappel Then
            rappel = True
            totalCoupon = i * coupon * 100
            remboursement = 100
            Exit For
        End If
    Next i

    If rappel = False Then
        If prixFinal(5) >= prixInitial * barriereProtection Then
            remboursement = 100
        Else
            remboursement = prixFinal(5) / prixInitial * 100
        End If
    End If
  
    With Sheets("Resultats")
        If rappel Then
            .Range("B1").Value = "Rappel à l'année " & i
        Else
            .Range("B1").Value = "Pas de rappel"
        End If
        .Range("B2").Value = totalCoupon
        .Range("B3").Value = remboursement
    End With

    MsgBox "Calcul terminé !"

End Sub
