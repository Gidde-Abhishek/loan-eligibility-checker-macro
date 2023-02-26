Sub CheckLoanEligibility()

    
    ' Set variables
    Dim LAI As String
    Dim LoanAmount As Double
    Dim LTV As Double
    Dim TotalPropertyValue As Double
    Dim PropertyNALand As String
    Dim PropertyLayoutPlan As String
    Dim PropertySanctionedPlan As String
    Dim PrimaryContact As String
    Dim CIBILScore As Integer
    Dim CoApplicant1 As String
    Dim CoApplicantCIBIL1 As Integer
    Dim CoApplicant2 As String
    Dim CoApplicantCIBIL2 As Integer
    Dim CoApplicant3Guarantor As String
    Dim CoApplicantCIBIL3 As Integer
    Dim BureauEligibility As String
    Dim EligibilityCheck As String
    
    ' Set up loop
    Dim i As Long
    Dim LastRow As Long
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    ' Loop through data
    For i = 2 To LastRow
        
        ' Get values from data
        LAI = Range("A" & i).Value
        LoanAmount = Range("B" & i).Value
        LTV = Range("C" & i).Value
        TotalPropertyValue = Range("D" & i).Value
        PropertyNALand = Range("E" & i).Value
        PropertyLayoutPlan = Range("F" & i).Value
        PropertySanctionedPlan = Range("G" & i).Value
        PrimaryContact = Range("H" & i).Value
        CIBILScore = Range("I" & i).Value
        CoApplicant1 = Range("J" & i).Value
        CoApplicantCIBIL1 = Range("K" & i).Value
        CoApplicant2 = Range("L" & i).Value
        CoApplicantCIBIL2 = Range("M" & i).Value
        CoApplicant3Guarantor = Range("N" & i).Value
        CoApplicantCIBIL3 = Range("O" & i).Value
        BureauEligibility = ""
        EligibilityCheck = ""
        
            ' Check loan eligibility
    If LoanAmount >= 500000 And LoanAmount <= 3500000 And _
        (TotalPropertyValue < 4500000 Or PropertySanctionedPlan = "SECO") And _
        PropertyNALand = "Yes" And _
        PropertyLayoutPlan = "Formal" And _
        (PropertySanctionedPlan = "Collector (Zilla Parishad) (ZP)" Or _
         PropertySanctionedPlan = "Gram Panchayat (GP)" Or _
         PropertySanctionedPlan = "Municipality/Town Planning (TP)") Then
        EligibilityCheck = "Eligible"
        Range("P" & i).Value = EligibilityCheck
        Range("P" & i).Interior.Color = RGB(146, 208, 80)
    Else
        EligibilityCheck = "Not Eligible"
         Range("P" & i).Interior.Color = xlNone
        Dim Reason As String
        ' Determine the reason for ineligibility
        If LoanAmount < 500000 Then
            Reason = "Loan amount less than minimum requirement"
        ElseIf LoanAmount > 3500000 Then
            Reason = "Loan amount greater than maximum limit"
        ElseIf TotalPropertyValue >= 4500000 And PropertySanctionedPlan <> "SECO" Then
            Reason = "Total property value exceeds limit and property plan is not SECO"
        ElseIf PropertyNALand <> "Yes" Then
            Reason = "Property is not NA land"
        ElseIf PropertyLayoutPlan <> "Formal" Then
            Reason = "Property layout plan is not formal"
        ElseIf PropertySanctionedPlan <> "Collector (Zilla Parishad) (ZP)" And _
               PropertySanctionedPlan <> "Gram Panchayat (GP)" And _
               PropertySanctionedPlan <> "Municipality/Town Planning (TP)" Then
            Reason = "Property is not sanctioned by eligible authorities"
        End If
        ' Write reason to worksheet
        Range("R" & i).Value = Reason
        Range("P" & i).Value = EligibilityCheck
        
    End If
    
    ' Check bureau eligibility
    If CIBILScore > 675 Or CIBILScore = -1 Then
        If CoApplicantCIBIL1 > 675 Or CoApplicantCIBIL1 = -1 Then
            If CoApplicantCIBIL2 > 675 Or CoApplicantCIBIL2 = -1 Then
                If CoApplicantCIBIL3 > 675 Or CoApplicantCIBIL3 = -1 Then
                    BureauEligibility = "Eligible"
                    ' Write eligibility result to worksheet
                    Range("Q" & i).Value = BureauEligibility
                End If
            End If
        End If
    End If
Next i

' Display message box to indicate completion of task
MsgBox "Loan eligibility check completed."
End Sub

