Attribute VB_Name = "Medix"
Option Explicit
Global CompanyName As String
Global CompanyAddress As String
Global BusinessYear As Integer
Global Access As String
Global Logon As String
Global LogonName As String
Global ClockedInStatus As String
Global SpcAccess As String
Global LGLogon As String
Global USRCLKTime As String
Global ComplaintCode As String
Global tmpPLNMinDurHosp As String
Global tmpPLNPhyVstPerDay As String
Global tmpPLNSepPeriodPrin As String
Global tmpPLNSepPeriodSup As String
Global tmpPLNRefRqdAdmin As String
Global tmpPLNRefRqdPreSpec  As String
Global tmpPLNLowerAgeLimit As String
Global tmpPLNUpperAgeLimit As String
Global tmpPLNStudentAge As String
Global tmpPLNChildUntilAge As String
Global tmpPolTermsStatus As Boolean
Global tmpPlnCode As String
Global tmpPayCode As String
Global tmpHosName As String
Dim TmpChkDigits As String
Global tmpCancerLTLmt As String
Global tmpKidneyLTLmt As String
Global tmpUSRMHBordx As String
Global tmpUSRBlockGL As String
Global tmpReimApp As Boolean

Public Function GenerateMBMNoPol(PayorCode, InsType, PolNO, tmpREnewal, FileStatus)
Dim StrRewSql As String
Dim tmpSeqNo, tmpKey, tmpHisSeqNo  As String
Dim RewSEQ As New ADODB.Recordset
Dim tmpRewSeqID, tmpRewSeqNo As String
Dim tmpINSCode As String
Dim RewMBMFound As Boolean
Dim RewUpdSql As String
Dim SEQHistory As New ADODB.Recordset
Dim HisSql As String
Dim HisFound As Boolean
    tmpSeqNo = ""
    If tmpREnewal = "N" And (FileStatus <> "R" And FileStatus <> "C") Then
        TmpChkDigits = ""
        TmpChkDigits = GenerateChkDigits
        Call Update2MBMSEQ05(PayorCode, tmpSeqNo)
        tmpSeqNo = tmpSeqNo & TmpChkDigits
    End If
    tmpKey = PayorCode & "*" & PolNO
    tmpRewSeqID = "01"
    RewMBMFound = False
    StrRewSql = ""
    StrRewSql = "SELECT RewMBMSEQKey, INSCode, RewSEQNo, RewSEQID From MBMSEQRenewPol WHERE RewMBMSEQKey='" & tmpKey & "'"
    RewSEQ.Open StrRewSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    Do While Not RewSEQ.EOF
        RewMBMFound = True
        tmpINSCode = IIf(Len(RewSEQ!INSCode), RewSEQ!INSCode, "")
        If tmpREnewal = "N" And (FileStatus <> "R" And FileStatus <> "C") Then
            
            tmpHisSeqNo = tmpSeqNo & "*" & RewSEQ!RewSEQID

            HisSql = "SELECT MBMKey, SEQNo From MBMSEQRenewHistory WHERE MBMKey='" & tmpKey & "'"
            SEQHistory.Open HisSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
            HisFound = False
            Do While Not SEQHistory.EOF
                HisFound = True
                Exit Do
            Loop
            SEQHistory.Close
            Set SEQHistory = Nothing
            If HisFound = False Then
                HisSql = "Insert Into MBMSEQRenewHistory(MBMKey, SEQNo) " & _
                         "Values('" & tmpKey & "','" & tmpHisSeqNo & "')"
            Else
                HisSql = "Update MBMSEQRenewHistory set SEQNo = SEQNo + '~" & tmpHisSeqNo & _
                         "' Where MBMKey = '" & tmpKey & "'"
            End If
            medixmasterconnMaint.Execute HisSql
        Else
            tmpHisSeqNo = RewSEQ!RewSEQNo & "*" & RewSEQ!RewSEQID
            tmpSeqNo = RewSEQ!RewSEQNo
            tmpRewSeqID = RewSEQ!RewSEQID + 1
        End If
        RewSEQ.MoveNext
        Exit Do
    Loop
    If RewMBMFound = True Then
        If tmpREnewal = "N" And (FileStatus <> "R" Or FileStatus <> "C") Then
            tmpINSCode = InsType
            RewUpdSql = "Update MBMSEQRenewPol set RewSEQID='" & Format(tmpRewSeqID, "00") & _
                        "', INSCode='" & tmpINSCode & _
                        "', RewSEQNo = '" & tmpSeqNo & "' WHERE RewMBMSEQKey='" & tmpKey & "'"
        Else
            RewUpdSql = "Update MBMSEQRenewPol set RewSEQID='" & Format(tmpRewSeqID, "00") & _
                    "', RewSEQNo = '" & tmpSeqNo & "' WHERE RewMBMSEQKey='" & tmpKey & "'"
        End If
    Else
        tmpINSCode = InsType
        If tmpREnewal <> "N" Then
            TmpChkDigits = ""
            Call Update2MBMSEQ05(PayorCode, tmpSeqNo)
            TmpChkDigits = GenerateChkDigits
            tmpSeqNo = tmpSeqNo & TmpChkDigits
            
        End If
        
        RewUpdSql = "Insert into MBMSEQRenewPol(RewMBMSEQKey, INSCode,RewSEQNo, RewSEQID)" & _
                    " Values ('" & tmpKey & "','" & tmpINSCode & "','" & tmpSeqNo & "','" & Format(tmpRewSeqID, "00") & "')"
    End If
    tmpSeqNo = Trim(tmpINSCode & tmpSeqNo)
    medixmasterconnMaint.Execute RewUpdSql
    GenerateMBMNoPol = tmpSeqNo & "*" & Format(tmpRewSeqID, "00")
End Function

Public Function FormLoadedByName(FormName As String) As Boolean
Dim i As Integer, fnamelc As String
    fnamelc = LCase$(FormName)
    FormLoadedByName = False
    For i = 0 To Forms.Count - 1
        If LCase$(Forms(i).Name) = fnamelc Then
            FormLoadedByName = True
            Exit Function
        End If
    Next
End Function

Public Function AgeCodeAutoNO(AGECode)
Dim AGE As New ADODB.Recordset
Dim AGEInsert As New ADODB.Recordset
Dim strsql As String
Dim AGENumber As String

AGENumber = "01"
strsql = "select * from AGE where AgeCode = '" & Trim(AGECode) & "'"
AGE.Open strsql, medixmasterconn, adOpenStatic, adLockBatchOptimistic
'currentNo = 1

If AGE.recordCount = 0 Then
AGEInsert.Open "Select * from AGE", medixmasterconn, adOpenStatic, adLockOptimistic
With AGEInsert
    .AddNew
    !AGECode = AGECode
    .Update
    End With
    AGEInsert.Close
    Set AGEInsert = Nothing
    End If
    AgeCodeAutoNO = AGECode
    AGE.Close
Set AGE = Nothing

End Function
Public Function GenerateReimAppNo()
    Dim SEQ As New ADODB.Recordset
    Dim SEQInsert As New ADODB.Recordset
    Dim currentNo, i As Integer
    Dim strsql As String
    Dim SEQNumber, SEQCode As String
    Dim CharacterArray
    Dim firstChar, currentChar As String
    

    
    SEQNumber = "00001"
    CharacterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType) & Trim(BusinessYear))
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType))
    'SEQCode = Trim(PayorCode & Trim(InsuredType))
    
    strsql = "Select * from APPREIMSEQ "
    SEQ.Open strsql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    currentNo = 1
    currentChar = "A"
    
    If SEQ.recordCount = 0 Then
        SEQInsert.Open "Select * from APPREIMSEQ ", medixmasterconnMaint, adOpenStatic, adLockOptimistic
        With SEQInsert
            .AddNew
            '!SEQCode = Trim(SEQCode)
            '!PAYCode = PayorCode
            '!PLNCode = PlanCode
            '!INSCode = InsuredType
            !SEQNumber = currentChar & SEQNumber
            SEQNumber = currentChar & SEQNumber
            .Update
        End With
        SEQInsert.Close
        Set SEQInsert = Nothing
    Else
        currentNo = Mid(SEQ!SEQNumber, 2, 5)
        firstChar = Mid(SEQ!SEQNumber, 1, 1)
        Do Until i = 24
            If firstChar = CharacterArray(i) Then
                If currentNo <> "99999" Then
                    currentChar = CharacterArray(i)
                    currentNo = currentNo + 1
                Else
                    currentChar = CharacterArray(i + 1)
                    currentNo = "00001"
                End If
                Exit Do
            End If
            i = i + 1
        Loop
        SEQNumber = currentChar & Format(currentNo, "00000")
        SEQ!SEQNumber = SEQNumber
        SEQ.Update
    End If

    GenerateReimAppNo = SEQNumber
    SEQ.Close
    Set SEQ = Nothing
End Function


Public Function GenerateMBMNo05(PayorCode, InsType, MBMDOB, MBMName, tmpREnewal)
Dim StrRewSql As String
Dim tmpSeqNo, tmpKey, tmpHisSeqNo  As String
Dim RewSEQ As New ADODB.Recordset
Dim tmpRewSeqID, tmpRewSeqNo As String
Dim tmpINSCode As String
Dim RewMBMFound As Boolean
Dim RewUpdSql As String
Dim SEQHistory As New ADODB.Recordset
Dim HisSql As String
Dim HisFound As Boolean
    tmpSeqNo = ""
    If tmpREnewal = "N" Then
        TmpChkDigits = ""
        TmpChkDigits = GenerateChkDigits
        Call Update2MBMSEQ05(PayorCode, tmpSeqNo)
        tmpSeqNo = tmpSeqNo & TmpChkDigits
    End If
    tmpKey = PayorCode & "*" & InsType & "*" & Format(MBMDOB, "yyyymmdd") & "*" & MBMName
    tmpRewSeqID = "01"
    RewMBMFound = False
    StrRewSql = "SELECT RewMBMSEQKey, INSCode, RewSEQNo, RewSEQID From MBMSEQRenew05 WHERE RewMBMSEQKey='" & tmpKey & "'"
    RewSEQ.Open StrRewSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    Do While Not RewSEQ.EOF
        RewMBMFound = True
        tmpINSCode = IIf(Len(RewSEQ!INSCode), RewSEQ!INSCode, "")
        If tmpREnewal = "N" Then
            
            tmpHisSeqNo = tmpSeqNo & "*" & RewSEQ!RewSEQID

            HisSql = "SELECT MBMKey, SEQNo From MBMSEQRenewHistory WHERE MBMKey='" & tmpKey & "'"
            SEQHistory.Open HisSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
            HisFound = False
            Do While Not SEQHistory.EOF
                HisFound = True
                Exit Do
            Loop
            SEQHistory.Close
            Set SEQHistory = Nothing
            If HisFound = False Then
                HisSql = "Insert Into MBMSEQRenewHistory(MBMKey, SEQNo) " & _
                         "Values('" & tmpKey & "','" & tmpHisSeqNo & "')"
            Else
                HisSql = "Update MBMSEQRenewHistory set SEQNo = SEQNo + '~" & tmpHisSeqNo & _
                         "' Where MBMKey = '" & tmpKey & "'"
            End If
            medixmasterconnMaint.Execute HisSql
        Else
            tmpHisSeqNo = RewSEQ!RewSEQNo & "*" & RewSEQ!RewSEQID
            tmpSeqNo = RewSEQ!RewSEQNo
            tmpRewSeqID = RewSEQ!RewSEQID + 1
        End If
        RewSEQ.MoveNext
        Exit Do
    Loop
    If RewMBMFound = True Then
        If tmpREnewal = "N" Then
            tmpINSCode = InsType
            RewUpdSql = "Update MBMSEQRenew05 set RewSEQID='" & Format(tmpRewSeqID, "00") & _
                        "', INSCode='" & tmpINSCode & _
                        "', RewSEQNo = '" & tmpSeqNo & "' WHERE RewMBMSEQKey='" & tmpKey & "'"
        Else
            RewUpdSql = "Update MBMSEQRenew05 set RewSEQID='" & Format(tmpRewSeqID, "00") & _
                    "', RewSEQNo = '" & tmpSeqNo & "' WHERE RewMBMSEQKey='" & tmpKey & "'"
        End If
    Else
        tmpINSCode = InsType
        If tmpREnewal <> "N" Then
            TmpChkDigits = ""
            Call Update2MBMSEQ05(PayorCode, tmpSeqNo)
            TmpChkDigits = GenerateChkDigits
            tmpSeqNo = tmpSeqNo & TmpChkDigits
        End If
        RewUpdSql = "Insert into MBMSEQRenew05(RewMBMSEQKey, INSCode,RewSEQNo, RewSEQID)" & _
                    " Values ('" & tmpKey & "','" & tmpINSCode & "','" & tmpSeqNo & "','" & Format(tmpRewSeqID, "00") & "')"
    End If
    tmpSeqNo = Trim(tmpINSCode & tmpSeqNo)
    medixmasterconnMaint.Execute RewUpdSql
    GenerateMBMNo05 = tmpSeqNo & "*" & Format(tmpRewSeqID, "00")
End Function
Public Function GenerateChkDigits()
Dim strsql As String
Dim SEQDigits As New ADODB.Recordset
Dim tmpSEQDigits As String
    strsql = "SELECT MBMChkDigits From MBMCheckDigits "
    SEQDigits.Open strsql, medixmasterconn, adOpenStatic, adLockOptimistic
        If SEQDigits!MBMChkDigits <> "99" Then
            tmpSEQDigits = Format(SEQDigits!MBMChkDigits + 1, "00")
            SEQDigits!MBMChkDigits = Format(SEQDigits!MBMChkDigits + 1, "00")
            SEQDigits.Update
        Else
            tmpSEQDigits = "01"
            SEQDigits!MBMChkDigits = tmpSEQDigits
            SEQDigits.Update
        End If
    SEQDigits.Close
    Set SEQDigits = Nothing
        GenerateChkDigits = tmpSEQDigits
End Function

Public Function GenerateOPSecurity()
Dim strsql As String
Dim SecDigits As New ADODB.Recordset
Dim tmpSecDigits As String
    strsql = "SELECT MBMSecurityCode From MBMSecurityDigits "
    SecDigits.Open strsql, medixmasterconn, adOpenStatic, adLockOptimistic
        If SecDigits!MBMSecurityCode <> "9999" Then
            tmpSecDigits = Format(SecDigits!MBMSecurityCode + 1, "0000")
            SecDigits!MBMSecurityCode = Format(SecDigits!MBMSecurityCode + 1, "0000")
            SecDigits.Update
        Else
            tmpSecDigits = "1001"
            SecDigits!MBMSecurityCode = tmpSecDigits
            SecDigits.Update
        End If
    SecDigits.Close
    Set SecDigits = Nothing
        GenerateOPSecurity = tmpSecDigits
End Function

Public Function CompanyInfo()
    Dim COM As New ADODB.Recordset

    COM.Open "Medi", medixmasterconn, adOpenKeyset, adLockOptimistic

    CompanyName = COM!COMName
    CompanyAddress = COM!COMAddress
    BusinessYear = COM!COMCurrentYear

    COM.Close
    
    
End Function

Public Function GetAge(EffectiveDate As Date, DateOfBirth As Date)
    Dim CheckDate As Long
    Dim CheckYear As Integer
    Dim CheckDOB As Integer
    
    CheckDate = Year(EffectiveDate) - Year(DateOfBirth)
        If CheckDate <= 17 Then
            If CheckDate <= 0 Then
                If CheckDate <= 0 Then
                    GetAge = "-1D"
                Else
                    GetAge = DateDiff("d", EffectiveDate, DateOfBirth) & "D"
                End If
            Else
                If Month(DateOfBirth) > Month(EffectiveDate) Then
                    GetAge = CheckYear - 1 & "Y"
                Else
                    If Day(DateOfBirth) > Day(EffectiveDate) Then
                        GetAge = CheckYear - 1 & "Y"
                    Else
                        GetAge = CheckYear & "Y"
                    End If
                    
                End If
            End If
    Else
        CheckYear = CheckDate
        'If Month(DateOfBirth) > Month(EffectiveDate) Then
        '    GetAge = CheckYear - 1 & "Y"
        'Else
        '    If Day(DateOfBirth) > Day(EffectiveDate) Then
        '        GetAge = CheckYear - 1 & "Y"
        '    Else
                GetAge = CheckYear & "Y"
        '    End If
            
        'End If
    End If
End Function

Public Function GetAgeBand(HealthValue As String, GetAgeValue As String, plancode As String, InsType As String)
On Error Resume Next
Dim strsql1, strsql2, strsql3 As String
Dim AgeOnly As Integer
Dim AGE As New ADODB.Recordset
Dim AgeCat As String
Dim AgeLenght As String
Dim AgeByPlnRec As New ADODB.Recordset
Dim StrByPln As String
Dim RecFound As Boolean
    strsql1 = ""
    strsql2 = ""
    strsql3 = ""
    AgeLenght = Len(GetAgeValue)
    AgeCat = Mid(GetAgeValue, AgeLenght, 1)
    AgeOnly = Mid(GetAgeValue, 1, InStr(1, GetAgeValue, AgeCat) - 1)
    If Trim(plancode) <> "" Then
        strsql2 = " AND PLNCode='" & Trim(plancode) & "'"
    End If
    If AgeCat = "Y" Then
        strsql1 = "Select AGECODE ,AGEDescription, AGEFrom from AGE " & _
                  "Where cast(substring(AGEfrom ,1  , LEN(AGEFROM)-1   ) as integer) <='" & AgeOnly & "' and  cast(substring(AGETo ,1  , LEN(AGETo)-1   ) as integer) >='" & AgeOnly & "'"
        strsql3 = ""
    Else
        strsql1 = "Select AGECODE ,AGEDescription, AgeFrom from AGE" & _
                  " Where cast(substring(AGEfrom ,1  , LEN(AGEFROM)-1   ) as integer) <= '" & AgeOnly & "'"
        strsql3 = " order by AgeFrom"
    End If
    StrByPln = strsql1 & strsql2 & strsql3
    
    'AGE.MaxRecords = 1
    AGE.Open StrByPln, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    RecFound = False
    If AGE.EOF = True Then
        AGE.Close
        Set AGE = Nothing
        'AGE.MaxRecords = 1
        strsql2 = " AND PLNCode IS NULL "
        strsql1 = strsql1 & strsql2 & strsql3
        AGE.Open strsql1, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If AGE.EOF = False Then
            RecFound = True
        End If
    Else
        RecFound = True
    End If
    If RecFound Then
        If AgeCat = "D" Then
            AgeLenght = Len(AGE!AGEFrom)
            If AgeOnly >= Mid(AGE!AGEFrom, 1, AgeLenght - 1) Then
                GetAgeBand = AGE!AGECode & " - " & AGE!AGEDescription
            End If
        Else
            GetAgeBand = AGE!AGECode & " - " & AGE!AGEDescription
        End If
    Else
        MsgBox "Age Band not found or Date out of Acceptable Range! " & _
            vbNewLine & "Please verify your maintenance table!", vbCritical + vbOKOnly
    End If
    AGE.Close
    Set AGE = Nothing
End Function

Public Function GetPremium(InsureType As String, PayorValue As String, AgeValue As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, Sex As String) As Double
'On Error Resume Next
    Dim prm As New ADODB.Recordset
    Dim strsql As String
    Dim PLN As New ADODB.Recordset
    Dim strsqlPLN As String
    Dim tmpPremGenderStatus As String
    
    strsql = "SELECT PRMAmount,PRMAmountFemale  From PRM "
    strsql = strsql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') "
    strsql = strsql & "AND (PLNCode = '" & Trim(PlanValue) & "') AND (AGECode = '" & Trim(AgeValue) & "') "
    strsql = strsql & "AND ( PRMEffDate  <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "') "
    If Trim(HealthValue) <> "S" Then
        strsql = strsql & "AND (PAYCode = '" & PayorValue & "')"
    End If
    strsql = strsql & " ORDER BY PRMEffDate DESC "
    prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'strsqlPLN = ""
    strsqlPLN = "SELECT PLNPremGenderStatus  From PLN "
    strsqlPLN = strsqlPLN & "WHERE (HLTCode = '" & Trim(HealthValue) & "') "
    strsqlPLN = strsqlPLN & "AND (PLNCode = '" & Trim(PlanValue) & "')  "
    If Trim(HealthValue) <> "S" Then
        strsqlPLN = strsqlPLN & "AND (PAYCode = '" & PayorValue & "')"
    End If
    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If PLN.recordCount Then
        tmpPremGenderStatus = IIf(Len(PLN!PLNPremGenderStatus), PLN!PLNPremGenderStatus, "N")
    Else
        tmpPremGenderStatus = ""
    End If
    
    If prm.recordCount Then
        'If PRM!SuppPrmStatus = True Then
        '    If PRM!SuppPRMAmount <> "" Then
        '        GetPremium = PRM!SuppPRMAmount
        '    Else
        '        MsgBox "Premium Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
        '    End If
        'Else
            If prm!PRMAmount <> "" Then
            
                If Trim(tmpPremGenderStatus) = "Y" Then
                    If Trim(Sex) = "F" Then
                        GetPremium = prm!PRMAmountFemale
                    Else
                        GetPremium = prm!PRMAmount
                    End If
                Else
                    GetPremium = prm!PRMAmount
                End If
            Else
                MsgBox "Premium Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
            End If
    'End If
    Else
        MsgBox "Age Out of Premium Range! Please check your Premium Table."
    End If
        
    prm.Close
    Set prm = Nothing
End Function

Public Function GetHISPremium(InsureType As String, PayorValue As String, AgeValue As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, tmpMBMDataType As String, CovID As String, tmpPLNPremSexStatus As String, tmpVSex As String) As Double
    Dim prm As New ADODB.Recordset
    Dim strsql As String
    Dim TmpSuppPrmStatus As String
    If tmpMBMDataType = "P" Then
        strsql = "SELECT SuppPrmStatus, PRMAmount as PrmAmt, PRMAmountFemale From PRM "
    Else
        strsql = "SELECT SuppPrmStatus, SuppPRMAmt as PrmAmt, PRMAmountFemale From PRM "
    End If
    TmpSuppPrmStatus = "A"
    strsql = strsql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') "
    strsql = strsql & "AND (PLNCode = '" & Trim(PlanValue) & "') AND (AGECode = '" & Trim(AgeValue) & "') "
    strsql = strsql & "AND ( PRMEffDate  <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "') "
    
    If Trim(HealthValue) <> "S" Then
        strsql = strsql & "AND (PAYCode = '" & PayorValue & "')"
    End If
    
    strsql = strsql & " ORDER BY PRMEffDate DESC "
    prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If prm.recordCount Then
        TmpSuppPrmStatus = IIf(Len(prm!SuppPrmStatus), prm!SuppPrmStatus, "A")
'        If IsNumeric(prm!PrmAmt) = True Then
        If Trim(tmpPLNPremSexStatus) = "Y" Then
            If Trim(tmpVSex) = "F" Then
                GetHISPremium = IIf(IsNumeric(prm!PRMAmountFemale), prm!PRMAmountFemale, 0)
            Else
                GetHISPremium = IIf(IsNumeric(prm!PrmAmt), prm!PrmAmt, 0)
            End If
        Else
            GetHISPremium = IIf(IsNumeric(prm!PrmAmt), prm!PrmAmt, 0)
        End If
'        Else
'            MsgBox "Premium Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
'        End If
    Else
        MsgBox "Age Out of Premium Range! Please check your Premium Table."
    End If
    
    prm.Close
    Set prm = Nothing
    If tmpMBMDataType = "C" Then   ' CoverPersons
        If TmpSuppPrmStatus = "A" And CovID <> "00" Then
            GetHISPremium = 0
        ElseIf TmpSuppPrmStatus = "B" And CovID <> "01" Then
            GetHISPremium = 0
        End If
    End If

End Function

Public Function GetCLNPremium(InsureType As String, PayorValue As String, AgeValue As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double
'On Error Resume Next
    Dim prm As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "SELECT PRMAmount  From PRM "
    strsql = strsql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') "
    strsql = strsql & "AND (PLNCode = '" & Trim(PlanValue) & "') AND (AGECode = '" & Trim(AgeValue) & "') "
    strsql = strsql & "AND ( PRMEffDate  <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "') "
    If Trim(HealthValue) <> "S" Then
        strsql = strsql & "AND (PAYCode = '" & PayorValue & "')"
    End If
    strsql = strsql & " ORDER BY PRMEffDate DESC"
    prm.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If prm.recordCount Then
        
            If prm!PRMAmount <> "" Then
                GetCLNPremium = prm!PRMAmount
            Else
'                MsgBox "Premium Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
            End If
    
    Else
'        MsgBox "Age Out of Premium Range! Please check your Premium Table."
    End If
        
    prm.Close
    Set prm = Nothing
End Function

Public Function GetCLNSVisitLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double
Dim AL As New ADODB.Recordset
Dim strsql As String
          
    strsql = "Select VisitLimit from VisitLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND VisitEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by VisitEffDate  desc"

    AL.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
         GetCLNSVisitLimit = AL!VisitLimit
    Else
'         MsgBox "Visit Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
End Function

Public Function GetCLNSMCO(InsureType As String, HealthValue As String, AdultOrChild As String, PayorEffDate As Date, PLNCode As String) As Double
Dim MCO As New ADODB.Recordset
Dim strsql As String
Dim MCONew As New ADODB.Recordset
Dim strsqlNEW As String
    
    strsqlNEW = "SELECT MCOIndex, * From MCO "
    strsqlNEW = strsqlNEW & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') AND (PLNCode = '" & Trim(PLNCode) & "') " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                
    MCO.Open strsqlNEW, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If MCO.recordCount < 1 Then
        MCO.Close
        Set MCO = Nothing
        strsql = "SELECT MCOIndex, * From MCO "
        strsql = strsql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') AND PLNCode IS NULL " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                
        MCO.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    End If
    If MCO.recordCount Then
        If AdultOrChild = "A" Then
            GetCLNSMCO = IIf(IsNumeric(MCO!MCOSuppAdultAmt) = True, MCO!MCOSuppAdultAmt, 0)
        ElseIf AdultOrChild = "C" Then
            GetCLNSMCO = IIf(IsNumeric(MCO!MCOSuppChildAmt) = True, MCO!MCOSuppChildAmt, 0)
        End If
    Else
'        MsgBox "MCO Information Is Not Correct" & vbNewLine & " Please Check your MCO table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    MCO.Close
    Set MCO = Nothing
End Function

Public Function GetCLNSuppLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double

    Dim AL As New ADODB.Recordset
    Dim strsql As String
  
    strsql = "Select suppLimit from AnnualLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnEffDate  desc"

    AL.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
            
         GetCLNSuppLimit = IIf(IsNull(AL!SuppLimit), 0, AL!SuppLimit)
    Else
'         MsgBox "Annual Supplementary Limit Not Found! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
End Function

Public Function GetCLNSPremium(InsureType As String, PayorValue As String, AgeValue As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double
'On Error Resume Next
    Dim prm As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "SELECT SuppPrmStatus, SuppPRMAmt  From PRM "
    strsql = strsql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') "
    strsql = strsql & "AND (PLNCode = '" & Trim(PlanValue) & "') AND (AGECode = '" & Trim(AgeValue) & "') "
    strsql = strsql & "AND (PAYCode = '" & Trim(PayorValue) & "') "
    strsql = strsql & "AND ( PRMEffDate  <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "') "
    strsql = strsql & " ORDER BY PRMEffDate DESC"

    prm.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If prm.recordCount Then
        If prm!SuppPrmAmt <> "" Then
            GetCLNSPremium = prm!SuppPrmAmt
        Else
'            MsgBox "Premium Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    Else
'        MsgBox "Age Out of Premium Range! Please check your Premium Table."
    End If
        
    prm.Close
    Set prm = Nothing
End Function

Public Function GetSPremium(InsureType As String, PayorValue As String, AgeValue As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double
    Dim prm As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "SELECT SuppPrmStatus, SuppPRMAmt  From PRM "
    strsql = strsql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') "
    strsql = strsql & "AND (PLNCode = '" & Trim(PlanValue) & "') AND (AGECode = '" & Trim(AgeValue) & "') "
    strsql = strsql & "AND ( PRMEffDate  <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "') "
    If Trim(HealthValue) <> "S" Then
        strsql = strsql & "AND (PAYCode = '" & PayorValue & "')"
    End If
    strsql = strsql & " ORDER BY PRMEffDate DESC "
    'PRM.Close
    'Set PRM = Nothing
    prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If prm.recordCount Then
        'If PRM!SuppPrmStatus = True Then
            If prm!SuppPrmAmt <> "" Then
                GetSPremium = prm!SuppPrmAmt
            Else
                MsgBox "Premium Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
            End If
        'Else
        '    If PRM!PRMAmount <> "" Then
        '        GetPremium = PRM!PRMAmount
        '    Else
        '        MsgBox "Premium Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
        '    End If
        'End If
    Else
        MsgBox "Age Out of Premium Range! Please check your Premium Table."
    End If
        
    prm.Close
    Set prm = Nothing
End Function

Public Function GetHISMCO(InsureType As String, HealthValue As String, AdultOrChild As String, PayorEffDate As Date, PLNCode As String, MBMDataType As String, CovID As String) As Double
Dim MCO As New ADODB.Recordset
Dim strsql, strsqlNEW As String
Dim MCONew As New ADODB.Recordset
Dim tmpSql, tmpField As String
Dim tmpCovIDChrg, tmpSuppMcoStatus As String

    tmpSql = "SELECT SuppMcoStatus, CovIDChrg, "
    tmpField = "MCOAmountAdult as MCOAmt From MCO"
    If MBMDataType = "P" Then
        If AdultOrChild = "C" Then
            tmpField = "MCOAmountChild as MCOAmt From MCO"
        End If
    ElseIf MBMDataType = "C" Then
        If AdultOrChild = "A" Then
            tmpField = "MCOSuppAdultAmt as MCOAmt From MCO"
        ElseIf AdultOrChild = "C" Then
            tmpField = "MCOSuppChildAmt as MCOAmt From MCO"
        End If
    End If
    tmpSql = tmpSql & tmpField
    
    strsqlNEW = tmpSql & " WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') AND (PLNCode = '" & Trim(PLNCode) & "') " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                
    MCONew.Open strsqlNEW, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    strsql = tmpSql & " WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') and PLNCode is Null " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                
    MCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    tmpSuppMcoStatus = "A"
    If MCONew.recordCount Then
        GetHISMCO = IIf(IsNumeric(MCONew!MCOAmt), MCONew!MCOAmt, 0)
        tmpSuppMcoStatus = IIf(Len(MCONew!SuppMCOStatus), MCONew!SuppMCOStatus, "A")
        tmpCovIDChrg = IIf(Len(MCONew!CovIDChrg), MCONew!CovIDChrg, "00")
    ElseIf MCO.recordCount Then
        GetHISMCO = IIf(IsNumeric(MCO!MCOAmt), MCO!MCOAmt, 0)
        tmpSuppMcoStatus = IIf(Len(MCO!SuppMCOStatus), MCO!SuppMCOStatus, "A")
        tmpCovIDChrg = IIf(Len(MCO!CovIDChrg), MCO!CovIDChrg, "00")
   ' Else
   '     MsgBox "MCO Information Is Not Correct" & vbNewLine & " Please Check your MCO table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    MCONew.Close
    Set MCONew = Nothing
    MCO.Close
    Set MCO = Nothing
    If MBMDataType = "C" Then
        If tmpSuppMcoStatus = "A" And CovID <> "00" Then
            GetHISMCO = 0
        ElseIf tmpSuppMcoStatus = "B" And CovID <> "01" Then
            GetHISMCO = 0
        ElseIf tmpSuppMcoStatus = "D" And CovID < tmpCovIDChrg Then
            GetHISMCO = 0
        End If
    End If
End Function

Public Function GetMCO(InsureType As String, HealthValue As String, AdultOrChild As String, PayorEffDate As Date, PLNCode As String) As Double
'On Error Resume Next
    Dim MCO As New ADODB.Recordset
    Dim strsql As String
    Dim MCONew As New ADODB.Recordset
    Dim strsqlNEW As String
    
    strsqlNEW = "SELECT MCOCode, * From MCO "
    strsqlNEW = strsqlNEW & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') AND (PLNCode = '" & Trim(PLNCode) & "') " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                
    MCONew.Open strsqlNEW, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    strsql = "SELECT MCOCode, * From MCO "
    strsql = strsql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') and PLNCode is Null " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                
    MCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MCONew.recordCount Then
        If AdultOrChild = "A" Then
            GetMCO = MCONew!MCOAmountAdult
        ElseIf AdultOrChild = "C" Then
            GetMCO = MCONew!MCOAmountChild
        'Else
        '    MsgBox "MCO Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    ElseIf MCO.recordCount Then
        If AdultOrChild = "A" Then
            GetMCO = MCO!MCOAmountAdult
        ElseIf AdultOrChild = "C" Then
            GetMCO = MCO!MCOAmountChild
        'Else
        '    MsgBox "MCO Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    'Else
    '    MsgBox "MCO Information Is Not Correct" & vbNewLine & " Please Check your MCO table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    MCONew.Close
    Set MCONew = Nothing
    MCO.Close
    Set MCO = Nothing
End Function

Public Function GetSuppMCO(InsureType As String, HealthValue As String, AdultOrChild As String, PayorEffDate As Date, PLNCode As String) As Double
'On Error Resume Next
Dim MCO As New ADODB.Recordset
Dim strsql As String
Dim MCONew As New ADODB.Recordset
Dim strsqlNEW As String
Dim strsql2 As String
Dim tmpMCO As String

    strsql = "SELECT MCOSuppAdultAmt AS SuppMCOAmt From MCO "
    If Trim(AdultOrChild) = "C" Then
        strsql = "SELECT MCOSuppChildAmt AS SuppMCOAmt From MCO "
    End If
    strsql = strsql & "WHERE INSCode='" & Trim(InsureType) & "' AND HLTCode='" & Trim(HealthValue) & "'"
    
    strsql2 = " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
    
    strsqlNEW = strsql & " AND PLNCode='" & Trim(PLNCode) & "'" & strsql2
    
    MCONew.Open strsqlNEW, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MCONew.recordCount Then
        GetSuppMCO = IIf(IsNumeric(MCONew!SuppMCOAmt), MCONew!SuppMCOAmt, 0)
    Else
        strsqlNEW = ""
        strsqlNEW = strsql & " AND PLNCode IS NULL " & strsql2
        MCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MCO.recordCount Then
            GetSuppMCO = IIf(IsNumeric(MCO!SuppMCOAmt), MCO!SuppMCOAmt, 0)
        'Else
        '    MsgBox "MCO Information Is Not Correct" & vbNewLine & " Please Check your MCO table", vbOKOnly + vbCritical, DialogBoxTitle
        End If
        MCO.Close
        Set MCO = Nothing
    End If
    MCONew.Close
    Set MCONew = Nothing
End Function


Public Function GetCISAnnLmt(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, tmpDataType As String, CovID As String) As Double
Dim AL As New ADODB.Recordset
Dim strsql As String
Dim tmpSuppLmtSta As String

    If tmpDataType = "P" Then
        strsql = "SELECT SuppLimitStatus, AnnLimit as AnnLmt From AnnualLimit "
    Else
        strsql = "SELECT SuppLimitStatus, SuppLimit as AnnLmt From AnnualLimit "
    End If
    tmpSuppLmtSta = "A"
    GetCISAnnLmt = 0
    strsql = strsql & " Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnEffDate  desc"
    AL.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
         GetCISAnnLmt = IIf(IsNumeric(AL!AnnLmt), AL!AnnLmt, 0)
         tmpSuppLmtSta = IIf(Len(AL!SuppLimitStatus), AL!SuppLimitStatus, "A")
    Else
         MsgBox "Clinical Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
    
    If tmpDataType = "C" Then   ' CoverPersons
        If tmpSuppLmtSta = "A" And CovID <> "00" Then
            GetCISAnnLmt = 0
        ElseIf tmpSuppLmtSta = "B" And CovID <> "01" Then
            GetCISAnnLmt = 0
        End If
    End If
End Function


Public Function GetCisMCO(InsureType As String, HealthValue As String, AdultOrChild As String, PayorEffDate As Date, PLNCode As String, MBMDataType As String, CovID As String) As Double
Dim MCO As New ADODB.Recordset
Dim strsql, strsqlNEW As String
Dim MCONew As New ADODB.Recordset
Dim tmpSql, tmpField, tmpSuppMcoStatus As String

    tmpSql = "SELECT SuppMcoStatus, "
    tmpField = "MCOAmtAdult as MCOAmt From MCO"
    If MBMDataType = "P" Then
        If AdultOrChild = "C" Then
            tmpField = "MCOAmtChild as MCOAmt From MCO"
        End If
    ElseIf MBMDataType = "C" Then
        If AdultOrChild = "A" Then
            tmpField = "MCOSuppAdultAmt as MCOAmt From MCO"
        ElseIf AdultOrChild = "C" Then
            tmpField = "MCOSuppChildAmt as MCOAmt From MCO"
        End If
    End If
    tmpSql = tmpSql & tmpField
    
    strsqlNEW = tmpSql & " WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') AND (PLNCode = '" & Trim(PLNCode) & "') " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                
    MCONew.Open strsqlNEW, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    strsql = tmpSql & " WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') and PLNCode is Null " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                
    MCO.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    tmpSuppMcoStatus = "A"
    If MCONew.recordCount Then
        GetCisMCO = IIf(IsNumeric(MCONew!MCOAmt), MCONew!MCOAmt, 0)
        tmpSuppMcoStatus = IIf(Len(MCONew!SuppMCOStatus), MCONew!SuppMCOStatus, "A")
'        tmpCovIDChrg = IIf(Len(MCONew!CovIDChrg), MCONew!CovIDChrg, "00")
    ElseIf MCO.recordCount Then
        GetCisMCO = IIf(IsNumeric(MCO!MCOAmt), MCO!MCOAmt, 0)
        tmpSuppMcoStatus = IIf(Len(MCO!SuppMCOStatus), MCO!SuppMCOStatus, "A")
'        tmpCovIDChrg = IIf(Len(MCO!CovIDChrg), MCO!CovIDChrg, "00")
    'Else
    '    MsgBox "Clinical MCO Information Is Not Correct" & vbNewLine & " Please Check your MCO table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    MCONew.Close
    Set MCONew = Nothing
    MCO.Close
    Set MCO = Nothing
    If MBMDataType = "C" Then
        If tmpSuppMcoStatus = "A" And CovID <> "00" Then
            GetCisMCO = 0
        ElseIf tmpSuppMcoStatus = "B" And CovID <> "01" Then
            GetCisMCO = 0
'        ElseIf tmpSuppMcoStatus = "D" And CovID < tmpCovIDChrg Then
'            GetHISMCO = 0
        End If
    End If
End Function

Public Function GetCISPremium(InsureType As String, PayorValue As String, AgeValue As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, tmpMBMDataType As String, CovID As String) As Double
    Dim prm As New ADODB.Recordset
    Dim strsql As String
    Dim TmpSuppPrmStatus As String
    If tmpMBMDataType = "P" Then
        strsql = "SELECT SuppPrmStatus, PRMAmount as PrmAmt From PRM "
    Else
        strsql = "SELECT SuppPrmStatus, SuppPRMAmt as PrmAmt From PRM "
    End If
    TmpSuppPrmStatus = "A"
    strsql = strsql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') "
    strsql = strsql & "AND (PLNCode = '" & Trim(PlanValue) & "') AND (AGECode = '" & Trim(AgeValue) & "') "
    strsql = strsql & "AND ( PRMEffDate  <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "') "
    If Trim(HealthValue) <> "S" Then
        strsql = strsql & "AND (PAYCode = '" & PayorValue & "')"
    End If
    strsql = strsql & " ORDER BY PRMEffDate DESC "
    prm.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If prm.recordCount Then
        TmpSuppPrmStatus = IIf(Len(prm!SuppPrmStatus), prm!SuppPrmStatus, "A")
        If IsNumeric(prm!PrmAmt) = True Then
            GetCISPremium = prm!PrmAmt
'            If tmpVPolDis > 0 Then
'                GetCISPremium = GetHISPremium * (100 - tmpVPolDis) / 100
'            End If
'            If tmpVRenewDis > 0 Then
'                GetCISPremium = GetHISPremium * (100 - tmpVRenewDis) / 100
'            End If
        Else
            MsgBox "Premium Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    Else
        MsgBox "Age Out of Premium Range! Please check your Premium Table."
    End If
    
    prm.Close
    Set prm = Nothing
    If tmpMBMDataType = "C" Then   ' CoverPersons
        If TmpSuppPrmStatus = "A" And CovID <> "00" Then
            GetCISPremium = 0
        ElseIf TmpSuppPrmStatus = "B" And CovID <> "01" Then
            GetCISPremium = 0
        End If
    End If

End Function

Public Function GetCLNMCO(InsureType As String, HealthValue As String, AdultOrChild As String, PayorEffDate As Date, PLNCode As String) As Double
Dim MCO As New ADODB.Recordset
Dim strsql As String
Dim MCONew As New ADODB.Recordset
Dim strsqlNEW As String
    
    strsqlNEW = "SELECT MCOIndex, * From MCO "
    strsqlNEW = strsqlNEW & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') AND (PLNCode = '" & Trim(PLNCode) & "') " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate DESC"
                
    MCONew.Open strsqlNEW, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    strsql = "SELECT MCOIndex, * From MCO "
    strsql = strsql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') AND PLNCode is Null " & _
            " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate DESC"
                
    MCO.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If MCONew.recordCount Then
        If AdultOrChild = "A" Then
            GetCLNMCO = IIf(IsNumeric(MCONew!MCOAmtAdult) = True, MCONew!MCOAmtAdult, 0)
        ElseIf AdultOrChild = "C" Then
            GetCLNMCO = IIf(IsNumeric(MCONew!MCOAmtChild) = True, MCONew!MCOAmtChild, 0)
        Else
'            MsgBox "MCO Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    ElseIf MCO.recordCount Then
        If AdultOrChild = "A" Then
            GetCLNMCO = IIf(IsNumeric(MCO!MCOAmtAdult) = True, MCO!MCOAmtAdult, 0)
        ElseIf AdultOrChild = "C" Then
            GetCLNMCO = IIf(IsNumeric(MCO!MCOAmtChild) = True, MCO!MCOAmtChild, 0)
        Else
'            MsgBox "MCO Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    Else
 '       MsgBox "MCO Information Is Not Correct" & vbNewLine & " Please Check your MCO table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    MCONew.Close
    Set MCONew = Nothing
    MCO.Close
    Set MCO = Nothing
End Function

Public Function GetSuppLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double

    Dim AL As New ADODB.Recordset
    Dim strsql As String
  
    strsql = "Select suppLimit from AnnualLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnualEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"

    AL.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
            
         GetSuppLimit = IIf(IsNull(AL!SuppLimit), 0, AL!SuppLimit)
    Else
         MsgBox "Annual Supplementary Limit Not Found! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
End Function


Public Function GetAnnualLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double

    Dim AL As New ADODB.Recordset
    Dim strsql As String
  
    strsql = "Select AnnLimit from AnnualLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnualEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"
    AL.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
         GetAnnualLimit = IIf(Len(AL!AnnLimit), AL!AnnLimit, 0)
    Else
         MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
End Function

Public Function GetHISAnnLmt(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, tmpMBMDataType As String, CovID As String) As Double
Dim AL As New ADODB.Recordset
Dim strsql As String
Dim tmpSuppLmtSta As String

    If tmpMBMDataType = "P" Then
        strsql = "SELECT SuppLimitStatus, AnnLimit as AnnLmt From AnnualLimit "
    Else
        strsql = "SELECT SuppLimitStatus, SuppLimit as AnnLmt From AnnualLimit "
    End If
    tmpSuppLmtSta = "A"
    GetHISAnnLmt = 0
    strsql = strsql & " Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnualEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"
    AL.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
         GetHISAnnLmt = IIf(IsNumeric(AL!AnnLmt), AL!AnnLmt, 0)
         tmpSuppLmtSta = IIf(Len(AL!SuppLimitStatus), AL!SuppLimitStatus, "A")
    Else
         MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table" & "-" & "(Plan Clode" & PlanValue & ")", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
    
    If tmpMBMDataType = "C" Then   ' CoverPersons
        If tmpSuppLmtSta = "A" And CovID <> "00" Then
            GetHISAnnLmt = 0
        ElseIf tmpSuppLmtSta = "B" And CovID <> "01" Then
            GetHISAnnLmt = 0
        End If
    End If
End Function

Public Function GetHISLTKidneyCancer(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, tmpMBMDataType As String, CovID As String) As Double
Dim AL As New ADODB.Recordset
Dim strsql As String
Dim tmpSuppLmtSta As String
    
    'If tmpMBMDataType = "P" Then
    strsql = "SELECT SuppLimitStatus, OPKidneyLTLmt,OPCancerTLmt as AnnLmt From AnnualLimit "
    'End If
    tmpSuppLmtSta = "A"
    tmpCancerLTLmt = 0
    tmpKidneyLTLmt = 0
    strsql = strsql & " Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnualEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"
    AL.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
        tmpCancerLTLmt = IIf(IsNumeric(AL!AnnLmt), AL!AnnLmt, 0)
        tmpKidneyLTLmt = IIf(IsNumeric(AL!OPKidneyLTLmt), AL!OPKidneyLTLmt, 0)
        tmpSuppLmtSta = IIf(Len(AL!SuppLimitStatus), AL!SuppLimitStatus, "A")
    Else
         MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
    
    'If tmpMBMDataType = "C" Then   ' CoverPersons
    '    If tmpSuppLmtSta = "A" And CovID <> "00" Then
    '        GetHISAnnLmt = 0
    '    ElseIf tmpSuppLmtSta = "B" And CovID <> "01" Then
    '        GetHISAnnLmt = 0
    '    End If
    'End If
End Function

Public Function GetHISLifeTimeLmt(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, tmpMBMDataType As String, CovID As String) As Double
Dim AL As New ADODB.Recordset
Dim strsql As String
Dim TmpSuppPrmStatus As String

    If tmpMBMDataType = "P" Then
        strsql = "SELECT SuppLimitStatus, LifeTimeLimit as LifeTimeLmt From AnnualLimit "
    Else
        strsql = "SELECT SuppLimitStatus, SuppLifeTimeLimit as LifeTimeLmt From AnnualLimit "
    End If
    TmpSuppPrmStatus = "A"
    GetHISLifeTimeLmt = 0
    strsql = strsql & " Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnualEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"
    AL.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
         GetHISLifeTimeLmt = IIf(IsNumeric(AL!LifeTimeLmt), AL!LifeTimeLmt, 0)
         TmpSuppPrmStatus = IIf(Len(AL!SuppLimitStatus), AL!SuppLimitStatus, "A")
    Else
         MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
    
    If tmpMBMDataType = "C" Then   ' CoverPersons
        If TmpSuppPrmStatus = "A" And CovID <> "00" Then
            GetHISLifeTimeLmt = 0
        ElseIf TmpSuppPrmStatus = "B" And CovID <> "01" Then
            GetHISLifeTimeLmt = 0
        End If
    End If
End Function
Public Function GetCLNAnnualLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double
Dim AL As New ADODB.Recordset
Dim strsql As String
          
    strsql = "Select AnnLimit from AnnualLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnEffDate  desc"
    AL.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
         GetCLNAnnualLimit = AL!AnnLimit
    Else
         GetCLNAnnualLimit = 0
'         MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
End Function

Public Function GetVisitLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double
Dim AL As New ADODB.Recordset
Dim strsql As String
          
    strsql = "Select VisitLimit from VisitLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND VisitEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by VisitEffDate  desc"
    AL.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
         GetVisitLimit = AL!VisitLimit
    Else
         GetVisitLimit = 0
'         MsgBox "Visit Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
End Function

Public Function GetCISLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, AnnLmtType As String) As Double
Dim AL As New ADODB.Recordset
Dim strsql As String
    strsql = "Select " & Trim(AnnLmtType) & " AS AnnLmt "
    
    strsql = strsql & " from AnnualLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & _
             "' And HLTCode = '" & Trim(HealthValue) & "' AND AnnEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "'" & _
             " order by AnnEffDate  desc"
    AL.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
        GetCISLimit = AL!AnnLmt
    Else
'         MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
End Function

Public Function GetProviderCharges(ServiceType As String, InsureType As String, HealthValue As String, AdultChild As String) As Double
On Error Resume Next
    Dim PRV As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "Select * from PRV Where SRVCode ='" & ServiceType & "' And INSCode ='" & InsureType & "' And HLTCode ='" & HealthValue & "'"
    PRV.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If PRV.recordCount = 0 Then
        MsgBox "Provider Charges Information Not Correct", vbCritical + vbOKOnly, DialogBoxTitle
    Else
        'If AdultChild = "C" Then
        '    GetProviderCharges = PRV!PRVProviderChargesChild & "-" & PRV!PRVCaseChargesChild
        'ElseIf AdultChild = "A" Then
            'GetProviderCharges = PRV!PRVProviderCharges & "-" & PRV!PRVCaseCharges
            GetProviderCharges = PRV!PRVProviderCharges
        'End If
    End If
    PRV.Close
    Set PRV = Nothing
End Function

Public Function CheckBatch(Payor As String, BordxDate As Date, BordxType As String) As String
    Dim BAT As New ADODB.Recordset
    Dim strsql As String
    Dim CurrentBatch As Integer
    
    strsql = "Select * from BAT Where PAYCode ='" & Payor & "' And BATBordxDate = '" & Format(BordxDate, "MM/DD/YYYY") & "' And BordxType = '" & BordxType & "'"
    BAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                                                    
    'If BordxType = "N" Then
    If BAT.recordCount = 0 Then
        BAT.AddNew
        With BAT
        If BordxType = "N" Then
            !BordxType = "N"
        ElseIf BordxType = "E" Then
            !BordxType = "E"
        End If
            !PAYCode = Payor
            !BATBordxDate = BordxDate
            !BATSequence = "1"
        End With
        BAT.Update
        CheckBatch = "1"
    Else
        With BAT
            CurrentBatch = !BATSequence + 1
            !BATSequence = CurrentBatch
            BAT.Update
        End With
        CheckBatch = CurrentBatch
    End If
    'End If
    
    'If BordxType = "E" Then
    '    If BAT.recordCount = 0 Then
    '        BAT.AddNew
    '        With BAT
    '            !PAYCode = Payor
    '            !BATBordxDate = BordxDate
    '            !BATSequence = "1"
    '        End With
    '        BAT.Update
    '        CheckBatch = "1"
    '    Else
    '        With BAT
    '            CurrentBatch = !BATSequence + 1
    '            !BATSequence = CurrentBatch
    '            BAT.Update
    '        End With
    '        CheckBatch = CurrentBatch
    '    End If
    'End If
    
End Function

Public Function l(stringvalue As String, stringlenght As Integer)
Dim stringformat
Dim LengLoop As Integer

    For LengLoop = 1 To stringlenght
        stringformat = stringformat + "1"
    Next LengLoop
    LSet stringformat = stringvalue
    l = stringformat
End Function

Public Function GetFirstAuthoriseCode(CodeType As String, Byear As String) As String
    Dim AUC As New ADODB.Recordset
    Dim strsql As String
    Dim AuthoriseCode As String
    Dim Code As String
    
    
    strsql = "Select * from AUC Where AUCType = '" & CodeType & "' and AUCBYear = '" & Mid(BusinessYear, 3, 2) & "'"
    AUC.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
    If AUC.recordCount = 0 Then
        With AUC
            .AddNew
            If CodeType = "F" Then
                !AUCSequence = "0001"
                AuthoriseCode = "0001"
                !AUCType = "F"
                !AUCBYear = Byear
            Else
                !AUCSequence = "0000001"
                AuthoriseCode = "0000001"
                !AUCType = "S"
                !AUCBYear = Byear
            End If
            !AUCVAlueDate = Date
            .Update
        End With
    Else
    
        With AUC
            If CodeType = "F" Then
                !AUCSequence = Format(!AUCSequence + 1, "0000")
                AuthoriseCode = Format(!AUCSequence + 1, "0000")
            Else
                !AUCSequence = Format(!AUCSequence + 1, "0000000")
                AuthoriseCode = Format(!AUCSequence + 1, "0000000")
            End If
        .Update
        End With
    End If

    GetFirstAuthoriseCode = AuthoriseCode
    AUC.Close
    Set AUC = Nothing
End Function

Public Function GenerateCaseNo(PayorCode)
    Dim CAQ As New ADODB.Recordset
    Dim strsql As String
    Dim CAQNumber As String
    Dim CAQCharacter As String
    'strsql = "Select * from CAQ Where PAYCode = '" & PayorCode & "' And PLNCode = '" & PlanCode & "'"
    strsql = "Select * from CAQ2013 Where PAYCode = '" & PayorCode & "'"
    CAQ.Open strsql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    
    If CAQ.recordCount = 0 Then
        With CAQ
            .AddNew
            '!CAQCode = Trim(PayorCode) & Trim(PlanCode) & BusinessYear
            !PAYCode = PayorCode
            '!PLNCode = PlanCode
            !CAQNumber = "00000001"
            .Update
            CAQNumber = "00000001"
        End With
    Else
        CAQ!CAQNumber = Format(CAQ!CAQNumber + 1, "00000000")
        CAQNumber = CAQ!CAQNumber
        'CAQ.Update
        
        'CAQ.Close
        Set CAQ = Nothing
        
        Dim strsql1 As String
        strsql1 = "EXEC SP_castrans '" & PayorCode & "'"
        CAQ.Open strsql1, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    End If
    'GenerateCaseNo = PayorCode & PlanCode & Mid(BusinessYear, 3, 2) & CAQNumber
    GenerateCaseNo = PayorCode & CAQNumber
    'CAQ.Close
    Set CAQ = Nothing
End Function

Public Function GetAdjustmentNo(MembershipNo As String, Policyno As String, CaseNo As String, AdjustmentType As String) As Integer
    Dim ADJ As New ADODB.Recordset
    Dim adjstr As String
    Dim NextSequence As Integer
    
    adjstr = "Select * from ADJ Where ADJMemberNo ='" & MembershipNo & "' "
    adjstr = adjstr & "And ADJPolicyNo ='" & Policyno & "' And ADJCaseNo ='" & CaseNo & "' And ADJType ='" & AdjustmentType & "'"
    ADJ.Open adjstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If ADJ.recordCount = 0 Then
        ADJ.AddNew
         With ADJ
            !ADJMemberno = MembershipNo
            !ADJPolicyNo = Policyno
            !ADJDate = Date
            !ADJType = AdjustmentType
            If AdjustmentType = "C" Then
                !ADJCaseNo = CaseNo
            End If
            !ADJSequence = 1
        End With
        ADJ.Update
        GetAdjustmentNo = 1
    Else
        NextSequence = ADJ!ADJSequence + 1
        ADJ!ADJSequence = NextSequence
        GetAdjustmentNo = NextSequence
        ADJ.Update
    End If
            
End Function

Public Function GetNextClaimSequence(CaseNo As String, MemberNo As String, CoveredID As String) As Integer
On Error Resume Next
    Dim CLV As New ADODB.Recordset
    Dim strClv As String
    Dim NextSequence As Integer
    
    strClv = "Select * from CLV Where CASNumber = '" & CaseNo & "' And MBMNumber = '" & MemberNo & "'"
    If CoveredID <> "" Then
        strClv = " And MBMCoveredID ='" & CoveredID & "'"
    End If
    CLV.Open strClv, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If CLV.recordCount = 0 Then
        CLV.AddNew
            With CLV
                !CasNumber = CaseNo
                !MBMNumber = MemberNo
                !MBMCoveredID = CoveredID
                !CLVSequence = 1
            End With
        CLV.Update
        GetNextClaimSequence = 1
    Else
        NextSequence = CLV!CLVSequence + 1
        CLV!CLVSequence = NextSequence
        CLV.Update
        GetNextClaimSequence = NextSequence
    End If
    
    CLV.Close
    
End Function


Public Sub EnableListBar(status As Boolean)
    'If status = False Then
    '    mdiMain.SSListBar1.Enabled = False
    'ElseIf status = True Then
    '    mdiMain.SSListBar1.Enabled = True
    'End If
End Sub

Public Function GetLastCoveredID(MembershipNo As String) As Integer
On Error Resume Next
    Dim MBMCovered As New ADODB.Recordset
    Dim strsql As String
    Dim tmpCover As Integer
    
    strsql = "SELECT MBMCNumber, MBMCMemberType, MAX(MBMCCoverID) AS CoveredID "
    strsql = strsql & "From UploadDB.dbo.MBMCoveredPersons GROUP BY MBMCNumber, MBMCMemberType "
    strsql = strsql & "HAVING (MBMCNumber = '" & MembershipNo & "')"
'    StrSql = StrSql & "(MBMCMemberType = 'C')"
    MBMCovered.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMCovered.EOF = False Then
        tmpCover = IIf(IsNumeric(MBMCovered!CoveredID), MBMCovered!CoveredID, 0)
        GetLastCoveredID = tmpCover + 1
    Else
        GetLastCoveredID = tmpCover + 1
    End If
    
    MBMCovered.Close
    Set MBMCovered = Nothing
End Function

Public Function MCORefund(ExpiryDate As Date, CancelDate As Date, Effdate As Date, MCOFees As Double, MBMCancelStatus As String) As Double
Dim CancelMth, EffMth As Integer

    'MCORefund = (CInt(ExpiryDate - CancelDate) * MCOFees) / 365
    'MCOPremiumReduction = (CInt(ExpiryDate - CancelDate) / (ExpiryDate - EffDate)) * Premium
    'If Trim(MBMCancelStatus) = "Y" Then
    '    CancelMth = Int(CInt(ExpiryDate - CancelDate) / 30)
    '    EffMth = Int((ExpiryDate - Effdate) / 30)
    '    MCORefund = (CancelMth / EffMth) * MCOFees
    'Else
        MCORefund = (CDbl(ExpiryDate - CancelDate + 1) / ((ExpiryDate - Effdate) + 1)) * MCOFees
    'End If
    If MCORefund < 0 Then
        MCORefund = 0
    End If
End Function


Public Function CheckAccess(AccessNo As Integer) As Boolean
On Error Resume Next
    If Mid(Access, AccessNo, 1) = 1 Then
        CheckAccess = True
    Else
        CheckAccess = False
    End If
End Function

Public Function MCOPremiumReduction(ExpiryDate As Date, CancelDate As Date, Effdate As Date, Premium As Double) As Double
'On Error Resume Next
    If DateDiff("d", ExpiryDate, CancelDate) < 0 Then
       ' MCOPremiumReduction = (CInt(ExpiryDate - CancelDate) / (ExpiryDate - Effdate)) * Premium
        MCOPremiumReduction = (CDbl(ExpiryDate - CancelDate) / ((ExpiryDate - Effdate) + 1)) * Premium
        'MCOPremiumReduction = (DateDiff("d", ExpiryDate, CancelDate) / 365) * Premium
    Else
        MCOPremiumReduction = 0
    End If
End Function

Public Sub EnterToTab(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        If KeyAscii = 13 Then   'Turn off the beep
            KeyAscii = 0
        End If
        SendKeys "{tab}"
    End If
End Sub

Public Function EAFeesReduction(ExpiryDate As Date, CancelDate As Date, Effdate As Date, ProviderFees As Double) As Double
'On Error Resume Next
'    If DateDiff("d", ExpiryDate, CancelDate) < 0 Then
'        EAFeesReduction = (DateDiff("d", ExpiryDate, CancelDate) * ProviderFees) / 365
'    Else
'        EAFeesReduction = 0
'    End If
    EAFeesReduction = (CInt(ExpiryDate - CancelDate) / (ExpiryDate - Effdate) * ProviderFees)
    If EAFeesReduction < 0 Then
        EAFeesReduction = 0
    End If
End Function

Public Function InsertMCOFees(ExpiryDate As Date, EffectiveDate As Date, MCOFees As Double) As Double
Dim OriDays As String
Dim ComDays As String
    
    'OriDays = CDbl(ExpiryDate) - CDbl(EffectiveDate)
    'ComDays = CDbl(ExpiryDate) - CDbl(Date)
    'If CDbl(ComDays) - CDbl(OriDays) >= 10 Then
    InsertMCOFees = ((CDbl(ExpiryDate - EffectiveDate + 1)) * MCOFees) / 365
    'End If
End Function

Public Function InsertPremium(ExpiryDate As Date, EffectiveDate As Date, Premium As Double) As Double
    InsertPremium = (CInt(ExpiryDate - EffectiveDate + 1) / 365) * Premium
End Function

Public Function InsertProvider(ExpiryDate As Date, EffectiveDate As Date, AccessFees As Double) As Double
    InsertProvider = ((CInt(ExpiryDate - EffectiveDate) + 1) / 365) * AccessFees
End Function

Public Function GenerateKIVCaseNo(PayorCode)
    Dim KIVSEQ As New ADODB.Recordset
    Dim strsql As String
    Dim KIVNumber As String
    Dim KIVCAQNumber As String
    'Dim KIVCharacter As String
    
    'strsql = "Select * from CAQ Where PAYCode = '" & PayorCode & "' And PLNCode = '" & PlanCode & "'"
    strsql = "Select * from KIV_CAS_SEQ Where PAYCode = '" & PayorCode & "'"
    KIVSEQ.Open strsql, medixmasterconn, adOpenStatic, adLockOptimistic
    
    If KIVSEQ.recordCount = 0 Then
        With KIVSEQ
            .AddNew
            '!CAQCode = Trim(PayorCode) & Trim(PlanCode) & BusinessYear
            !PAYCode = PayorCode
            '!PLNCode = PlanCode
            !KIVCAQNumber = "00001"
            .Update
            KIVCAQNumber = "00001"
        End With
    Else
        KIVSEQ!KIVCAQNumber = Format(KIVSEQ!KIVCAQNumber + 1, "00000")
        KIVCAQNumber = KIVSEQ!KIVCAQNumber
        KIVSEQ.Update
    End If
    'GenerateCaseNo = PayorCode & PlanCode & Mid(BusinessYear, 3, 2) & CAQNumber
    GenerateKIVCaseNo = "T" & PayorCode & KIVCAQNumber
    KIVSEQ.Close
    Set KIVSEQ = Nothing
End Function


Public Function GenerateINVMSeqNo(BatchDate)
    Dim INVQ As New ADODB.Recordset
    Dim strsql As String
    Dim INVSEQNO As String
    
    strsql = "Select * from INVAcMSeq Where INVAcDate = '" & Format(BatchDate, "mm/dd/yyyy") & "'"
    INVQ.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic

    If INVQ.recordCount = 0 Then
        With INVQ
            .AddNew
            '!PAYCode = PAYCode
            !INVAcDate = BatchDate
            !INVAcBatch = "1"
            .Update
            INVSEQNO = Format(BatchDate, "ddmmyyyy") & "-" & "1"
        End With
    Else
        With INVQ
            !INVAcBatch = !INVAcBatch + 1
            INVSEQNO = Format(BatchDate, "ddmmyyyy") & "-" & !INVAcBatch
        End With
         INVQ.Update
    End If

    GenerateINVMSeqNo = "M" & INVSEQNO
    INVQ.Close
    Set INVQ = Nothing
End Function

Public Function GenerateINVHSeqNo(BatchDate)
Dim INVQ As New ADODB.Recordset
Dim strsql As String
Dim INVSEQNO As String
    
    strsql = " Select * from INVACHSEQ Where INVAcDate = '" & Format(BatchDate, "mm/dd/yyyy") & "'"
    INVQ.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic
   '
    If INVQ.recordCount = 0 Then
        With INVQ
            .AddNew
            '!PAYCode = PAYCode
            !INVAcDate = BatchDate
            !INVAcBatch = "1"
            .Update
            'INVSEQNO = PAYCode & Format(BatchDate, "ddmmyyyy") & "-" & "1"
            INVSEQNO = Format(BatchDate, "ddmmyyyy") & "-" & "1"
        End With
    Else
        With INVQ
            !INVAcBatch = !INVAcBatch + 1
            'INVSEQNO = PAYCode & Format(BatchDate, "ddmmyyyy") & "-" & !INVAcBatch
            INVSEQNO = Format(BatchDate, "ddmmyyyy") & "-" & !INVAcBatch
        End With
        INVQ.Update
        
    End If

    GenerateINVHSeqNo = "H" & INVSEQNO
    INVQ.Close
    Set INVQ = Nothing
End Function

Public Function GenerateINVHMSeqNo(BatchDate, DCNote)
    Dim INVQ As New ADODB.Recordset
    Dim strsql As String
    Dim INVSEQNO As String
    
    strsql = "Select * from INVAcHMSeq Where INVAcDate = '" & Format(BatchDate, "mm/dd/yyyy") & "' And INVdcnote = '" & DCNote & "'"
    INVQ.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic

    If INVQ.recordCount = 0 Then
        With INVQ
            .AddNew
            !INVDCnote = DCNote
            '!PAYCode = PAYCode
            !INVAcDate = BatchDate
            !INVAcBatch = "1"
            .Update
            INVSEQNO = DCNote & Format(BatchDate, "ddmmyyyy") & "-" & "1"
        End With
    Else
        With INVQ
            !INVAcBatch = !INVAcBatch + 1
            INVSEQNO = DCNote & Format(BatchDate, "ddmmyyyy") & "-" & !INVAcBatch
        End With
         INVQ.Update
    End If

    GenerateINVHMSeqNo = "H" & INVSEQNO
    INVQ.Close
    Set INVQ = Nothing
End Function

Public Function GetLGCode(CaseNumber As String)
    Dim LG As New ADODB.Recordset
    Dim strsql As String
    Dim LGSeq As String
    
    strsql = "Select * from LGSeq Where CASNumber = '" & CaseNumber & "'"
    LG.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    
    If LG.recordCount = 0 Then
        With LG
            .AddNew
                !CasNumber = CaseNumber
                !LGSequence = "1"
                LGSeq = !LGSequence
            .Update
        End With
    Else
        With LG
                !CasNumber = CaseNumber
                LGSeq = Format(!LGSequence + 1)
                !LGSequence = LGSeq
        .Update
        End With
    End If
    'GenerateDebitNo = "D" & payor & SEQNO
    GetLGCode = "L" & CaseNumber & "-" & LGSeq
    LG.Close
    Set LG = Nothing
End Function

Public Function GetFollowUpCode(CaseNumber As String)
Dim FOL As New ADODB.Recordset
Dim strsql As String
Dim FOLSeq As String
    
    strsql = "Select * from FollowUpSeq Where CASNumber = '" & CaseNumber & "'"
    FOL.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    
    If FOL.recordCount = 0 Then
        With FOL
            .AddNew
                !CasNumber = CaseNumber
                !FOLSeqNo = "1"
                FOLSeq = !FOLSeqNo
            .Update
        End With
    Else
        With FOL
                !CasNumber = CaseNumber
                FOLSeq = Format(!FOLSeqNo + 1)
                !FOLSeqNo = FOLSeq
        .Update
        End With
    End If
    GetFollowUpCode = "F" & CaseNumber & "-" & FOLSeq
    FOL.Close
    Set FOL = Nothing
End Function

Public Function GenerateBNFVersion(BNFCode As String, HLTCode As String, INSCode As String, PLNCode As String, Effdate As Date)
    Dim bnf As New ADODB.Recordset
    Dim strsql As String
    Dim BNFSEQNO As String
    
    strsql = "Select BNFCode from BNF where BNFCode = '" & BNFCode & "' And HLTCode = '" & HLTCode & "' And INSCode = '" & INSCode & "' And PLNCode = '" & PLNCode & "' " ' And BNFEffDate = '" & EffDate & "'"
    bnf.Open strsql, medixmasterconn, adOpenStatic, adLockOptimistic
    'BNF.Open "BNFSEQ", medixmasterconn, adOpenStatic, adLockOptimistic
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType) & Trim(BusinessYear))

    'If BNF!BNFCode = Trim(BNFCode) And BNF!HLTCode = Trim(HLTCode) And BNF!INSCode = Trim(INSCode) And BNF!PLNCode = Trim(PLNCode) And BNF!BNFEffDate <> EffDate Then
    'If BNF.recordCount = 0 Then
    '    With BNF
    '        .AddNew
    '        !SEQNO = "1"
    '        .Update
    '        BNFSEQNO = "1"
    '    End With
    'Else
    If Not bnf.EOF Then
   '     BNF!SEQNO = Format(BNF!SEQNO + 1)
        BNFSEQNO = bnf.recordCount + 1
     '   BNF!BNFCode = BNFCode
      '  BNF!HLTCode = HLTCode
       ' BNF!INSCode = INSCode
        'BNF!PLNCode = PLNCode
    '    BNF!BNFEffDate = Format(EffDate, "mm/dd/yyyy")
'        BNF.Update
    
    Else
'    BNF.AddNew
        bnf!SEQNo = "1"
     '   BNFSEQNO = Trim(BNF!SEQNO)
  '      BNF!BNFCode = BNFCode
   '     BNF!HLTCode = Trim(HLTCode)
    '    BNF!INSCode = INSCode
      '  BNF!PLNCode = PLNCode
     '   BNF!BNFEffDate = Format(EffDate, "mm/dd/yyyy")
       ' BNF.Update
    End If

    GenerateBNFVersion = Trim(BNFSEQNO)
    bnf.Close
    Set bnf = Nothing
End Function

Public Function GenerateCLMBordxNo(PAYCode)
    Dim CLMBordxSEQ As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim strsql As String
    Dim strsql2 As String
    Dim CMLBordxNo As String
    Dim CAQCharacter As String
    Dim HLTCode As String
    Dim CharacterArray
    Dim firstChar, currentChar As String

    strsql = "Select * from Claim_BordxSEQ2013 Where PAYCode = '" & Trim(PAYCode) & "'"
    strsql2 = "Select * from PAY Where PAYCode = '" & Trim(PAYCode) & "'"
    CLMBordxSEQ.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic
    PAY.Open strsql2, medixmasterconn, adOpenStatic, adLockOptimistic
    
    HLTCode = PAY!HLTCode
    If CLMBordxSEQ.recordCount = 0 Then
        With CLMBordxSEQ
            .AddNew
            !HLTCode = HLTCode
            !PAYCode = PAYCode
            !CLMBordxNo = "0000001"
            .Update
            CMLBordxNo = "0000001"
        End With
    Else
        CLMBordxSEQ!CLMBordxNo = Format(CLMBordxSEQ!CLMBordxNo + 1, "0000000")
        CMLBordxNo = CLMBordxSEQ!CLMBordxNo
        CLMBordxSEQ.Update
    End If
    
    GenerateCLMBordxNo = HLTCode & PAYCode & CMLBordxNo
    CLMBordxSEQ.Close
    Set CLMBordxSEQ = Nothing
    
    'SEQNumber = "0001"
    'CharacterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

    
    'strsql = "Select * from CLAIM_BordxSEQ Where PAYCode = '" & Trim(PAYCode) & "'"
    'strsql2 = "Select * from PAY Where PAYCode = '" & Trim(PAYCode) & "'"
    
    'CLMBordxSEQ.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic
    'PAY.Open strsql2, medixmasterconn, adOpenStatic, adLockOptimistic
    
    'HLTCode = PAY!HLTCode

    
    'currentNo = 1
    'currentChar = "A"
    
    'If CLMBordxSEQ.recordCount = 0 Then
    '    With CLMBordxSEQ
    '        .AddNew
    '        !HLTCode = HLTCode
    '        !PAYCode = PAYCode
    '        !CLMBordxNo = currentChar & "0001"
    '        .Update
    '        CMLBordxNo = currentChar & "0001"
    '
    '    End With
    'Else
    '    currentNo = Mid(CLMBordxSEQ!CLMBordxNo, 2, 4)
    '    firstChar = Mid(CLMBordxSEQ!CLMBordxNo, 1, 1)
    '    Do Until i = 24
    '        If firstChar = CharacterArray(i) Then
    '            If currentNo <> "99999" Then
    '                currentChar = CharacterArray(i)
    '                currentNo = currentNo + 1
    '            Else
    '                currentChar = CharacterArray(i + 1)
    '                currentNo = "00001"
    '            End If
    '            Exit Do
    '        End If
    '        i = i + 1
    '    Loop
    '    SEQNumber = currentChar & Format(currentNo, "00000")
    '    SEQ!SEQNumber = SEQNumber
    '    SEQ.Update
    'End If'

    'GenerateMembershipNo = SEQNumber
    'SEQ.Close
    'Set SEQ = Nothing
    
    
    
    
    
    
End Function

''' Worksheet  Debit & Credit Notes

Function GenerateDebitNo(Payor As String) As String
    Dim DRQ As New ADODB.Recordset
    Dim strsql As String
    Dim SEQNo As String
    
    strsql = "select DRIndex , DRSEQ , Payor from DRSeq where Payor ='" & Payor & "'"
    DRQ.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic
    
    If DRQ.recordCount = 0 Then
        With DRQ
            .AddNew
            !DRSEQ = "00001"
            !Payor = Payor
            .Update
            SEQNo = "00001"
        End With
    Else
        With DRQ
            !DRSEQ = Format(!DRSEQ + 1, "00000")
            SEQNo = !DRSEQ
            .Update
        End With
    End If

    GenerateDebitNo = "D" & Payor & Format(SEQNo, "00000")
    DRQ.Close
    Set DRQ = Nothing
End Function

Function GenerateCreditNo(Payor As String) As String
    Dim CRQ As New ADODB.Recordset
    Dim strsql As String
    Dim SEQNo As String
    
    strsql = "select CRIndex , CRSEQ , Payor from CRSeq where Payor ='" & Payor & "'"
    CRQ.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic
    
    If CRQ.recordCount = 0 Then
        With CRQ
            .AddNew
            !CRSEQ = "00001"
            !Payor = Payor
            .Update
            SEQNo = "00001"
        End With
    Else
        With CRQ
            !CRSEQ = Format(!CRSEQ + 1, "00000")
            SEQNo = !CRSEQ
            .Update
        End With
    End If

    GenerateCreditNo = "C" & Payor & Format(SEQNo, "00000")
    CRQ.Close
    Set CRQ = Nothing
End Function

Function GenerateHDebitNo(Payor As String) As String
    Dim DRQ As New ADODB.Recordset
    Dim strsql As String
    Dim SEQNo As String
    
    strsql = "select DRIndex , DRSEQ , Payor from DRHSeq where Payor ='" & Payor & "'"
    DRQ.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic
    
    If DRQ.recordCount = 0 Then
        With DRQ
            .AddNew
            !DRSEQ = "00001"
            !Payor = Payor
            .Update
            SEQNo = "00001"
        End With
    Else
        With DRQ
            !DRSEQ = Format(!DRSEQ + 1, "00000")
            SEQNo = !DRSEQ
            .Update
        End With
    End If

    GenerateHDebitNo = "H" & "D" & Payor & Format(SEQNo, "00000")
    DRQ.Close
    Set DRQ = Nothing
End Function

Public Function GetExcessCode(CaseNumber As String)
    Dim EXC As New ADODB.Recordset
    Dim strsql As String
    Dim EXCSeq As String
    
    strsql = "Select * from EXC_SEQ Where CASNumber = '" & Trim(CaseNumber) & "'"
    EXC.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    
    If EXC.recordCount = 0 Then
        With EXC
            .AddNew
                !CasNumber = CaseNumber
                !EXCSequence = "1"
                EXCSeq = !EXCSequence
            .Update
        End With
    Else
        With EXC
                !CasNumber = CaseNumber
                EXCSeq = Format(!EXCSequence + 1)
                !EXCSequence = EXCSeq
        .Update
        End With
    End If
    'GenerateDebitNo = "D" & payor & SEQNO
    GetExcessCode = "E" & CaseNumber & "-" & EXCSeq
    EXC.Close
    Set EXC = Nothing
End Function


Public Function AccessFeeRefund(ExpiryDate As Date, CancelDate As Date, Effdate As Date, MBMAccessFees As Double) As Double
    
    'MCORefund = (CInt(ExpiryDate - CancelDate) * MCOFees) / 365
    'MCOPremiumReduction = (CInt(ExpiryDate - CancelDate) / (ExpiryDate - EffDate)) * Premium

    AccessFeeRefund = (CInt(ExpiryDate - CancelDate) / (ExpiryDate - Effdate)) * MBMAccessFees
    If AccessFeeRefund < 0 Then
        AccessFeeRefund = 0
    End If
End Function

Public Function GetKIVLGCode(CaseNumber As String)
    Dim KIVLG As New ADODB.Recordset
    Dim strsql As String
    Dim KIVLGSeq As String
    
    strsql = "Select * from KIV_LGSeq Where CASNumber = '" & CaseNumber & "'"
    KIVLG.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
    If KIVLG.recordCount = 0 Then
        With KIVLG
            .AddNew
                !CasNumber = CaseNumber
                !LGSequence = "1"
                KIVLGSeq = !LGSequence
            .Update
        End With
    Else
        With KIVLG
                !CasNumber = CaseNumber
                KIVLGSeq = Format(!LGSequence + 1)
                !LGSequence = KIVLGSeq
        .Update
        End With
    End If
    'GenerateDebitNo = "D" & payor & SEQNO
    GetKIVLGCode = "L" & CaseNumber & "-" & KIVLGSeq
    KIVLG.Close
    Set KIVLG = Nothing
End Function

Public Function isTime(str As String) As Boolean
    Dim Correct As Boolean
    isTime = False
    Correct = False
    If IsNumeric(Left(str, 2)) Then
        If Left(str, 2) <= 24 Then
            Correct = True
        End If
    End If
    If Correct = True Then
        Correct = False
        If Right(str, 2) < 60 Then
            Correct = True
        End If
    End If
    If Correct = True Then
        isTime = True
    End If
End Function

Public Function GenerateDPNo()
    Dim PAQ As New ADODB.Recordset
    Dim strsql As String
    Dim PAYSEQNO As String
    
    'strsql = "Select * from PAQ Where PAYCode = '" & PayorCode & "'"
    PAQ.Open "PAQ", medixmasterconn, adOpenStatic, adLockOptimistic
    
    If PAQ.recordCount = 0 Then
        With PAQ
            .AddNew
            '!PayCode = PayorCode
            !PAYSEQNO = "0000001"
            .Update
            PAYSEQNO = "0000001"
        End With
    Else
        PAQ!PAYSEQNO = Format(PAQ!PAYSEQNO + 1, "0000000")
        PAYSEQNO = PAQ!PAYSEQNO
        'KIVNumber = KIVSEQ!KIVCAQNumber
        PAQ.Update
    End If

    'GenerateCaseNo = PayorCode & PlanCode & Mid(BusinessYear, 3, 2) & CAQNumber
    GenerateDPNo = "DP - " & PAYSEQNO
    PAQ.Close
    Set PAQ = Nothing
End Function

Function GenerateHCreditNo(Payor As String) As String
    Dim CRQ As New ADODB.Recordset
    Dim strsql As String
    Dim SEQNo As String
    
    strsql = "select CRIndex , CRSEQ , Payor from CRHSeq where Payor ='" & Payor & "'"
    CRQ.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic
    
    If CRQ.recordCount = 0 Then
        With CRQ
            .AddNew
            !CRSEQ = "00001"
            !Payor = Payor
            .Update
            SEQNo = "00001"
        End With
    Else
        With CRQ
            !CRSEQ = Format(!CRSEQ + 1, "00000")
            SEQNo = !CRSEQ
            .Update
        End With
    End If

    GenerateHCreditNo = "H" & "C" & Payor & Format(SEQNo, "00000")
    CRQ.Close
    Set CRQ = Nothing
End Function

Public Sub SearchForm(ScanCode As String)
On Error GoTo errOpen
    
    If CheckAccess(1) = True And Trim(ScanCode) = "LOM001" Then 'Membership Registration
        If FormLoadedByName("frmMemberEntry") = True Then
            frmMemberEntry.SetFocus
        Else
            frmMemberEntry.Show
        End If
    
    ElseIf CheckAccess(2) = True And Trim(ScanCode) = "LOM002" Then  'Membership Adjustment
        
        If FormLoadedByName("frmNewMBMAdjustment") = True Then
            frmNewMBMAdjustment.SetFocus
        Else
            frmNewMBMAdjustment.Show
        End If
        
    ElseIf CheckAccess(3) = True And Trim(ScanCode) = "LOM003" Then  'Membership Adjustment
        
        If FormLoadedByName("frmNewMBMEnquiry") = True Then
            frmNewMBMEnquiry.SetFocus
        Else
            frmNewMBMEnquiry.Show
        End If
        
    ElseIf CheckAccess(6) = True And Trim(ScanCode) = "LOM004" Then 'Membership Tracking
        If FormLoadedByName("frmNewMBMTracking") = True Then
            FrmNewMBMTracking.SetFocus
        Else
            FrmNewMBMTracking.Show
        End If
        
    ElseIf CheckAccess(9) = True And Trim(ScanCode) = "LOC001" Then 'Case Registration
       If FormLoadedByName("frmCaseAdmission") = True Then
            frmCaseAdmission.SetFocus
        Else
            frmCaseAdmission.Show
        End If
        
    ElseIf CheckAccess(10) = True And Trim(ScanCode) = "LOC002" Then 'Case Adjustment
        If FormLoadedByName("frmCaseAdjustment2016") = True Then
            frmCaseAdjustment2016.SetFocus
        Else
            frmCaseAdjustment2016.Show
        End If
        
    ElseIf CheckAccess(11) = True And Trim(ScanCode) = "LOC003" Then 'Case Inquiry
        If FormLoadedByName("frmCaseInquiry") = True Then
            frmCaseInquiry.SetFocus
        Else
            frmCaseInquiry.Show
        End If
    
    ElseIf CheckAccess(12) = True And Trim(ScanCode) = "LOC004" Then 'Case Tracking
        If FormLoadedByName("frmCaseTracking") = True Then
            FrmCaseTracking.SetFocus
        Else
            FrmCaseTracking.Show
        End If
    
    'ElseIf CheckAccess(22) = True And Trim(ScanCode) = "LOW001" Then 'Worksheet Registration
    '    If FormLoadedByName("frmNewWorksheet") = True Then
    '        frmNewWorksheet.SetFocus
    '    Else
    '        frmNewWorksheet.Show
    '    End If
        
   ' ElseIf CheckAccess(23) = True And Trim(ScanCode) = "LOW002" Then 'Worksheet Adjustment
   '     If FormLoadedByName("frmNewWSAdjust") = True Then
   '         frmNewWSAdjust.SetFocus
   '     Else
   '         frmNewWSAdjust.Show
   '     End If
        
    ElseIf CheckAccess(24) = True And Trim(ScanCode) = "LOW003" Then 'Worksheet Tracking
        If FormLoadedByName("FrmWorksheetTracking") = True Then
            FrmWorksheetTracking.SetFocus
        Else
            FrmWorksheetTracking.Show
        End If
        
    ElseIf CheckAccess(27) = True And Trim(ScanCode) = "LOB001" Then 'Worksheet Adjustment
        If FormLoadedByName("FrmWksBordx") = True Then
            FrmWksBordx.SetFocus
        Else
            FrmWksBordx.Show
        End If
        
    ElseIf CheckAccess(28) = True And Trim(ScanCode) = "LOB002" Then 'Remove Worksheet
        If FormLoadedByName("frmBordxRevise") = True Then
            FrmBordxRevise.SetFocus
        Else
            FrmBordxRevise.Show
        End If
        
    ElseIf CheckAccess(29) = True And Trim(ScanCode) = "LOB003" Then 'Print Bordx
        If FormLoadedByName("frmPrintBordx") = True Then
            FrmPrintBordx.SetFocus
        Else
            FrmPrintBordx.Show
        End If
        
    ElseIf CheckAccess(40) = True And Trim(ScanCode) = "LOB004" Then 'Bordx Tracking
        If FormLoadedByName("frmBordxTracking2") = True Then
            FrmBordxTracking2.SetFocus
        Else
            FrmBordxTracking2.Show
        End If
        
    ElseIf CheckAccess(33) = True And Trim(ScanCode) = "LOB005" Then 'Claim Tracking
        If FormLoadedByName("FrmClaimRpt") = True Then
            FrmClaimRpt.SetFocus
        Else
            FrmClaimRpt.Show
        End If
        
    ElseIf CheckAccess(90) = True And Trim(ScanCode) = "LOB006" Then 'Claim Tracking 2
        If FormLoadedByName("FrmClaimRpt2") = True Then
            FrmClaimRpt2.SetFocus
        Else
            FrmClaimRpt2.Show
        End If
    
    ElseIf CheckAccess(90) = True And Trim(ScanCode) = "LOB007" Then 'Claim Tracking 3
        If FormLoadedByName("FrmClaimRpt3") = True Then
            FrmClaimRpt3.SetFocus
        Else
            FrmClaimRpt3.Show
        End If
        
    ElseIf CheckAccess(90) = True And Trim(ScanCode) = "LOB008" Then 'Claim Tracking 4
        If FormLoadedByName("FrmClaimRpt4") = True Then
            FrmClaimRpt4.SetFocus
        Else
            FrmClaimRpt4.Show
        End If
        
    ElseIf CheckAccess(29) = True And Trim(ScanCode) = "LOB009" Then 'Print Bordx Excess
        If FormLoadedByName("frmPrintBordxEx") = True Then
            FrmPrintBordxEx.SetFocus
        Else
            FrmPrintBordxEx.Show
        End If
    
    ElseIf CheckAccess(43) = True And Trim(ScanCode) = "LOP001" Then 'Payment to Hospital
        If FormLoadedByName("frmPaymentHospital") = True Then
            FrmPaymentHospital.SetFocus
        Else
            FrmPaymentHospital.Show
        End If
        
    ElseIf CheckAccess(44) = True And Trim(ScanCode) = "LOP002" Then 'Payment to Member
        If FormLoadedByName("frmPayMember") = True Then
            FrmPayMember.SetFocus
        Else
            FrmPayMember.Show
        End If
        
    ElseIf CheckAccess(46) = True And Trim(ScanCode) = "LOR001" Then 'Receipt from Member
        If FormLoadedByName("frmReceiptMember") = True Then
            FrmReceiptMember.SetFocus
        Else
            FrmReceiptMember.Show
        End If
        
    ElseIf CheckAccess(47) = True And Trim(ScanCode) = "LOR002" Then 'Receipt from Payor
        If FormLoadedByName("frmReceiptPayor") = True Then
            FrmReceiptPayor.SetFocus
        Else
            FrmReceiptPayor.Show
        End If
        
    ElseIf CheckAccess(47) = True And Trim(ScanCode) = "LOR003" Then 'Receipt from MNRB
        If FormLoadedByName("frmReceiptMNRB") = True Then
            FrmReceiptMNRB.SetFocus
        Else
            FrmReceiptMNRB.Show
        End If
        
    ElseIf CheckAccess(25) = True And Trim(ScanCode) = "LOI001" Then 'Invoice Entry
        If FormLoadedByName("frmInvoice") = True Then
            frmInvoice.SetFocus
        Else
            frmInvoice.Show
        End If
        
    ElseIf CheckAccess(26) = True And Trim(ScanCode) = "LOI002" Then 'Invoice Tracking
        If FormLoadedByName("frmInvoiceTracking2") = True Then
            FrmInvoiceTracking2.SetFocus
        Else
            FrmInvoiceTracking2.Show
        End If
        
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If

    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If


End Sub

Public Function GenerateSPVNo()
    Dim PVQ As New ADODB.Recordset
    Dim strsql As String
    Dim PVSEQNO As String
    Dim PVYear As String
    
    'strsql = "Select * from PVSeq Where PVYear = '" & Format(Date, "YYYY") & "'"
    PVQ.Open "PVSeq2013", medixmasterconnPayment, adOpenStatic, adLockOptimistic
    
    If PVQ.recordCount = 0 Then
        With PVQ
            .AddNew
            !PVSEQNO = "000000001"
            .Update
            PVSEQNO = "000000001"
        End With
    Else
        PVQ!PVSEQNO = Format(PVQ!PVSEQNO + 1, "000000000")
        PVSEQNO = PVQ!PVSEQNO
        PVQ.Update
    End If
    
    'If HTLCode = "S" Then
        GenerateSPVNo = "CA" & PVSEQNO
    'Else
    '    GeneratePVNo = "CB" & PVSEQNO
    'End If
    PVQ.Close
    Set PVQ = Nothing
End Function

Public Function GenerateNPVNo()
    Dim PVQ As New ADODB.Recordset
    Dim strsql As String
    Dim PVSEQNO As String
    Dim PVYear As String
    
    'strsql = "Select * from PVSeq Where PVYear = '" & Format(Date, "YYYY") & "'"
    PVQ.Open "PVNSeq2013", medixmasterconnPayment, adOpenStatic, adLockOptimistic
    
    If PVQ.recordCount = 0 Then
        With PVQ
            .AddNew
            !PVSEQNO = "00000001"
            .Update
            PVSEQNO = "00000001"
        End With
    Else
        PVQ!PVSEQNO = Format(PVQ!PVSEQNO + 1, "00000000")
        PVSEQNO = PVQ!PVSEQNO
        PVQ.Update
    End If

    GenerateNPVNo = "CB" & PVSEQNO
    PVQ.Close
    Set PVQ = Nothing
End Function

Public Function GeneratePaymentPVNo(tmpHltCode As String, tmpMedicaLife As Boolean, tmpPayCode As String)
Dim PVQ As New ADODB.Recordset
Dim strsql As String
Dim PVSEQNO As String
Dim tmpPrefix As String
    If Trim(tmpPayCode) = "MX" Then
        PVQ.Open "PVMedicaMXSeq", medixmasterconnPayment, adOpenStatic, adLockOptimistic
        tmpPrefix = "CG"
    ElseIf tmpHltCode = "S" Then
        PVQ.Open "PVSeq2013", medixmasterconnPayment, adOpenStatic, adLockOptimistic
        tmpPrefix = "CA"
    ElseIf tmpHltCode = "N" And tmpMedicaLife = False Then
        PVQ.Open "PVNSeq2013", medixmasterconnPayment, adOpenStatic, adLockOptimistic
        tmpPrefix = "CB"
    ElseIf tmpHltCode = "N" And tmpMedicaLife = True Then
        PVQ.Open "PVMedicaSeq2013", medixmasterconnPayment, adOpenStatic, adLockOptimistic
        tmpPrefix = "CC"

        
    End If
    
    If PVQ.recordCount = 0 Then
        With PVQ
            .AddNew
            !PVSEQNO = "00000001"
            .Update
            PVSEQNO = "00000001"
        End With
    Else
        PVQ!PVSEQNO = Format(PVQ!PVSEQNO + 1, "00000000")
        PVSEQNO = PVQ!PVSEQNO
        PVQ.Update
    End If
    
    GeneratePaymentPVNo = tmpPrefix & PVSEQNO
    PVQ.Close
    Set PVQ = Nothing
End Function


Public Function GetCLMAsciiMedBatch(Payor As String, SystemDate As Date) As String
    Dim BAT As New ADODB.Recordset
    Dim strsql As String
    Dim CurrentBatch As String
    
    strsql = "Select * from CLMASCIIMedGen Where 0=0"
    BAT.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                                                                                    
    'If BordxType = "N" Then
    'If BAT.recordCount = 0 Then
    '    BAT.AddNew
    '    With BAT
    '        !SEQPayCode = Payor
    '        !SEQBatchDate = SystemDate
    '        !SEQBatchNo = "00001"
    '        CurrentBatch = "00001"
    '    End With
    '    BAT.Update
        'CheckBatch = "1"
    'Else
        With BAT
            !SEQBatchNo = Format(!SEQBatchNo + 1, "00000")
            CurrentBatch = !SEQBatchNo
            BAT.Update
        End With
    'End If
    GetCLMAsciiMedBatch = CurrentBatch
    BAT.Close
    Set BAT = Nothing
    
End Function

Public Function GetCLMAsciiEtiqaBatch(Payor As String, SystemDate As Date) As String
    Dim BAT As New ADODB.Recordset
    Dim strsql As String
    Dim CurrentBatch As String
    
    strsql = "Select * from CLMASCIIEtiqa Where 0=0"
    BAT.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                                                                                    
    'If BordxType = "N" Then
    'If BAT.recordCount = 0 Then
    '    BAT.AddNew
    '    With BAT
    '        !SEQPayCode = Payor
    '        !SEQBatchDate = SystemDate
    '        !SEQBatchNo = "00001"
    '        CurrentBatch = "00001"
    '    End With
    '    BAT.Update
        'CheckBatch = "1"
    'Else
        With BAT
            !SEQBatchNo = Format(!SEQBatchNo + 1, "00000")
            CurrentBatch = !SEQBatchNo
            BAT.Update
        End With
    'End If
    GetCLMAsciiEtiqaBatch = CurrentBatch
    BAT.Close
    Set BAT = Nothing
    
End Function
Public Function GetCLMAsciiBatch(Payor As String, SystemDate As Date) As String
    Dim BAT As New ADODB.Recordset
    Dim strsql As String
    Dim CurrentBatch As String
    
    strsql = "Select * from CLMAsciiSEQ Where SEQPayCode ='" & Payor & "' And SEQBatchDate = '" & Format(SystemDate, "MM/DD/YYYY") & "'"
    BAT.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                                                                                    
    'If BordxType = "N" Then
    If BAT.recordCount = 0 Then
        BAT.AddNew
        With BAT
            !SEQPayCode = Payor
            !SEQBatchDate = SystemDate
            !SEQBatchNo = "00001"
            CurrentBatch = "00001"
        End With
        BAT.Update
        'CheckBatch = "1"
    Else
        With BAT
            !SEQBatchNo = Format(!SEQBatchNo + 1, "00000")
            CurrentBatch = !SEQBatchNo
            BAT.Update
        End With
    End If
    GetCLMAsciiBatch = CurrentBatch
    BAT.Close
    Set BAT = Nothing
    
End Function
Public Function GetCASAsciiBatch(Payor As String, SystemDate As Date) As String
    Dim BAT As New ADODB.Recordset
    Dim strsql As String
    Dim CurrentBatch As String
    
    strsql = "Select * from CASAsciiSEQ Where SEQPayCode ='" & Payor & "' And SEQBatchDate = '" & Format(SystemDate, "MM/DD/YYYY") & "'"
    BAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                                                    
    'If BordxType = "N" Then
    If BAT.recordCount = 0 Then
        BAT.AddNew
        With BAT
            !SEQPayCode = Payor
            !SEQBatchDate = SystemDate
            !SEQBatchNo = "00001"
            CurrentBatch = "00001"
        End With
        BAT.Update
        'CheckBatch = "1"
    Else
        With BAT
            !SEQBatchNo = Format(!SEQBatchNo + 1, "00000")
            CurrentBatch = !SEQBatchNo
            BAT.Update
        End With
    End If
    GetCASAsciiBatch = CurrentBatch
    BAT.Close
    Set BAT = Nothing
    
End Function
Public Function GenerateSRecNo(tmpRCInd As String)
Dim RECQ As New ADODB.Recordset
Dim strsql As String
Dim RECSEQNO As String
    
    strsql = "Select RCInd,RecSeqNo from RecSeq Where RCInd ='" & Trim(tmpRCInd) & "'"
    RECQ.Open strsql, medixmasterconnPayment, adOpenStatic, adLockOptimistic
    
    If RECQ.recordCount = 0 Then
        With RECQ
            .AddNew
            !RECSEQNO = "10001"
            !RCInd = tmpRCInd
            .Update
            RECSEQNO = "10001"
        End With
    Else
        RECQ!RECSEQNO = Format(RECQ!RECSEQNO + 1, "00000")
        RECSEQNO = RECQ!RCInd & RECQ!RECSEQNO
        RECQ.Update
    End If
        
    GenerateSRecNo = Trim(RECSEQNO)
    
    RECQ.Close
    Set RECQ = Nothing
End Function

Public Function GenerateNRecNo()
Dim RECQ As New ADODB.Recordset
Dim strsql As String
Dim RECSEQNO As String
    
    
    RECQ.Open "RecNSeq", medixmasterconnPayment, adOpenStatic, adLockOptimistic
    
    If RECQ.recordCount = 0 Then
        With RECQ
            .AddNew
            !RECSEQNO = "0000001"
            .Update
            RECSEQNO = "0000001"
        End With
    Else
        RECQ!RECSEQNO = Format(RECQ!RECSEQNO + 1, "0000000")
        RECSEQNO = RECQ!RECSEQNO
        RECQ.Update
    End If

    GenerateNRecNo = RECSEQNO
    RECQ.Close
    Set RECQ = Nothing
End Function

Public Function GenerateFTRecNo(PAYCode As String)
Dim RECQ As New ADODB.Recordset
Dim strsql As String
Dim RECSEQNO As String
Dim RECTransCode As String

    strsql = "Select CLMPayor,CLMTransCode,CLMTransSeq from CLMFloatTransSeq Where CLMPayor ='" & PAYCode & "'"
    RECQ.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    
    If RECQ.recordCount = 0 Then
        With RECQ
            .AddNew
            !CLMTransSeq = "10001"
            !CLMTransCode = IIf(Trim(PAYCode) = "XG", "FTG", "FTL")
            !CLMPayor = PAYCode
            .Update
            RECSEQNO = RECQ!CLMTransCode & "10001"
        End With
    Else
        RECQ!CLMTransSeq = Format(RECQ!CLMTransSeq + 1, "00000")
        RECSEQNO = RECQ!CLMTransCode & RECQ!CLMTransSeq
        RECQ.Update
    End If

    GenerateFTRecNo = RECSEQNO
    RECQ.Close
    Set RECQ = Nothing
End Function

Public Function GenerateFTPVNo(PAYCode As String)
Dim RECQ As New ADODB.Recordset
Dim strsql As String
Dim RECPVNO As String
Dim RECTransCode As String

    strsql = "Select CLMPayor,CLMPVCode,CLMPVSeq from CLMFloatPVSeq Where CLMPayor ='" & PAYCode & "'"
    RECQ.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    
    If RECQ.recordCount = 0 Then
        With RECQ
            .AddNew
            !CLMPVSeq = "10001"
            !CLMPVCode = IIf(Trim(PAYCode) = "XG", "CE", "CD")
            !CLMPayor = PAYCode
            .Update
            RECPVNO = !CLMPVCode & "10001"
        End With
    Else
        RECQ!CLMPVSeq = Format(RECQ!CLMPVSeq + 1, "00000")
        RECPVNO = RECQ!CLMPVCode & RECQ!CLMPVSeq
        RECQ.Update
    End If

    GenerateFTPVNo = Trim(RECPVNO)
    RECQ.Close
    Set RECQ = Nothing
End Function

Public Function GenerateTUNo(FTCode As String)
Dim RECQ As New ADODB.Recordset
Dim strsql As String
Dim RECSEQNO As String
Dim RECTransCode As String

    strsql = "Select CLMFTCode,CLMFTTopupCode,CLMFTTopupNumber from CLMFloatTopupSeq Where CLMFTCode ='" & FTCode & "'"
    RECQ.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    
    If RECQ.recordCount = 0 Then
        With RECQ
            .AddNew
            !CLMFTTopupNumber = "10001"
            !CLMFTTopupCode = IIf(Trim(FTCode) = "FTL", "FRL", "FRG")
            !CLMFTCode = FTCode
            .Update
            RECSEQNO = RECQ!CLMFTTopupCode & "10001"
        End With
    Else
        RECQ!CLMFTTopupNumber = Format(RECQ!CLMFTTopupNumber + 1, "00000")
        RECSEQNO = RECQ!CLMFTTopupCode & RECQ!CLMFTTopupNumber
        
        
        'RECQ!CLMTransSeq = Format(RECQ!CLMTransSeq + 1, "0000")
        'RECSEQNO = RECQ!CLMTransCode & RECQ!CLMTransSeq
        RECQ.Update
    End If

    GenerateTUNo = RECSEQNO
    RECQ.Close
    Set RECQ = Nothing
End Function

Public Function GenerateTURecNo(TUCode As String)
Dim RECQ As New ADODB.Recordset
Dim strsql As String
Dim RECSEQNO As String
Dim RECTransCode As String
strsql = ""
    strsql = "Select CLMTopupCode,CLMRecTopupCode,CLMRecTopupSeq from CLMFloatTopupRecSeq Where CLMTopupCode ='" & TUCode & "'"
    RECQ.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    
    If RECQ.recordCount = 0 Then
        With RECQ
            .AddNew
            !CLMRecTopupSeq = "10001"
            !CLMRecTopupCode = IIf(Trim(TUCode) = "TUG", "RE", "RD")
            !CLMTopupCode = TUCode
            .Update
            RECSEQNO = RECQ!CLMRecTopupCode & "10001"
        End With
    Else
        RECQ!CLMRecTopupSeq = Format(RECQ!CLMRecTopupSeq + 1, "00000")
        RECSEQNO = RECQ!CLMRecTopupCode & RECQ!CLMRecTopupSeq
        RECQ.Update
    End If

    GenerateTURecNo = RECSEQNO
    RECQ.Close
    Set RECQ = Nothing
End Function

Public Function GenerateCLNSeq(MixClinic)
    Dim CLNSeq As New ADODB.Recordset
    Dim strsql As String
    Dim CLNCode As String
    
    strsql = "Select * from CLNSeq Where CLNCode = '" & Trim(MixClinic) & "'"
    CLNSeq.Open strsql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    
    If CLNSeq.recordCount = 0 Then
        With CLNSeq
            .AddNew
            !CLNCode = ChkStr(Trim(MixClinic))
            !CLNSeq = "001"
            .Update
            CLNCode = ChkStr(Trim(MixClinic)) & "001"
        End With
    Else
        CLNSeq!CLNSeq = Format(CLNSeq!CLNSeq + 1, "000")
        CLNSeq!CLNCode = ChkStr(Trim(MixClinic))
        CLNCode = CLNSeq!CLNSeq
        CLNSeq.Update
    End If

    'GenerateCaseNo = PayorCode & PlanCode & Mid(BusinessYear, 3, 2) & CAQNumber
    GenerateCLNSeq = ChkStr(Trim(MixClinic)) & CLNCode
    CLNSeq.Close
    Set CLNSeq = Nothing
End Function

Public Function GenerateCoGrpSeq(GrpCompany)
Dim COMSeq As New ADODB.Recordset
Dim strsql As String
Dim GRPCoCode As String

strsql = "Select * from CoGrpSeq where CoGrpCode = '" & Trim(GrpCompany) & "'"
COMSeq.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic ', adLockBatchOptimistic 'adOpenStatic, adLockOptimistic

If COMSeq.EOF = True Then
'If COMSeq.recordCount = 0 Then
    With COMSeq
        .AddNew
        !CoGRPCode = ChkStr(Trim(GrpCompany))
        !CoGrpSeq = "001"
        .Update
        GRPCoCode = ChkStr(Trim(GrpCompany)) & "001"
        GenerateCoGrpSeq = GRPCoCode
    End With
Else
    COMSeq!CoGrpSeq = Format(COMSeq!CoGrpSeq + 1, "000")
    COMSeq!CoGRPCode = ChkStr(Trim(GrpCompany))
    GRPCoCode = COMSeq!CoGrpSeq
    COMSeq.Update
    GenerateCoGrpSeq = ChkStr(Trim(GrpCompany)) & GRPCoCode
End If

COMSeq.Close
Set COMSeq = Nothing

End Function

Public Function GenerateCasePass(DPO)
Dim CasePASSSeq As New ADODB.Recordset
Dim strsql As String
Dim BatchNo As String
    
    strsql = " Select * From CASPASSSeq " & _
             " Where DPO = '" & Format(DPO, "mm/dd/yyyy") & "'" & _
             " AND PASSBy = '" & Trim(Logon) & "'"
             
    CasePASSSeq.Open strsql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    
    If CasePASSSeq.recordCount = 0 Then
        With CasePASSSeq
            .AddNew
            !DPO = Format(DPO, "dd/mm/yyyy")
            !Passby = Trim(Logon)
            !PASSSeq = "1"
            .Update
            BatchNo = "1"
        End With
    Else
        CasePASSSeq!PASSSeq = CasePASSSeq!PASSSeq + 1
        BatchNo = CasePASSSeq!PASSSeq
        CasePASSSeq.Update
    End If

    
    GenerateCasePass = BatchNo
    CasePASSSeq.Close
    Set CasePASSSeq = Nothing
End Function

Function GenerateMCOCreditNo()
    Dim CRQ As New ADODB.Recordset
    Dim strsql As String
    Dim SEQNo As String
    
    strsql = "select CRNote from MCOCRNoteSeq where 0=0"
    CRQ.Open strsql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    
    If CRQ.recordCount = 0 Then
        With CRQ
            .AddNew
            !CRNote = "00001"
            .Update
            SEQNo = "00001"
        End With
    Else
        With CRQ
            !CRNote = Format(!CRNote + 1, "00000")
            SEQNo = !CRNote
            .Update
        End With
    End If

    GenerateMCOCreditNo = "C" & Format(SEQNo, "00000")
    CRQ.Close
    Set CRQ = Nothing
End Function

Function GenerateContraNo()
    Dim CRQ As New ADODB.Recordset
    Dim strsql As String
    Dim SEQNo As String
    
    strsql = "select CONAllcSeq from ContraAllcSeq where 0=0"
    CRQ.Open strsql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    'CRQ.Close
    'Set CRQ = Nothing
    If CRQ.recordCount = 0 Then
        With CRQ
            .AddNew
            !CONAllcSeq = "100001"
            .Update
            SEQNo = "100001"
        End With
    Else
        With CRQ
            !CONAllcSeq = Format(!CONAllcSeq + 1, "000000")
            SEQNo = !CONAllcSeq
            .Update
        End With
    End If

    GenerateContraNo = Format(SEQNo, "000000")
    CRQ.Close
    Set CRQ = Nothing
End Function

Public Function GenerateAllocNo(AllocCode As String)
    Dim SEQ As New ADODB.Recordset
    Dim strsql As String
    Dim SEQNo As String
    
    strsql = "Select * from ContraAllcSeq Where ALLCode = '" & Trim(AllocCode) & "'"
    SEQ.Open strsql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    
    If SEQ.recordCount = 0 Then
        With SEQ
            .AddNew
            !ALLCode = AllocCode
            !CONAllcSeq = "00001"
            .Update
            SEQNo = "00001"
        End With
    Else
        SEQ!CONAllcSeq = Format(SEQ!CONAllcSeq + 1, "00000")
        SEQNo = SEQ!CONAllcSeq
        SEQ.Update
    End If

    GenerateAllocNo = AllocCode & SEQNo
    SEQ.Close
    Set SEQ = Nothing
End Function

Public Function GetLifeTimeLimit(TmpLimitVar As String, InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double
Dim AL As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "Select " & Trim(TmpLimitVar) & " as LTimeLimit from AnnualLimit "
    strsql = strsql & " Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnualEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"
    AL.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If AL.recordCount Then
         GetLifeTimeLimit = CDbl(AL!LTimeLimit)
    Else
         MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
End Function

Public Function GeneratePLNSeq(PLNCode)
Dim PLANSeq As New ADODB.Recordset
Dim strsql As String
Dim plancode As String


Dim PLNSeq As String
Dim SubSeq As String
Dim PLNNum As String



strsql = "Select PLNSeq,SubSeq,PLNNum from PLNSeq2"
'strsql = "Select * From PLNSeq where PLNSeqCode = '" & Trim(PLNCode) & "'"
PLANSeq.Open strsql, medixmasterconn, adOpenStatic, adLockOptimistic

If PLANSeq.recordCount = 0 Then
    With PLANSeq
        .AddNew
        !PLNSeqCode = ChkStr(Trim(PLNCode))
        !PLNNoSeq = "001"
        .Update
        plancode = ChkStr(Trim(PLNCode)) & "001"
        GeneratePLNSeq = plancode
    End With
Else

    PLNSeq = PLANSeq!PLNSeq
    SubSeq = PLANSeq!SubSeq
    PLNNum = PLANSeq!PLNNum
    
    If PLNNum = "99" Then
        If SubSeq = "Z" Then
            PLANSeq!PLNSeq = Chr(Asc(PLANSeq!PLNSeq) + 1)
            PLANSeq!SubSeq = "A"
        Else
            PLANSeq!SubSeq = Chr(Asc(PLANSeq!SubSeq) + 1)
        End If
        PLANSeq!PLNNum = "01"
    Else
        PLANSeq!PLNNum = Format(PLANSeq!PLNNum + 1, "00")
        
    End If
        PLANSeq!PLNNum = PLANSeq!PLNNum
        PLANSeq!PLNSeq = PLANSeq!PLNSeq
        PLANSeq!SubSeq = PLANSeq!SubSeq
    
    'PLANSeq!PLNNoSeq = Format(PLANSeq!PLNNoSeq + 1, "000")
    'PLANSeq!PLNSeqCode = ChkStr(Trim(PLNCode))
    'plancode = PLANSeq!PLNNoSeq
    PLANSeq.Update
    
    GeneratePLNSeq = ChkStr(Trim(PLNSeq)) & ChkStr(Trim(SubSeq)) & PLNNum
End If

PLANSeq.Close
Set PLANSeq = Nothing

End Function

Public Function GenerateReceiptMCONo()
Dim MCOQ As New ADODB.Recordset
Dim strsql As String
Dim MCOSEQNO As String
    
    MCOQ.Open "RECOthersSeq", medixmasterconnPayment, adOpenStatic, adLockOptimistic
    
    If MCOQ.recordCount = 0 Then
        With MCOQ
            .AddNew
            !RecOthersSeqNo = "00001"
            .Update
            MCOSEQNO = "00001"
        End With
    Else
        MCOQ!RecOthersSeqNo = Format(MCOQ!RecOthersSeqNo + 1, "00000")
        MCOSEQNO = MCOQ!RecOthersSeqNo
        MCOQ.Update
    End If

    GenerateReceiptMCONo = "MB" & MCOSEQNO
    MCOQ.Close
    Set MCOQ = Nothing
End Function

Public Sub GetBenefitLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, KidneyAnnLmt As Double, CancerAnnLmt As Double, HomeNursAnnLmt As Double, tmpSuppBNFStatus As String)
Dim BNFLimitRec As New ADODB.Recordset
Dim strsql As String
  
    strsql = "Select SuppALBStatus as SuppBNFStatus, KidDialysisLmt, CancerLmt, HomeNursLmt from BnfAnnLmt Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND ALBEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by ALBEffDate desc"
    BNFLimitRec.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    If BNFLimitRec.EOF = False Then
        KidneyAnnLmt = IIf(IsNumeric(BNFLimitRec!KidDialysisLmt), BNFLimitRec!KidDialysisLmt, 0)
        CancerAnnLmt = IIf(IsNumeric(BNFLimitRec!CancerLmt), BNFLimitRec!CancerLmt, 0)
        HomeNursAnnLmt = IIf(IsNumeric(BNFLimitRec!HomeNursLmt), BNFLimitRec!HomeNursLmt, 0)
        tmpSuppBNFStatus = IIf(Len(Trim(BNFLimitRec!SuppBNFStatus)), Trim(BNFLimitRec!SuppBNFStatus), "A")
    Else
         MsgBox "Benefit Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    BNFLimitRec.Close
    Set BNFLimitRec = Nothing
    
End Sub

Public Function GenerateComplaintNo()
Dim MBMComplaint As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim strAppend As String
Dim ComplaintNo As String

        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Search_MBM_Complaint"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append prm
        MBMComplaint.CursorLocation = adUseClient
        MBMComplaint.CursorType = adOpenForwardOnly
        MBMComplaint.CacheSize = 100
        Set MBMComplaint = cmd.Execute
        
        If MBMComplaint.EOF Then
            ComplaintNo = "10001"
        Else
            Do Until MBMComplaint.EOF
                ComplaintNo = MBMComplaint!MBMComplaintNo
                ComplaintNo = ComplaintNo + 1
                MBMComplaint.MoveNext
            Loop
        End If
        ComplaintCode = ComplaintNo
        MBMComplaint.Close
        Set MBMComplaint = Nothing
End Function

Public Function GenerateComplaintSeq(strAppend)
Dim rst As New ADODB.Recordset
Dim FollowSeq As String
Dim strsql As String
    
    strsql = "Select * From MBMComplaintFollowSeq Where MBMComplaintNo = '" & Trim(strAppend) & "'"
    rst.Open strsql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    
    If rst.EOF Then
        FollowSeq = "1"
        With rst
            .AddNew
                !MBMComplaintNo = strAppend
                !FollowUpSeq = FollowSeq
            .Update
        End With
    Else
        FollowSeq = rst!FollowUpSeq + 1
        With rst
            !MBMComplaintNo = strAppend
            !FollowUpSeq = FollowSeq
            .Update
        End With
    End If
    rst.Close
    Set rst = Nothing

End Function

Private Sub Update2MBMSEQ05(PayorCode, tmpSeqNo)
Dim SEQ As New ADODB.Recordset
Dim CurrNo As Long
Dim CurrChar As String
Dim CharSetNo As String
Dim PaySeqFound As Boolean
Dim strsql, PaySeqSql As String

        tmpSeqNo = "00001"
        PaySeqFound = False
        strsql = "Select PAYCode, SEQNumber from MBMSEQ05 where PAYCode='" & Trim(PayorCode) & "' "
        SEQ.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
        Do While Not SEQ.EOF
            PaySeqFound = True
            CurrNo = Mid(SEQ!SEQNumber, 2, 5)
            CurrChar = Mid(SEQ!SEQNumber, 1, 1)
            SEQ.MoveNext
            Exit Do
        Loop
        SEQ.Close
        Set SEQ = Nothing
        If PaySeqFound = False Then ' new record
            tmpSeqNo = "A00001"
            PaySeqSql = "Insert into MBMSEQ05(PAYCode, SEQNumber)" & _
                    " Values ('" & PayorCode & "','" & tmpSeqNo & "')"
        Else
            If CurrNo <> 99999 Then
                CurrNo = CurrNo + 1
            Else
                CharSetNo = Asc(CurrChar) + 1
                CurrChar = Chr(CharSetNo)
                CurrNo = "00001"
            End If
            tmpSeqNo = CurrChar & Format(CurrNo, "00000")
            PaySeqSql = "Update MBMSEQ05 set SEQNumber='" & tmpSeqNo & "' where PAYCode='" & Trim(PayorCode) & "'"
        End If
        medixmasterconnMaint.Execute PaySeqSql

End Sub

Public Function GenSUNPostingBatchNo()
Dim SeqRec As New ADODB.Recordset
Dim strsql As String
Dim NewSeqCode As String
Dim NewRec As Boolean

    SeqRec.Open "SunPostBatch", medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    NewRec = True
    NewSeqCode = "1001"
    Do While Not SeqRec.EOF
        NewRec = False
        NewSeqCode = SeqRec!BatchNo + 1
        SeqRec.MoveNext
    Loop
    SeqRec.Close
    Set SeqRec = Nothing
    If NewRec = True Then
        strsql = "Insert Into SunPostBatch(BatchNo) values('" & NewSeqCode & "')"
    Else
        strsql = "Update SunPostBatch SET BatchNo = '" & NewSeqCode & "'"
    End If
    medixmasterconnWS.Execute strsql
    GenSUNPostingBatchNo = Trim(NewSeqCode)
End Function

Public Function GetProductType(tmpPlan)
Dim PLNRec As New ADODB.Recordset
Dim PLNSql As String

    GetProductType = ""
    If Trim(tmpPlan) <> "" Then
        PLNSql = "SELECT PLNProCat FROM PLN WHERE PLNCode='" & Trim(tmpPlan) & "'"
        PLNRec.Open PLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If PLNRec.recordCount > 0 Then
            If Trim(PLNRec!PLNProCat) <> "" Then
                GetProductType = Trim(PLNRec!PLNProCat)
            End If
        End If
        PLNRec.Close
        Set PLNRec = Nothing
    End If

End Function

Public Function GetProductCat(tmpHLTType, tmpPayCode)
Dim PLNRec As New ADODB.Recordset
Dim PLNSql As String

    GetProductCat = ""
    'If Trim(tmpPlan) <> "" Then
        PLNSql = "SELECT PLNProCat FROM PLN WHERE HLTCode ='" & Trim(tmpHLTType) & "' And PayCode ='" & Trim(tmpPayCode) & "'"
        PLNRec.Open PLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If PLNRec.recordCount > 0 Then
            If Trim(PLNRec!PLNProCat) <> "" Then
                GetProductCat = Trim(PLNRec!PLNProCat)
            End If
        End If
        PLNRec.Close
        Set PLNRec = Nothing
    'End If

End Function

Public Sub REceiptInfo(tmpCASNo As String, tmpVEr As String, tmpReceiptNo As String, tmpReceiptAmt As Double)
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpRcptNo As String
Dim tmpAmt As Double

    RecSql = " AND CASNumber='" & Trim(tmpCASNo) & "' AND ClaimVersion='" & Trim(tmpVEr) & "'"
    tmpRcptNo = ""
    tmpAmt = 0
    Set cmd1.ActiveConnection = medixmasterconnPayment
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Payment"
    Set Prm1 = cmd1.CreateParameter("PaymentType", adVarChar, adParamInput, 2000, "Receipt")
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, RecSql)
    cmd1.Parameters.Append Prm2
    ReceiptRec.CursorLocation = adUseClient
    ReceiptRec.CursorType = adOpenForwardOnly
    ReceiptRec.CacheSize = 100
    Set ReceiptRec = cmd1.Execute
    Do While Not ReceiptRec.EOF
        With ReceiptRec
            If Trim(!ReceiptNoMNRB & !ReceiptNoPAY & !ReceiptNoMBM) <> "" Then
                If Trim(!ReceiptNoMNRB) <> "" Then
                    tmpRcptNo = Trim(!ReceiptNoMNRB)
                ElseIf Trim(!ReceiptNoPAY) <> "" Then
                    tmpRcptNo = Trim(!ReceiptNoPAY)
                End If
                If Trim(tmpRcptNo) <> "" Then
                    If Trim(!ReceiptNoMBM) <> "" Then
                        tmpRcptNo = tmpRcptNo & ", " & Trim(!ReceiptNoMBM)
                    End If
                Else
                    tmpRcptNo = IIf(Trim(!ReceiptNoMBM) <> "", Trim(!ReceiptNoMBM), "")
                End If
                tmpAmt = IIf(IsNumeric(!CheckAmtMNRB), !CheckAmtMNRB, 0)
                tmpAmt = tmpAmt + IIf(IsNumeric(!CheckAmtPAY), !CheckAmtPAY, 0)
                tmpAmt = tmpAmt + IIf(IsNumeric(!ReceiptAmtMBM), !ReceiptAmtMBM, 0)
            End If
            .MoveNext
        End With
    Loop
    ReceiptRec.Close
    Set ReceiptRec = Nothing
    tmpReceiptNo = tmpRcptNo
    tmpReceiptAmt = tmpAmt
End Sub

Public Function PolTermsListing(tmpPlnCode As String, tmpHltCode As String, tmpPolTermsStatus As Boolean, _
tmpPLNMinDurHosp As String, tmpPLNPhyVstPerDay As String, _
tmpPLNSepPeriodPrin As String, tmpPLNSepPeriodSup As String, _
tmpPLNRefRqdAdmin As String, tmpPLNRefRqdPreSpec As String, _
tmpPLNLowerAgeLimit As String, tmpPLNUpperAgeLimit As String, _
tmpPLNStudentAge As String, tmpPLNChildUntilAge As String, tmpPayCode As String)

    Dim POLTERM As New ADODB.Recordset
    Dim PolList As ListItem
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
            If Trim(tmpHltCode) = "S" Then
                strAppendCase = "6"
                strAppend = " AND PLNCode = '" & tmpPlnCode & "'"
    
                Set cmd.ActiveConnection = medixmasterconn
                cmd.CommandText = "Case_Registration"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd.Parameters.Append Prm2
                POLTERM.CursorLocation = adUseClient
                POLTERM.CursorType = adOpenForwardOnly
                POLTERM.CacheSize = 100
                
                Set POLTERM = cmd.Execute
             ElseIf Trim(tmpHltCode) = "N" Then
                strAppendCase = "6"
                strAppend = " AND PLNCode ='" & tmpPlnCode & "' AND PAYCode ='" & tmpPayCode & "'"
            
                Set cmd.ActiveConnection = medixmasterconn
                cmd.CommandText = "Case_Registration"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd.Parameters.Append Prm2
                
                POLTERM.CursorLocation = adUseClient
                POLTERM.CursorType = adOpenForwardOnly
                POLTERM.CacheSize = 100
                
                Set POLTERM = cmd.Execute
            End If
            If POLTERM.EOF = False Then
                tmpPolTermsStatus = True
                    With POLTERM
                        tmpPLNRefRqdAdmin = IIf(Len(!PLNRefRqdAdmin), !PLNRefRqdAdmin, "")
                        tmpPLNRefRqdPreSpec = IIf(Len(!PLNRefRqdPreSpec), !PLNRefRqdPreSpec, "")
                        tmpPLNMinDurHosp = IIf(Len(!PLNMinDurHosp), !PLNMinDurHosp, "")
                        tmpPLNPhyVstPerDay = IIf(Len(!PLNPhyVstPerDay), !PLNPhyVstPerDay, "")
                        tmpPLNSepPeriodPrin = IIf(Len(!PLNSepPeriodPrin), !PLNSepPeriodPrin, "")
                        tmpPLNSepPeriodSup = IIf(Len(!PLNSepPeriodSup), !PLNSepPeriodSup, "")
                        tmpPLNLowerAgeLimit = IIf(Len(!PLNLowerAgeLimit), !PLNLowerAgeLimit, "")
                        tmpPLNUpperAgeLimit = IIf(Len(!PLNUpperAgeLimit), !PLNUpperAgeLimit, "")
                        tmpPLNStudentAge = IIf(Len(!PLNStudentAge), !PLNStudentAge, "")
                        tmpPLNChildUntilAge = IIf(Len(!PLNChildUntilAge), !PLNChildUntilAge, "")
                    End With
            Else
                tmpPolTermsStatus = False
            End If
            POLTERM.Close
            Set POLTERM = Nothing
End Function

Public Function PolicyTermsListing(tmpPlnCode As String, tmpfrmTermsStatus As Integer, tmpPayCode As String, tmpHltCode As String)
Dim PolList As MSFlexGrid
'dim PolList as ListItem

    Call PolTermsListing(tmpPlnCode, tmpHltCode, tmpPolTermsStatus, tmpPLNMinDurHosp, tmpPLNPhyVstPerDay, tmpPLNSepPeriodPrin, _
    tmpPLNSepPeriodSup, tmpPLNRefRqdAdmin, tmpPLNRefRqdPreSpec, _
    tmpPLNLowerAgeLimit, tmpPLNUpperAgeLimit, tmpPLNStudentAge, tmpPLNChildUntilAge, tmpPayCode)
    
        If tmpfrmTermsStatus = 1 Then
            'Set PolList = frmCaseAdmission.lstPolicyTerm.listitems.Add(, , tmpPLNRefRqdAdmin)
            Set PolList = frmCaseAdmission.Fgrid
            PolList.row = 1
            PolList.col = 1

            PolList = tmpPLNRefRqdAdmin
        ElseIf tmpfrmTermsStatus = 2 Then
            'Set PolList = frmCaseAdjustment2016.lstPolicyTerm.listitems.Add(, , tmpPLNRefRqdAdmin)
            Set PolList = frmCaseAdjustment2016.Fgrid
            PolList.row = 1
            PolList.col = 1
            
            PolList = tmpPLNRefRqdAdmin
            
        ElseIf tmpfrmTermsStatus = 3 Then
            'Set PolList = frmCaseInquiry.lstPolicyTerm.listitems.Add(, , tmpPLNRefRqdAdmin)
            Set PolList = frmCaseInquiry2016.Fgrid
            PolList.row = 1
            PolList.col = 1
            PolList = tmpPLNRefRqdAdmin
            
        ElseIf tmpfrmTermsStatus = 4 Then
            'Set PolList = frmNewMBMEnquiry.lstPolicyTerm.listitems.Add(, , tmpPLNRefRqdAdmin)
            Set PolList = frmNewMBMEnquiry.Fgrid
            PolList.row = 1
            PolList.col = 1
            PolList = tmpPLNRefRqdAdmin
        End If
            
            PolList.row = 2
            PolList = tmpPLNRefRqdPreSpec
            PolList.row = 3
            If Len(tmpPLNMinDurHosp) Then
                PolList = tmpPLNMinDurHosp
            End If
            PolList.row = 4
            If Len(tmpPLNPhyVstPerDay) Then
                PolList = tmpPLNPhyVstPerDay
            End If
            PolList.row = 5
            PolList = tmpPLNSepPeriodPrin
            PolList.row = 6
            PolList = tmpPLNSepPeriodSup
            PolList.row = 7
            PolList = tmpPLNLowerAgeLimit
            PolList.row = 8
            PolList = tmpPLNUpperAgeLimit
            PolList.row = 9
            PolList = tmpPLNStudentAge
            PolList.row = 10
            PolList = tmpPLNChildUntilAge
            
            'PolList.SubItems(1) = tmpPLNRefRqdPreSpec
            'If Len(tmpPLNMinDurHosp) Then
            '    PolList.SubItems(2) = tmpPLNMinDurHosp
            'End If
            'If Len(tmpPLNPhyVstPerDay) Then
            '    PolList.SubItems(3) = tmpPLNPhyVstPerDay
            'End If
            'PolList.SubItems(4) = tmpPLNSepPeriodPrin
            'PolList.SubItems(5) = tmpPLNSepPeriodSup
            'PolList.SubItems(6) = tmpPLNLowerAgeLimit
            'PolList.SubItems(7) = tmpPLNUpperAgeLimit
            'PolList.SubItems(8) = tmpPLNStudentAge
            'PolList.SubItems(9) = tmpPLNChildUntilAge
    
End Function

Public Function ScheduleSurg(tmpPayCode As String, tmpPlnCode As String, tmpfrmTermsStatus As Integer)
    Dim strAppend As String
    Dim Search As New ADODB.Recordset
    Dim SearchList As ListItem
    
    Screen.MousePointer = vbHourglass
     
        If Trim(tmpPayCode) = "" And Trim(tmpPlnCode) = "" Then
            Exit Function
        ElseIf Len(Trim(tmpPayCode)) Or Len(Trim(tmpPlnCode)) Then
            strAppend = " Select PAYCode, PLNCode, SurgMainCat, SurgSubCat , SurgPerc From PLNSurgOperation Where PAYCode = '" & Trim(tmpPayCode) & "'" & _
            " And PLNCode = '" & Trim(tmpPlnCode) & "'"
            Search.Open strAppend, medixmasterconn, adOpenKeyset, adLockOptimistic
        End If
        
        If Search.EOF = True Then
            Search.Close
            Set Search = Nothing
            Exit Function
        End If
        If tmpfrmTermsStatus = 1 Then
            Set SearchList = frmCaseAdmission.lstSurgProc.ListItems.Add(, , Search!PAYCode)
        ElseIf tmpfrmTermsStatus = 2 Then
            Set SearchList = frmCaseAdjustment2016.lstSurgProc.ListItems.Add(, , Search!PAYCode)
        ElseIf tmpfrmTermsStatus = 3 Then
            Set SearchList = frmCaseInquiry2016.lstSurgProc.ListItems.Add(, , Search!PAYCode)
        End If
            SearchList.SubItems(1) = Search!PLNCode
            SearchList.SubItems(2) = Search!SurgMainCat
            SearchList.SubItems(3) = Search!SurgSubCat
            SearchList.SubItems(4) = Search!SurgPerc
            
        Search.Close
        Set Search = Nothing
        Screen.MousePointer = Default
End Function

'Public Function PanelHospListing(TmpPLNCode As String, tmpfrmTermsStatus As Integer, _
'tmpPAYCode As String, tmpHosName As String)
'Dim HospNHospNameList As ListItem

    
    'Call PanelHospital(TmpPLNCode, tmpPolTermsStatus, tmpPAYCode, tmpHosName)
    
        'If tmpfrmTermsStatus = 1 Then
        
            'Set HospNameList = frmCaseAdmission.lstPanelHosp.listitems.Add(, , tmpPAYCode)
            'HospNameList.SubItems(1) = tmpHosName
       ' ElseIf tmpfrmTermsStatus = 2 Then
       '     Set HospNameList = frmCaseAdjustment2016.lstPanelHosp.listitems.Add(, , tmpPAYCode)
            
        'ElseIf tmpfrmTermsStatus = 3 Then
         '   Set HospNameList = frmCaseInquiry.lstPanelHosp.listitems.Add(, , tmpPAYCode)
        
        'ElseIf tmpfrmTermsStatus = 4 Then
         '   Set HospNameList = frmNewMBMEnquiry.lstPanelHosp.listitems.Add(, , tmpPAYCode)
        
        
        'End If
            
            
'End Function


'Public Function PanelHospListing(tmpPLNCode As String, tmpfrmTermsStatus As Integer, _
'tmpfrmPanelHospStatus As Integer, tmpPAYCode As String, tmpHOSCode As String, tmpHosName As String, tmpHLTCode As String)

'        Dim strsql As String
'        Dim HOS As New ADODB.Recordset
'        Dim PanelHOS As New ADODB.Recordset
'        Dim PanelHOSList As ListItem
'        Dim cmd As New ADODB.Command
'        Dim Prm1 As ADODB.Parameter
'        Dim StrAppend As String
'        Dim HospNameList As ListItem
         
            
'           If tmpfrmTermsStatus = 1 Then
            
'                If tmpHLTCode = "S" Then
'                    StrAppend = " AND PLNCode ='" & tmpPLNCode & "'"
'                Else
'                    StrAppend = " AND PLNCode ='" & tmpPLNCode & "' And PAYCode ='" & tmpPAYCode & "'"
'                End If
'
'                    Set cmd.ActiveConnection = medixmasterconn
'                    cmd.CommandText = "Search_PanelHospitals"
'                   cmd.CommandType = adCmdStoredProc
'                    cmd.CommandTimeout = 300
'                    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
'                    cmd.Parameters.Append Prm1
'                    HOS.CursorLocation = adUseClient
'                    HOS.CursorType = adOpenForwardOnly
'                    HOS.CacheSize = 100
                    
'                    Set PanelHOS = cmd.Execute
'
'                    If PanelHOS.EOF = False Then
'                        tmpPolTermsStatus = True
'                        Do While Not PanelHOS.EOF
'
'                            With PanelHOS
'                            tmpHOSCode = IIf(Len(!HOSCode), !HOSCode, "")
'                            tmpHosName = IIf(Len(!HOSName), !HOSName, "")
'                                If tmpfrmPanelHospStatus = 1 Then
'                                    Set HospNameList = frmCaseAdmission.lstPanelHosp.listitems.Add(, , tmpHOSCode)
'                                    HospNameList.SubItems(1) = tmpHosName
                                
'                                ElseIf tmpfrmPanelHospStatus = 2 Then
'                                    Set HospNameList = frmCaseAdjustment2016.lstPanelHosp.listitems.Add(, , tmpHOSCode)
'                                    HospNameList.SubItems(1) = tmpHosName
                                    
'                                ElseIf tmpfrmPanelHospStatus = 3 Then
'                                    Set HospNameList = frmNewMBMEnquiry.lstPanelHosp.listitems.Add(, , tmpHOSCode)
'                                    HospNameList.SubItems(1) = tmpHosName
                                    
'                                ElseIf tmpfrmPanelHospStatus = 4 Then
'                                    Set HospNameList = frmCaseInquiry.lstPanelHosp.listitems.Add(, , tmpHOSCode)
'                                   HospNameList.SubItems(1) = tmpHosName
                                    
'                                End If
'
'                            End With
'                        PanelHOS.MoveNext
'                        Loop
'                    End If
            
'            ElseIf tmpfrmTermsStatus = 2 Then
'
'                If tmpHLTCode = "S" Then
'                    StrAppend = " AND PLNCode ='" & tmpPLNCode & "' AND HOS.HOSCode ='" & tmpHOSCode & "'"
'                Else
'                    StrAppend = " AND PLNCode ='" & tmpPLNCode & "' And PAYCode ='" & tmpPAYCode & "'AND HOS.HOSCode ='" & tmpHOSCode & "'"
'                End If

            
'                Set cmd.ActiveConnection = medixmasterconn
'                cmd.CommandText = "Search_PanelHospitals"
'                cmd.CommandType = adCmdStoredProc
'                cmd.CommandTimeout = 300
'                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
'                cmd.Parameters.Append Prm1
'                HOS.CursorLocation = adUseClient
'                HOS.CursorType = adOpenForwardOnly
'                HOS.CacheSize = 100
                
'                Set PanelHOS = cmd.Execute
            
                
'                If PanelHOS.EOF = True Then
'                    MsgBox "Non-Panel Hosptal", vbCritical
'                End If
'
'            Else
'                tmpPolTermsStatus = False
'            End If
            
'        PanelHOS.Close
'        Set PanelHOS = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing

'End Function


Public Function GetAgeDays(EffectiveDate As Date, DateOfBirth As Date)
    Dim CheckDate As Long
    Dim CheckYear As Integer
    Dim CheckDOB As Integer
    
    CheckDate = CDate(EffectiveDate) - CDate(DateOfBirth)
        
        If CheckDate <= 364 Then
            'If CheckDate <= 0 Then
            '    If CheckDate <= 0 Then
            '        GetAge = "-1D"
            '    Else
                GetAgeDays = CheckDate & "D"
            '    End If
            Else
                'If Month(DateOfBirth) > Month(EffectiveDate) Then
                '    GetAge = CheckYear - 1 & "Y"
                'Else
                GetAgeDays = Format(EffectiveDate, "yyyy") - Format(DateOfBirth, "yyyy") & "Y"
            End If
   ' Else
   '     CheckYear = CheckDate
        'If Month(DateOfBirth) > Month(EffectiveDate) Then
        '    GetAge = CheckYear - 1 & "Y"
        'Else
        '    If Day(DateOfBirth) > Day(EffectiveDate) Then
        '        GetAge = CheckYear - 1 & "Y"
        '    Else
    '            GetAgeDays = CheckYear & "Y"
        '    End If
            
        'End If
   ' End If
End Function

Public Function GenerateMembershipNo(PayorCode, InsuredType)
    Dim SEQ As New ADODB.Recordset
    Dim SEQInsert As New ADODB.Recordset
    Dim currentNo, i As Integer
    Dim strsql As String
    Dim SEQNumber, SEQCode As String
    Dim CharacterArray
    Dim firstChar, currentChar As String
    

    
    SEQNumber = "00001"
    CharacterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType) & Trim(BusinessYear))
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType))
    SEQCode = Trim(PayorCode & Trim(InsuredType))
    
    strsql = "Select * from MBMSEQ where SEQCOde='" & Trim(SEQCode) & "' "
    SEQ.Open strsql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    currentNo = 1
    currentChar = "A"
    
    If SEQ.recordCount = 0 Then
        SEQInsert.Open "Select * from MBMSEQ ", medixmasterconnMaint, adOpenStatic, adLockOptimistic
        With SEQInsert
            .AddNew
            !SEQCode = Trim(SEQCode)
            !PAYCode = PayorCode
            '!PLNCode = PlanCode
            !INSCode = InsuredType
            !SEQNumber = currentChar & SEQNumber
            SEQNumber = currentChar & SEQNumber
            .Update
        End With
        SEQInsert.Close
        Set SEQInsert = Nothing
    Else
        currentNo = Mid(SEQ!SEQNumber, 2, 5)
        firstChar = Mid(SEQ!SEQNumber, 1, 1)
        Do Until i = 24
            If firstChar = CharacterArray(i) Then
                If currentNo <> "99999" Then
                    currentChar = CharacterArray(i)
                    currentNo = currentNo + 1
                Else
                    currentChar = CharacterArray(i + 1)
                    currentNo = "00001"
                End If
                Exit Do
            End If
            i = i + 1
        Loop
        SEQNumber = currentChar & Format(currentNo, "00000")
        SEQ!SEQNumber = SEQNumber
        SEQ.Update
    End If

    GenerateMembershipNo = SEQNumber
    SEQ.Close
    Set SEQ = Nothing
End Function

