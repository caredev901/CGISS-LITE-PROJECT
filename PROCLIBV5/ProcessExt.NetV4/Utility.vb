Imports System.Configuration
Imports CareWebLibrary
Imports CareWebLibrary.DataTypeEnum
Module Utility
    ' Constant
    Public Const C_MaxRecords = 500
    Public Const CCode = "15-DEMO"

    Public Function tovalSF(ByVal Nvalue As Object) As String
        Dim str As String

        str = tostr(toval(Nvalue))
        tovalSF = str.Replace(",", ".")
    End Function
    Public Function RoundUp(ByVal Value As Double, ByVal RFactor As Double) As Double
        RoundUp = Rounding(Value, RFactor, 0.5)
    End Function
    Public Function RoundDn(ByVal Value As Double, ByVal RFactor As Double) As Double
        RoundDn = Rounding(Value, RFactor, -0.5)
    End Function
    Public Function toval(ByVal Value As Object, Optional ByVal Digit As Integer = 20) As Double
        Dim nValue As Double
        Try
            If IsNothing(Value) Or IsDBNull(Value) Then
                nValue = 0
            Else
                If TypeName(Value) = "String" Then
                    If Value = "True" Then
                        nValue = -1
                    ElseIf Value = "False" Then
                        nValue = 0
                    Else
                        If IsNumeric(Value) Then
                            nValue = CDbl(Value)
                        Else
                            nValue = 0
                        End If
                    End If
                Else
                    If IsNumeric(Value) Then
                        nValue = CDbl(Value)
                    Else
                        nValue = 0
                    End If
                End If
            End If
        Catch ex As Exception
            nValue = 0
        End Try
        If Digit <> 20 And Digit >= 0 Then
            nValue = Math.Round(nValue, Digit, MidpointRounding.AwayFromZero)
        End If
        toval = nValue
    End Function
    Public Function tostr(ByVal Value As Object, Optional ByVal lstr As Integer = 0) As String
        Dim str As String
        Try
            If IsNothing(Value) Or IsDBNull(Value) Then
                str = ""
            Else
                str = CStr(Value)
                If lstr <> 0 Then
                    If Len(str) < lstr Then
                        str = str & Space(lstr - Len(str))
                    Else
                        str = Mid(str, 1, lstr)
                    End If
                End If
                If Left(str, 5).ToLower <> "{\rtf" Then
                    str = str.Replace("'", "`")
                End If
            End If
        Catch ex As Exception
            str = ""
        End Try
        tostr = str
    End Function
    Public Function Time() As String
        Time = Microsoft.VisualBasic.Format(Date.Now, "HH:mm:ss")
    End Function
    Public Function todate(ByVal Value As Object) As Date
        Try
            If IsNothing(Value) Or IsDBNull(Value) Then
                todate = Today.Date
            Else
                If IsDate(Value) Then
                    todate = CDate(Value).Date
                Else
                    todate = Today.Date
                End If
            End If
        Catch ex As Exception
            todate = Today.Date
        End Try
    End Function
    Public Function Rounding(ByVal Value As Double, ByVal RFactor As Double, Optional ByVal RValue As Double = 0) As Double
        If RFactor = 0 Then
            Rounding = Value
        Else
            If toval(Value / RFactor, 0) * RFactor = Value Then
                Rounding = Value
            Else
                If RValue = 0 Then
                    Rounding = toval(Value / RFactor, 0) * RFactor
                Else
                    Rounding = toval(Value / RFactor + RValue, 0) * RFactor
                End If
            End If
        End If
    End Function
    Public Function RoundCCY(ByVal Value As Double, ByVal CCode As String, Optional ByVal RValue As Double = 0) As Double
        ' This Fixed Code will be replaced by definition in Table Currency
        If RValue <= 0 Then
            RValue = 1
            If Not (CCode = "IDR" Or CCode = "JPY" Or CCode = "ITL" Or CCode = "TWD") Then
                RValue = RValue / 100
                If RValue > 0.5 Then RValue = 0.5
            End If
        Else
            If Not (CCode = "IDR" Or CCode = "JPY" Or CCode = "ITL" Or CCode = "TWD") Then
                RValue = RValue / 100
                If RValue > 0.5 Then RValue = 0.5
            End If
        End If
        RoundCCY = Rounding(Value, RValue)
    End Function

    Public Function GetFieldIndex(ByVal FName As Object, ByRef findex() As String) As Integer
        On Error Resume Next
        Dim Index As Integer

        If TypeName(FName) = "String" Then
            Dim i As Integer
            Dim j As Integer
            Dim SName As String
            SName = UCase(FName)
            i = 0
            j = UBound(findex)
            Do While i < j
                If findex(i) = SName Then
                    Exit Do
                End If
                i = i + 1
            Loop
            Index = i
        Else
            Index = FName
        End If
        GetFieldIndex = Index
    End Function

    Public Function GetFieldTrailler(ByVal fvalue As String) As String
        GetFieldTrailler = Mid(fvalue, 1, 1)
    End Function

    Public Function RemFieldTrailler(ByVal fvalue As String) As String
        RemFieldTrailler = Mid(fvalue, 2, Len(fvalue))
    End Function

    Public Function SetFieldValue(ByVal fvalue As Object) As String
        Select Case TypeName(fvalue)
            Case "Date"
                If CType(fvalue, Date).Hour = 0 And CType(fvalue, Date).Minute = 0 Then
                    SetFieldValue = "D" & Microsoft.VisualBasic.Format(fvalue, "yyyy/MM/dd")
                Else
                    SetFieldValue = "D" & Microsoft.VisualBasic.Format(fvalue, "yyyy/MM/dd HH:mm")
                End If
            Case "String"
                SetFieldValue = "S" & fvalue
            Case "Boolean"
                SetFieldValue = "B" & IIf(fvalue, "1", "0")
            Case "Double", "Single"
                SetFieldValue = "N" & tovalSF(fvalue)
            Case "Integer", "Long"
                SetFieldValue = "N" & fvalue
            Case Else
                SetFieldValue = "S" & fvalue
        End Select
    End Function
    Public Function GetFieldValue(ByVal fvalue As String) As Object
        Dim fType As String
        fType = GetFieldTrailler(fvalue)
        fvalue = RemFieldTrailler(fvalue)
        Select Case fType
            Case "D"
                GetFieldValue = SQLConvToDate(fvalue)
            Case "N"
                GetFieldValue = Val(fvalue)
            Case "B"
                GetFieldValue = IIf(fvalue = "1", True, False)
            Case Else
                GetFieldValue = fvalue
        End Select
    End Function
    Public Function SQLConvToDate(ByVal fvalue As String) As Date
        If fvalue.Length > 10 Then
            SQLConvToDate = DateSerial(Mid(fvalue, 1, 4), Mid(fvalue, 6, 2), Mid(fvalue, 9, 2)) & " " & Mid(fvalue, 12, 5)
        Else
            SQLConvToDate = DateSerial(Mid(fvalue, 1, 4), Mid(fvalue, 6, 2), Mid(fvalue, 9, 2))
        End If
    End Function
    Public Function Format(ByVal Expression As Object, ByVal Style As String) As String
        If Expression.GetType().Name = "String" Then
            Format = Expression
        Else
            Style = Style.ToUpper
            Style = Replace(Style, "D", "d")
            Style = Replace(Style, "Y", "y")
            Format = Microsoft.VisualBasic.Format(Expression, Style)
        End If
    End Function
    Public Function First_Date_Of_Period(ByVal tdate As Date, ByVal Period As Integer) As Date
        Dim nMonth As Integer

        If Period <> 0 Then
            nMonth = ((Month(tdate) - 1) \ Period) * Period + 1
            First_Date_Of_Period = DateSerial(Year(tdate), nMonth, 1)
        Else
            First_Date_Of_Period = tdate
        End If
    End Function
    Public Function Last_date_of_period(ByVal tdate As Date, ByVal Period As Integer) As Date
        If Period <> 0 Then
            Last_date_of_period = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, Period, First_Date_Of_Period(tdate, Period)))
        Else
            Last_date_of_period = tdate
        End If
    End Function
    Public Function SetParamValue(ByVal Param As Object, Optional ByVal RecurringF As Boolean = False) As String
        Try
            If IsNothing(Param) Then
                SetParamValue = ""
            ElseIf TypeName(Param) = "Object" Then
                SetParamValue = "(...)"
            ElseIf (TypeName(Param) = "Date") Then
                If Not RecurringF Then
                    SetParamValue = Format(Param, "YYYY/MM/DD")
                Else
                    SetParamValue = Format(DateAdd(DateInterval.Day, 1, Param), "YYYY/MM/DD")
                End If
            ElseIf TypeName(Param) = "Object()" Or TypeName(Param) = "String()" Then
                SetParamValue = SetParamValue(getInParamValue(Param, 0))
            Else
                SetParamValue = tostr(Param)
            End If
        Catch ex As Exception
            SetParamValue = ""
        End Try
    End Function
    Public Function AdoDataType(ByVal TypeName As String) As DataTypeEnum
        Select Case TypeName
            Case "Int64"
                AdoDataType = adBigInt '20
            Case "Char[]"
                AdoDataType = adVarWChar '202
            Case "Byte[]"
                AdoDataType = adBinary '128
            Case "Boolean"
                AdoDataType = adBoolean '11
            Case "String"
                AdoDataType = adVarChar '200
            Case "DateTime"
                AdoDataType = adDBDate  '133
            Case "Decimal"
                AdoDataType = adDecimal '14
            Case "Double"
                AdoDataType = adDouble  '5
            Case "Int32"
                AdoDataType = adInteger '3
            Case "Single"
                AdoDataType = adSingle  '4
            Case "Int16"
                AdoDataType = adSmallInt    '2
            Case "Byte"
                AdoDataType = adTinyInt '16
            Case "Object"
                AdoDataType = adVariant '12
            Case "UInt64"
                AdoDataType = adUnsignedBigInt  '21
            Case "UInt32"
                AdoDataType = adUnsignedInt '19
            Case "UInt16"
                AdoDataType = adUnsignedSmallInt    '18
            Case Else
                AdoDataType = adVarChar '200
        End Select
    End Function
    Public Function Encrypt(ByVal plainText As String, ByVal passPhrase As String, Optional ByVal KeySize As Integer = -1) As String
        Dim C As New CareWebLibrary.Crypto
        Encrypt = C.Encrypt(plainText, passPhrase, KeySize)
    End Function
    Public Function Decrypt(ByVal cipherText As String, ByVal passPhrase As String, Optional ByVal KeySize As Integer = -1) As String
        Dim C As New CareWebLibrary.Crypto
        Decrypt = C.Decrypt(cipherText, passPhrase, KeySize)
    End Function

    Public Function getInParamValue(ByVal InParam As Object, ByVal index As Integer) As Object
        If index <= InParam.GetUpperBound(0) And index >= InParam.GetLowerBound(0) Then
            Return InParam(index)
        Else
            Return ""
        End If
    End Function
    Public Function GetAppSetting(ByVal AppSettingID As String) As String
        Return ConfigurationManager.AppSettings(AppSettingID)
    End Function
    Public Function GetCertificateNumber(ByVal CertNum As Long) As String
        Dim CFormat As String = GetAppSetting("CertificateNumberFormat")
        If CFormat = "" Then CFormat = "000000"
        GetCertificateNumber = Format(CertNum, CFormat)
    End Function
End Module


