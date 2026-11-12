Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Net
Imports System.Net.Mail
Imports System.ComponentModel
Imports CareWebLibrary
Imports CareWebLibrary.CursorTypeEnum
Imports CareWebLibrary.EditTypeEnum
Imports System.Math

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class WSUnderwriting
    Inherits System.Web.Services.WebService

    Dim db As ClassLib
    Dim dbName As String
    Dim dbServer As String
    Dim dbSysUser As String
    Dim dbSysPassword As String
    Dim dbUser As String
    Dim dbUserType As String
    Dim dbPassword As String
    Dim dbTimeOut As Integer
    Dim dbTrusted As Boolean
    Dim StrMsg As String
    Dim WI As System.Security.Principal.WindowsIdentity
    Dim WIC As System.Security.Principal.WindowsImpersonationContext
    Public Sub New()
        db = Nothing
        dbUser = ""
        dbPassword = ""
    End Sub
    Private Function OpenConnection(ByRef _dbUser As String, ByRef _dbPassword As String, ByRef ErrMsg As String) As Boolean
        Try
            If db Is Nothing Then
                db = New ClassLib
                dbUser = _dbUser
                dbPassword = _dbPassword
                dbName = GetAppSetting("dbName")
                dbServer = GetAppSetting("dbServer")
                dbSysUser = GetAppSetting("dbSysUser")
                dbSysPassword = GetAppSetting("dbSysPassword")
                dbTimeOut = toval(GetAppSetting("dbTimeout"))
                dbTrusted = toval(GetAppSetting("dbTrusted"))
                If dbTrusted Then
                    WI = CType(HttpContext.Current.User.Identity, System.Security.Principal.WindowsIdentity)
                    WIC = WI.Impersonate
                End If
                If Not db.OpenSession("", _dbUser, _dbPassword, dbName, dbServer, dbSysUser, dbSysPassword, dbTimeOut, False, dbTrusted) Then Throw New Exception("Opening Session Failed !")
                dbTrusted = db.Get_DbTrusted
                dbUserType = db.Var_db_User_Type
                If dbTrusted Then
                    dbUser = db.Var_db_user
                    dbPassword = db.Var_db_Password
                    _dbUser = dbUser
                    _dbPassword = dbPassword
                End If
            End If
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            CloseConnection()
            db = Nothing
            Return False
        End Try
    End Function
    Private Function CloseConnection() As Boolean
        Try
            If Not db Is Nothing Then
                db.CloseSession()
                db = Nothing
                dbUser = ""
                dbPassword = ""
            End If
        Catch ex As Exception
            db = Nothing
        End Try
        Return True
    End Function  
    <WebMethod(EnableSession:=False)> _
    Public Function SearchPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim Value1 As String = getInParamValue(InParam, 0)   ' Insured
            Dim Value2 As String = getInParamValue(InParam, 1)   ' Cover No
            Dim Value3 As String = getInParamValue(InParam, 2)   ' Cover Date, Format "MM/DD/YYYY"
            Dim Value4 As String = getInParamValue(InParam, 3)   ' Expiry Date, Format "MM/DD/YYYY"
            Dim Value5 As String = getInParamValue(InParam, 4)   ' Policy No
            Dim Value6 As String = getInParamValue(InParam, 5)   ' Certificate No
            Dim Value7 As String = getInParamValue(InParam, 6)   ' Policy Date Start, Format "MM/DD/YYYY"
            Dim Value8 As String = getInParamValue(InParam, 7)   ' AType
            Dim Value9 As String = getInParamValue(InParam, 8)   ' AStatus
            Dim Value10 As String = getInParamValue(InParam, 9)  ' RegNo / CI No.
            Dim Value11 As String = getInParamValue(InParam, 10) ' Policy Holder
            Dim Value12 As String = getInParamValue(InParam, 11) ' TOC
            Dim Value13 As String = getInParamValue(InParam, 12) ' Ref No.
            Dim Value14 As String = getInParamValue(InParam, 13) ' Matters (Pending Matters)
            Dim Value15 As String = getInParamValue(InParam, 14) ' Service No
            Dim Value16 As String = getInParamValue(InParam, 15) ' Policy Date End, Format "MM/DD/YYYY"
            Dim Value17 As String = getInParamValue(InParam, 16) ' Register Date Start, Format "MM/DD/YYYY"
            Dim Value18 As String = getInParamValue(InParam, 17) ' Register Date End, Format "MM/DD/YYYY"
            Dim AInfo As String = getInParamValue(InParam, 18)
            Dim DetailF As Boolean = toval(getInParamValue(InParam, 19))    ' With Detail
            Dim ProformaF As Boolean = toval(getInParamValue(InParam, 20))  ' Proforma Policy
            Dim CardID As String = getInParamValue(InParam, 21)
            Dim FacSessionNo As String = getInParamValue(InParam, 22)   ' Facultative Slip No
            Dim BSName As String = getInParamValue(InParam, 23)     ' Business source Name
            Dim Segment As String = getInParamValue(InParam, 24)     ' Segment
            Dim MO As String = getInParamValue(InParam, 25)     ' Marketing Officer
            Dim RenewalF As Boolean = toval(getInParamValue(InParam, 26))   ' Renewal Polis Flag
            Dim CRefNo As String = getInParamValue(InParam, 27)      ' cover RefNo / Original Document No.
            Dim Value19 As String = getInParamValue(InParam, 28)     ' Batch No
            Dim BackDatedF As Boolean = toval(getInParamValue(InParam, 29)) ' Backdated Policy
            Dim ConjunctionF As Boolean = toval(getInParamValue(InParam, 30)) ' Conjunction Policy
            Dim Location As String = getInParamValue(InParam, 31)        ' Policy Location
            Dim TBAF As Boolean = toval(getInParamValue(InParam, 32))    ' TBA Policy
            Dim PolicyType As String = tostr(getInParamValue(InParam, 33))    ' Policy Type
            Dim CollateralNo As String = getInParamValue(InParam, 34)       ' Collateral No
            Dim NotPrintLogF As Boolean = toval(getInParamValue(InParam, 35))    ' NotPrintLogF
            Dim STNCF As Boolean = toval(getInParamValue(InParam, 36))    ' Subject Not To Claim
            Dim Value20 As String = IIf(BackDatedF, "1", "")
            Dim Value21 As String = IIf(ConjunctionF, "1", "")
            Dim Value22 As String = IIf(TBAF, "1", "")
            Dim Value23 As String = IIf(NotPrintLogF, "1", "")
            Dim Value24 As String = getInParamValue(InParam, 37)     ' Application Source
            Dim IDNo As String = getInParamValue(InParam, 38)  ' KTP ID or TAX ID
            Dim Value25 As String = getInParamValue(InParam, 39)     ' Object Search
            Dim ISType As String = getInParamValue(InParam, 40)     ' Insurance Source Type

            AInfo = Replace(AInfo, "like '%", "like '")

            Dim SQL As String
            SQL = "select distinct Cast(0 as bit) as ProcessFlag, [COVERNO],[Insured],[CNO],[CStatus],[lANO],[ANO]," & _
            "[REGNO],[ADATE],[USERID],[Currency], [ASTATUS]," & _
            "[ATYPE],[SDate],[EDate],RYear,[ISType],[Renewal],[ESeqNo],[PolicyNo],[CertificateNo],[REFNO], " & _
            "PCALC.TSI, PCALC.OURSHARE, PCALC.GROSS"
            If DetailF Then
                If AInfo = "" And Value25 = "" Then
                    SQL = SQL & " from listqry_3('" & Value1 & "','" & Value2 & "','" & Value3 & "','" & Value4 & "'," & _
                    "'" & Value5 & "','" & Value6 & "','" & Value7 & "','" & Value8 & "'," & _
                    "'" & Value9 & "','" & Value10 & "','" & Value11 & "','" & Value12 & "','" & Value13 & "','" & Value14 & "','" & Value15 & "','" & Value16 & "','" & Value17 & "','" & Value18 & "','" & Value19 & "','" & Value20 & "','" & Value21 & "','" & Value22 & "','" & Value23 & "','" & Value24 & "') as q"
                Else
                    If Value25 = "" Then
                        SQL = SQL & " from listqry_3_a('" & Value1 & "','" & Value2 & "','" & Value3 & "','" & Value4 & "'," & _
                        "'" & Value5 & "','" & Value6 & "','" & Value7 & "','" & Value8 & "'," & _
                        "'" & Value9 & "','" & Value10 & "','" & Value11 & "','" & Value12 & "','" & Value13 & "','" & Value14 & "','" & Value15 & "','" & Value16 & "','" & Value17 & "','" & Value18 & "','" & Value19 & "','" & Value20 & "','" & Value21 & "','" & Value22 & "','" & Value23 & "','" & Value24 & "') as q"
                    Else
                        SQL = SQL & " from listqry_3_b('" & Value1 & "','" & Value2 & "','" & Value3 & "','" & Value4 & "'," & _
                        "'" & Value5 & "','" & Value6 & "','" & Value7 & "','" & Value8 & "'," & _
                        "'" & Value9 & "','" & Value10 & "','" & Value11 & "','" & Value12 & "','" & Value13 & "','" & Value14 & "','" & Value15 & "','" & Value16 & "','" & Value17 & "','" & Value18 & "','" & Value19 & "','" & Value20 & "','" & Value21 & "','" & Value22 & "','" & Value23 & "','" & Value24 & "','" & Value25 & "') as q"
                    End If
                End If
            Else
                If AInfo = "" And Value25 = "" Then
                    SQL = SQL & " from listqry_3b('" & Value1 & "','" & Value2 & "','" & Value3 & "','" & Value4 & "'," & _
                    "'" & Value5 & "','" & Value6 & "','" & Value7 & "','" & Value8 & "'," & _
                    "'" & Value9 & "','" & Value10 & "','" & Value11 & "','" & Value12 & "','" & Value13 & "','" & Value14 & "','" & Value15 & "','" & Value16 & "','" & Value17 & "','" & Value18 & "','" & Value19 & "','" & Value20 & "','" & Value21 & "','" & Value22 & "','" & Value23 & "','" & Value24 & "') as q"
                Else
                    If Value25 = "" Then
                        SQL = SQL & " from listqry_3b_a('" & Value1 & "','" & Value2 & "','" & Value3 & "','" & Value4 & "'," & _
                        "'" & Value5 & "','" & Value6 & "','" & Value7 & "','" & Value8 & "'," & _
                        "'" & Value9 & "','" & Value10 & "','" & Value11 & "','" & Value12 & "','" & Value13 & "','" & Value14 & "','" & Value15 & "','" & Value16 & "','" & Value17 & "','" & Value18 & "','" & Value19 & "','" & Value20 & "','" & Value21 & "','" & Value22 & "','" & Value23 & "','" & Value24 & "') as q"
                    Else
                        SQL = SQL & " from listqry_3b_b('" & Value1 & "','" & Value2 & "','" & Value3 & "','" & Value4 & "'," & _
                        "'" & Value5 & "','" & Value6 & "','" & Value7 & "','" & Value8 & "'," & _
                        "'" & Value9 & "','" & Value10 & "','" & Value11 & "','" & Value12 & "','" & Value13 & "','" & Value14 & "','" & Value15 & "','" & Value16 & "','" & Value17 & "','" & Value18 & "','" & Value19 & "','" & Value20 & "','" & Value21 & "','" & Value22 & "','" & Value23 & "','" & Value24 & "','" & Value25 & "') as q"
                    End If
                End If
            End If

            SQL = SQL & " OUTER APPLY (SELECT SUM(PCALC.TSI*PCALC.RATE*CASE WHEN ACCEPTANCE.QTY=0 THEN 1 ELSE ACCEPTANCE.QTY END) TSI, SUM(PCALC.OURSHARE*PCALC.RATE*CASE WHEN ACCEPTANCE.QTY=0 THEN 1 ELSE ACCEPTANCE.QTY END) OURSHARE, SUM(PCALC.GROSS*PCALC.RATE*CASE WHEN ACCEPTANCE.QTY=0 THEN 1 ELSE ACCEPTANCE.QTY END) GROSS FROM  PCALC INNER JOIN ACCEPTANCE ON (ACCEPTANCE.ANO=PCALC.ANO AND ACCEPTANCE.ANO=Q.ANO)) PCALC"

            If db.GetUserType <> "SUPERVISOR" Then
                Dim ALLOWQRYALLBRANCH As Boolean = db.GetPrivileges("ALLOWQRYALLBRANCH")
                Dim ALLOWQRYALLTOC As Boolean = db.GetPrivileges("ALLOWQRYALLTOC")

                If ALLOWQRYALLBRANCH And ALLOWQRYALLTOC Then
                Else
                    If ALLOWQRYALLBRANCH Then
                        SQL = SQL & " CROSS APPLY ("
                        SQL = SQL & " Select TOP 1 A.TOC from Acceptance A "
                        SQL = SQL & " Outer Apply (Select S.TOC from sysUserTOC S Where S.ID='" & dbUser & "' and S.TOC=A.TOC) S1 "
                        SQL = SQL & " Outer Apply (Select S.TOC from sysUserTOCBranch S Where S.ID='" & dbUser & "' and S.TOC=A.TOC and S.Branch=A.Branch) S2 "
                        SQL = SQL & " Where q.ANO=A.ANO and coalesce(S1.TOC, S2.TOC) = A.TOC"
                        SQL = SQL & " ) U"
                    ElseIf ALLOWQRYALLTOC Then
                        SQL = SQL & " CROSS APPLY ("
                        SQL = SQL & " Select TOP 1 A.Branch from Acceptance A "
                        SQL = SQL & " Outer Apply (Select S.Branch from sysUserBranch S Where S.ID='" & dbUser & "' and S.Branch=A.Branch and S.AllowedF=1) S1 "
                        SQL = SQL & " Outer Apply (Select S.Branch from sysUserTOCBranch S Where S.ID='" & dbUser & "' and S.TOC=A.TOC and S.Branch=A.Branch) S2 "
                        SQL = SQL & " Where q.ANO=A.ANO and coalesce(S1.Branch, S2.Branch) = A.Branch"
                        SQL = SQL & " ) U"
                    Else
                        SQL = SQL & " CROSS APPLY ("
                        SQL = SQL & " Select TOP 1 A.TOC, A.Branch from Acceptance A "
                        SQL = SQL & " Outer Apply (Select S.TOC from SysUserTOC S Where S.ID='" & dbUser & "' and S.TOC=A.TOC) S1 "
                        SQL = SQL & " Outer Apply (Select S.Branch from sysUserBranch S Where S.ID='" & dbUser & "' and S.Branch=A.Branch and S.AllowedF=1) S2"
                        SQL = SQL & " Outer Apply (Select S.TOC, S.Branch from sysUserTOCBranch S Where S.ID='" & dbUser & "' and S.TOC=A.TOC and S.Branch=A.Branch) S3 "
                        SQL = SQL & " Where q.ANO=A.ANO and coalesce(S1.TOC, S3.TOC) = A.TOC And coalesce(S2.Branch, S3.Branch) = A.Branch"
                        SQL = SQL & " ) U"
                    End If
                End If
                SQL = SQL & " CROSS APPLY (Select TOP 1 C.MO from Cover C, FN_UserMO('" & dbUser & "') S Where q.CNO=C.CNO and C.MO=S.ID) U3"
            End If

            If AInfo <> "" Then
                SQL = SQL & " where " & AInfo
            End If

            Dim SQLQuery As String = ""

            If ProformaF Then
                If SQLQuery <> "" Then
                    SQLQuery = SQLQuery & " and "
                End If
                SQLQuery = SQLQuery & " Q.ProformaF=1"
            End If

            If CardID <> "" Then
                If SQLQuery <> "" Then
                    SQLQuery = SQLQuery & " and "
                End If
                SQLQuery = SQLQuery & " Exists (Select ANO from fn_oneCard('" & CardID & "') f Where f.ANO=Q.ANO)"
            End If

            If FacSessionNo <> "" Or BSName <> "" Or Location <> "" Or STNCF Then
                If SQLQuery <> "" Then
                    SQLQuery = SQLQuery & " and "
                End If
                If FacSessionNo <> "" Then
                    SQLQuery = SQLQuery & " Exists (Select A.ANO from Acceptance A Left Join FPart F on (F.ANO=A.ANO) Left Join Profile P on (P.ID=A.Source) Where (A.ANO=Q.ANO)"
                Else
                    SQLQuery = SQLQuery & " Exists (Select A.ANO from Acceptance A Left Join Profile P on (P.ID=A.Source) Where (A.ANO=Q.ANO)"
                End If

                If FacSessionNo <> "" Then
                    SQLQuery = SQLQuery & " and F.SlipNo Like '" & FacSessionNo & "%'"
                End If

                If BSName <> "" Then
                    SQLQuery = SQLQuery & " and P.Name Like '" & BSName & "%'"
                End If

                If Location <> "" Then
                    SQLQuery = SQLQuery & " and A.Location Like '" & Location & "%'"
                End If

                If STNCF Then
                    SQLQuery = SQLQuery & " and A.AType not in ('A','C') and NCDate>dbo.fn_Get_ExtPeriodDate(A.ExtPeriod, A.ExtPeriodInDaysF, A.PPDate)"
                End If

                SQLQuery = SQLQuery & ")"
            End If

            If Segment <> "" Or MO <> "" Or CRefNo <> "" Or PolicyType <> "" Then
                If SQLQuery <> "" Then
                    SQLQuery = SQLQuery & " and "
                End If
                SQLQuery = SQLQuery & " Exists (Select TOP 1 C.CNO from Cover C Where q.CNO=C.CNO"

                If Segment <> "" Then
                    SQLQuery = SQLQuery & " and C.Segment='" & Segment & "'"
                End If

                If MO <> "" Then
                    SQLQuery = SQLQuery & " and C.MO='" & MO & "'"
                End If

                If CRefNo <> "" Then
                    SQLQuery = SQLQuery & " and C.RefNo Like '" & CRefNo & "%'"
                End If

                If PolicyType <> "" Then
                    SQLQuery = SQLQuery & " and C.PolicyType Like '" & PolicyType & "%'"
                End If

                SQLQuery = SQLQuery & ")"
            End If

            If RenewalF Then
                If SQLQuery <> "" Then
                    SQLQuery = SQLQuery & " and "
                End If
                SQLQuery = SQLQuery & " Q.Renewal<>0"
            End If

            If CollateralNo <> "" Then
                If SQLQuery <> "" Then
                    SQLQuery = SQLQuery & " and "
                End If
                SQLQuery = SQLQuery & " Exists (Select Top 1 A1.ANO from Acceptance A1, AccCollateral A2 Where A1.ANO=q.ANO and A2.ANO=A1.OANO and A2.CollateralNo='" & CollateralNo & "')"
            End If

            If IDNo <> "" Then
                If SQLQuery <> "" Then
                    SQLQuery = SQLQuery & " and "
                End If
                SQLQuery = SQLQuery & " Exists (Select TOP 1 P1.ID_NO, P.TAXID from Acceptance A Inner Join VProfile P on (P.ID=A.AID) Left Join VProfile_Ext P1 on (P1.ID=P.ID) Where A.ANO=Q.ANO and (coalesce(P1.ID_NO,'') Like '" & IDNo & "%' or P.TaxID Like '" & IDNo & "%'))"
            End If

            If ISType <> "" Then
                If SQLQuery <> "" Then
                    SQLQuery = SQLQuery & " and "
                End If
                SQLQuery = SQLQuery & " Q.ISType='" & ISType & "'"
            End If

            If SQLQuery <> "" Then
                If AInfo = "" Then
                    SQL = SQL & " Where " & SQLQuery
                Else
                    SQL = SQL & " and " & SQLQuery
                End If
            End If

            SQL = SQL & " Order By Q.ANO"

            Dim rs As RemoteSet
            Dim C_MaxRecords As Long
            C_MaxRecords = toval(GetAppSetting("CareMaxRecords"))
            rs = db.OpenRemoteSet(SQL, , , , C_MaxRecords)
            Dataset = rs.GetDataSet

            If Dataset Is Nothing Then Throw New Exception(ErrMsg)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SearchCertificate(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long = getInParamValue(InParam, 0)
            Dim LANO As Long = getInParamValue(InParam, 1)
            Dim CNO As Long = getInParamValue(InParam, 2)
            Dim QueryOpt As Integer = getInParamValue(InParam, 3)  ' 0 Marked By Query, 1 : Retrieve All, 2 : by Certificate Date, 3 : by Certificate No.
            Dim Query As String = getInParamValue(InParam, 4)
            Dim SDate As Date = getInParamValue(InParam, 5)
            Dim EDate As Date = getInParamValue(InParam, 6)
            Dim FCertNo As String = getInParamValue(InParam, 7)
            Dim LCertNo As String = getInParamValue(InParam, 8)
            Dim FBatchNo As String = getInParamValue(InParam, 9)
            Dim LBatchNo As String = getInParamValue(InParam, 10)

            Dim rs As RemoteSet
            Select Case QueryOpt
                Case 0
                    rs = db.OpenRemoteSet("Select * from (" & Query & ") Q Order By Q.CertificateNo")
                Case 1
                    rs = db.OpenQueryDef("LISTDET01", adOpenForwardonly, LANO)
                Case 2
                    rs = db.OpenQueryDef("LISTDET03", adOpenForwardonly, LANO, SDate, EDate)
                Case 3
                    rs = db.OpenQueryDef("LISTDET04", adOpenForwardonly, LANO, FCertNo, LCertNo)
                Case 4
                    rs = db.OpenQueryDef("LISTDET05", adOpenForwardonly, LANO, FBatchNo, LBatchNo)
            End Select
            Dataset = rs.GetDataSet

            If Dataset Is Nothing Then Throw New Exception(ErrMsg)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenQueryDef("QPolicy_Data", adOpenKeyset, ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetCertificate(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenQueryDef("qCertificate_Data", adOpenKeyset, ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function CheckDigit(ByVal PolicyNo As String) As String
        Dim nTemp As Integer = 0
        Dim nRes As Integer = 0

        Dim i As Integer
        For i = 1 To 20
            nTemp = nTemp + toval(Mid(PolicyNo, i, 1)) * IIf(i Mod 2 = 0, 2, 1)
        Next

        If nTemp <= 9 Then
            nRes = nTemp
        Else
            Do While Len(Trim(Str(nTemp))) > 1
                nRes = Val(Mid(Trim(Str(nTemp)), Len(Trim(Str(nTemp))), 1))
                nTemp = Val(Mid(Trim(Str(nTemp)), 1, Len(Trim(Str(nTemp))) - 1))
                nTemp = nTemp + nRes
            Loop
            nRes = nTemp
        End If
        CheckDigit = Trim(Str(nRes))
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetNewPolicyNo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter
            Dim Branch As String
            Dim CT As String
            Dim TOC As String
            Dim ISType As String
            Dim PolicyType As String
            Dim Product As String

            Branch = getInParamValue(InParam, 0)        ' Branch
            CT = getInParamValue(InParam, 1)            ' Profit/Cost Center
            TOC = getInParamValue(InParam, 2)           ' Sub Class Code
            ISType = getInParamValue(InParam, 3)        ' See ISType Table (Ex. 'I' : Direct,'L': Coins Leader, 'C': Coins Member, etc)
            PolicyType = getInParamValue(InParam, 4)    ' see PolicyType Table (Ex. '01' : Individual Policy, '02' : File Multiple, etc)
            Product = getInParamValue(InParam, 5)        ' CGISS/CHISS

            Dim PolicyNo As String = ""
            If Product.ToUpper = "CGISS" Then
                Dim BAlias As String
                Dim TAlias As String = ""
                Dim POL_AutoNum As String = ""
                Dim Prefix As String
                Dim InFix As String
                Dim rs As RemoteSet
                Dim Separator As String
                Dim ADate As Date
                ADate = Today.Date

                BAlias = Get_Branch_Alias_By_TOC(Branch, TOC, TAlias, POL_AutoNum)
                If PolicyType = "" Then
                    rs = db.OpenRemoteSet("Select P1.PolicyType from PolicyType P1, PTClass P2, TOC Where P1.SType='S' and (P1.ISType in ('D','A') and '" & ISType & "' in ('I','C','L')) and P2.PolicyType=P1.PolicyType and P2.COB=TOC.COB and (P2.TOC='<All>' or P2.TOC=TOC.TOC) and TOC.TOC ='" & TOC & "'", adOpenForwardonly)
                    PolicyType = rs.fields("PolicyType")
                End If
                rs = db.OpenRemoteSet("Select P.Prefix, P.InFix, P.Separator, P.BranchDigit, P.TOCDigit, P.CTDigit, P.CheckDigitF, I.Prefix as IPrefix, I.InFix as IInFix from PolicyType P, ISType I Where I.ISType='" & ISType & "' and P.PolicyType='" & PolicyType & "'")
                Separator = tostr(rs.fields("Separator"))
                Prefix = IIf(tostr(rs.fields("Prefix")) <> "", tostr(rs.fields("Prefix")), tostr(rs.fields("IPrefix")))
                InFix = IIf(tostr(rs.fields("InFix")) <> "", tostr(rs.fields("InFix")), tostr(rs.fields("IInFix")))
                If Not rs.Eof Then
                    If BAlias = Branch Then
                        PolicyNo = db.GetNextDocNo(POL_AutoNum, ADate, IIf(Prefix <> "", Prefix & Separator, "") & IIf(rs.fields("BranchDigit") <> 0, Mid(BAlias, 1, rs.fields("BranchDigit")) & Separator, "") & IIf(rs.fields("CTDigit") <> 0, Mid(CT, 1, rs.fields("CTDigit")) & Separator, "") & IIf(InFix <> "", InFix & Separator, "") & IIf(rs.fields("TOCDigit") <> 0, Mid(TAlias, 1, rs.fields("TOCDigit")) & Separator, ""))
                    Else
                        PolicyNo = db.GetNextDocNo(POL_AutoNum, ADate, IIf(Prefix <> "", Prefix & Separator, "") & BAlias)
                    End If
                Else
                    PolicyNo = db.GetNextDocNo(POL_AutoNum, ADate, BAlias & Mid(TAlias, 1, 2))
                End If
                If PolicyNo = "" Then
                    StrMsg = "Generate Policy No. Failed !"
                    Throw New Exception(StrMsg)
                End If
                If toval(rs.fields("CheckDigitF")) Then PolicyNo = PolicyNo & CheckDigit(PolicyNo)

            ElseIf Product.ToUpper = "CHISS" Then
                PolicyNo = db.GetNextDocNo("POLH", Today.Date, Branch & TOC)
                If PolicyNo = "" Then
                    StrMsg = "Generate Policy No. Failed !"
                    Throw New Exception(StrMsg)
                End If
            End If

            ReDim OutParam(0)
            OutParam(0) = PolicyNo

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SavePolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long
            Dim IncludeExtCoverF As Boolean
            Dim ResetF As Boolean
            Dim AdjustInterestF As Boolean
            Dim AdjustRiskF As Boolean
            Dim AdjustTCF As Boolean
            Dim Grace As Integer
            Dim KeepExistingF As Boolean
            Dim NotApplyConjunctPolicyDataF As Boolean

            ' getInParamValue(InParam, 0) = CNO
            ANO = getInParamValue(InParam, 1)
            ' getInParamValue(InParam, 1) = ANO
            ' getInParamValue(InParam, 2) = CANO
            LANO = getInParamValue(InParam, 3)
            ' getInParamValue(InParam, 3) = LANO
            ' getInParamValue(InParam, 4) = SNO
            ' getInParamValue(InParam, 5) = PolicyType
            ' getInParamValue(InParam, 6) = TOC
            ' getInParamValue(InParam, 7) = VOYAGE
            ' getInParamValue(InParam, 8) = ISType
            ' getInParamValue(InParam, 9) = ReinstatementF
            ' getInParamValue(InParam, 10) = SType
            ' getInParamValue(InParam, 11) = ProformaF
            ' getInParamValue(InParam, 12) = BasisList.ListIndex
            ' getInParamValue(InParam, 13) = RenewableF
            ' getInParamValue(InParam, 14) = WPC
            ' getInParamValue(InParam, 15) = OverWritePayorF
            ' getInParamValue(InParam, 16) = OverWriteDAF
            ' getInParamValue(InParam, 17) = OverWriteCAF
            ' getInParamValue(InParam, 18) = RegID
            ' getInParamValue(InParam, 19) = FLDID
            ' getInParamValue(InParam, 20) = ValueID
            ' getInParamValue(InParam, 21) = ValueDesc
            IncludeExtCoverF = getInParamValue(InParam, 22)
            ResetF = toval(getInParamValue(InParam, 23))
            AdjustInterestF = toval(getInParamValue(InParam, 24))
            AdjustRiskF = toval(getInParamValue(InParam, 25))
            AdjustTCF = toval(getInParamValue(InParam, 26))
            Grace = toval(getInParamValue(InParam, 27))
            KeepExistingF = toval(getInParamValue(InParam, 28))
            NotApplyConjunctPolicyDataF = toval(getInParamValue(InParam, 29))


            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim PANO As Long
            Dim pSDate As Date
            Dim pEDate As Date
            Dim OpenFlag As Boolean
            rs = db.OpenRemoteSet("Select A.AStatus, A.PANO, A.SDate, A.EDate, C.OpenFlag from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) where A.ANO=" & ANO)
            If Not rs.Eof Then
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If
            PANO = toval(rs.fields("PANO"))
            pSDate = todate(rs.fields("SDate"))
            pEDate = todate(rs.fields("EDate"))
            OpenFlag = toval(rs.fields("OpenFlag"))

            If LANO = ANO And Not OpenFlag Then
                rs = db.OpenRemoteSet("Select Top 1 Acceptance.SDate, Acceptance.CERTIFICATENO from Acceptance Where AType in ('N','E') and AStatus in ('W','T','I') and ANO<>LANO and LANO<>-1 and ((not (SDate between '" & Format(GetFieldValue(fbuffer(GetFieldIndex("SDate", findex))), "MM/DD/YYYY") & "' and '" & Format(GetFieldValue(fbuffer(GetFieldIndex("EDate", findex))), "MM/DD/YYYY") & "')) or not ((EDate between '" & Format(GetFieldValue(fbuffer(GetFieldIndex("SDate", findex))), "MM/DD/YYYY") & "' and '" & Format(GetFieldValue(fbuffer(GetFieldIndex("EDate", findex))), "MM/DD/YYYY") & "'))) and Acceptance.LANO=" & PANO & " and not exists (Select A.pANO from Acceptance A Inner join Acceptance mA on (mA.ANO=A.LANO) Where A.AType in ('N','E') and A.AStatus='W' and A.pANO=Acceptance.ANO and mA.ANO=" & ANO & ")")
                If Not rs.Eof Then
                    Throw New Exception("Policy Inception Period Date Doesn't Include Detail Inception Date (CertificateNo : " & rs.fields("CertificateNo") & ", Inception Date :" & rs.fields("SDate") & ") !")
                End If

                rs = db.OpenRemoteSet("Select Top 1 Acceptance.SDate, Acceptance.CERTIFICATENO from Acceptance Where AType in ('N','E') and AStatus in ('W','T','I') and ANO<>LANO and LANO<>-1 and ((not (SDate between '" & Format(GetFieldValue(fbuffer(GetFieldIndex("SDate", findex))), "MM/DD/YYYY") & "' and '" & Format(GetFieldValue(fbuffer(GetFieldIndex("EDate", findex))), "MM/DD/YYYY") & "') and Acceptance.SDate<>'" & Format(pSDate, "MM/DD/YYYY") & "') or (not (EDate between '" & Format(GetFieldValue(fbuffer(GetFieldIndex("SDate", findex))), "MM/DD/YYYY") & "' and '" & Format(GetFieldValue(fbuffer(GetFieldIndex("EDate", findex))), "MM/DD/YYYY") & "') and Acceptance.EDate<>'" & Format(pEDate, "MM/DD/YYYY") & "')) and Acceptance.LANO=" & ANO)
                If Not rs.Eof Then
                    Throw New Exception("Policy Inception Period Date Doesn't Include Detail Inception Date (CertificateNo : " & rs.fields("CertificateNo") & ", Inception Date :" & rs.fields("SDate") & ") !")
                End If
            End If

            ' Check Extended Period
            Dim ExtPeriod As Integer = toval(GetFieldValue(fbuffer(GetFieldIndex("ExtPeriod", findex))))
            If ExtPeriod <> 0 Then
                Dim EXTPERIODINDAYSF As Boolean = toval(GetFieldValue(fbuffer(GetFieldIndex("EXTPERIODINDAYSF", findex))))
                Dim EDate As Date = todate(GetFieldValue(fbuffer(GetFieldIndex("EDate", findex))))

                rs = db.OpenRemoteSet("Select dbo.fn_Get_ExtPeriodDate(" & ExtPeriod & "," & IIf(EXTPERIODINDAYSF = 0, 0, 1) & ",'" & Format(EDate, "MM/DD/YYYY") & "')")
                If TypeName(rs) = "Nothing" Then
                    Throw New Exception("Invalid Extended Period !")
                End If
            End If

            ' filter Input Data
            ReDim Preserve InParam(26)

            Dim TOPRO As String
            Dim TANO As Long = -1
            Dim WordingID As String = ""
            TOPRO = GetFieldValue(fbuffer(GetFieldIndex("TOPRO", findex)))
            If TOPRO <> "" AndAlso (ANO <= 0 Or ResetF) Then
                rs = db.OpenRemoteSet("Select IncludeExtCoverF, WordingID, PolicyNo from TOPRO Where TOPRO='" & TOPRO & "'")
                IncludeExtCoverF = IncludeExtCoverF Or toval(rs.fields("IncludeExtCoverF"))
                WordingID = tostr(rs.fields("WordingID"))

                If rs.fields("PolicyNo") <> "" Then ' Templace Policy No
                    rs = db.OpenRemoteSet("Select A.ANO from Acceptance A, Cover C Where C.CNO=A.CNO and C.ProformaF=1 and A.AStatus='I' and A.LANO=-1 and A.PolicyNo='" & rs.fields("PolicyNo") & "'")
                    If Not rs.Eof Then
                        TANO = rs.fields("ANO")
                    End If
                End If
                If GetFieldValue(fbuffer(GetFieldIndex("PWordingID", findex))) = "" Then
                    fbuffer(GetFieldIndex("PWordingID", findex)) = SetFieldValue(WordingID)
                End If
            End If

            InParam(22) = TANO ' Template Policy
            InParam(23) = CBool(toval(GetAppSetting("AdjustFacultativePeriodF")))
            InParam(24) = Grace
            InParam(25) = KeepExistingF
            InParam(26) = NotApplyConjunctPolicyDataF
            fbuffer(GetFieldIndex("PolicyNo", findex)) = SetFieldValue(IIf(GetFieldValue(fbuffer(GetFieldIndex("PolicyNo", findex))) = "<Auto Generated>", "", GetFieldValue(fbuffer(GetFieldIndex("PolicyNo", findex)))))
            fbuffer(GetFieldIndex("RegNo", findex)) = SetFieldValue(IIf(GetFieldValue(fbuffer(GetFieldIndex("RegNo", findex))) = "<Auto Generated>", "", GetFieldValue(fbuffer(GetFieldIndex("RegNo", findex)))))
            ANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "SAVE_POLICY_DATA", InParam, findex, fbuffer)
            If ANO = 0 Then
                Throw New Exception(Me.ErrDesc)
            End If



            If TOPRO <> "" Then
                Dim tParam(10) As Object
                tParam(0) = GetFieldValue(fbuffer(GetFieldIndex("TOPRO", findex)))
                tParam(1) = ""  ' UnUsed
                tParam(2) = ANO
                tParam(3) = GetFieldValue(fbuffer(GetFieldIndex("Currency", findex)))
                tParam(4) = GetFieldValue(fbuffer(GetFieldIndex("TSI", findex)))
                tParam(5) = GetFieldValue(fbuffer(GetFieldIndex("ATSI", findex)))
                tParam(6) = IncludeExtCoverF
                tParam(7) = IncludeExtCoverF
                tParam(8) = True
                tParam(9) = ResetF
                tParam(10) = GetFieldValue(fbuffer(GetFieldIndex("Premium", findex)))
                If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "TOP_PROCESS", tParam, AdjustInterestF, AdjustRiskF, AdjustTCF) Then Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = ANO

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GenerateSequenceNo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()
            Dim SequenceNo As String
            Dim IName As String = getInParamValue(InParam, 0)   ' Ex. COCNO, CNCNO
            Dim InFix As String = getInParamValue(InParam, 1)   ' Passing InFix, Ex : YYYY 
            Dim Suffix As String = getInParamValue(InParam, 2)   ' Passing Sufix
            Dim LeadingZero As String = getInParamValue(InParam, 3)   ' LeadingZero 000000
            ' Default
            If LeadingZero = "" Then LeadingZero = "000000000000"

            SequenceNo = GetAppSetting(IName & "PREFIX") & InFix & Format(db.GetIdentity(IName + InFix) + 1, LeadingZero) & Suffix

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = SequenceNo

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveCertificate(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim RegID As String
            Dim FLDID As String
            Dim ValueID As String
            Dim IncludeExtCoverF As Boolean
            Dim ResetF As Boolean
            Dim AdjustInterestF As Boolean
            Dim AdjustRiskF As Boolean
            Dim AdjustTCF As Boolean
            Dim KeepExistingF As Boolean
            Dim NotApplyConjunctPolicyDataF As Boolean

            ' getInParamValue(InParam, 0) = CNO
            ANO = getInParamValue(InParam, 1)
            ' getInParamValue(InParam, 1) = ANO
            ' getInParamValue(InParam, 2) = CANO
            ' getInParamValue(InParam, 3) = LANO
            ' getInParamValue(InParam, 4) = RenewableF
            ' getInParamValue(InParam, 5) = VOYAGE
            ' getInParamValue(InParam, 6) = OpenFlag
            ' getInParamValue(InParam, 7) = M_SDate
            ' getInParamValue(InParam, 8) = M_EDate
            ' getInParamValue(InParam, 9) = M_PDate
            ' getInParamValue(InParam, 10) = OverWriteCAF
            RegID = getInParamValue(InParam, 11)
            FLDID = getInParamValue(InParam, 12)
            ValueID = getInParamValue(InParam, 13)
            ' getInParamValue(InParam, 11) = RegID
            ' getInParamValue(InParam, 12) = FLDID
            ' getInParamValue(InParam, 13) = ValueID
            ' getInParamValue(InParam, 14) = ValueDesc
            ' getInParamValue(InParam, 15) = False
            ' getInParamValue(InParam, 16) = UnUsed
            IncludeExtCoverF = getInParamValue(InParam, 17)
            ResetF = getInParamValue(InParam, 18)
            AdjustInterestF = toval(getInParamValue(InParam, 19))
            AdjustRiskF = toval(getInParamValue(InParam, 20))
            AdjustTCF = toval(getInParamValue(InParam, 21))
            KeepExistingF = toval(getInParamValue(InParam, 22))
            NotApplyConjunctPolicyDataF = toval(getInParamValue(InParam, 23))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If Not rs.Eof Then
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Certificate Already Processed !")
                End If
            End If

            ' Check Extended Period
            Dim ExtPeriod As Integer = toval(GetFieldValue(fbuffer(GetFieldIndex("ExtPeriod", findex))))
            If ExtPeriod <> 0 Then
                Dim EXTPERIODINDAYSF As Boolean = toval(GetFieldValue(fbuffer(GetFieldIndex("EXTPERIODINDAYSF", findex))))
                Dim EDate As Date = todate(GetFieldValue(fbuffer(GetFieldIndex("EDate", findex))))
                rs = db.OpenRemoteSet("Select dbo.fn_Get_ExtPeriodDate(" & ExtPeriod & "," & IIf(EXTPERIODINDAYSF = 0, 0, 1) & ",'" & Format(EDate, "MM/DD/YYYY") & "')")
                If TypeName(rs) = "Nothing" Then
                    Throw New Exception("Invalid Extended Period !")
                End If
            End If

            ' filter Input Data
            ReDim Preserve InParam(21)

            Dim TOPRO As String
            Dim TANO As Long = -1
            Dim WordingID As String = ""
            TOPRO = GetFieldValue(fbuffer(GetFieldIndex("TOPRO", findex)))
            If TOPRO <> "" AndAlso (ANO <= 0 Or ResetF) Then
                rs = db.OpenRemoteSet("Select IncludeExtCoverF, WordingID, PolicyNo from TOPRO Where TOPRO='" & TOPRO & "'")
                IncludeExtCoverF = IncludeExtCoverF Or toval(rs.fields("IncludeExtCoverF"))
                WordingID = tostr(rs.fields("WordingID"))

                If rs.fields("PolicyNo") <> "" Then ' Templace Policy No
                    rs = db.OpenRemoteSet("Select A.ANO from Acceptance A, Cover C Where C.CNO=A.CNO and C.ProformaF=1 and A.AStatus='I' and A.LANO=-1 and A.PolicyNo='" & rs.fields("PolicyNo") & "'")
                    If Not rs.Eof Then
                        TANO = rs.fields("ANO")
                    End If
                End If

                If GetFieldValue(fbuffer(GetFieldIndex("PWordingID", findex))) = "" Then
                    fbuffer(GetFieldIndex("PWordingID", findex)) = SetFieldValue(WordingID)
                End If
            End If

            InParam(16) = TANO ' Template Policy
            InParam(17) = CBool(toval(GetAppSetting("AdjustFacultativePeriodF")))
            InParam(18) = KeepExistingF
            InParam(21) = NotApplyConjunctPolicyDataF
            ANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "SAVE_CERTIFICATE_DATA", InParam, findex, fbuffer)
            If ANO = 0 Then
                Throw New Exception(Me.ErrDesc)
            End If

            If TOPRO <> "" Then
                Dim tParam(10) As Object
                tParam(0) = GetFieldValue(fbuffer(GetFieldIndex("TOPRO", findex)))
                tParam(1) = ""  ' UnUsed
                tParam(2) = ANO
                tParam(3) = GetFieldValue(fbuffer(GetFieldIndex("Currency", findex)))
                tParam(4) = GetFieldValue(fbuffer(GetFieldIndex("TSI", findex)))
                tParam(5) = GetFieldValue(fbuffer(GetFieldIndex("ATSI", findex)))
                tParam(6) = IncludeExtCoverF
                tParam(7) = IncludeExtCoverF
                tParam(8) = True
                tParam(9) = ResetF
                tParam(10) = GetFieldValue(fbuffer(GetFieldIndex("Premium", findex)))
                If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "TOP_PROCESS", tParam, AdjustInterestF, AdjustRiskF, AdjustTCF) Then Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = ANO

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DropCertificate(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select LANO, CNO from Acceptance Where AStatus='W' and ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Please Check Certificate Data Status !"
                Throw New Exception(ErrMsg)
            End If
            Dim LANO As Long
            Dim CNO As Long
            Dim OpenFlag As Boolean
            LANO = toval(rs.fields("LANO"))
            CNO = toval(rs.fields("CNO"))

            rs = db.OpenRemoteSet("Select OpenFlag from Cover Where CNO=" & CNO)
            OpenFlag = rs.fields("OpenFlag")

            rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1 Left Join Acceptance pA1 on (pA1.ANO=A1.pANO), Acceptance A2 Inner Join Cover C on (C.CNO=A2.CNO) Left Join RArrHeader R on (R.ANO=A2.ANO) Where (R.Status='A' or A2.AStatus='W') and A2.CANO=A1.ANO and coalesce(pA1.AType,'') Not in ('N','E') and A1.AType in ('N','E') and A1.ANO=" & ANO)
            If Not rs.Eof Then
                ErrMsg = "Must Drop the Co. Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                Throw New Exception(ErrMsg)
            End If

            rs = db.OpenRemoteSet("Select Top 1 UserID from SysRequest Where Status<>'P' and Param1='SEND' and Param2 in ('UND','UND-MW','IT','MAR') and Param3=" & ANO & " and ActionCode<>''")
            If Not rs.Eof Then
                ErrMsg = "Already In Action Flow Request By User : " & rs.fields("UserID") & " !"
                Throw New Exception(ErrMsg)
            End If

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "DROPACCEPTANCEDETAIL", ANO) Then
                Throw New Exception(Me.ErrDesc)
            End If

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select LANO, CNO, AType from Acceptance Where AStatus='W' and ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Please Check Policy Data Status !"
                Throw New Exception(ErrMsg)
            End If
            Dim LANO As Long
            Dim CNO As Long
            Dim AType As String
            Dim OpenFlag As Boolean
            Dim pCNO As Long

            LANO = toval(rs.fields("LANO"))
            CNO = toval(rs.fields("CNO"))
            AType = tostr(rs.fields("AType"))

            rs = db.OpenRemoteSet("Select OpenFlag, ASource, CStatus, pCNO from Cover Where CNO=" & CNO)
            OpenFlag = rs.fields("OpenFlag")
            pCNO = rs.fields("pCNO")

            If rs.fields("ASource") = "MKT" And rs.fields("CStatus") = "R" Then
                ErrMsg = "Can't Drop Quotation Data !"
                Throw New Exception(ErrMsg)
            End If

            If AType = "N" And (ANO = LANO Or LANO = -1) Then
                rs = db.OpenRemoteSet("Select Top 1 C.CStatus, C.ASource from Cover C, Acceptance A where (A.ANO=A.LANO or A.LANO=-1) and A.AType='N' and A.CNO=C.CNO and C.CStatus='A' and C.ASource='MKT' and C.CNO=" & pCNO)
                If Not rs.Eof Then
                    ErrMsg = "Can't Drop Policy Data Sourcing from Quotation, use Close Policy Instead !"
                    Throw New Exception(ErrMsg)
                End If
            End If

            If ANO = LANO And Not OpenFlag Then
                ' Check Detail Acceptance for Conjunction Policies
                rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1 Left Join Acceptance pA1 on (pA1.ANO=A1.pANO), Acceptance A2 Inner Join Cover C on (C.CNO=A2.CNO) Left Join RArrHeader R on (R.ANO=A2.ANO) Where (R.Status='A' or A2.AStatus='W') and A2.CANO=A1.ANO and coalesce(pA1.AType,'') Not in ('N','E') and A1.AType in ('N','E') and A1.LANO=" & ANO)
                If Not rs.Eof Then
                    ErrMsg = "Must Drop the Co. Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                    Throw New Exception(ErrMsg)
                End If
            Else
                rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1 Left Join Acceptance pA1 on (pA1.ANO=A1.pANO), Acceptance A2 Inner Join Cover C on (C.CNO=A2.CNO) Left Join RArrHeader R on (R.ANO=A2.ANO) Where (R.Status='A' or A2.AStatus='W') and A2.CANO=A1.ANO and coalesce(pA1.AType,'') Not in ('N','E') and A1.AType in ('N','E') and A1.ANO=" & ANO)
                If Not rs.Eof Then
                    ErrMsg = "Must Drop the Co. Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                    Throw New Exception(ErrMsg)
                End If
            End If

            rs = db.OpenRemoteSet("Select Top 1 UserID from SysRequest Where Status<>'P' and Param1='SEND' and Param2 in ('UND','UND-MW','IT','MAR') and Param3=" & ANO & " and ActionCode<>''")
            If Not rs.Eof Then
                ErrMsg = "Already In Action Flow Request By User : " & rs.fields("UserID") & " !"
                Throw New Exception(ErrMsg)
            End If

            If ANO = LANO Then
                If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "DROPMASTER", CNO, ANO) Then
                    Throw New Exception(Me.ErrDesc)
                End If
            Else
                If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "DROPACCEPTANCE", CNO, ANO) Then
                    Throw New Exception(Me.ErrDesc)
                End If
            End If

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListRiskAccumulation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim RegID As String = ""
            Dim FLDID As String = ""
            Dim ValueID As String = ""
            Dim RYear As Integer = 0

            ANO = getInParamValue(InParam, 0)
            RegID = getInParamValue(InParam, 1)
            FLDID = getInParamValue(InParam, 2)
            ValueID = getInParamValue(InParam, 3)

            Dim rs As RemoteSet
            Dim rsR As RemoteSet

            Dataset = New DataSet
            Dim dt As DataTable = New DataTable
            Dim NewRow As DataRow
            dt.Columns.Add("Column0", Type.GetType("System.String"))
            dt.Columns.Add("Column1", Type.GetType("System.String"))
            dt.Columns.Add("Column2", Type.GetType("System.String"))
            dt.Columns.Add("Column3", Type.GetType("System.String"))
            dt.Columns.Add("Column4", Type.GetType("System.String"))
            Dataset.Tables.Add(dt)

            Dim RiskAccumulationByTreatyYearF As Boolean
            rs = db.OpenRemoteSet("Select RiskAccumulationByTreatyYearF from systab")
            RiskAccumulationByTreatyYearF = toval(rs.fields("RiskAccumulationByTreatyYearF"))

            If ANO > 0 Then
                rsR = db.OpenQueryDef("QACCREG", adOpenForwardonly, ANO)
                RegID = tostr(rsR.fields("RegID"))
                FLDID = tostr(rsR.fields("FLDID"))
                ValueID = tostr(rsR.fields("ValueID"))
                If RiskAccumulationByTreatyYearF Then
                    RYear = toval(rsR.fields("RYear"))
                End If
            End If

            Dim FLDTag As String
            FLDTag = Get_FLDTAG(FLDID)

            rsR = db.OpenQueryDef("QGenDTab", adOpenKeyset, FLDID, ValueID)
            NewRow = Dataset.Tables(0).NewRow
            NewRow("Column1") = "Acc. Reg :"
            NewRow("Column2") = ValueID
            NewRow("Column3") = tostr(rsR.fields("Description"))
            Dataset.Tables(0).Rows.Add(NewRow)

            NewRow = Dataset.Tables(0).NewRow
            NewRow("Column0") = "RISK INFO..."
            Dataset.Tables(0).Rows.Add(NewRow)

            NewRow = Dataset.Tables(0).NewRow
            NewRow("Column1") = "Object Information"
            Dataset.Tables(0).Rows.Add(NewRow)

            Dim PFLDID As String
            Dim i As Integer

            i = 1
            Do While i <= 15
                PFLDID = tostr(rsR.fields("FLDID" & i))
                If PFLDID <> "" Then
                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column1") = Get_FLDTAG(PFLDID) & ":"
                    NewRow("Column2") = tostr(rsR.fields("VALUEID" & i))
                    NewRow("Column3") = tostr(rsR.fields("VALUEDESC" & i))
                    Dataset.Tables(0).Rows.Add(NewRow)
                Else
                    Exit Do
                End If
                i = i + 1
            Loop

            If FLDID <> "" And ValueID <> "" Then
                Dim TSI As Double
                rsR = db.OpenQueryDef("QAccumulation", adOpenKeyset, RegID, FLDID, ValueID, "Q", RYear)
                TSI = toval(rsR.fields("TSI"))
                NewRow = Dataset.Tables(0).NewRow
                If toval(rsR.fields("ReAllocF")) Then
                    NewRow("Column0") = "(Need Reallocation)"
                End If
                NewRow("Column1") = "Total Sum Insured:"
                NewRow("Column2") = Format(TSI, "###,###,##0.00")
                Dataset.Tables(0).Rows.Add(NewRow)

                rsR = db.OpenRemoteSet("Select * from QAccumulation_Control('" & RegID & "','" & FLDID & "','" & ValueID & "'," & RYear & ") A")
                TSI = toval(rsR.fields("TSI"))
                NewRow = Dataset.Tables(0).NewRow
                NewRow("Column1") = "Zone " & tostr(rsR.fields("Description")) & " :"
                NewRow("Column2") = Format(TSI, "###,###,##0.00")
                Dataset.Tables(0).Rows.Add(NewRow)
            End If

            Dim TLimit As Double
            Dim RID As String

            NewRow = Dataset.Tables(0).NewRow
            NewRow("Column0") = "POLICY INFO..."
            NewRow("Column1") = "Number Of Policies :"

            ' Show Policies Info
            If FLDID <> "" And ValueID <> "" Then
                rs = db.OpenRemoteSet("Select A1.ANO, A1.PolicyNo+'-'+A1.CertificateNo as PolicyNo, A1.ValueDesc, A1.SDate, A1.EDate, A1.PDate, A1.AName from QAccumulation_Detail('" & RegID & "','" & FLDID & "','" & ValueID & "'," & RYear & ") A1")
            Else
                rs = db.OpenRemoteSet("Select A1.ANO, A1.PolicyNo+'-'+A1.CertificateNo as PolicyNo, A1.ValueDesc, A1.SDate, A1.EDate, A1.PDate, A1.AName from Acceptance A1, RArrHeader R1 Where R1.ANO=A1.ANO and R1.Status='A' and A1.ANO=" & ANO)
            End If
            NewRow("Column2") = Format(rs.RecordCount, "#,##0")
            Dataset.Tables(0).Rows.Add(NewRow)

            Dim pANO As Long
            Dim Amount_C As Double
            Dim Amount_D As Double
            Dim Amount_Q As Double
            Dim Amount_XP As Double
            Dim Amount_S1 As Double
            Dim Amount_S2 As Double
            Dim Amount_S3 As Double
            Dim Amount_A As Double

            If Not rs.Eof Then
                Amount_Q = Get_Accumulation(RegID, FLDID, ValueID, "Q", RYear)
                Amount_C = Get_Accumulation(RegID, FLDID, ValueID, "C", RYear)
                Amount_D = Get_Accumulation(RegID, FLDID, ValueID, "D", RYear)
                Amount_XP = Get_Accumulation(RegID, FLDID, ValueID, "XP", RYear)
                Amount_S1 = Get_Accumulation(RegID, FLDID, ValueID, "S1", RYear)
                Amount_S2 = Get_Accumulation(RegID, FLDID, ValueID, "S2", RYear)
                Amount_S3 = Get_Accumulation(RegID, FLDID, ValueID, "S3", RYear)
                Amount_A = Get_Accumulation(RegID, FLDID, ValueID, "A", RYear)
            End If
            Do While Not rs.Eof
                pANO = rs.fields("ANO")

                NewRow = Dataset.Tables(0).NewRow
                Dataset.Tables(0).Rows.Add(NewRow)

                NewRow = Dataset.Tables(0).NewRow
                NewRow("Column0") = IIf(pANO = ANO, "(this)", "")
                NewRow("Column1") = "Policy No :"
                NewRow("Column3") = rs.fields("PolicyNo")
                Dataset.Tables(0).Rows.Add(NewRow)

                NewRow = Dataset.Tables(0).NewRow
                NewRow("Column0") = ""
                NewRow("Column1") = "Insured Name :"
                NewRow("Column3") = rs.fields("AName")
                Dataset.Tables(0).Rows.Add(NewRow)

                NewRow = Dataset.Tables(0).NewRow
                NewRow("Column0") = ""
                NewRow("Column1") = "Ins. Period :"
                NewRow("Column3") = Format(rs.fields("SDate"), "dd-mmm-yyyy") & " - " & Format(rs.fields("Edate"), "dd-mmm-yyyy")
                Dataset.Tables(0).Rows.Add(NewRow)

                NewRow = Dataset.Tables(0).NewRow
                NewRow("Column0") = ""
                NewRow("Column1") = "Last Eff. Period :"
                NewRow("Column3") = Format(rs.fields("PDate"), "dd-mmm-yyyy")
                Dataset.Tables(0).Rows.Add(NewRow)

                NewRow = Dataset.Tables(0).NewRow
                NewRow("Column0") = ""
                NewRow("Column1") = FLDTag & " :"
                NewRow("Column3") = rs.fields("ValueDesc")
                Dataset.Tables(0).Rows.Add(NewRow)

                NewRow = Dataset.Tables(0).NewRow
                NewRow("Column1") = "Spreading Of Risk"
                NewRow("Column3") = "Reinsurance ID"
                NewRow("Column4") = "Risk Accumulation"
                Dataset.Tables(0).Rows.Add(NewRow)

                rsR = db.OpenQueryDef("QRArrHeader", adOpenKeyset, pANO)
                If rsR.fields("Status") = "A" Then   ' Booked
                    TLimit = toval(rsR.fields("Limit"))

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = IIf(rsR.fields("ManualF"), "(Manual Mode)", "(Auto Mode)")
                    NewRow("Column1") = "BPPDAN"
                    NewRow("Column2") = Format(toval(rsR.fields("CAMOUNT") - rsR.fields("CCAMOUNT")), "###,###,##0.00")
                    RID = tostr(rsR.fields("CID"))
                    NewRow("Column3") = RID
                    NewRow("Column4") = Format(Amount_C, "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "Decree"
                    NewRow("Column2") = Format(toval(rsR.fields("DAMOUNT") - rsR.fields("CDAMOUNT")), "###,###,##0.00")
                    RID = tostr(rsR.fields("DID"))
                    NewRow("Column3") = RID
                    NewRow("Column4") = Format(Amount_D, "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "MAIPARK"
                    NewRow("Column2") = Format(toval(rsR.fields("XPAMOUNT") - rsR.fields("CXPAMOUNT")), "###,###,##0.00")
                    RID = tostr(rsR.fields("XPID"))
                    NewRow("Column3") = RID
                    NewRow("Column4") = Format(Amount_XP, "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "Quota Share(+OR)"
                    NewRow("Column2") = Format(toval(rsR.fields("QAMOUNT") + rsR.fields("QSRAMOUNT") - rsR.fields("CQAMOUNT")), "###,###,##0.00")
                    RID = tostr(rsR.fields("QID"))
                    NewRow("Column3") = RID
                    NewRow("Column4") = Format(Amount_Q, "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "1st Surplus"
                    NewRow("Column2") = Format(toval(rsR.fields("SP1AMOUNT") - rsR.fields("CSP1AMOUNT")), "###,###,##0.00")
                    RID = tostr(rsR.fields("SP1ID"))
                    NewRow("Column3") = RID
                    NewRow("Column4") = Format(Amount_S1, "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "2nd Surplus"
                    NewRow("Column2") = Format(toval(rsR.fields("SP2AMOUNT") - rsR.fields("CSP2AMOUNT")), "###,###,##0.00")
                    RID = tostr(rsR.fields("SP2ID"))
                    NewRow("Column3") = RID
                    NewRow("Column4") = Format(Amount_S2, "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "3rd Surplus"
                    NewRow("Column2") = Format(toval(rsR.fields("SP3AMOUNT") - rsR.fields("CSP3AMOUNT")), "###,###,##0.00")
                    RID = tostr(rsR.fields("SP3ID"))
                    NewRow("Column3") = RID
                    NewRow("Column4") = Format(Amount_S3, "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "Auto Facultative"
                    NewRow("Column2") = Format(toval(rsR.fields("AAMOUNT") + -rsR.fields("CAAMOUNT")), "###,###,##0.00")
                    RID = tostr(rsR.fields("AID"))
                    NewRow("Column3") = RID
                    NewRow("Column4") = Format(Amount_A, "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "FACULTATIVE"
                    NewRow("Column2") = Format(toval(rsR.fields("FAMOUNT")), "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "EXCESS+SHORTFALL"
                    NewRow("Column2") = Format(toval(rsR.fields("RAMOUNT")) + toval(rsR.fields("SAMOUNT")), "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                    NewRow = Dataset.Tables(0).NewRow
                    NewRow("Column0") = ""
                    NewRow("Column1") = "TOTAL"
                    NewRow("Column2") = Format(toval(rsR.fields("TAMOUNT")), "###,###,##0.00")
                    Dataset.Tables(0).Rows.Add(NewRow)

                End If
                rs.MoveNext()
            Loop

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListTreatyReallocation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim RegID As String
            Dim FLDID As String
            Dim ValueID As String

            ANO = getInParamValue(InParam, 0)
            RegID = getInParamValue(InParam, 1)
            FLDID = getInParamValue(InParam, 2)
            ValueID = getInParamValue(InParam, 3)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select Cast(0 as bit) as ProcessFlag, '" & RegID & "' as REGID, FLDID, VALUEID, DESCRIPTION, VALUEID as NEWVALUEID, DESCRIPTION as NEWDESCRIPTION from GenDTab Where FLDID='" & FLDID & "' and ValueID='" & ValueID & "'")
            If rs.Eof Then
                rs = db.OpenRemoteSet("Select Cast(0 as bit) as ProcessFlag, '" & RegID & "' as REGID, FLDID, VALUEID, VALUEDESC AS DESCRIPTION, VALUEID as NEWVALUEID, VALUEDESC AS NEWDESCRIPTION from Acceptance Where ANO=" & ANO)
            End If
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ProcessTreatyReallocation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ADate As Date
            Dim RegID As String
            Dim FLDID As String
            Dim ValueID As String
            Dim NewValueID As String
            Dim AutoInforceF As Boolean
            Dim RYear As Integer

            ADate = getInParamValue(InParam, 0)
            RegID = getInParamValue(InParam, 1)
            FLDID = getInParamValue(InParam, 2)
            ValueID = getInParamValue(InParam, 3)
            NewValueID = getInParamValue(InParam, 4)
            AutoInforceF = toval(getInParamValue(InParam, 5))
            RYear = toval(getInParamValue(InParam, 6))
            If Not db.GetLockedObject("Policy" & Left(RegID & FLDID & ValueID, 40)) Then Throw New Exception(db.ErrDesc)

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "TREATY_ADJUSTMENT", InParam) Then Throw New Exception(Me.ErrDesc)

            If Not db.ReleaseLockedObject("Policy" & Left(RegID & FLDID & ValueID, 40)) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function PremiumCalculationAll(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim LANO As Long         ' Master ANO or LANO
            Dim ResetF As Boolean

            LANO = getInParamValue(InParam, 0)
            ResetF = toval(getInParamValue(InParam, 1))

            If Not Me.Direct_Call_Ext_Process("Process.NetV3.Proclib", "Premium_Calc_All", LANO, ResetF) Then Throw New Exception(Me.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function PreRemoveRiskAccumulationConditionsOk(ByVal RegID As String, ByVal FLDID As String, ByVal ValueID As String, ByVal RYear As Integer, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet

        On Error GoTo errmsg

        rs = db.OpenRemoteSet("Select A.PolicyNo, A.AStatus, A.RegNo from QAccumulation_Detail('" & RegID & "','" & FLDID & "','" & ValueID & "'," & RYear & ") A Where A.AStatus not in ('W','C') ")
        If Not rs.Eof Then
            ErrMsg = "Please Check Policy Data Status with Register/CI No. " & rs.fields("RegNo") & ""
            GoTo errmsg
        End If

        PreRemoveRiskAccumulationConditionsOk = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PreRemoveRiskAccumulationConditionsOk = False
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function RemoveRiskAccumulation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim RegID As String
            Dim FLDID As String
            Dim ValueID As String
            Dim RYear As Integer

            RegID = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            RYear = toval(getInParamValue(InParam, 3))

            If Not db.GetLockedObject("Policy" & Left(RegID & FLDID & ValueID, 40)) Then Throw New Exception(db.ErrDesc)

            If Not PreRemoveRiskAccumulationConditionsOk(RegID, FLDID, ValueID, RYear, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            Dim rs As RemoteSet
            If RYear = 0 Then
                If Not db.Execute("Delete from AccRegDet Where RegID='" & RegID & "' and FLDID='" & FLDID & "' and ValueID='" & ValueID & "'") Then Throw New Exception(db.ErrDesc)
            Else
                If Not db.Execute("Delete from AccRegDet Where RegID='" & RegID & "' and FLDID='" & FLDID & "' and ValueID='" & ValueID & "' and RYear=" & RYear) Then Throw New Exception(db.ErrDesc)
            End If
            rs = db.OpenRemoteSet("Select A.ANO, A.RegNo, A.AStatus from QAccumulation_Detail('" & RegID & "','" & FLDID & "','" & ValueID & "'," & RYear & ") A")
            Do While Not rs.Eof
                If rs.fields("AStatus") <> "W" And rs.fields("AStatus") <> "C" Then
                    ErrMsg = "Please Check Policy Data Status with Register/CI No. " & rs.fields("RegNo") & ""
                    Throw New Exception(ErrMsg)
                End If

                If rs.fields("AStatus") = "W" Then
                    If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & rs.fields("ANO")) Then Throw New Exception(db.ErrDesc)
                    If db.RecordsAffected = 0 Then
                        ErrMsg = "Please Check Policy Data Status with Register/CI No. " & rs.fields("RegNo") & ""
                        Throw New Exception(ErrMsg)
                    End If
                End If

                If Not db.Execute("Update RArrHeader Set Status='R' Where Status='A' and ANO=" & rs.fields("ANO")) Then Throw New Exception(db.ErrDesc)
                If db.RecordsAffected = 0 Then
                    ErrMsg = "Please Check Policy Data Status with Register/CI No. " & rs.fields("RegNo") & ""
                    Throw New Exception(ErrMsg)
                End If
                rs.MoveNext()
            Loop

            If Not db.ReleaseLockedObject("Policy" & Left(RegID & FLDID & ValueID, 40)) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function RecalculateRiskAccumulation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim RegID As String
            Dim FLDID As String
            Dim ValueID As String
            Dim RYear As Integer

            RegID = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            RYear = toval(getInParamValue(InParam, 3))

            If Not db.GetLockedObject("Policy" & Left(RegID & FLDID & ValueID, 40)) Then Throw New Exception(db.ErrDesc)

            Dim rs As RemoteSet
            Dim RLimit As Double
            rs = db.OpenQueryDef("Treaty_Adjustment_Limit", adOpenForwardonly, RegID, FLDID, ValueID, RYear)
            RLimit = toval(rs.fields("Limit"))

            rs = db.OpenRemoteSet("Select A.ANO from QAccumulation_Detail('" & RegID & "','" & FLDID & "','" & ValueID & "'," & RYear & ") A Where A.AStatus='W' and A.AType not in ('A','C')")
            Do While Not rs.Eof
                If Not Me.Direct_Call_Ext_Process("Process.NetV3.Proclib", "Reinsurance_Cession", toval(rs.fields("ANO")), RLimit, , True) Then Throw New Exception(Me.ErrDesc)
                rs.MoveNext()
            Loop

            If Not db.ReleaseLockedObject("Policy" & Left(RegID & FLDID & ValueID, 40)) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetListInterest(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim sRecNo As Long
            Dim eRecNo As Long

            ANO = getInParamValue(InParam, 0)
            sRecNo = toval(getInParamValue(InParam, 1))
            eRecNo = toval(getInParamValue(InParam, 2))

            Dim rs As RemoteSet
            rs = db.OpenQueryDef("GetListInterest", adOpenKeyset, ANO, sRecNo, eRecNo)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetInterest(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef Dataset2 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim INO As Long
            INO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet

            rs = db.OpenQueryDef("QInterest_Data", adOpenKeyset, INO)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            rs = db.OpenQueryDef("QICoverDet", adOpenKeyset, INO, 0)
            Dataset = rs.GetDataSet

            rs = db.OpenRemoteSet("Select * from ICoverNOC where INO= " & INO & " Order By OrderNo")
            Dataset2 = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveInterest(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByVal Dataset2 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim INO As Long
            Dim ANO As Long
            ' Get Input Parameter

            INO = getInParamValue(InParam, 0)
            ANO = getInParamValue(InParam, 1)
            ' getInParamValue(InParam, 0) = INO
            ' getInParamValue(InParam, 1) = ANO
            ' getInParamValue(InParam, 2) = ICoverDetF
            ' getInParamValue(InParam, 3) = ICoverExtF

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus, A.OANO, A.TOC, TOC.CheckDuplicateInvoiceNoF, A.FLDID, A.VALUEID from Acceptance A Inner Join TOC on (TOC.TOC=A.TOC) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            If toval(rs.fields("CheckDuplicateInvoiceNoF")) Then
                Dim SQL As String
                SQL = ""

                SQL = SQL & " Declare @P_OANO as Int"
                SQL = SQL & " Declare @P_TOC as varchar(10)"
                SQL = SQL & " Declare @P_FLDID as varchar(255)"
                SQL = SQL & " Declare @P_VALUEID as varchar(255)"
                SQL = SQL & " Declare @P_RefNo as varchar(255)"
                SQL = SQL & " "
                SQL = SQL & " Select @P_OANO=" & toval(rs.fields("OANO"))
                SQL = SQL & " Select @P_TOC='" & tostr(rs.fields("TOC")) & "'"
                SQL = SQL & " Select @P_FLDID='" & tostr(rs.fields("FLDID")) & "'"
                SQL = SQL & " Select @P_VALUEID='" & tostr(rs.fields("ValueID")) & "'"
                SQL = SQL & " Select @P_REFNO='" & GetFieldValue(fbuffer(GetFieldIndex("RefNo", findex))) & "'"
                SQL = SQL & " "
                SQL = SQL & " Select Top 1 A.PolicyNo, A.CertificateNo"
                SQL = SQL & " from Acceptance A"
                SQL = SQL & " Inner Join ICover I on (I.ANO=A.ANO)"
                SQL = SQL & " Where not (A.AStatus='C' and A.pAStatus='W') and A.AStatus not in ('W','T')"
                SQL = SQL & " and A.OANO<>@P_OANO and A.TOC=@P_TOC and A.FLDID=@P_FLDID and A.ValueID=@P_VALUEID and I.RefNo=@P_RefNO"

                rs = db.OpenRemoteSet(SQL)
                If Not rs.Eof Then
                    Throw New Exception("Please Check the Policy Data " & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & " Duplicate Invoice No !")
                End If

            End If


            Const MaxCols = 6
            Dim Bufferdata(Dataset.Tables(0).Rows.Count - 1, MaxCols - 1) As Object

            Dim i As Integer
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Bufferdata(i, 0) = tostr(Dataset.Tables(0).Rows(i)("Description"))
                Bufferdata(i, 1) = toval(Dataset.Tables(0).Rows(i)("Qty"))
                Bufferdata(i, 2) = tostr(Dataset.Tables(0).Rows(i)("Unit"))
                Bufferdata(i, 3) = toval(Dataset.Tables(0).Rows(i)("UValue"))
                Bufferdata(i, 4) = toval(Dataset.Tables(0).Rows(i)("Amount"))
                Bufferdata(i, 5) = tostr(Dataset.Tables(0).Rows(i)("Memo"))
            Next

            Const MaxCols_2 = 9
            Dim bufferdata_2(Dataset2.Tables(0).Rows.Count - 1, MaxCols_2 - 1) As Object
            For i = 0 To Dataset2.Tables(0).Rows.Count - 1
                bufferdata_2(i, 0) = tostr(Dataset2.Tables(0).Rows(i)("NOC"))
                bufferdata_2(i, 1) = tostr(Dataset2.Tables(0).Rows(i)("Description"))
                bufferdata_2(i, 2) = toval(Dataset2.Tables(0).Rows(i)("MaxFreq"))
                bufferdata_2(i, 3) = toval(Dataset2.Tables(0).Rows(i)("MaxQty"))
                bufferdata_2(i, 4) = tostr(Dataset2.Tables(0).Rows(i)("Currency"))
                bufferdata_2(i, 5) = toval(Dataset2.Tables(0).Rows(i)("PCT"))
                bufferdata_2(i, 6) = toval(Dataset2.Tables(0).Rows(i)("Amount"))
                bufferdata_2(i, 7) = toval(Dataset2.Tables(0).Rows(i)("Limit"))
                bufferdata_2(i, 8) = toval(Dataset2.Tables(0).Rows(i)("PolicyYear"))
            Next

            INO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Save_Interest", InParam, findex, fbuffer, TrimBuffer(Bufferdata), TrimBuffer(bufferdata_2))
            If INO = 0 Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = INO

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropInterest(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim INO As Long
            INO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * From ICover Where INO=" & INO & "")
            If rs.Eof Then
                ErrMsg = "No Interest Data !"
                Throw New Exception(ErrMsg)
            End If

            Dim ANO As Long
            ANO = rs.fields("ANO")
            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetDefaultInterestDetail(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset2 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim TOI As String

            ANO = getInParamValue(InParam, 0)
            TOI = getInParamValue(InParam, 1)

            Dim rs As RemoteSet

            rs = db.OpenRemoteSet("Select Cast(0 as int) as INO, cast(0 as int) as ANO, I1.NOC, I1.Currency, I1.Description, I1.MaxFreq, I1.MaxQty, I1.PCT, I1.Amount, I1.OrderNo, I1.Limit, I1.Indemnity, I1.PolicyYear from ICoverNOC I1, ICover I2, Acceptance A1 Where A1.ANO=" & ANO & " and I2.ANO=A1.LANO and I2.TOI='" & TOI & "' and I1.INO=I2.INO Order By I1.OrderNo")
            If rs.Eof Then
                rs = db.OpenRemoteSet("Select Cast(0 as int) as INO, cast(0 as int) as ANO, NOC.NOC, '' as Currency, NOC.Description, NOC.MaxFreq, NOC.MaxQty, NOC.PCT, NOC.Amount, cast(0 AS int) as OrderNo, NOC.Limit, cast(0 as float) as Indemnity, cast(0 as int) as PolicyYear from NOC, ACCEPTANCE WHERE NOC.DEFAULTF=1 and EXISTS (SELECT P.TOI FROM RCLASS R, PCLASS P WHERE R.TOC=ACCEPTANCE.TOC AND R.TOR=P.TOR AND P.TOI=NOC.TOI AND P.TOI=NOC.TOI) AND NOC.TOI='" & TOI & "' AND ACCEPTANCE.ANO=" & ANO)
            End If

            Dataset2 = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListCoverage(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            Dim PSDate As Date
            Dim PEDate As Date

            rs = db.OpenRemoteSet("Select SDate, EDate from Acceptance Where ANO=" & ANO)
            PSDate = rs.fields("SDate")
            PEDate = rs.fields("EDate")

            SQL = "Select Cast(0 as bit) as ProcessFlag, cast(0 as int) as PolicyYear, RCover.*, RCover.Rate * (100 - RCover.DISCRATE) / 100 * RCover.Loading as NetRate"
            SQL = SQL & " from"
            SQL = SQL & " ("
            SQL = SQL & " Select RCOVER.*,"
            SQL = SQL & " Case "
            SQL = SQL & " When RCover.DiscUnit='M' Then"
            SQL = SQL & " (RCOVER.DISCOUNT+RCOVER.DISCOUNT_EXT)/10"
            SQL = SQL & " When RCover.DiscUnit='F' Then"
            SQL = SQL & " (RCOVER.DISCOUNT+RCOVER.DISCOUNT_EXT)"
            SQL = SQL & " /"
            SQL = SQL & " ("
            SQL = SQL & " CASE "
            SQL = SQL & " WHEN RCOVER.UNIT='L' Then 0 "
            SQL = SQL & " WHEN RCOVER.UNIT='C' THEN RCOVER.RATE/100"
            SQL = SQL & " WHEN RCOVER.UNIT='M' THEN RCOVER.RATE/1000"
            SQL = SQL & " WHEN RCOVER.UNIT='F' THEN RCOVER.RATE"
            SQL = SQL & " End"
            SQL = SQL & " *"
            SQL = SQL & " CASE"
            SQL = SQL & " WHEN RCOVER.UNIT='C' or RCOVER.UNIT='M' THEN"
            SQL = SQL & " ("
            SQL = SQL & " Select SUM(ICover.SI)"
            SQL = SQL & " FROM ACCEPTANCE"
            SQL = SQL & " INNER JOIN ICOVER ON (ICOVER.ANO=ACCEPTANCE.ANO)"
            SQL = SQL & " INNER JOIN RCLASS ON (RCLASS.TOC=ACCEPTANCE.TOC)"
            SQL = SQL & " INNER JOIN RATECOV ON (RATECOV.CODE=RCOVER.CODE And RATECOV.TOR=RCLASS.TOR)"
            SQL = SQL & " INNER JOIN PCLASS ON (PCLASS.TOR=RATECOV.TOR And PCLASS.TOI=ICOVER.TOI)"
            SQL = SQL & " WHERE(ACCEPTANCE.ANO = RCOVER.ANO)"
            SQL = SQL & " )"
            SQL = SQL & " ELSE	1"
            SQL = SQL & " End"
            SQL = SQL & " )*100"
            SQL = SQL & " ELSE"
            SQL = SQL & " (RCOVER.DISCOUNT+RCOVER.DISCOUNT_EXT)"
            SQL = SQL & " END as DISCRATE"
            SQL = SQL & " from RCOVER"
            SQL = SQL & " WHERE RCOVER.ANO=" & ANO
            SQL = SQL & " ) RCover"
            SQL = SQL & " Order By RCOVER.Code, RCOVER.SDate, RCOVER.PCTLoss"

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet

            Dim i As Integer

            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If DateDiff(DateInterval.Month, PSDate, Dataset.Tables(0).Rows(i)("SDate")) = 0 And DateDiff(DateInterval.Month, PEDate, Dataset.Tables(0).Rows(i)("EDate")) = 0 Then
                    Dataset.Tables(0).Rows(i)("PolicyYear") = 0
                Else
                    Dataset.Tables(0).Rows(i)("PolicyYear") = DateDiff(DateInterval.Month, PSDate, Dataset.Tables(0).Rows(i)("SDate")) \ 12 + 1
                End If
            Next


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetCoverage(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Code As String
            Dim SDate As Date
            ANO = getInParamValue(InParam, 0)
            Code = getInParamValue(InParam, 1)
            SDate = getInParamValue(InParam, 2)

            Dim rs As RemoteSet

            rs = db.OpenQueryDef("QRCover", adOpenKeyset, ANO, Code, SDate)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            rs = db.OpenRemoteSet("Select * from RCoverDet Where ANO=" & ANO & " and Code='" & Code & "' and SDate='" & Format(SDate, "MM/DD/YYYY") & "'")
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveCoverage(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim PCode As String
            Dim PSDate As String
            ' getInParamValue(InParam, 0) = ANO
            ' getInParamValue(InParam, 1) = Code
            ' getInParamValue(InParam, 2) = SDate
            ' getInParamValue(InParam, 3) = Portion
            ANO = getInParamValue(InParam, 0)
            PCode = getInParamValue(InParam, 4)
            PSDate = getInParamValue(InParam, 5)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            fbuffer(GetFieldIndex("Code", findex)) = SetFieldValue(PCode)
            fbuffer(GetFieldIndex("SDate", findex)) = SetFieldValue(PSDate)

            Const MaxCols = 6
            Dim Bufferdata(Dataset.Tables(0).Rows.Count - 1, MaxCols - 1) As Object

            Dim i As Integer
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Bufferdata(i, 0) = todate(Dataset.Tables(0).Rows(i)("TransDate"))
                Bufferdata(i, 1) = tostr(Dataset.Tables(0).Rows(i)("Voyage"))
                Bufferdata(i, 2) = tostr(Dataset.Tables(0).Rows(i)("Vessel"))
                Bufferdata(i, 3) = tostr(Dataset.Tables(0).Rows(i)("Conveyance"))
                Bufferdata(i, 4) = tostr(Dataset.Tables(0).Rows(i)("AtAndFrom"))
                Bufferdata(i, 5) = tostr(Dataset.Tables(0).Rows(i)("TransTo"))
            Next

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "save_coverage", InParam, findex, fbuffer, TrimBuffer(Bufferdata, 1)) Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SplitRiskCoverage(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim RCode As String
            Dim SDate As Date
            ANO = getInParamValue(InParam, 0)
            RCode = getInParamValue(InParam, 1)
            SDate = getInParamValue(InParam, 2)
            ' InParam 3 - ... PCTLOss Year 1 - n

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            Dim rsS As RemoteSet
            Dim rsD As RemoteSet
            Dim EDate As Date
            Dim RSDate As Date
            Dim REDate As Date
            Dim PCalc As String
            Dim Rate As Double
            Dim PCTLoss As Double

            Dim i As Integer
            rsS = db.OpenRemoteSet("Select * From RCover Where ANO=" & ANO & " and Code='" & RCode & "' and SDate='" & Format(SDate, "MM/DD/YYYY") & "'")
            rsD = db.OpenRemoteSet("Select * from RCover Where 1<>1")
            If Not rsS.Eof Then
                SDate = rsS.fields("SDate")
                EDate = rsS.fields("EDate")
                PCalc = rsS.fields("PCalc")
                Rate = rsS.fields("Rate")
                RSDate = SDate
                i = 1
                Do While DateDiff(DateInterval.Day, RSDate, EDate) > 0
                    PCTLoss = toval(getInParamValue(InParam, 2 + i))
                    If PCTLoss = 0 Then PCTLoss = rsS.fields("PCTLOSS")

                    REDate = DateAdd(DateInterval.Month, 12 * i, SDate)
                    If REDate > EDate Then REDate = EDate
                    If RSDate = SDate Then
                        rsS.Edit()
                        rsS.fields("EDate") = REDate
                        If PCalc = "F" Or PCalc = "S" Then
                            If Day(SDate) = Day(EDate) Then
                                rsS.fields("Rate") = Rate * DateDiff(DateInterval.Month, RSDate, REDate) / DateDiff(DateInterval.Month, SDate, EDate)

                            Else
                                rsS.fields("Rate") = Rate * DateDiff(DateInterval.Day, RSDate, REDate) / DateDiff(DateInterval.Day, SDate, EDate)
                            End If
                        End If
                        rsS.fields("PCTLoss") = PCTLoss
                        rsS.fields("PAddition") = 100 - PCTLoss
                        If Not rsS.Update Then Throw New Exception(db.ErrDesc)
                    Else
                        rsD.AddNew()
                        If Not rsD.PutRecord(rsS.GetRecord()) Then Throw New Exception(db.ErrDesc)
                        rsD.fields("SDate") = RSDate
                        rsD.fields("EDate") = REDate
                        rsD.fields("NCDate") = rsD.fields("SDate")
                        If PCalc = "F" Or PCalc = "S" Then
                            If Day(SDate) = Day(EDate) Then
                                rsD.fields("Rate") = Rate * DateDiff(DateInterval.Month, RSDate, REDate) / DateDiff(DateInterval.Month, SDate, EDate)
                            Else
                                rsD.fields("Rate") = Rate * DateDiff(DateInterval.Day, RSDate, REDate) / DateDiff(DateInterval.Day, SDate, EDate)
                            End If
                        End If
                        rsD.fields("PCTLoss") = PCTLoss
                        rsD.fields("PAddition") = 100 - PCTLoss
                        If Not rsD.Update Then Throw New Exception(db.ErrDesc)
                    End If

                    i = i + 1
                    RSDate = REDate
                Loop
            End If


            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
       Public Function LoadICoverFromExcel(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef dsExcel As DataSet, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            Dim ANO As Long
            Dim Mode As String
            Dim rs As RemoteSet
            Dim rsTOI As RemoteSet

            ANO = toval(getInParamValue(InParam, 0))
            Mode = toval(getInParamValue(InParam, 1)) ' Mode : 0 (Reset All), 1 (Add), 2(Update)

            Dim TOI As String = ""
            Dim Category As String
            Dim Currency As String = ""
            Dim Remark As String
            Dim RefNo As String
            Dim PCTADJ As Double
            Dim Amount As Double
            Dim Amount_AOE As Double
            Dim SI As Double
            Dim StrMsg As String = ""
            Dim SQL As String

            Dim NewRow As DataRow

            SQL = "Select Cast(0 as bit) as ProcessFlag, "
            SQL = SQL & " I1.TOI, I1.Currency, I1.Remark, I1.Category, I1.PCTADJ, I1.Amount, I1.Amount_AOE, I1.SI, I1.TOPROF, I1.RefNo, I1.RefDate, I1.RefDueDate, I1.DocNo, I1.DocDate,"
            SQL = SQL & " I2.ValueID1 as OBJ_INFO_01, I2.ValueDesc1 as Description1, "
            SQL = SQL & " I2.ValueID2 as OBJ_INFO_02, I2.ValueDesc2 as Description2, "
            SQL = SQL & " I2.ValueID3 as OBJ_INFO_03, I2.ValueDesc3 as Description3, "
            SQL = SQL & " I2.ValueID4 as OBJ_INFO_04, I2.ValueDesc4 as Description4, "
            SQL = SQL & " I2.ValueID5 as OBJ_INFO_05, I2.ValueDesc5 as Description5, "
            SQL = SQL & " I2.ValueID6 as OBJ_INFO_06, I2.ValueDesc6 as Description6, "
            SQL = SQL & " I2.ValueID7 as OBJ_INFO_07, I2.ValueDesc7 as Description7, "
            SQL = SQL & " I2.ValueID8 as OBJ_INFO_08, I2.ValueDesc8 as Description8, "
            SQL = SQL & " I2.ValueID9 as OBJ_INFO_09, I2.ValueDesc9 as Description9, "
            SQL = SQL & " I2.ValueID10 as OBJ_INFO_10, I2.ValueDesc10 as Description10, "
            SQL = SQL & " I2.ValueID11 as OBJ_INFO_11, I2.ValueDesc11 as Description11, "
            SQL = SQL & " I2.ValueID12 as OBJ_INFO_12, I2.ValueDesc12 as Description12, "
            SQL = SQL & " I2.ValueID13 as OBJ_INFO_13, I2.ValueDesc13 as Description13, "
            SQL = SQL & " I2.ValueID14 as OBJ_INFO_14, I2.ValueDesc14 as Description14, "
            SQL = SQL & " I2.ValueID15 as OBJ_INFO_15, I2.ValueDesc15 as Description15 "
            SQL = SQL & " From ICover I1"
            SQL = SQL & " Inner Join IINFO I2 on (I2.INO=I1.INO)"
            SQL = SQL & " WHERE 1<>1"

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet

            ' Verify Data
            For i = 0 To dsExcel.Tables(0).Rows.Count - 1
                If TOI <> tostr(GetItemValue("TOI", dsExcel, i)) Then
                    TOI = tostr(GetItemValue("TOI", dsExcel, i))
                    rsTOI = db.OpenRemoteSet("Select TOI.TOI, TOI.Description, TOI.Adjustable, TOI.DispModel, PCLASS.CATEGORY from TOI INNER JOIN RCOVER ON (RCOVER.ANO=" & ANO & ") INNER JOIN RATECOV ON (RATECOV.CODE=RCOVER.CODE) INNER JOIN PCLASS ON (TOI.TOI=PCLASS.TOI AND RATECOV.TOR=PCLASS.TOR) WHERE TOI.TOI='" & TOI & "'")
                    If rsTOI.Eof Then
                        StrMsg = StrMsg & "Line : " & i + 1 & ", Incorrect Type Of Interest (" & TOI & ") !" & vbCrLf
                    End If
                End If

                RefNo = tostr(GetItemValue("RefNo", dsExcel, i))
                Remark = tostr(GetItemValue("Remark", dsExcel, i))
                Category = tostr(GetItemValue("Category", dsExcel, i))
                PCTADJ = toval(GetItemValue("PCTADJ", dsExcel, i))
                Amount = toval(GetItemValue("Amount", dsExcel, i))
                Amount_AOE = toval(GetItemValue("Amount_AOE", dsExcel, i))
                SI = toval(GetItemValue("SI", dsExcel, i))

                If Currency <> tostr(GetItemValue("Currency", dsExcel, i)) Then
                    Currency = tostr(GetItemValue("Currency", dsExcel, i))
                    rs = db.OpenRemoteSet("Select Currency from Currency Where Currency.AllowedF=1 and Currency='" & Currency & "'")
                    If rs.Eof Then
                        StrMsg = StrMsg & "Line : " & i + 1 & ", Incorrect Currency (" & Currency & ") !" & vbCrLf
                    End If
                End If

                If Mode = 2 Then
                    rs = db.OpenRemoteSet("Select INO from ICover Where ANO=" & ANO & " and TOI='" & TOI & "' and RefNo='" & RefNo & "'")
                    If rs.Eof Then
                        StrMsg = StrMsg & "Line : " & i + 1 & ", No Interest Found, Type Of Interesst (" & TOI & "), Reference No (" & RefNo & ") !" & vbCrLf
                    End If
                End If

                If Remark = "" Then
                    Remark = rsTOI.fields("Description")
                End If
                If Category = "" Then
                    Category = rsTOI.fields("Category")
                End If

                If Not toval(rsTOI.fields("Adjustable")) Then
                    PCTADJ = 0
                End If

                If toval(rsTOI.fields("Adjustable")) AndAlso (PCTADJ > 100 Or PCTADJ <= 0) Then
                    StrMsg = StrMsg & "Line : " & i + 1 & ", Incorrect Adjustment Value Must Be Greater than 0 and Less than or Equal To 100 !" & vbCrLf
                End If

                If Category <> "M" AndAlso Category <> "A" Then
                    StrMsg = StrMsg & "Line : " & i + 1 & ", Incorrect Interest Category (Additional/Main) !" & vbCrLf
                End If

                If Amount < 0 Then
                    StrMsg = StrMsg & "Line : " & i + 1 & ", Incorrect Limit Of Liability !" & vbCrLf
                End If

                If Amount_AOE < 0 Then
                    StrMsg = StrMsg & "Line : " & i + 1 & ", Incorrect Max Liability !" & vbCrLf
                End If

                If SI < 0 Then
                    StrMsg = StrMsg & "Line : " & i + 1 & ", Incorrect Sum Insured !" & vbCrLf
                End If

                NewRow = Dataset.Tables(0).NewRow
                NewRow("TOI") = TOI
                NewRow("Currency") = Currency
                NewRow("Remark") = Remark
                NewRow("Category") = Category
                NewRow("PCTADJ") = PCTADJ
                NewRow("Amount") = Amount
                NewRow("Amount_AOE") = Amount_AOE
                NewRow("SI") = SI
                NewRow("TOPROF") = toval(GetItemValue("TOPROF", dsExcel, i))
                NewRow("RefNo") = tostr(GetItemValue("RefNo", dsExcel, i))
                NewRow("RefDate") = todate(GetItemValue("RefDate", dsExcel, i))
                NewRow("RefDueDate") = todate(GetItemValue("RefDueDate", dsExcel, i))
                NewRow("DocNo") = tostr(GetItemValue("DocNo", dsExcel, i))
                NewRow("DocDate") = todate(GetItemValue("DocDate", dsExcel, i))

                NewRow("OBJ_INFO_01") = tostr(GetItemValue("OBJ_INFO_01", dsExcel, i))
                NewRow("Description1") = tostr(GetItemValue("Description1", dsExcel, i))
                NewRow("OBJ_INFO_02") = tostr(GetItemValue("OBJ_INFO_02", dsExcel, i))
                NewRow("Description2") = tostr(GetItemValue("Description2", dsExcel, i))
                NewRow("OBJ_INFO_03") = tostr(GetItemValue("OBJ_INFO_03", dsExcel, i))
                NewRow("Description3") = tostr(GetItemValue("Description3", dsExcel, i))
                NewRow("OBJ_INFO_04") = tostr(GetItemValue("OBJ_INFO_04", dsExcel, i))
                NewRow("Description4") = tostr(GetItemValue("Description4", dsExcel, i))
                NewRow("OBJ_INFO_05") = tostr(GetItemValue("OBJ_INFO_05", dsExcel, i))
                NewRow("Description5") = tostr(GetItemValue("Description5", dsExcel, i))
                NewRow("OBJ_INFO_06") = tostr(GetItemValue("OBJ_INFO_06", dsExcel, i))
                NewRow("Description6") = tostr(GetItemValue("Description6", dsExcel, i))
                NewRow("OBJ_INFO_07") = tostr(GetItemValue("OBJ_INFO_07", dsExcel, i))
                NewRow("Description7") = tostr(GetItemValue("Description7", dsExcel, i))
                NewRow("OBJ_INFO_08") = tostr(GetItemValue("OBJ_INFO_08", dsExcel, i))
                NewRow("Description8") = tostr(GetItemValue("Description8", dsExcel, i))
                NewRow("OBJ_INFO_09") = tostr(GetItemValue("OBJ_INFO_09", dsExcel, i))
                NewRow("Description9") = tostr(GetItemValue("Description9", dsExcel, i))
                NewRow("OBJ_INFO_10") = tostr(GetItemValue("OBJ_INFO_10", dsExcel, i))
                NewRow("Description10") = tostr(GetItemValue("Description10", dsExcel, i))
                NewRow("OBJ_INFO_11") = tostr(GetItemValue("OBJ_INFO_11", dsExcel, i))
                NewRow("Description11") = tostr(GetItemValue("Description11", dsExcel, i))
                NewRow("OBJ_INFO_12") = tostr(GetItemValue("OBJ_INFO_12", dsExcel, i))
                NewRow("Description12") = tostr(GetItemValue("Description12", dsExcel, i))
                NewRow("OBJ_INFO_13") = tostr(GetItemValue("OBJ_INFO_13", dsExcel, i))
                NewRow("Description13") = tostr(GetItemValue("Description13", dsExcel, i))
                NewRow("OBJ_INFO_14") = tostr(GetItemValue("Description13", dsExcel, i))
                NewRow("Description14") = tostr(GetItemValue("Description14", dsExcel, i))
                NewRow("OBJ_INFO_15") = tostr(GetItemValue("OBJ_INFO_15", dsExcel, i))
                NewRow("Description15") = tostr(GetItemValue("Description15", dsExcel, i))
                Dataset.Tables(0).Rows.Add(NewRow)
            Next

            If StrMsg <> "" Then
                Throw New Exception(StrMsg)
            End If

            If Not SaveListInterest(dbUser, dbPassword, InParam, OutParam, Nothing, Nothing, Dataset, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            ' Set Output Parameter

            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveListInterest(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()
            ' Get Input Parameter
            ' getInParamValue(InParam, 0) = ANO
            ' getInParamValue(InParam, 1) = ' Mode : 0 (Reset All), 1 (Add), 2(Update)

            Dim ANO As Long = getInParamValue(InParam, 0)
            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Save_Multiple_Interest", InParam, Dataset) Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveListCoverage(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            ' getInParamValue(InParam, 0) = ANO
            ' getInParamValue(InParam, 1) = PSDate   ' Acceptance SDate
            ' getInParamValue(InParam, 2) = PEDate   ' Acceptance EDate
            ' getInParamValue(InParam, 3) = Voyage
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus, A.TOC, TOC.PTerm from Acceptance A Inner Join TOC on (TOC.TOC=A.TOC) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If
            ReDim Preserve InParam(5)
            InParam(4) = toval(rs.fields("PTerm"))
            InParam(5) = tostr(rs.fields("TOC"))

            Const MaxCols = 16
            Dim Bufferdata(Dataset.Tables(0).Rows.Count - 1, MaxCols - 1) As Object

            Dim i As Integer
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Bufferdata(i, 0) = tostr(Dataset.Tables(0).Rows(i)("PolicyYear"))
                Bufferdata(i, 1) = tostr(Dataset.Tables(0).Rows(i)("Code"))
                Bufferdata(i, 2) = tostr(Dataset.Tables(0).Rows(i)("Remark"))
                Bufferdata(i, 3) = toval(Dataset.Tables(0).Rows(i)("Rate"))
                Bufferdata(i, 4) = tostr(Dataset.Tables(0).Rows(i)("Unit"))
                Bufferdata(i, 5) = toval(Dataset.Tables(0).Rows(i)("Discount"))
                Bufferdata(i, 6) = toval(Dataset.Tables(0).Rows(i)("PCTLoss"))
                Bufferdata(i, 7) = tostr(Dataset.Tables(0).Rows(i)("Category"))
                Bufferdata(i, 8) = toval(Dataset.Tables(0).Rows(i)("PML"))
                Bufferdata(i, 9) = toval(Dataset.Tables(0).Rows(i)("PAddition"))
                Bufferdata(i, 10) = tostr(Dataset.Tables(0).Rows(i)("PCalc"))
                Bufferdata(i, 11) = toval(Dataset.Tables(0).Rows(i)("Loading"))
                Bufferdata(i, 12) = toval(Dataset.Tables(0).Rows(i)("TOPROF"))
                Bufferdata(i, 13) = IIf(tostr(Dataset.Tables(0).Rows(i)("NCDate")) = "", "", todate(Dataset.Tables(0).Rows(i)("NCDate")))
                Bufferdata(i, 14) = toval(Dataset.Tables(0).Rows(i)("PRATEPCT"))
                Bufferdata(i, 15) = toval(Dataset.Tables(0).Rows(i)("PProrata"))
            Next

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Save_Multiple_Coverage", InParam, TrimBuffer(Bufferdata, 1)) Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropCoverage(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Code As String
            Dim SDate As Date
            ANO = getInParamValue(InParam, 0)
            Code = getInParamValue(InParam, 1)
            SDate = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * From RCover Where ANO=" & ANO & " and Code='" & Code & "' and SDate='" & Format(SDate, "MM/DD/YYYY") & "'")
            If rs.Eof Then
                ErrMsg = "No Coverage Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            ' Delete First
            If Not db.Execute("Delete A from RateTabCla R, AccCla A Where A.OrderNo=R.OrderNo and A.ANO=" & ANO & " and R.Code='" & Code & "'") Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Delete A from RateTabCon R, AccCon A Where A.OrderNo=R.OrderNo and A.ANO=" & ANO & " and R.Code='" & Code & "'") Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Delete A from RateTabEx R, AccEx A Where A.OrderNo=R.OrderNo and A.ANO=" & ANO & " and R.Code='" & Code & "'") Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListDeductible(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet

            SQL = "Select Cast(0 as bit) as ProcessFlag, AccDed.*"
            SQL = SQL & " From AccDed"
            SQL = SQL & " WHERE ANO=" & ANO
            SQL = SQL & " Order By DNO"

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetDeductible(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim TOR As String
            Dim DCode As String
            Dim Section As String
            Dim PolicyYear As Integer

            ANO = getInParamValue(InParam, 0)
            TOR = getInParamValue(InParam, 1)
            DCode = getInParamValue(InParam, 2)
            Section = getInParamValue(InParam, 3)
            PolicyYear = getInParamValue(InParam, 4)

            Dim rs As RemoteSet

            rs = db.OpenQueryDef("QAccDed", adOpenKeyset, ANO, TOR, DCode, Section, PolicyYear)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveDeductible(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim PTOR As String
            Dim PDCode As String
            Dim PSection As String
            Dim PPolicyYear As Integer
            ' getInParamValue(InParam, 0) = ANO
            ' getInParamValue(InParam, 1) = TOR
            ' getInParamValue(InParam, 2) = DCode
            ' getInParamValue(InParam, 3) = Section
            ' getInParamValue(InParam, 4) = PolicyYear
            ' getInParamValue(InParam, 5) = SDate
            ' getInParamValue(InParam, 6) = EDate
            ANO = getInParamValue(InParam, 0)
            PTOR = getInParamValue(InParam, 7)
            PDCode = getInParamValue(InParam, 8)
            PSection = getInParamValue(InParam, 9)
            PPolicyYear = getInParamValue(InParam, 10)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            fbuffer(GetFieldIndex("TOR", findex)) = SetFieldValue(PTOR)
            fbuffer(GetFieldIndex("DCode", findex)) = SetFieldValue(PDCode)
            fbuffer(GetFieldIndex("Section", findex)) = SetFieldValue(PSection)
            fbuffer(GetFieldIndex("PolicyYear", findex)) = SetFieldValue(PPolicyYear)

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Save_Deductible", InParam, findex, fbuffer) Then Throw New Exception(Me.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropDeductible(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim TOR As String
            Dim DCode As String
            Dim Section As String
            Dim PolicyYear As Integer

            ANO = getInParamValue(InParam, 0)
            TOR = getInParamValue(InParam, 1)
            DCode = getInParamValue(InParam, 2)
            Section = getInParamValue(InParam, 3)
            PolicyYear = getInParamValue(InParam, 4)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * From AccDed Where ANO=" & ANO & " and TOR='" & TOR & "' and DCode='" & DCode & "' and Section='" & Section & "' and PolicyYear=" & PolicyYear)
            If rs.Eof Then
                ErrMsg = "No Deductible Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListBusinessSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            SQL = ""
            SQL = SQL & "SELECT Cast(0 as bit) as ProcessFlag, ACCBS.ID, ACCBS.BSID, PROFILE.NAME, ACCBS.BSType, BSType.Description as BSTypeDesc, ACCBS.FEE, ACCBS.DISCOUNT, ACCBS.UPFEE"
            SQL = SQL & " FROM ACCBS"
            SQL = SQL & " Left Join BSType on (BSType.BSType=ACCBS.BSTYPE)"
            SQL = SQL & " Inner Join Profile on (Profile.ID=AccBS.ID)"
            SQL = SQL & " WHERE ACCBS.ANO=" & ANO

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetBusinessSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim BSType As String
            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccBS Where ANO=" & ANO & " and ID='" & ID & "' and BSType='" & BSType & "'")
            findex = rs.fields_index
            fbuffer = rs.GetRecord



            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveBusinessSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim BSType As String
            Dim PID As String
            Dim PBSType As String
            Dim OverWritePayorF As Boolean
            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)
            ' Former Value
            PID = getInParamValue(InParam, 3)
            PBSType = getInParamValue(InParam, 4)
            OverWritePayorF = toval(getInParamValue(InParam, 5))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim PDate As Date
            Dim MO As String
            rs = db.OpenRemoteSet("Select A.AStatus,A.PDate, A.SDate, C.MO from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If
            PDate = todate(rs.fields("PDate"))
            MO = tostr(rs.fields("MO"))

            Dim EditMode As Integer

            Dim BSID As String = ""
            rs = db.OpenQueryDef("GetBusinessSourcePaidTo", adOpenKeyset, ID, PDate)
            If Not rs.Eof AndAlso rs.fields("ID") <> ID Then
                BSID = rs.fields("ID")
            End If
            If BSID = "" Then BSID = ID

            ' Check Business Source
            rs = db.OpenRemoteSet("Select BS.RestrictedF, BS.RestrictedRemarks, Profile.LicenseExpiry from FN_UserBusinessSource('" & MO & "','" & Format(PDate, "MM/DD/YYYY") & "') BS Inner Join Profile on (Profile.ID=BS.ID) Where BS.ID='" & ID & "'")
            If rs.Eof Then
                Throw New Exception("Not Allowed for Business Source Type Profile-" & ID & " (Agent/Broker) !")
            Else
                If toval(rs.fields("RestrictedF")) Then
                    If tostr(rs.fields("RestrictedRemarks")) = "License has expired" Then
                        If db.GetPrivileges("FORCELICENSEEXPIRY") And dbUserType <> "SUPERVISOR" Then
                            Throw New Exception("Please Check for Business Source Profile-" & ID & ", License has Expired on " & Format(rs.fields("LicenseExpiry"), "d") & " !")
                        End If
                    Else
                        Throw New Exception("Not Allowed for Restricted Profile-" & ID & " (Agent/Broker) !")
                    End If
                End If
            End If

            rs = db.OpenRemoteSet("Select * from AccBS Where ANO=" & ANO & " and ID='" & PID & "' and BSType='" & PBSType & "'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("ANO") = ANO
            rs.fields("ID") = ID
            rs.fields("BSType") = BSType
            If rs.fields("OBSType") = "" Then rs.fields("OBSType") = rs.fields("BSType")
            rs.fields("BSID") = BSID
            rs.fields("LUserID") = dbUser
            rs.fields("LDate") = Today.Date
            If Not rs.Update Then Throw New Exception(db.ErrDesc)
            If Not Create_Payor(ANO, ID, , , OverWritePayorF) Then Throw New Exception(db.ErrDesc)


            If Not rs.Eof Then
                If Not rs.Update Then Throw New Exception(db.ErrDesc)
            End If

            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and AType not in ('A','C') and LANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropBusinessSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim BSType As String
            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccBS Where ANO=" & ANO & " and ID='" & ID & "' and BSType='" & BSType & "'")
            If rs.Eof Then
                ErrMsg = "No Business Source Data !"
                Throw New Exception(ErrMsg)
            End If
            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and AType not in ('A','C') and LANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListAlternateBusinessSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            SQL = ""
            SQL = SQL & " SELECT Cast(0 as bit) as ProcessFlag, ACCBS.ID, ACCBS.BSID, PROFILE.NAME, ACCBS.BSType, BSType.Description as BSTypeDesc, ACCBS.FEE, ACCBS.DISCOUNT, ACCBS.UPFEE"
            SQL = SQL & " FROM AACCBS ACCBS"
            SQL = SQL & " Left Join BSType on (BSType.BSType=ACCBS.BSTYPE)"
            SQL = SQL & " Inner Join Profile on (Profile.ID=AccBS.ID)"
            SQL = SQL & " WHERE ACCBS.ANO=" & ANO

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetAlternateBusinessSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim BSType As String
            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AAccBS Where ANO=" & ANO & " and ID='" & ID & "' and BSType='" & BSType & "'")
            findex = rs.fields_index
            fbuffer = rs.GetRecord


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveAlternateBusinessSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim BSType As String
            Dim PID As String
            Dim PBSType As String
            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)
            ' Former Value
            PID = getInParamValue(InParam, 3)
            PBSType = getInParamValue(InParam, 4)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim PDate As Date
            Dim MO As String
            rs = db.OpenRemoteSet("Select A.AStatus, A.PDate, C.MO from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If
            PDate = todate(rs.fields("PDate"))
            MO = tostr(rs.fields("MO"))

            Dim EditMode As Integer

            Dim BSID As String = ID
            rs = db.OpenQueryDef("GetBusinessSourcePaidTo", adOpenKeyset, ID, PDate)
            If Not rs.Eof AndAlso rs.fields("ID") <> ID Then
                BSID = rs.fields("ID")
            End If

            ' Check Business Source
            rs = db.OpenRemoteSet("Select BS.RestrictedF, BS.RestrictedRemarks, Profile.LicenseExpiry from FN_UserBusinessSource('" & MO & "','" & Format(PDate, "MM/DD/YYYY") & "') BS Inner Join Profile on (Profile.ID=BS.ID) Where BS.ID='" & ID & "'")
            If rs.Eof Then
                Throw New Exception("Not Allowed for Business Source Type Profile-" & ID & " (Alternate Agent/Broker) !")
            Else
                If toval(rs.fields("RestrictedF")) Then
                    If tostr(rs.fields("RestrictedRemarks")) = "License has expired" Then
                        If db.GetPrivileges("FORCELICENSEEXPIRY") And dbUserType <> "SUPERVISOR" Then
                            Throw New Exception("Please Check for Business Source Profile-" & ID & ", License has Expired on " & Format(rs.fields("LicenseExpiry"), "d") & " !")
                        End If
                    Else
                        Throw New Exception("Not Allowed for Restricted Profile-" & ID & " (Alternate Agent/Broker) !")
                    End If
                End If
            End If

            rs = db.OpenRemoteSet("Select * from AAccBS Where ANO=" & ANO & " and ID='" & PID & "' and BSType='" & PBSType & "'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("ANO") = ANO
            rs.fields("ID") = ID
            rs.fields("BSType") = BSType
            If rs.fields("OBSType") = "" Then rs.fields("OBSType") = rs.fields("BSType")
            rs.fields("BSID") = BSID
            rs.fields("LUserID") = dbUser
            rs.fields("LDate") = Today.Date
            If Not rs.Update Then Throw New Exception(db.ErrDesc)



            If Not rs.Eof Then
                If Not rs.Update Then Throw New Exception(db.ErrDesc)
            End If

            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropAlternateBusinessSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim BSType As String
            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AAccBS Where ANO=" & ANO & " and ID='" & ID & "' and BSType='" & BSType & "'")
            If rs.Eof Then
                ErrMsg = "No Business Source Data !"
                Throw New Exception(ErrMsg)
            End If
            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListFeeDuty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            SQL = "SELECT Cast(0 as bit) as ProcessFlag, AccFD.*, coalesce(Profile.Name,'<Default Payor>') as Name"
            SQL = SQL & " FROM ACCFD Left Join Profile On (Profile.ID=AccFD.ID)  "
            SQL = SQL & " WHERE ANO=" & ANO

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetFeeDuty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FDID As String
            Dim ID As String
            Dim PolicyYear As Integer
            ANO = getInParamValue(InParam, 0)
            FDID = getInParamValue(InParam, 1)
            ID = getInParamValue(InParam, 2)
            PolicyYear = getInParamValue(InParam, 3)

            Dim rs As RemoteSet

            rs = db.OpenQueryDef("QAccFD", adOpenKeyset, ANO, FDID, ID, PolicyYear)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveFeeDuty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FDID As String
            Dim ID As String
            Dim PolicyYear As Integer
            Dim PFDID As String
            Dim PID As String
            Dim PPolicyYear As Integer
            ANO = getInParamValue(InParam, 0)
            FDID = getInParamValue(InParam, 1)
            ID = getInParamValue(InParam, 2)
            PolicyYear = getInParamValue(InParam, 3)
            ' Former Value
            PFDID = getInParamValue(InParam, 4)
            PID = getInParamValue(InParam, 5)
            PPolicyYear = getInParamValue(InParam, 6)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from AccFD Where ANO=" & ANO & " and FDID='" & PFDID & "' and ID='" & PID & "' and PolicyYear=" & PPolicyYear & "")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("ANO") = ANO
            rs.fields("FDID") = FDID
            rs.fields("ID") = ID
            rs.fields("PolicyYear") = PolicyYear
            rs.fields("LUserID") = dbUser
            rs.fields("LDate") = Today.Date
            rs.fields("AutoGeneratedF") = False
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropFeeDuty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FDID As String
            Dim ID As String
            Dim PolicyYear As Integer
            ANO = getInParamValue(InParam, 0)
            FDID = getInParamValue(InParam, 1)
            ID = getInParamValue(InParam, 2)
            PolicyYear = getInParamValue(InParam, 3)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccFD Where ANO=" & ANO & " and FDID='" & FDID & "' and ID='" & ID & "' and PolicyYear=" & PolicyYear & "")
            If rs.Eof Then
                ErrMsg = "No Fee Duty Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetComment(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            Dim SQL As String

            SQL = "Select A1.Notes, A1.AType, A1.EMethod, A1.Reason, A1.Matters, A1.SysUserID, A1.Comment, A1.Remarks, A1.Memo, A1.PDate, A1.PPDate, A1.ADate, A1.AStatus, A1.NCDate, "
            SQL = SQL & " A1.LargeF, A1.UniqueF, A1.RenewableF, A1.NoRefundF, A1.BackDatedF, A1.AdjustableF, A1.EEndorseF, A1.BulkF, C1.TBAF,"
            SQL = SQL & " C1.CDate, C1.Validation, C1.Insured, C1.MO, C1.FollowUp, C1.Comment as CComment, C1.Memo as CMemo, C1.CStatus,"
            SQL = SQL & " S1.ServiceNo, S1.SStatus, S1.ServiceInfo, S1.Remarks as SRemarks, S1.Memo as SMemo, S1.Comment as SComment, S1.InfoF, S1.SurveyF "
            SQL = SQL & " from Acceptance A1 inner join Cover C1 on (C1.CNO=A1.CNO) Left Join Service S1 on (S1.SType<>'Claim' and ((S1.PANO<>-1 and S1.PANO=A1.ANO) or (S1.ANO<>-1 and S1.ANO=A1.ANO))) "
            SQL = SQL & " Where A1.ANO=" & ANO
            SQL = SQL & " Order by S1.SNO Desc"

            rs = db.OpenRemoteSet(SQL)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveComment(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim PID As Long
            Dim Matters As String
            Dim Notes As String
            ANO = getInParamValue(InParam, 0)
            PID = getInParamValue(InParam, 1)   ' Refer to Current PNotes PID
            Matters = getInParamValue(InParam, 2)   ' Matters ID
            Notes = getInParamValue(InParam, 3)   ' Notes

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim CNO As Long
            Dim OANO As Long
            Dim EditMode As Integer
            Dim PMatters As String
            Dim PNotes As String

            ' Underwriting Section
            rs = db.OpenRemoteSet("Select * from Acceptance Where ANO=" & ANO)
            CNO = toval(rs.fields("CNO"))
            OANO = toval(rs.fields("OANO"))
            PMatters = tostr(rs.fields("Matters"))
            PNotes = tostr(rs.fields("Notes"))
            rs.Edit()

            If rs.fields("SysUserID") <> GetFieldValue(fbuffer(GetFieldIndex("SysUserID", findex))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "SysUserID", rs.fields("SysUserID"), GetFieldValue(fbuffer(GetFieldIndex("SysUserID", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("SysUserID") = GetFieldValue(fbuffer(GetFieldIndex("SysUserID", findex)))
            End If

            If rs.fields("Comment") <> GetFieldValue(fbuffer(GetFieldIndex("Comment", findex))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "Comment", rs.fields("Comment"), GetFieldValue(fbuffer(GetFieldIndex("Comment", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("Comment") = GetFieldValue(fbuffer(GetFieldIndex("Comment", findex)))
            End If

            If rs.fields("Remarks") <> GetFieldValue(fbuffer(GetFieldIndex("Remarks", findex))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "Remarks", rs.fields("Remarks"), GetFieldValue(fbuffer(GetFieldIndex("Remarks", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("Remarks") = GetFieldValue(fbuffer(GetFieldIndex("Remarks", findex)))
            End If

            If rs.fields("Memo") <> GetFieldValue(fbuffer(GetFieldIndex("Memo", findex))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "Memo", rs.fields("Memo"), GetFieldValue(fbuffer(GetFieldIndex("Memo", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("Memo") = GetFieldValue(fbuffer(GetFieldIndex("Memo", findex)))
            End If

            If rs.fields("Notes") <> GetFieldValue(fbuffer(GetFieldIndex("Notes", findex))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "Notes", rs.fields("Notes"), GetFieldValue(fbuffer(GetFieldIndex("Notes", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("Notes") = GetFieldValue(fbuffer(GetFieldIndex("Notes", findex)))
            End If

            If rs.fields("Reason") <> GetFieldValue(fbuffer(GetFieldIndex("Reason", findex))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "Reason", rs.fields("Reason"), GetFieldValue(fbuffer(GetFieldIndex("Reason", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("Reason") = GetFieldValue(fbuffer(GetFieldIndex("Reason", findex)))
            End If

            If rs.fields("Matters") <> GetFieldValue(fbuffer(GetFieldIndex("Matters", findex))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "Matters", rs.fields("Matters"), GetFieldValue(fbuffer(GetFieldIndex("Matters", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("Matters") = GetFieldValue(fbuffer(GetFieldIndex("Matters", findex)))
            End If

            If toval(rs.fields("LargeF")) <> toval(GetFieldValue(fbuffer(GetFieldIndex("LargeF", findex)))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "LargeF", rs.fields("LargeF"), GetFieldValue(fbuffer(GetFieldIndex("LargeF", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("LargeF") = GetFieldValue(fbuffer(GetFieldIndex("LargeF", findex)))
            End If

            If toval(rs.fields("UniqueF")) <> toval(GetFieldValue(fbuffer(GetFieldIndex("UniqueF", findex)))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "UniqueF", rs.fields("UniqueF"), GetFieldValue(fbuffer(GetFieldIndex("UniqueF", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("UniqueF") = GetFieldValue(fbuffer(GetFieldIndex("UniqueF", findex)))
            End If

            If toval(rs.fields("RenewableF")) <> toval(GetFieldValue(fbuffer(GetFieldIndex("RenewableF", findex)))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "RenewableF", rs.fields("RenewableF"), GetFieldValue(fbuffer(GetFieldIndex("RenewableF", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("RenewableF") = GetFieldValue(fbuffer(GetFieldIndex("RenewableF", findex)))
            End If

            If toval(rs.fields("NoRefundF")) <> toval(GetFieldValue(fbuffer(GetFieldIndex("NoRefundF", findex)))) Then
                If Not Save_Log_Object("Acceptance", "ANO", ANO, "RegNo", rs.fields("RegNo"), "PolicyNo", rs.fields("PolicyNo"), "CertificateNo", rs.fields("CertificateNo"), "NoRefundF", rs.fields("NoRefundF"), GetFieldValue(fbuffer(GetFieldIndex("NoRefundF", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("NoRefundF") = GetFieldValue(fbuffer(GetFieldIndex("NoRefundF", findex)))
            End If

            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            If Matters <> "" And Not (Matters = PMatters And Notes = PNotes) Then
                rs = db.OpenRemoteSet("Select * from PNOTES Where PID=" & PID)
                If rs.Eof Then
                    rs.AddNew()
                Else
                    rs.Edit()
                End If

                EditMode = rs.editmode

                If EditMode = AdEditAdd Then
                    rs.fields("ANO") = OANO
                End If
                rs.fields("SystemF") = False
                rs.fields("Date") = Today.Date
                rs.fields("Time") = Time()
                rs.fields("UserID") = dbUser
                rs.fields("Category") = "POLICY"
                rs.fields("Matters") = Matters
                rs.fields("Notes") = Notes
                If Not rs.Update Then Throw New Exception(rs.ErrDesc)

                If EditMode = AdEditAdd Then
                    rs = db.OpenRemoteSet("Select Top 1 PID from PNOTES Where ANO=" & OANO & " Order by PID Desc")
                    PID = toval(rs.fields("PID"))
                End If

            End If

            ' Marketing Section
            rs = db.OpenRemoteSet("Select * from Cover Where CNO=" & CNO)
            If rs.Eof Then
                ErrMsg = "No Quotation Data !"
                Throw New Exception(ErrMsg)
            End If
            rs.Edit()
            If toval(rs.fields("TBAF")) <> toval(GetFieldValue(fbuffer(GetFieldIndex("TBAF", findex)))) Then
                If Not Save_Log_Object("Cover", "CNO", CNO, "CoverNo", rs.fields("CoverNo"), "", "", "", "", "TBAF", rs.fields("TBAF"), GetFieldValue(fbuffer(GetFieldIndex("TBAF", findex))), "Update") Then Throw New Exception("Can't Save to Log !")
                rs.fields("TBAF") = GetFieldValue(fbuffer(GetFieldIndex("TBAF", findex)))
            End If
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = PID

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetBankersClause(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            Dim SQL As String

            SQL = "Select BankersID, BankersClause from Acceptance A1 "
            SQL = SQL & " Where A1.ANO=" & ANO

            rs = db.OpenRemoteSet(SQL)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveBankersClauses(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim BankersClause As String
            Dim NewBankersClause As String

            ' Underwriting Section
            rs = db.OpenRemoteSet("Select * from Acceptance Where ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If
            BankersClause = tostr(rs.fields("BankersClause"))
            NewBankersClause = GetFieldValue(fbuffer(GetFieldIndex("BankersClause", findex)))
            rs.Edit()
            rs.fields("BankersID") = GetFieldValue(fbuffer(GetFieldIndex("BankersID", findex)))
            rs.fields("BankersClause") = NewBankersClause
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            Dim SQL As String
            If Trim(BankersClause) <> "" Then
                SQL = "Update AccCla Set Description=replace(cast(Description as VARCHAR(MAX)),'" & BankersClause & "','<<BANKERS CLAUSE>>'), Description_1=replace(cast(Description_1 as VARCHAR(MAX)),'" & BankersClause & "','<<BANKERS CLAUSE>>') from AccCla Where ANO=" & ANO
                If Not db.Execute(SQL) Then Throw New Exception(db.ErrDesc)
            End If
            If Trim(NewBankersClause) <> "" Then
                SQL = "Update AccCla Set Description=replace(cast(Description as VARCHAR(MAX)),'<<BANKERS CLAUSE>>','" & NewBankersClause & "'), Description_1=replace(cast(Description_1 as VARCHAR(MAX)),'<<BANKERS CLAUSE>>','" & NewBankersClause & "') from AccCla Where ANO=" & ANO
                If Not db.Execute(SQL) Then Throw New Exception(db.ErrDesc)
            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetDeliveryAddress(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccDA Where ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveDeliveryAddress(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim AStatus As String
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            End If
            AStatus = rs.fields("AStatus")

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from AccDA Where ANO=" & ANO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                If AStatus <> "W" AndAlso EditMode = AdEditInProgress Then
                    If Not Save_Log_Object("ACCDA", "ANO", ANO, "", "", "", "", "", "", findex(i), rs.fields(findex(i)), GetFieldValue(fbuffer(i)), "Update") Then Throw New Exception("Can't Save to Log !")
                End If
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop

            rs.fields("ANO") = ANO
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetCorrespondenceAddress(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccCA Where ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveCorrespondenceAddress(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim AStatus As String
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            End If
            AStatus = rs.fields("AStatus")

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from AccCA Where ANO=" & ANO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                If AStatus <> "W" AndAlso EditMode = AdEditInProgress Then
                    If Not Save_Log_Object("ACCCA", "ANO", ANO, "", "", "", "", "", "", findex(i), rs.fields(findex(i)), GetFieldValue(fbuffer(i)), "Update") Then Throw New Exception("Can't Save to Log !")
                End If
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop

            rs.fields("ANO") = ANO
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetSignatory(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccSGN Where ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveSignatory(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from AccSGN Where ANO=" & ANO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            ' Check Status
            If EditMode = AdEditInProgress Then
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop

            rs.fields("ANO") = ANO
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter            
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetDeposit(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccDeposit Where ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveDeposit(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from AccDeposit Where ANO=" & ANO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            ' Check Status
            If EditMode = AdEditInProgress Then
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop

            rs.fields("ANO") = ANO
            rs.fields("Last_Opr") = dbUser
            rs.fields("Last_Update") = Today.Date
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    Private Function Check_Calculation(ByVal ANO As Long, ByRef ErrMsg As String) As Boolean
        Try
            Dim rs As RemoteSet

            rs = db.OpenRemoteSet("Select LANO, PStatus, AStatus from Acceptance Where ANO=" & ANO)
            If rs.fields("AStatus") = "W" Then
                If rs.fields("LANO") <> ANO Then
                    If tostr(rs.fields("PStatus")) = "I" Then
                        If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "Premium_Calc", ANO) Then
                            Throw New Exception(Me.ErrDesc)
                        End If
                    End If
                Else
                    If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "PREMIUM_CALC_ALL", ANO) Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPremiumCalculation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)
            If Not Check_Calculation(ANO, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If
            Dim rs As RemoteSet
            rs = db.OpenQueryDef("View05L1", adOpenKeyset, ANO)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetRiskCalculation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)
            If Not Check_Calculation(ANO, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If
            Dim rs As RemoteSet
            rs = db.OpenQueryDef("View05L2", adOpenKeyset, ANO)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetInterestCalculation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)
            If Not Check_Calculation(ANO, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If
            Dim rs As RemoteSet
            rs = db.OpenQueryDef("View05L3", adOpenKeyset, ANO)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPayor(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef DataSet As DataSet, ByRef DataSet2 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from Payor Where ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            rs = db.OpenRemoteSet("Select * from PayorDet Where ANO=" & ANO)
            DataSet = rs.GetDataSet

            rs = db.OpenRemoteSet("Select * from AccDiscount Where ANO=" & ANO)
            DataSet2 = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SavePayor(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal DataSet As DataSet, ByVal DataSet2 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim AllowPDueGTFDue As Boolean
            ANO = getInParamValue(InParam, 0)
            AllowPDueGTFDue = toval(getInParamValue(InParam, 1))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim LANO As Long
            Dim ADate As Date
            Dim Effective As Date
            Dim TOCMinLapseFacDue As Integer
            Dim A_PolicyNo As String
            Dim AType As String

            rs = db.OpenRemoteSet("Select A.AType, A.AStatus, A.PolicyNo, A.CertificateNo, A.LANO, A.PDate, A.ADate, TOC.MinLapseFacDue as TOCMinLapseFacDue from Acceptance A, TOC Where TOC.TOC=A.TOC and A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If
            AType = rs.fields("AType")
            LANO = toval(rs.fields("LANO"))
            ADate = todate(rs.fields("ADate"))
            Effective = todate(rs.fields("PDate"))
            TOCMinLapseFacDue = toval(rs.fields("TOCMinLapseFacDue"))
            If rs.fields("PolicyNo") <> "" And rs.fields("CertificateNo") <> "" Then
                A_PolicyNo = rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo")
            Else
                A_PolicyNo = ""
            End If

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from Payor Where ANO=" & ANO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            If AType <> "A" AndAlso AType <> "C" Then
                ' Update Previous Virtual Account Information
                If rs.fields("RefType") = "VIRTUAL" And rs.fields("AccountNo") <> "" And A_PolicyNo <> "" Then
                    If Not db.Execute("Update V1 Set AllocatedF=0, A_PolicyNo='' from VAccountD V1, VAccount V2 Where V1.AllocatedF=1 and V2.VID=V1.VID and V1.AccountNo='" & rs.fields("AccountNo") & "' and V2.Bank='" & rs.fields("Bank") & "' and V1.PayorID=''") Then Throw New Exception(db.ErrDesc)
                End If
            End If

            ' Check Status
            If EditMode = AdEditInProgress Then
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            If Not db.GetPrivileges("ALLOWPASSSCANPAN") Then
                If isContainRestrictedPAN(rs.fields("AccountNo")) Then
                    Throw New Exception("Contain Restricted PAN (Payment Account Number) in field AccountNo!")
                End If
                If isContainRestrictedPAN(rs.fields("REFNO")) Then
                    Throw New Exception("Contain Restricted PAN (Payment Account Number) in field REFNO!")
                End If
                If isContainRestrictedPAN(rs.fields("REMARKS")) Then
                    Throw New Exception("Contain Restricted PAN (Payment Account Number) in field REMARKS!")
                End If
            End If
            If EditMode = AdEditAdd Then
                rs.fields("ANO") = ANO
            End If
            If rs.fields("RefType") = "TRANSFER" Then
                rs.fields("Bank") = ""
                rs.fields("BankBranch") = ""
            ElseIf rs.fields("RefType") = "AUTODEBIT" Or rs.fields("RefType") = "CREDIT CARD" Or rs.fields("RefType") = "VIRTUAL" Then
            Else
                rs.fields("Bank") = ""
                rs.fields("BankBranch") = ""
                rs.fields("AccountNo") = ""
                rs.fields("BankAccount") = ""
                rs.fields("BankName") = ""
            End If
            rs.fields("DefaultPayorF") = False
            rs.fields("LUserID") = dbUser
            rs.fields("LDate") = Today.Date
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            If AType <> "A" AndAlso AType <> "C" Then
                If Not toval(rs.fields("SIScheduleF")) Then
                    If Not db.Execute("Delete from Payor_Schedule where ANO=" & ANO & "") Then Throw New Exception(db.ErrDesc)
                End If

                If Not AllowPDueGTFDue Then
                    If Not Compare_DueDate_Payor_and_Facultative(ANO, LANO, TOCMinLapseFacDue, ADate, Effective, "", ErrMsg) Then Throw New Exception(ErrMsg)
                End If

                ' Update Current Virtual Account Information
                If rs.fields("RefType") = "VIRTUAL" And rs.fields("AccountNo") <> "" And A_PolicyNo <> "" Then
                    If Not db.Execute("Update V1 Set AllocatedF=1, A_PolicyNo='" & A_PolicyNo & "' from VAccountD V1, VAccount V2 WHere V1.AllocatedF=0 and V2.VID=V1.VID and V1.AccountNo='" & rs.fields("AccountNo") & "' and V2.Bank='" & rs.fields("Bank") & "' and V1.PayorID=''") Then Throw New Exception(db.ErrDesc)
                End If

                ' Delete First
                If Not db.Execute("Delete from PayorDet Where ANO=" & ANO) Then Throw New Exception(db.ErrDesc)

                rs = db.OpenRemoteSet("Select * from PayorDet Where 1<>1")
                For i = 0 To DataSet.Tables(0).Rows.Count - 1
                    If toval(DataSet.Tables(0).Rows(i)("Period")) >= 0 And tostr(DataSet.Tables(0).Rows(i)("Code")) <> "" Then
                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("Period") = toval(DataSet.Tables(0).Rows(i)("Period"))
                        rs.fields("Code") = tostr(DataSet.Tables(0).Rows(i)("Code"))
                        rs.fields("DiscAmount") = toval(DataSet.Tables(0).Rows(i)("DiscAmount"))
                        rs.fields("Discount") = toval(DataSet.Tables(0).Rows(i)("Discount"))
                        rs.fields("NoGDiscountF") = toval(DataSet.Tables(0).Rows(i)("NoGDiscountF"))
                        If Not rs.UpdateRow Then Throw New Exception(db.ErrDesc)
                    End If
                Next
                If Not rs.Eof Then
                    If Not rs.Update Then Throw New Exception(db.ErrDesc)
                End If


                ' Delete First
                If Not db.Execute("Delete from AccDiscount Where ANO=" & ANO) Then Throw New Exception(db.ErrDesc)

                rs = db.OpenRemoteSet("Select * from AccDiscount Where 1<>1")
                For i = 0 To DataSet2.Tables(0).Rows.Count - 1
                    If tostr(DataSet2.Tables(0).Rows(i)("DType")) <> "" Then
                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("DType") = tostr(DataSet2.Tables(0).Rows(i)("DType"))
                        rs.fields("Discount") = toval(DataSet2.Tables(0).Rows(i)("Discount"))
                        rs.fields("DiscAmount") = toval(DataSet2.Tables(0).Rows(i)("DiscAmount"))
                        If Not rs.UpdateRow Then Throw New Exception(db.ErrDesc)
                    End If
                Next
                If Not rs.Eof Then
                    If Not rs.Update Then Throw New Exception(db.ErrDesc)
                End If

                If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and AType not in ('A','C') and LANO=" & ANO) Then Throw New Exception(db.ErrDesc)
                If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and AType not in ('A','C') and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)

            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetAlternatePayor(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef DataSet As DataSet, ByRef DataSet2 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from APayor Where ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            rs = db.OpenRemoteSet("Select * from APayorDet Where ANO=" & ANO)
            DataSet = rs.GetDataSet

            rs = db.OpenRemoteSet("Select * from AccDiscount Where ANO=" & ANO)
            DataSet2 = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveAlternatePayor(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal DataSet As DataSet, ByVal DataSet2 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim AllowPDueGTFDue As Boolean
            ANO = getInParamValue(InParam, 0)
            AllowPDueGTFDue = toval(getInParamValue(InParam, 1))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim LANO As Long
            Dim ADate As Date
            Dim Effective As Date
            Dim TOCMinLapseFacDue As Integer
            Dim A_PolicyNo As String
            rs = db.OpenRemoteSet("Select A.AStatus, A.PolicyNo, A.CertificateNo, A.LANO, A.PDate, A.ADate, TOC.MinLapseFacDue as TOCMinLapseFacDue from Acceptance A, TOC Where TOC.TOC=A.TOC and A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            LANO = toval(rs.fields("LANO"))
            Effective = todate(rs.fields("PDate"))
            ADate = todate(rs.fields("ADate"))
            TOCMinLapseFacDue = toval(rs.fields("TOCMinLapseFacDue"))
            If rs.fields("PolicyNo") <> "" And rs.fields("CertificateNo") <> "" Then
                A_PolicyNo = rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo")
            Else
                A_PolicyNo = ""
            End If

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from APayor Where ANO=" & ANO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            ' Update Previous Virtual Account Information
            If rs.fields("RefType") = "VIRTUAL" And rs.fields("AccountNo") <> "" And A_PolicyNo <> "" Then
                If Not db.Execute("Update V1 Set AllocatedF=0, A_PolicyNo='' from VAccountD V1, VAccount V2 Where V1.AllocatedF=1 and V2.VID=V1.VID and V1.AccountNo='" & rs.fields("AccountNo") & "' and V2.Bank='" & rs.fields("Bank") & "' and V1.PayorID=''") Then Throw New Exception(db.ErrDesc)
            End If

            ' Check Status
            If EditMode = AdEditInProgress Then
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            If Not db.GetPrivileges("ALLOWPASSSCANPAN") Then
                If isContainRestrictedPAN(rs.fields("AccountNo")) Then
                    Throw New Exception("Contain Restricted PAN (Payment Account Number) in field AccountNo!")
                End If
                If isContainRestrictedPAN(rs.fields("REFNO")) Then
                    Throw New Exception("Contain Restricted PAN (Payment Account Number) in field REFNO!")
                End If
                If isContainRestrictedPAN(rs.fields("REMARKS")) Then
                    Throw New Exception("Contain Restricted PAN (Payment Account Number) in field REMARKS!")
                End If
            End If
            If EditMode = AdEditAdd Then
                rs.fields("ANO") = ANO
            End If
            If rs.fields("RefType") = "TRANSFER" Then
                rs.fields("Bank") = ""
                rs.fields("BankBranch") = ""
            ElseIf rs.fields("RefType") = "AUTODEBIT" Or rs.fields("RefType") = "CREDIT CARD" Or rs.fields("RefType") = "VIRTUAL" Then
            Else
                rs.fields("Bank") = ""
                rs.fields("BankBranch") = ""
                rs.fields("AccountNo") = ""
                rs.fields("BankAccount") = ""
                rs.fields("BankName") = ""
            End If
            rs.fields("DefaultPayorF") = False
            rs.fields("LUserID") = dbUser
            rs.fields("LDate") = Today.Date
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            If Not toval(rs.fields("SIScheduleF")) Then
                If Not db.Execute("Delete from Payor_Schedule where ANO=" & ANO & "") Then Throw New Exception(db.ErrDesc)
            End If

            If Not AllowPDueGTFDue Then
                If Not Compare_DueDate_Payor_and_Facultative(ANO, LANO, TOCMinLapseFacDue, ADate, Effective, "", ErrMsg) Then Throw New Exception(ErrMsg)
            End If

            ' Update Current Virtual Account Information
            If rs.fields("RefType") = "VIRTUAL" And rs.fields("AccountNo") <> "" And A_PolicyNo <> "" Then
                If Not db.Execute("Update V1 Set AllocatedF=1, A_PolicyNo='" & A_PolicyNo & "' from VAccountD V1, VAccount V2 WHere V1.AllocatedF=0 and V2.VID=V1.VID and V1.AccountNo='" & rs.fields("AccountNo") & "' and V2.Bank='" & rs.fields("Bank") & "' and V1.PayorID=''") Then Throw New Exception(db.ErrDesc)
            End If

            ' Delete First
            If Not db.Execute("Delete from APayorDet Where ANO=" & ANO) Then Throw New Exception(db.ErrDesc)

            rs = db.OpenRemoteSet("Select * from APayorDet Where 1<>1")
            For i = 0 To DataSet.Tables(0).Rows.Count - 1
                If toval(DataSet.Tables(0).Rows(i)("Period")) >= 0 And tostr(DataSet.Tables(0).Rows(i)("Code")) <> "" Then
                    rs.AddNew()
                    rs.fields("ANO") = ANO
                    rs.fields("Period") = toval(DataSet.Tables(0).Rows(i)("Period"))
                    rs.fields("Code") = tostr(DataSet.Tables(0).Rows(i)("Code"))
                    rs.fields("DiscAmount") = toval(DataSet.Tables(0).Rows(i)("DiscAmount"))
                    rs.fields("Discount") = toval(DataSet.Tables(0).Rows(i)("Discount"))
                    rs.fields("NoGDiscountF") = toval(DataSet.Tables(0).Rows(i)("NoGDiscountF"))
                    If Not rs.UpdateRow Then Throw New Exception(db.ErrDesc)
                End If
            Next
            If Not rs.Eof Then
                If Not rs.Update Then Throw New Exception(db.ErrDesc)
            End If

            ' Delete First
            If Not db.Execute("Delete from AccDiscount Where ANO=" & ANO) Then Throw New Exception(db.ErrDesc)

            rs = db.OpenRemoteSet("Select * from AccDiscount Where 1<>1")
            For i = 0 To DataSet2.Tables(0).Rows.Count - 1
                If tostr(DataSet2.Tables(0).Rows(i)("DType")) <> "" Then
                    rs.AddNew()
                    rs.fields("ANO") = ANO
                    rs.fields("DType") = tostr(DataSet2.Tables(0).Rows(i)("DType"))
                    rs.fields("Discount") = toval(DataSet2.Tables(0).Rows(i)("Discount"))
                    rs.fields("DiscAmount") = toval(DataSet2.Tables(0).Rows(i)("DiscAmount"))
                    If Not rs.UpdateRow Then Throw New Exception(db.ErrDesc)
                End If
            Next
            If Not rs.Eof Then
                If Not rs.Update Then Throw New Exception(db.ErrDesc)
            End If

            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Certificate Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DropAlternatePayor(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from APayor Where ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "No Alternate Payor Data !"
                Throw New Exception(ErrMsg)
            End If
            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Delete from Payor_Schedule Where ANO=" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Certificate Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function Create_Payment_Schedule(ByVal ANO As Long, ByVal DecreaseF As Boolean, ByVal PremiumDecreaseF As Boolean, ByVal Term As Integer, ByVal IRate As Double, ByRef errmsg As String) As Boolean
        Try
            Dim PayNo As Integer
            Dim PayDate As Date
            Dim LastPayDate As Date

            Dim SI As Double

            Dim rs As RemoteSet

            Dim TotalPremium As Double
            Dim CPRefund As Double
            Dim CPFormula As String
            Dim PPeriod As Integer
            Dim Prorata As Double
            Dim TotalSI As Double

            errmsg = ""
            TotalSI = 0

            rs = db.OpenRemoteSet("Select P.NumCCY, P.Gross, P.Rate, Case When P.Description='Premium by Decreasing SI' Then P.OGross else P.Gross End as OGross, TSI from fn_PCalc(" & ANO & ") P")
            If toval(rs.fields("NumCCY")) = 1 Then
                TotalPremium = toval(rs.fields("Gross")) / toval(rs.fields("Rate"))
            Else
                TotalPremium = toval(rs.fields("Gross"))
            End If

            Dim TSI As Double
            Dim pRate As Double

            ' Annual Rate
            TSI = toval(rs.fields("TSI"))
            pRate = toval(rs.fields("OGross")) / TSI * 100

            Dim Premium As Double
            Dim LastSI As Double
            Dim EXCLUSIVEF As Boolean
            Dim P_STime As String
            Dim P_ETime As String
            Dim ExtRSDate As Date
            Dim ExtREDate As Date
            Dim ExtRPeriod As Integer
            Dim EMethod As String
            Dim SDate As Date
            Dim EDate As Date
            Dim PDate As Date
            Dim ExtPeriod As Integer
            Dim EXTPERIODINDAYSF As Boolean

            rs = db.OpenRemoteSet("Select Top 1 A.SDate, A.EDate, A.PDate, A.STime, A.ETime, A.ExclusiveF, A.ExtPeriod, A.EXTPERIODINDAYSF, A.ExtRPeriod, A.ExtRSDate, A.ExtREDate, A.EMethod, R.CPRefund, R.CPFormula, C.Prorata, C.PPeriod from RCover R, Acceptance A, Cover C Where C.CNO=A.CNO and R.ANO=A.ANO and R.Category='M' and A.ANO=" & ANO)
            SDate = todate(rs.fields("SDate"))
            EDate = todate(rs.fields("EDate"))
            PDate = todate(rs.fields("PDate"))
            ExtPeriod = toval(rs.fields("ExtPeriod"))
            EXTPERIODINDAYSF = toval(rs.fields("EXTPERIODINDAYSF"))
            PPeriod = toval(rs.fields("PPERIOD"))
            Prorata = toval(rs.fields("Prorata"))
            P_STime = tostr(rs.fields("STime"))
            P_ETime = tostr(rs.fields("ETime"))
            EXCLUSIVEF = toval(rs.fields("EXCLUSIVEF"))

            CPRefund = toval(rs.fields("CPRefund"))
            If CPRefund = 0 Then CPRefund = 100
            CPFormula = tostr(rs.fields("CPFormula"))
            ExtRPeriod = toval(rs.fields("ExtRPeriod"))
            ExtRSDate = todate(rs.fields("ExtRSDate"))
            ExtREDate = todate(rs.fields("ExtREDate"))
            EMethod = tostr(rs.fields("EMethod"))

            Dim Tenor1 As Integer
            Dim TenorR As Integer
            Dim Tenor2 As Integer ' Tenor1 + Extended Period
            Dim Tenor As Integer

            If Not EXCLUSIVEF And P_STime < P_ETime Then
                Tenor1 = Premium_Calc_Prorata(SDate, DateAdd(DateInterval.Day, 1, EDate), Prorata, 12, PPeriod, "") * 12 / Term
                Tenor2 = Premium_Calc_Prorata(SDate, DateAdd(DateInterval.Day, 1, IIf(EXTPERIODINDAYSF, DateAdd(DateInterval.Day, ExtPeriod, EDate), DateAdd(DateInterval.Month, ExtPeriod, EDate))), Prorata, 12, PPeriod, "") * 12 / Term
            Else
                Tenor1 = Premium_Calc_Prorata(SDate, EDate, Prorata, 12, PPeriod, "") * 12 / Term
                Tenor2 = Premium_Calc_Prorata(SDate, IIf(EXTPERIODINDAYSF, DateAdd(DateInterval.Day, ExtPeriod, EDate), DateAdd(DateInterval.Month, ExtPeriod, EDate)), Prorata, 12, PPeriod, "") * 12 / Term
            End If

            If EMethod = "S" Then
                If Not EXCLUSIVEF And P_STime < P_ETime Then
                    TenorR = Premium_Calc_Prorata(SDate, DateAdd(DateInterval.Day, 1, PDate), Prorata, 12, PPeriod, "") * 12 / Term
                Else
                    TenorR = Premium_Calc_Prorata(SDate, PDate, Prorata, 12, PPeriod, "") * 12 / Term
                End If
            Else
                TenorR = 0
            End If

            Tenor = Tenor1 - TenorR
            Premium = TotalPremium / Tenor

            If Not db.Execute("Delete from Payor_Schedule where ANO=" & ANO & "") Then Throw New Exception(db.ErrDesc)

            If DecreaseF Then
                Dim rsTerm As RemoteSet
                rsTerm = db.OpenRemoteSet("Select Top 1 A.TOPRO from TOP_RTERM T1, Acceptance A Where A.ANO=" & ANO & " and T1.TOPRO=A.TOPRO")
                If Not rsTerm.Eof Then
                    Dim TOPRO As String
                    TOPRO = rsTerm.fields("TOPRO")
                    rsTerm = db.OpenRemoteSet("Select T1.PTERM, T1.PUnit, T2.* from TOP_RTERM T1, TOP_RTERMDET T2 Where T2.TOPRO=T1.TOPRO and T2.RATE=T1.RATE AND T1.TOPRO='" & TOPRO & "' and T1.Rate='" & tovalSF(IRate) & "' Order by T2.[Month]")
                    If rsTerm.Eof Then
                        errmsg = "No Interest Rate (" & tovalSF(IRate) & ") Defined !"
                        Throw New Exception(errmsg)
                    End If
                    Dim PTerm As Integer
                    Dim PUnit As Integer
                    PTerm = toval(rsTerm.fields("PTerm"))
                    PUnit = toval(rsTerm.fields("PUnit"))
                    LastSI = 0

                    If PTerm = 0 Then 'Annually
                        SI = toval(rsTerm.fields(tostr(Term * Tenor1 / 12))) / PUnit * TSI
                    Else
                        SI = toval(rsTerm.fields(tostr(Term * Tenor1 / PTerm))) / PUnit * TSI
                    End If
                    If SI = 0 Then
                        errmsg = "No Unit Tenor (" & tostr(Tenor1) & ") Defined !"
                        Throw New Exception(errmsg)
                    End If

                    rs = db.OpenRemoteSet("Select * from Payor_Schedule where 1<>1")
                    PayNo = 1
                    Do While PayNo <= (Tenor2 - Tenor1)
                        PayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1), SDate)
                        LastPayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1 + 1), SDate)
                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("InstNo") = PayNo
                        rs.fields("InstDate") = PayDate
                        rs.fields("InstLDate") = LastPayDate
                        rs.fields("InstPCT") = 0
                        rs.fields("InstAmount") = 0
                        rs.fields("InstSI") = SI
                        rs.fields("InstSIPCT") = SI / TSI * 100
                        TotalSI = TotalSI + SI

                        If Not EXCLUSIVEF And P_STime < P_ETime Then
                            Prorata = Calc_Prorata(IIf(LastPayDate < DateAdd(DateInterval.Day, 1, EDate), LastPayDate, DateAdd(DateInterval.Day, 1, EDate)), DateAdd(DateInterval.Day, 1, EDate), SDate, DateAdd(DateInterval.Day, 1, EDate), PPeriod)
                        Else
                            Prorata = Calc_Prorata(IIf(LastPayDate < EDate, LastPayDate, EDate), EDate, SDate, EDate, PPeriod)
                        End If

                        If UCase(CPFormula) = "PRORATE SQ" Then
                            Prorata = CPRefund / 100 * Prorata ^ 2
                        Else
                            Prorata = CPRefund / 100 * Prorata
                        End If
                        rs.fields("CPRefund") = TotalPremium * Prorata
                        If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)
                        LastSI = SI
                        PayNo = PayNo + 1
                    Loop

                    Do While Not rsTerm.Eof And toval(rsTerm.fields("Month")) <= Tenor1
                        PayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1), SDate)
                        LastPayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1 + 1), SDate)
                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("InstNo") = PayNo
                        rs.fields("InstDate") = PayDate
                        rs.fields("InstLDate") = LastPayDate
                        rs.fields("InstPCT") = Premium / TotalPremium * 100
                        rs.fields("InstAmount") = Premium
                        If PTerm = 0 Then 'Annually
                            SI = toval(rsTerm.fields(tostr(RoundUp(Term * Tenor1 / 12, 1)))) / PUnit * TSI
                        Else
                            SI = toval(rsTerm.fields(tostr(RoundUp(Term * Tenor1 / PTerm, 1)))) / PUnit * TSI
                        End If
                        If SI = 0 Then SI = LastSI
                        rs.fields("InstSI") = SI
                        rs.fields("InstSIPCT") = SI / TSI * 100
                        TotalSI = TotalSI + SI

                        If Not EXCLUSIVEF And P_STime < P_ETime Then
                            Prorata = Calc_Prorata(IIf(LastPayDate < DateAdd(DateInterval.Day, 1, EDate), LastPayDate, DateAdd(DateInterval.Day, 1, EDate)), DateAdd(DateInterval.Day, 1, EDate), SDate, DateAdd(DateInterval.Day, 1, EDate), PPeriod)
                        Else
                            Prorata = Calc_Prorata(IIf(LastPayDate < EDate, LastPayDate, EDate), EDate, SDate, EDate, PPeriod)
                        End If

                        If UCase(CPFormula) = "PRORATE SQ" Then
                            Prorata = CPRefund / 100 * Prorata ^ 2
                        Else
                            Prorata = CPRefund / 100 * Prorata
                        End If
                        rs.fields("CPRefund") = TotalPremium * Prorata
                        If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)
                        LastSI = SI
                        PayNo = PayNo + 1

                        If Not (ExtRPeriod <> 0 And LastPayDate > ExtRSDate And LastPayDate <= ExtREDate) Then
                            rsTerm.MoveNext()
                        End If
                    Loop

                    Do While PayNo <= Tenor2
                        PayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1), SDate)
                        LastPayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1 + 1), SDate)
                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("InstNo") = PayNo
                        rs.fields("InstDate") = PayDate
                        rs.fields("InstLDate") = LastPayDate
                        rs.fields("InstPCT") = 0
                        rs.fields("InstAmount") = 0
                        SI = LastSI
                        rs.fields("InstSI") = SI
                        rs.fields("InstSIPCT") = SI / TSI * 100
                        TotalSI = TotalSI + SI

                        If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)
                        PayNo = PayNo + 1
                    Loop
                Else
                    SI = TSI
                    LastSI = 0

                    rs = db.OpenRemoteSet("Select * from Payor_Schedule where 1<>1")

                    PayNo = 1
                    Do While PayNo <= TenorR
                        PayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1), SDate)
                        LastPayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1 + 1), SDate)
                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("InstNo") = PayNo
                        rs.fields("InstDate") = PayDate
                        rs.fields("InstLDate") = LastPayDate
                        rs.fields("InstPCT") = 0
                        rs.fields("InstAmount") = 0
                        rs.fields("InstSI") = SI
                        rs.fields("InstSIPCT") = SI / TSI * 100
                        TotalSI = TotalSI + SI

                        If Not EXCLUSIVEF And P_STime < P_ETime Then
                            Prorata = Calc_Prorata(IIf(LastPayDate < DateAdd(DateInterval.Day, 1, EDate), LastPayDate, DateAdd(DateInterval.Day, 1, EDate)), DateAdd(DateInterval.Day, 1, EDate), SDate, DateAdd(DateInterval.Day, 1, EDate), PPeriod)
                        Else
                            Prorata = Calc_Prorata(IIf(LastPayDate < EDate, LastPayDate, EDate), EDate, SDate, EDate, PPeriod)
                        End If

                        If UCase(CPFormula) = "PRORATE SQ" Then
                            Prorata = CPRefund / 100 * Prorata ^ 2
                        Else
                            Prorata = CPRefund / 100 * Prorata
                        End If
                        rs.fields("CPRefund") = TotalPremium * Prorata
                        If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)
                        LastSI = SI
                        PayNo = PayNo + 1
                    Loop

                    Do While PayNo <= Tenor1
                        If PremiumDecreaseF Then
                            If ((PayNo - 1) \ 12) * 12 = (PayNo - 1) Then
                                TotalPremium = SI * pRate / 100
                                If (Tenor1 - (PayNo - 1)) >= 12 Then
                                    Premium = TotalPremium / 12
                                Else
                                    Premium = TotalPremium / (Tenor1 - (PayNo - 1))
                                End If
                            End If
                        End If

                        PayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1), SDate)
                        LastPayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1 + 1), SDate)

                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("InstNo") = PayNo
                        rs.fields("InstDate") = PayDate
                        rs.fields("InstLDate") = LastPayDate
                        rs.fields("InstPCT") = Premium / TotalPremium * 100
                        rs.fields("InstAmount") = Premium
                        rs.fields("InstSI") = SI
                        rs.fields("InstSIPCT") = SI / TSI * 100
                        TotalSI = TotalSI + SI
                        If Not EXCLUSIVEF And P_STime < P_ETime Then
                            Prorata = Calc_Prorata(IIf(LastPayDate < DateAdd(DateInterval.Day, 1, EDate), LastPayDate, DateAdd(DateInterval.Day, 1, EDate)), DateAdd(DateInterval.Day, 1, EDate), SDate, DateAdd(DateInterval.Day, 1, EDate), PPeriod)
                        Else
                            Prorata = Calc_Prorata(IIf(LastPayDate < EDate, LastPayDate, EDate), EDate, SDate, EDate, PPeriod)
                        End If
                        If UCase(CPFormula) = "PRORATE SQ" Then
                            Prorata = CPRefund / 100 * Prorata ^ 2
                        Else
                            Prorata = CPRefund / 100 * Prorata
                        End If
                        rs.fields("CPRefund") = TotalPremium * Prorata
                        If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)

                        LastSI = SI
                        SI = SI - (-Pmt(IRate / 100 / 12, Tenor, TSI) - IRate / 100 / 12 * SI)

                        PayNo = PayNo + 1
                    Loop

                    Do While PayNo <= Tenor2
                        PayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1), SDate)
                        LastPayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1 + 1), SDate)
                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("InstNo") = PayNo
                        rs.fields("InstDate") = PayDate
                        rs.fields("InstLDate") = LastPayDate
                        rs.fields("InstPCT") = 0
                        rs.fields("InstAmount") = 0
                        rs.fields("InstSI") = 0
                        rs.fields("InstSIPCT") = 0
                        If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)
                        PayNo = PayNo + 1
                    Loop

                End If

            Else
                SI = TSI
                LastSI = 0

                rs = db.OpenRemoteSet("Select * from Payor_Schedule where 1<>1")

                PayNo = 1
                Do While PayNo <= TenorR
                    PayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1), SDate)
                    LastPayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1 + 1), SDate)
                    rs.AddNew()
                    rs.fields("ANO") = ANO
                    rs.fields("InstNo") = PayNo
                    rs.fields("InstDate") = PayDate
                    rs.fields("InstLDate") = LastPayDate
                    rs.fields("InstPCT") = 0
                    rs.fields("InstAmount") = 0
                    rs.fields("InstSI") = SI
                    rs.fields("InstSIPCT") = SI / TSI * 100
                    TotalSI = TotalSI + SI

                    If Not EXCLUSIVEF And P_STime < P_ETime Then
                        Prorata = Calc_Prorata(IIf(LastPayDate < DateAdd(DateInterval.Day, 1, EDate), LastPayDate, DateAdd(DateInterval.Day, 1, EDate)), DateAdd(DateInterval.Day, 1, EDate), SDate, DateAdd(DateInterval.Day, 1, EDate), PPeriod)
                    Else
                        Prorata = Calc_Prorata(IIf(LastPayDate < EDate, LastPayDate, EDate), EDate, SDate, EDate, PPeriod)
                    End If

                    If UCase(CPFormula) = "PRORATE SQ" Then
                        Prorata = CPRefund / 100 * Prorata ^ 2
                    Else
                        Prorata = CPRefund / 100 * Prorata
                    End If
                    rs.fields("CPRefund") = TotalPremium * Prorata
                    If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)
                    LastSI = SI
                    PayNo = PayNo + 1
                Loop

                Do While PayNo <= Tenor1
                    PayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1), SDate)
                    LastPayDate = DateAdd(DateInterval.Month, Term * (PayNo - 1 + 1), SDate)

                    rs.AddNew()
                    rs.fields("ANO") = ANO
                    rs.fields("InstNo") = PayNo
                    rs.fields("InstDate") = PayDate
                    rs.fields("InstLDate") = LastPayDate
                    rs.fields("InstPCT") = Premium / TotalPremium * 100
                    rs.fields("InstAmount") = Premium
                    rs.fields("InstSI") = SI
                    rs.fields("InstSIPCT") = SI / TSI * 100
                    TotalSI = TotalSI + SI
                    If Not EXCLUSIVEF And P_STime < P_ETime Then
                        Prorata = Calc_Prorata(IIf(LastPayDate < DateAdd(DateInterval.Day, 1, EDate), LastPayDate, DateAdd(DateInterval.Day, 1, EDate)), DateAdd(DateInterval.Day, 1, EDate), SDate, DateAdd(DateInterval.Day, 1, EDate), PPeriod)
                    Else
                        Prorata = Calc_Prorata(IIf(LastPayDate < EDate, LastPayDate, EDate), EDate, SDate, EDate, PPeriod)
                    End If
                    If UCase(CPFormula) = "PRORATE SQ" Then
                        Prorata = CPRefund / 100 * Prorata ^ 2
                    Else
                        Prorata = CPRefund / 100 * Prorata
                    End If
                    rs.fields("CPRefund") = TotalPremium * Prorata
                    If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)
                    LastSI = SI
                    PayNo = PayNo + 1
                Loop
            End If

            rs.MoveFirst()
            If Not rs.Eof Then If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            Dim eFactor As Double
            eFactor = 0
            TotalPremium = 0
            rs.MoveFirst()
            Do While Not rs.Eof
                rs.Edit()
                eFactor = eFactor + toval(rs.fields("InstSI")) / TotalSI
                rs.fields("eFactor") = eFactor

                If PremiumDecreaseF Then
                    TotalPremium = TotalPremium + rs.fields("InstAmount")
                End If

                If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)
                rs.MoveNext()
            Loop
            rs.MoveFirst()
            If Not rs.Eof Then If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            If PremiumDecreaseF Then
                If Not db.Execute("Update PCalc Set Gross=" & tovalSF(TotalPremium) & "/Rate, Description='Premium by Decreasing SI' from PCalc Where ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            End If
            Return True
        Catch ex As Exception
            errmsg = ex.Message
            Return False
        End Try

    End Function
    Private Function Calc_Prorata(ByVal PSDate As Date, ByVal PEDate As Date, ByVal SDate As Date, ByVal EDate As Date, ByVal PPeriod As Integer) As Double
        Dim Denominator As Double
        Dim Nominator As Double
        If PPeriod = 3 Then       ' Half Monthly Prorate
            Nominator = DateDiff(DateInterval.Month, PSDate, PEDate)
            If PSDate <> SDate Or PEDate <> EDate Then
                If PSDate.Day > PEDate.Day Then
                    Nominator = Nominator - 1
                    If Nominator < 0 Then Nominator = 0
                End If
            End If
            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Month, Nominator, PSDate), PEDate) > 15 Then
                Nominator = Nominator + 1
            End If
            Denominator = Premium_Calc_Prorata(SDate, EDate, 0, 1, PPeriod, "")
        ElseIf PPeriod = 2 Then     ' Monthly Prorate
            Nominator = DateDiff(DateInterval.Month, PSDate, PEDate)
            If PSDate <> SDate Or PEDate <> EDate Then
                If PSDate.Day > PEDate.Day Then
                    Nominator = Nominator - 1
                    If Nominator < 0 Then Nominator = 0
                End If
            End If
            Denominator = DateDiff(DateInterval.Month, SDate, EDate)
        ElseIf PSDate.Day = PEDate.Day And PSDate.Day = EDate.Day And SDate.Day = EDate.Day And PPeriod <> 1 And DateDiff(DateInterval.Month, SDate, EDate) <> 0 Then ' and not Daily Basis
            Nominator = DateDiff(DateInterval.Month, PSDate, PEDate)
            Denominator = DateDiff(DateInterval.Month, SDate, EDate)
        Else
            Nominator = DateDiff(DateInterval.Day, PSDate, PEDate)
            Denominator = DateDiff(DateInterval.Day, SDate, EDate)
        End If
        If Nominator <> 0 And Denominator <> 0 Then
            Calc_Prorata = Nominator / Denominator
        Else
            Calc_Prorata = 0
        End If
    End Function

    Private Function Premium_Calc_Prorata(ByVal PSDate As Date, ByVal PEDate As Date, ByVal DaysProrata As Integer, ByVal MonthsProrata As Integer, ByVal PPeriod As Integer, ByRef sProrata As String) As Double
        Dim Prorata As Double
        If PPeriod = 3 Then     ' Half Monthly Prorate
            Prorata = DateDiff(DateInterval.Month, PSDate, PEDate)
            If PSDate.Day > PEDate.Day Then
                Prorata = Prorata - 1
                If Prorata < 0 Then Prorata = 0
            End If
            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Month, Prorata, PSDate), PEDate) >= 15 Then
                Prorata = Prorata + 1
            End If
            sProrata = Format(Prorata, "#0") & "/" & MonthsProrata
            Prorata = Prorata / MonthsProrata
        ElseIf PPeriod = 2 Then     ' Monthly Prorate
            Prorata = DateDiff(DateInterval.Month, PSDate, PEDate)
            If DateAdd(DateInterval.Month, Prorata, PSDate) < PEDate AndAlso Not (Last_date_of_period(PSDate, 1) = PSDate) Then
                Prorata = Prorata + 1
            End If
            sProrata = Format(Prorata, "#0") & "/" & MonthsProrata
            Prorata = Prorata / MonthsProrata
        ElseIf DaysProrata <> 0 And (PSDate.Day <> PEDate.Day Or PPeriod = 1) Then
            sProrata = Format(DateDiff(DateInterval.Day, PSDate, PEDate), "#0") & "/" & DaysProrata
            Prorata = DateDiff(DateInterval.Day, PSDate, PEDate) / DaysProrata
        Else
            sProrata = Format(DateDiff(DateInterval.Month, PSDate, PEDate), "#0") & "/" & MonthsProrata
            Prorata = DateDiff(DateInterval.Month, PSDate, PEDate) / MonthsProrata
        End If
        Premium_Calc_Prorata = Prorata
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function CreatePayorSchedule(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef DataSet As DataSet, ByRef ErrMsg As String) As Boolean
        Try

            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long = getInParamValue(InParam, 0)
            Dim DecreaseF As Boolean = toval(getInParamValue(InParam, 1))
            Dim Term As Integer = toval(getInParamValue(InParam, 2))
            Dim IRate As Double = toval(getInParamValue(InParam, 3))
            Dim DecreasePremiumF As Boolean = toval(getInParamValue(InParam, 4))

            If Not Create_Payment_Schedule(ANO, DecreaseF, DecreasePremiumF, Term, IRate, ErrMsg) Then
                Throw New Exception("Create Payment Schedule Failed ! " & ErrMsg)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from Payor_Schedule where ANO=" & ANO)
            DataSet = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPayorSchedule(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef DataSet As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from Payor_Schedule where ANO=" & ANO)
            DataSet = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function CalculatePolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean

        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ResetF As Boolean
            ANO = getInParamValue(InParam, 0)
            ResetF = getInParamValue(InParam, 1)


            If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "PREMIUM_CALC", ANO, , ResetF) Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPremium(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Currency As String
            ANO = getInParamValue(InParam, 0)
            Currency = getInParamValue(InParam, 1)


            Dim rs As RemoteSet

            rs = db.OpenRemoteSet("Select * from PCalc Where ANO=" & ANO & " and Currency='" & Currency & "'")
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SavePremium(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Currency As String
            ANO = getInParamValue(InParam, 0)
            Currency = getInParamValue(InParam, 1)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim LANO As Long
            rs = db.OpenRemoteSet("Select A.AStatus, A.LANO from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            LANO = rs.fields("LANO")

            rs = db.OpenRemoteSet("Select * from PCalc Where ANO=" & ANO & " and Currency='" & Currency & "'")
            If rs.Eof Then
                ErrMsg = "No Premium Data !"
                Throw New Exception(ErrMsg)
            End If
            rs.Edit()

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("CRate") = toval(rs.fields("Gross")) / toval(rs.fields("TSI")) * 100

            If toval(rs.fields("Gross")) < 0 Then
                Throw New Exception("Incorrect Gross Premium , can't be less than 0 !")
            End If
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            Dim DecreaseF As Boolean = False
            Dim DecreasePremiumF As Boolean = False
            Dim AutoInstallmentPolicyYearF As Boolean = False
            Dim SIScheduleF As Boolean = False
            Dim IRate As Double = 0
            Dim Term As Integer = 0

            If ANO <> LANO Then
                If LANO <> -1 Then
                    rs = db.OpenRemoteSet("Select AutoInstallmentPolicyYearF, SIScheduleF, DecreaseF, DecreasePremiumF, IRate, Period from APayor Where ANO=" & ANO)
                Else
                    rs = db.OpenRemoteSet("Select AutoInstallmentPolicyYearF, SIScheduleF, DecreaseF, DecreasePremiumF, IRate, Period from Payor Where ANO=" & ANO)
                End If
                AutoInstallmentPolicyYearF = toval(rs.fields("AutoInstallmentPolicyYearF"))
                SIScheduleF = toval(rs.fields("SIScheduleF"))
                DecreaseF = toval(rs.fields("DecreaseF"))
                DecreasePremiumF = toval(rs.fields("DecreasePremiumF"))
                IRate = toval(rs.fields("IRate"))
                Term = toval(rs.fields("Period"))
            End If

            If SIScheduleF Then
                If Not CreatePayorSchedule(ANO, DecreaseF, DecreasePremiumF, Term, IRate) Then
                    StrMsg = "Can't Create Payor Schedule !" & StrMsg
                    Throw New Exception(StrMsg)
                End If
            End If
            If AutoInstallmentPolicyYearF Then
                If Not CreatePayorAutoInstallment(ANO, LANO) Then
                    StrMsg = "Can't Create Payor Auto Installment !" & StrMsg
                    Throw New Exception(StrMsg)
                End If
            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function CreatePayorAutoInstallment(ByVal ANO As Long, ByVal LANO As Long) As Boolean
        Try
            Dim rs As RemoteSet
            If LANO <> -1 Then
                rs = db.OpenRemoteSet("Select * from APayor Where ANO=" & ANO)
            Else
                rs = db.OpenRemoteSet("Select * from Payor Where ANO=" & ANO)
            End If

            Dim SQL As String
            Dim rsI As RemoteSet

            SQL = "Select R1.Date, R1.Gross/R2.Gross*100 as PCT"
            SQL = SQL & " from Acceptance A "
            SQL = SQL & " Cross Apply ("
            SQL = SQL & " Select dbo.fn_Get_UWYear(A.Sdate, A.EDate, R.SDATE) as UWYear, min(R.SDate) as Date, sum(R.Gross*P.Rate) as Gross"
            SQL = SQL & " from RCalc R, PCalc P"
            SQL = SQL & " Where R.ANO = A.ANO And P.ANO = R.ANO And P.Currency = R.Currency"
            SQL = SQL & " Group by dbo.fn_Get_UWYear(A.Sdate, A.EDate, R.SDATE)"
            SQL = SQL & " ) R1"
            SQL = SQL & " Cross Apply ("
            SQL = SQL & " Select sum(R.Gross*P.Rate) as Gross"
            SQL = SQL & " from RCalc R, PCalc P"
            SQL = SQL & " Where R.ANO = A.ANO And P.ANO = R.ANO And P.Currency = R.Currency"
            SQL = SQL & " ) R2"
            SQL = SQL & " Where A.ANO=" & ANO
            SQL = SQL & " Order by R1.Date"
            rsI = db.OpenRemoteSet(SQL)
            If Not rsI.Eof Then
                rs.Edit()
                rs.fields("Period") = 12  ' Annual
                rs.fields("Installment") = rsI.RecordCount
                For i = 1 To rsI.RecordCount
                    rs.fields("IDate_" & i) = rsI.fields("Date")
                    rs.fields("IPCT_" & i) = rsI.fields("PCT")
                    rsI.MoveNext()
                Next i
                If Not rs.Update Then Throw New Exception(rs.ErrDesc)
            End If
            Return True
        Catch ex As Exception
            StrMsg = ex.Message
            Return False
        End Try

    End Function
    Private Function CreatePayorSchedule(ByVal ANO As Long, ByVal DecreaseF As Boolean, ByVal PremiumDecreaseF As Boolean, ByVal Term As Integer, ByVal IRate As Double) As Boolean
        Try
            Dim ErrMsg As String
            If Not Create_Payment_Schedule(ANO, DecreaseF, PremiumDecreaseF, Term, IRate, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If
            Return True
        Catch ex As Exception
            StrMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function UndoInforcePolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Dim ANO As Long
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim LANO As Long


            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)


            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not PreUndoInforceConditionsOK(ANO, LANO, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            If ANO = LANO Then
                If Not Me.Direct_Call_Process_Execute("PROCESS.NETV3.PROCLIB", "REOPENMASTER", ANO) Then
                    Throw New Exception(Me.ErrDesc)
                End If
            Else
                If Not Me.Direct_Call_Process_Execute("PROCESS.NETV3.PROCLIB", "REOPEN", ANO) Then
                    Throw New Exception(Me.ErrDesc)
                End If
            End If

            ' Set Output Parameter

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message            
            db.RollBack()
            If db.CheckLockedObject("Policy" & ANO) Then db.ReleaseLockedObject("Policy" & ANO)
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ReviewPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Dim ANO As Long
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim LANO As Long
            Dim ADate As Date
            Dim PDate As Date
            Dim PolicyNo As String
            Dim BackDatedF As Boolean

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)
            ADate = getInParamValue(InParam, 2)  ' Inforce Date
            PDate = getInParamValue(InParam, 3)  ' Effective Date
            PolicyNo = getInParamValue(InParam, 4)
            BackDatedF = getInParamValue(InParam, 5)

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not PreInforceConditionsOK(ANO, ADate, PolicyNo, BackDatedF, False, False, True, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            If ANO = LANO Then
                If Not Me.Direct_Call_Process_Execute("PROCESS.NETV3.PROCLIB", "TEMPORARY_INFORCEMASTER", ANO, dbUser, BackDatedF) Then
                    Throw New Exception(Me.ErrDesc)
                End If
            Else
                If Not Me.Direct_Call_Process_Execute("PROCESS.NETV3.PROCLIB", "TEMPORARY_INFORCE", ANO, , , dbUser, BackDatedF) Then
                    Throw New Exception(Me.ErrDesc)
                End If
            End If

            ' Set Output Parameter

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            If db.CheckLockedObject("Policy" & ANO) Then db.ReleaseLockedObject("Policy" & ANO)
            CloseConnection()
            Return False
        End Try
    End Function

    Private Function Insert_Auto_Receipt_Payment(ByVal ANO As Long, ByVal Voucher As String, ByVal Account As String, ByVal Currency As String) As Boolean
        On Error GoTo errmsg
        Dim SQL As String
        Dim rs As RemoteSet

        SQL = ""

        SQL = SQL & " Declare @P_ANO as Int"
        SQL = SQL & " Declare @P_Currency as varchar(10)"
        SQL = SQL & " Select @P_ANO=" & ANO
        SQL = SQL & " Select @P_Currency='" & Currency & "'"

        SQL = SQL & " Select Top 1 nVoucher.Voucher, nVoucher.Date, "
        SQL = SQL & " case "
        SQL = SQL & " When nVoucher.INSTALLMENT<>0 then"
        SQL = SQL & "    nVoucher.AmountDue_1 - nVoucher.Payment_CC - nVoucher.Payment_OC"
        SQL = SQL & " else"
        SQL = SQL & "    nVoucher.Nominal_CC - nVoucher.Diff_CC - nVoucher.Payment_CC - nVoucher.Payment_OC"
        SQL = SQL & " End"
        SQL = SQL & " as Outstanding"
        SQL = SQL & " from Acceptance A1, AccAss A2, Admlink A3, nVoucher"
        SQL = SQL & " Where A2.ANO=A1.ANO and A3.AdmNo=A2.AdmNo and A3.Type in ('DI','IC','IR','IT','IX') and nVoucher.Voucher=A3.Voucher "
        SQL = SQL & " and A1.ANO=@P_ANO and nVoucher.Currency=@P_Currency"
        SQL = SQL & " and "
        SQL = SQL & " ("
        SQL = SQL & " case "
        SQL = SQL & " When nVoucher.INSTALLMENT<>0 then"
        SQL = SQL & "    nVoucher.AmountDue_1 - nVoucher.Payment_CC - nVoucher.Payment_OC"
        SQL = SQL & " else"
        SQL = SQL & "    nVoucher.Nominal_CC - nVoucher.Diff_CC - nVoucher.Payment_CC - nVoucher.Payment_OC"
        SQL = SQL & " End"
        SQL = SQL & " )>0"

        rs = db.OpenRemoteSet(SQL)

        If Not rs.Eof Then
            Dim nVoucher As String = rs.fields("Voucher")
            Dim Payment As Double = rs.fields("Outstanding")
            Dim Payment_Date As Date = rs.fields("Date")

            rs = db.OpenRemoteSet("Select * from RPBucket Where Voucher='" & Voucher & "' and nVoucher='" & nVoucher & "' and Status='R'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            rs.fields("Voucher") = Voucher
            rs.fields("Account") = Account
            rs.fields("NVoucher") = nVoucher
            rs.fields("Payment") = Payment
            rs.fields("Payment_Date") = Payment_Date
            rs.fields("Rate") = 1
            rs.fields("NetPaidF") = False
            rs.fields("Status") = "R"
            rs.fields("Last_Opr") = dbUser
            rs.fields("Last_Date") = Today.Date
            If Not rs.Update Then GoTo errmsg
        End If

        Insert_Auto_Receipt_Payment = True
        Exit Function
errmsg:
        Insert_Auto_Receipt_Payment = False
        Exit Function
    End Function
    Private Function Insert_Refund_Auto_Payment(ByVal ANO As Long, ByVal ProcName As String, ByVal ComAcc As String, ByVal RefType As String, ByVal RefNo As String, ByVal VType As String, ByVal CT As String, ByVal Remarks As String) As Boolean
        On Error GoTo errmsg
        Dim SQL As String
        Dim rs As RemoteSet
        Dim rsBucket As RemoteSet
        SQL = ""
        SQL = SQL & " Select A_PolicyNo, Voucher, Currency, Max(Date) as Date, sum(Outstanding) as Outstanding"
        SQL = SQL & " from"
        SQL = SQL & " ("
        SQL = SQL & " Select Distinct A3.A_PolicyNo, V.Date, V.Voucher, V.Currency, Case When V.DebtorF=1 Then (V.Nominal_CC-V.Diff_CC-V.Payment_CC) Else -(V.Nominal_CC-V.Diff_CC-V.Payment_CC) End as Outstanding"
        SQL = SQL & " from Acceptance A1, AccAss A2, Admlink A3, nVoucher V"
        SQL = SQL & " Where A1.AType='C' and A1.ANO = " & ANO & " And A2.ANO = A1.ANO And A3.AdmNo = A2.AdmNo And V.Voucher = A3.Voucher"
        SQL = SQL & " and A3.Type in ('DI','IC','IR','IT','IX') "
        SQL = SQL & " ) R"
        SQL = SQL & " Group by A_PolicyNo, Voucher, Currency"

        rs = db.OpenRemoteSet(SQL)
        Do While Not rs.Eof
            rsBucket = db.OpenRemoteSet("Select * from ARBucket Where A_PolicyNo='" & rs.fields("A_PolicyNo") & "' and Voucher='" & rs.fields("Voucher") & "' and Subject='Premium' and Currency='" & rs.fields("Currency") & "' and Status='R'")
            If rsBucket.Eof Then
                rsBucket.AddNew()
            Else
                rsBucket.Edit()
            End If
            rsBucket.fields("A_PolicyNo") = rs.fields("A_PolicyNo")
            rsBucket.fields("Voucher") = rs.fields("Voucher")
            rsBucket.fields("Subject") = "Premium"
            rsBucket.fields("Currency") = rs.fields("Currency")
            rsBucket.fields("AmountPaid") = toval(rs.fields("Outstanding"))
            rsBucket.fields("RefDate") = Today.Date
            rsBucket.fields("Status") = "R"
            rsBucket.fields("Rate") = 0
            rsBucket.fields("PDate") = todate(rs.fields("Date"))
            rsBucket.fields("ProcName") = ProcName
            rsBucket.fields("ComAcc") = ComAcc
            rsBucket.fields("RefType") = RefType
            rsBucket.fields("RefNo") = RefNo
            rsBucket.fields("VType") = VType
            rsBucket.fields("CT") = CT
            rsBucket.fields("Remarks") = IIf(Remarks = "", "Refund Auto Payment", Remarks)
            rsBucket.fields("Last_Opr") = dbUser
            rsBucket.fields("Last_Update") = Today.Date
            If Not rsBucket.Update Then GoTo errmsg
            rs.MoveNext()
        Loop

        Insert_Refund_Auto_Payment = True
        Exit Function
errmsg:
        Insert_Refund_Auto_Payment = False
        Exit Function
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function InforcePolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Dim ANO As Long
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim LANO As Long
            Dim ADate As Date
            Dim PDate As Date
            Dim NCDate As Date
            Dim sysUserID As String
            Dim Remarks As String
            Dim Memo As String
            Dim PolicyNo As String
            Dim Renewal As String
            Dim ESeqNo As Long
            Dim EESeqNo As Long
            Dim GracePeriod As Integer
            Dim BackDatedF As Boolean
            Dim ExpiredF As Boolean
            Dim BucketF As Boolean
            Dim AUserID As String
            Dim AutoPrintF As Boolean
            Dim AutoPrintVersion As String
            Dim AutoPrintRepName As String
            Dim AutoPaymentF As Boolean
            Dim AutoPaymentType As String
            Dim AutoPaymentComAcc As String
            Dim AutoPaymentRefType As String
            Dim AutoPaymentRefNo As String
            Dim AutoPaymentVType As String
            Dim AutoPaymentCT As String
            Dim AutoPaymentRemarks As String
            Dim BulkF As Boolean
            Dim BatchNo As String
            Dim WPC As Integer
            Dim CollectionF As Boolean
            Dim CollectionOpt As String
            Dim AutoGenerateRenewalF As Boolean
            Dim SplitDNCNF As Boolean

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)
            ADate = getInParamValue(InParam, 2)  ' Inforce Date
            PDate = getInParamValue(InParam, 3)  ' Effective Date
            NCDate = getInParamValue(InParam, 4)
            sysUserID = getInParamValue(InParam, 5)
            Remarks = getInParamValue(InParam, 6)
            Memo = getInParamValue(InParam, 7)
            PolicyNo = getInParamValue(InParam, 8)
            Renewal = getInParamValue(InParam, 9)
            ESeqNo = getInParamValue(InParam, 10)
            EESeqNo = getInParamValue(InParam, 11)
            GracePeriod = getInParamValue(InParam, 12)
            BackDatedF = getInParamValue(InParam, 13)
            ExpiredF = getInParamValue(InParam, 14)
            BucketF = toval(getInParamValue(InParam, 15))
            AUserID = tostr(getInParamValue(InParam, 16))
            If AUserID = "" Then AUserID = dbUser
            AutoPrintF = toval(getInParamValue(InParam, 17))
            AutoPrintVersion = tostr(getInParamValue(InParam, 18))
            BulkF = toval(getInParamValue(InParam, 19))
            BatchNo = tostr(getInParamValue(InParam, 20))   ' Bulk Inforce By Batch No
            WPC = toval(getInParamValue(InParam, 21))
            CollectionF = toval(getInParamValue(InParam, 22))   ' Auto Generate Collection
            CollectionOpt = tostr(getInParamValue(InParam, 23))   ' Collection Option : GROSS or NET 
            AutoGenerateRenewalF = toval(getInParamValue(InParam, 24))   ' Auto Generate Renewal 
            SplitDNCNF = toval(getInParamValue(InParam, 25))   ' Split DN CN for Endorsement 
            AutoPrintRepName = tostr(getInParamValue(InParam, 26))

            AutoPaymentF = toval(getInParamValue(InParam, 27))
            AutoPaymentType = tostr(getInParamValue(InParam, 28)) 'Request / Default'
            AutoPaymentComAcc = tostr(getInParamValue(InParam, 29)) 'Company Account Used for Auto Payment (if Empty Used Default Company Account Transfer=1)
            AutoPaymentRefType = tostr(getInParamValue(InParam, 30))  ' RefType Used for Auto Payment
            AutoPaymentRefNo = tostr(getInParamValue(InParam, 31))  ' RefNo Used for Auto Payment
            AutoPaymentVType = tostr(getInParamValue(InParam, 32))  ' Payment Voucher Type Used for Auto Payment
            AutoPaymentCT = tostr(getInParamValue(InParam, 33))  ' Profit/Cost Center Used for Auto Payment
            AutoPaymentRemarks = tostr(getInParamValue(InParam, 34))  ' Remarks Used for Auto Payment

            Dim rs As RemoteSet
            Dim rsAcc As RemoteSet
            Dim RANO As Long
            Dim RegNo As String
            Dim CertificateNo As String
            Dim pRenewal As Integer
            Dim AType As String
            rs = db.OpenRemoteSet("Select RANO, RegNo, CertificateNo, Renewal, AType from Acceptance Where ANO=" & ANO)
            RANO = toval(rs.fields("RANO"))
            RegNo = tostr(rs.fields("RegNo"))
            CertificateNo = tostr(rs.fields("CertificateNo"))
            pRenewal = toval(rs.fields("Renewal"))
            AType = tostr(rs.fields("AType"))

            rs = db.OpenRemoteSet("Select DueByEffectiveF from SysTab")
            Dim DueByEffectiveF As Boolean = toval(rs.fields("DueByEffectiveF"))

            db.Var_db_user = AUserID
            db.Var_db_User_Type = IIf(db.GetPrivileges("ALLOWALLPRIVILEGESINFORCE"), "SUPERVISOR", dbUserType)

            'Cash Before Cover
            Dim pVoucher As String = ""
            Dim pAccount As String = ""
            Dim pCurrency As String = ""
            If Not BulkF Then
                If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

                Dim rsPayor As RemoteSet
                rsPayor = db.OpenQueryDef("QPayor", adOpenKeyset, ANO)
                If ANO <> LANO And LANO <> -1 Then
                    rsPayor = db.OpenQueryDef("QAPayor", adOpenKeyset, ANO)
                End If

                If Not rsPayor.Eof Then
                    Dim Grace As Integer
                    If rsPayor.fields("Installment") <> 0 Then
                        If Not db.GetPrivileges("ALLOWCHANGEFIRSTINSTALLMENTDUE") Then
                            If Not DueByEffectiveF Then
                                Grace = DateDiff(DateInterval.Day, ADate, rsPayor.fields("IDate_1"))
                            Else
                                Grace = rsPayor.fields("Grace")
                            End If
                        Else
                            Grace = GracePeriod
                        End If
                    Else
                        Grace = GracePeriod
                    End If
                    If (rsPayor.fields("Grace") <> Grace Or rsPayor.fields("WPC") <> WPC) Then
                        rsPayor.Edit()
                        rsPayor.fields("Grace") = Grace
                        rsPayor.fields("WPC") = WPC
                        If Not rsPayor.Update Then Throw New Exception(db.ErrDesc)
                    End If
                End If
                pVoucher = tostr(rsPayor.fields("pVoucher"))
                pAccount = tostr(rsPayor.fields("pAccount"))
                pCurrency = tostr(rsPayor.fields("pCurrency"))

                If ExpiredF Then
                    Dim rsR As RemoteSet
                    rsR = db.OpenRemoteSet("Select AStatus, ANO, LANO, CNO from Acceptance Where ANO=" & RANO)
                    If Not rsR.Eof AndAlso rsR.fields("AStatus") = "I" Then
                        If Not PreExpireConditionsOk(RANO, ErrMsg) Then
                            Throw New Exception(ErrMsg)
                        End If
                        If rsR.fields("ANO") = rsR.fields("LANO") Then
                            If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "EXPIREMASTER", rsR.fields("ANO"), rsR.fields("CNO"), Today.Date) Then Throw New Exception(Me.ErrDesc)
                        Else
                            If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "EXPIRE", rsR.fields("ANO"), rsR.fields("CNO"), Today.Date) Then Throw New Exception(Me.ErrDesc)
                        End If
                    End If
                End If

                If Not PreInforceConditionsOK(ANO, ADate, PolicyNo, BackDatedF, False, BucketF, False, ErrMsg) Then
                    Throw New Exception(ErrMsg)
                End If

                rsAcc = db.OpenQueryDef("QAcceptance", adOpenKeyset, ANO)
                rsAcc.Edit()
                rsAcc.fields("ADate") = ADate
                rsAcc.fields("NCDate") = NCDate
                rsAcc.fields("SysUserID") = sysUserID
                rsAcc.fields("Remarks") = Remarks
                rsAcc.fields("Memo") = Memo
                If rsAcc.fields("PolicyNo") = "" Then
                    rsAcc.fields("PolicyNo") = PolicyNo
                End If
                rsAcc.fields("AutoGenerateRenewalF") = AutoGenerateRenewalF
                rsAcc.fields("SplitDNCNF") = SplitDNCNF
                If AType = "N" And RANO = -1 And (BackDatedF Or db.GetPrivileges("POLICYNOENTRY")) Then
                    If ESeqNo < EESeqNo Then ESeqNo = EESeqNo
                    rsAcc.fields("ESeqNo") = ESeqNo
                    rsAcc.fields("EESeqNo") = EESeqNo
                    rsAcc.fields("Renewal") = Renewal
                    ' Change Renewal No in Master Policy also for detail
                    If pRenewal <> Renewal And ANO = LANO Then
                        If Not db.Execute("Update Acceptance Set Renewal=" & Renewal & " from Acceptance Where AStatus='W' and ANO<>LANO and LANO=" & LANO) Then Throw New Exception(db.ErrDesc)
                    End If
                End If
                If Not rsAcc.Update Then Throw New Exception(db.ErrDesc)

                If Not BucketF Then
                    If ANO = LANO Then
                        If Not db.Execute("Update Acceptance Set SysUserID='" & sysUserID & "', NCDate=Case When SDate>'" & Format(NCDate, "MM/DD/YYYY") & "' Then SDate else '" & Format(NCDate, "MM/DD/YYYY") & "' End, AutoGenerateRenewalF='" & IIf(AutoGenerateRenewalF, 1, 0) & "', SplitDNCNF='" & IIf(SplitDNCNF, 1, 0) & "' Where ANO<>LANO and LANO=" & LANO) Then
                            Throw New Exception(rs.ErrDesc)
                        End If

                        If Not Me.Direct_Call_Process_Execute("PROCESS.NETV3.PROCLIB", "INFORCEMASTER", ANO, AUserID, BackDatedF) Then
                            Throw New Exception(Me.ErrDesc)
                        End If
                    Else
                        If Not Me.Direct_Call_Process_Execute("PROCESS.NETV3.PROCLIB", "INFORCE", ANO, AUserID, BackDatedF) Then
                            Throw New Exception(Me.ErrDesc)
                        End If
                    End If
                    ' Requery
                    rsAcc = db.OpenRemoteSet("Select PolicyNo from Acceptance Where ANO=" & ANO)
                    PolicyNo = tostr(rsAcc.fields("PolicyNo"))
                    ' Set Output Parameter
                    ReDim OutParam(0)
                    OutParam(0) = PolicyNo

                    ' Close Pending Auto Inforce
                    rs = db.OpenRemoteSet("Select PID from AIBucket Where Status='R' and RegNo='" & RegNo & "' and CertificateNo='" & CertificateNo & "'")
                    If Not rs.Eof Then
                        If Not db.Execute("Update AIBucket Set Status='C', ErrMsg='Inforced By User', PDate='" & Format(Today.Date, "MM/DD/YYYY") & "' Where PID=" & rs.fields("PID")) Then Throw New Exception(db.ErrDesc)
                    End If
                Else
                    If ANO = LANO Then
                        If Not db.Execute("Update Acceptance Set SysUserID='" & sysUserID & "', NCDate=Case When SDate>'" & Format(NCDate, "MM/DD/YYYY") & "' Then SDate else '" & Format(NCDate, "MM/DD/YYYY") & "' End, AutoGenerateRenewalF='" & IIf(AutoGenerateRenewalF, 1, 0) & "', SplitDNCNF='" & IIf(SplitDNCNF, 1, 0) & "' Where ANO<>LANO and LANO=" & LANO) Then
                            Throw New Exception(rs.ErrDesc)
                        End If
                    End If
                    ' insert into Auto Inforce Bucket
                    rs = db.OpenRemoteSet("Select * from AIBucket Where Status='R' and RegNo='" & RegNo & "' and CertificateNo='" & CertificateNo & "'")
                    If rs.Eof Then
                        rs.AddNew()
                    Else
                        rs.Edit()
                    End If
                    rs.fields("RegNo") = RegNo
                    rs.fields("CertificateNo") = CertificateNo
                    rs.fields("ADate") = ADate
                    rs.fields("BatchNo") = ""
                    rs.fields("Status") = "R"
                    rs.fields("Last_Opr") = dbUser
                    rs.fields("Last_Update") = Today.Date
                    If Not rs.Update Then Throw New Exception(db.ErrDesc)

                    ' Set Output Parameter
                    ReDim OutParam(0)
                    OutParam(0) = "Auto Inforce Process !"
                End If

                ' Auto Print
                If AutoPrintF Then
                    Insert_Policy_Auto_Print(ANO, AutoPrintVersion, AutoPrintRepName)
                End If

                If AutoPaymentF Then
                    Insert_Refund_Auto_Payment(ANO, IIf(UCase(AutoPaymentType) = "REQUEST", "AutoReceiptRequest_Process", ""), AutoPaymentComAcc, AutoPaymentRefType, AutoPaymentRefNo, AutoPaymentVType, AutoPaymentCT, AutoPaymentRemarks)
                End If

                ' Insert Auto Receipt Payment
                If pVoucher <> "" AndAlso pAccount <> "" AndAlso pCurrency <> "" Then
                    Insert_Auto_Receipt_Payment(ANO, pVoucher, pAccount, pCurrency)
                End If

                If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            Else    ' Bulk Inforce
                Dim NoRec As Integer = 0
                Dim ListANO() As Long = Nothing
                If BatchNo <> "" Then
                    rs = db.OpenRemoteSet("Select ANO, PDate from Acceptance Where AStatus in ('W','T') and BatchNo='" & BatchNo & "' and LANO=" & LANO)
                Else
                    rs = db.OpenRemoteSet("Select ANO, PDate from Acceptance Where AStatus in ('W','T') and LANO=" & LANO)
                End If
                Do While Not rs.Eof
                    If Not db.GetLockedObject("Policy" & rs.fields("ANO")) Then Throw New Exception(db.ErrDesc)

                    If Not PreInforceConditionsOK(rs.fields("ANO"), ADate, PolicyNo, BackDatedF, False, BucketF, False, ErrMsg) Then
                        Throw New Exception(ErrMsg)
                    End If

                    If Not db.Execute("Update Acceptance Set SysUserID='" & sysUserID & "', AutoGenerateRenewalF='" & IIf(AutoGenerateRenewalF, 1, 0) & "', SplitDNCNF='" & IIf(SplitDNCNF, 1, 0) & "' Where ANO=" & rs.fields("ANO")) Then
                        Throw New Exception(rs.ErrDesc)
                    End If

                    ReDim Preserve ListANO(NoRec)
                    ListANO(NoRec) = toval(rs.fields("ANO"))
                    NoRec = NoRec + 1

                    rs.MoveNext()
                Loop

                If NoRec > 0 Then
                    If Not BucketF Then
                        If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "INFORCEBULK", ListANO, AUserID, BackDatedF, ADate, BulkF) Then
                            Throw New Exception(Me.ErrDesc)
                        End If

                        ' Close Pending Auto Inforce
                        rs = db.OpenRemoteSet("Select PID from AIBucket Where Status='R' and RegNo='" & RegNo & "' and CertificateNo='" & CertificateNo & "'")
                        If rs.Eof Then
                            rs = db.OpenRemoteSet("Select PID from AIBucket Where Status='R' and RegNo='" & RegNo & "' and CertificateNo='' and BatchNo='" & BatchNo & "'")
                        End If
                        If Not rs.Eof Then
                            If Not db.Execute("Update AIBucket Set Status='C', ErrMsg='Inforced By User', PDate='" & Format(Today.Date, "MM/DD/YYYY") & "' Where PID=" & rs.fields("PID")) Then Throw New Exception(db.ErrDesc)
                        End If

                        ' Set Output Parameter
                        ReDim OutParam(0)
                        OutParam(0) = NoRec & " record(s) Processed !"
                    Else
                        rs = db.OpenRemoteSet("Select * from AIBucket Where Status='R' and RegNo='" & RegNo & "' and CertificateNo='" & CertificateNo & "'")
                        If rs.Eof Then
                            rs = db.OpenRemoteSet("Select * from AIBucket Where Status='R' and RegNo='" & RegNo & "' and CertificateNo='' and BatchNo='" & BatchNo & "'")
                        End If
                        If rs.Eof Then
                            rs.AddNew()
                        Else
                            rs.Edit()
                        End If
                        rs.fields("RegNo") = RegNo
                        rs.fields("CertificateNo") = ""
                        rs.fields("ADate") = ADate
                        rs.fields("BatchNo") = BatchNo
                        rs.fields("Status") = "R"
                        rs.fields("Last_Opr") = dbUser
                        rs.fields("Last_Update") = Today.Date
                        If Not rs.Update Then Throw New Exception(rs.ErrDesc)

                        ' Set Output Parameter
                        ReDim OutParam(0)
                        OutParam(0) = "Auto Inforce Process !"
                    End If

                    For Each ANO In ListANO
                        ' Auto Print
                        If AutoPrintF Then
                            Insert_Policy_Auto_Print(ANO, AutoPrintVersion, AutoPrintRepName)
                        End If

                        If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
                    Next
                Else
                    ' Set Output Parameter
                    ReDim OutParam(0)
                    OutParam(0) = "No Data Processed !"
                End If

            End If

            db.Var_db_user = dbUser
            db.Var_db_User_Type = dbUserType

            If Not BucketF AndAlso CollectionF Then
                If Not AutoGenerateCollection(ANO, CollectionOpt, ErrMsg) Then Throw New Exception(ErrMsg)
            End If

            ' Auto Done Task if ActionList ActionEndF=1
            rs = db.OpenRemoteSet("Select S.ID from SysRequest S Left Join SysRequest pS on (pS.ID=s.PID) Inner Join ActionList A on (A.ActionCode=(case When s.ActionCode='' and coalesce(pS.ActionCode, '')<>'' then  coalesce(pS.ActionCode, '') else s.ActionCode end) and A.ActionEndF=1 and (A.RCondition=1 or A.ApprovedFName='ALLOWPASSAPPLIMIT')) Where S.Status='' and S.Param2 in ('UND','UND-MW') and S.Param3='" & ANO & "'")
            If Not rs.Eof Then
                Dim ws As New WSCommon
                Dim IParam(7) As Object
                Dim OParam() As Object = Nothing
                IParam(0) = rs.fields("ID")
                IParam(5) = "A"
                IParam(7) = True
                If Not ws.DoneSMSTask(db, dbUser, dbPassword, IParam, OParam, ErrMsg) Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If CBool(toval(GetAppSetting("AutoPostingJournalDNCNF"))) Then
                Dim SQL As String

                If BulkF Or ANO <> LANO Then
                    SQL = "Select Distinct V.Voucher, V.Branch"
                    SQL = SQL & " from Acceptance A"
                    SQL = SQL & " Inner Join AccAss A1 on (A1.ANO=A.ANO)"
                    SQL = SQL & " Inner Join Admlink A2 on (A2.AdmNo=A1.AdmNo)"
                    SQL = SQL & " Inner Join nVoucher V on (V.Voucher=A2.Voucher)"
                    SQL = SQL & " Where V.Posting_Flag = 0"
                    If BulkF Then
                        If BatchNo <> "" Then
                            SQL = SQL & " and A.RegNo='" & RegNo & "' and A.BatchNo='" & BatchNo & "' and A.LANO=" & LANO
                        Else
                            SQL = SQL & " and A.RegNo='" & RegNo & "' and A.LANO=" & LANO
                        End If
                    Else
                        SQL = SQL & " and A.ANO=" & ANO
                    End If
                Else
                    SQL = "Select Distinct V.Voucher, V.Branch"
                    SQL = SQL & " from Acceptance mA"
                    SQL = SQL & " Inner Join Acceptance A on (A.LANO=mA.ANO and A.RegNo=mA.RegNo)"
                    SQL = SQL & " Inner Join AccAss A1 on (A1.ANO=A.ANO)"
                    SQL = SQL & " Inner Join Admlink A2 on (A2.AdmNo=A1.AdmNo)"
                    SQL = SQL & " Inner Join nVoucher V on (V.Voucher=A2.Voucher)"
                    SQL = SQL & " Where V.Posting_Flag = 0 "
                    SQL = SQL & " and mA.ANO = " & ANO
                End If

                rs = db.OpenRemoteSet(SQL)
                Do While Not rs.Eof
                    If Not AutoPostingJournal(rs.fields("Voucher"), "N", rs.fields("Branch"), ErrMsg) Then Throw New Exception(ErrMsg)
                    rs.MoveNext()
                Loop
            End If

            db.Committrans()

            If GetAppSetting("VABank") <> "" Then
                Try
                    Dim SQL As String
                    SQL = "Select Distinct Vd.VID, Vd.AccountNo"
                    SQL = SQL & " from Admlink A1, AccAss A2, VAccount Vh, VAccountD Vd "
                    SQL = SQL & " Where A1.Bank='" & GetAppSetting("VABank") & "' and A1.RefType='VIRTUAL' and Vh.Bank=A1.Bank and Vd.VID=Vh.VID and Vd.AccountNo=A1.AccountNo and Vd.A_PolicyNo<>'' and A1.AdmNo=A2.AdmNO and A2.ANO=" & ANO

                    rs = db.OpenRemoteSet(SQL)
                    Do While Not rs.Eof

                        Dim Param(1) As Object
                        Param(0) = rs.fields("VID")
                        Param(1) = rs.fields("AccountNo")

                        Dim wsExtra = New WSExtras
                        wsExtra.ActivateVAccount(dbUser, dbPassword, Param, ErrMsg)

                        rs.MoveNext()
                    Loop
                Catch ex As Exception

                End Try
            End If

            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            If db.CheckLockedObject("Policy" & ANO) Then db.ReleaseLockedObject("Policy" & ANO)
            CloseConnection()
            PNotes(ANO, "POLICY", "Policy Inforce : " & ErrMsg)
            Return False
        End Try
    End Function
    Private Function AutoPostingJournal(ByVal Voucher As String, ByVal SL_Cat As String, ByVal Branch As String, ByRef ErrMsg As String) As Boolean
        Try
            Dim InParam(2) As Object
            InParam(0) = Voucher
            InParam(1) = SL_Cat
            InParam(2) = Branch

            Dim findex() As String = Nothing
            Dim fbuffer() As String = Nothing
            Dim Dataset As DataSet = Nothing

            Dim wsAccounting As New WSAccounting
            If Not wsAccounting.GetJournal(db, InParam, findex, fbuffer, Dataset, ErrMsg) Then Throw New Exception(ErrMsg)
            If Not wsAccounting.PostingJournal(db, InParam, Dataset, ErrMsg) Then Throw New Exception(ErrMsg)

            AutoPostingJournal = True

        Catch ex As Exception
            ErrMsg = ex.Message
            AutoPostingJournal = False
        End Try
    End Function
    Private Function AutoGenerateCollection(ByVal ANO As Long, ByVal GrossNetOpt As String, ByRef ErrMsg As String) As Boolean
        Try
            Dim DataSet As DataSet
            Dim rs As RemoteSet = Nothing
            Select Case GrossNetOpt.ToUpper
                Case "GROSS"
                    rs = db.OpenRemoteSet("Select * from CL_I_COLLECT_PERIOD_GROSS_ALL_BYPOLICYNO(" & ANO & ")")
                Case "NET"
                    rs = db.OpenRemoteSet("Select * from CL_I_COLLECT_PERIOD_NET_ALL_BYPOLICYNO(" & ANO & ")")

            End Select
            DataSet = rs.GetDataSet

            Dim CollNo As String
            Dim rsH As RemoteSet
            Dim rsA As RemoteSet
            Dim rsD As RemoteSet
            Dim rsV As RemoteSet
            Dim RefNo As String
            Dim i As Long
            Dim Ccy As String
            Dim Nominal As Double
            Dim CollType As String
            Dim AutoNum As String
            Dim Branch As String
            Dim CollDate As Date

            Branch = rs.fields("Branch")
            CollDate = rs.fields("Date")

            rs = db.OpenRemoteSet("Select CollType, AutoNum from CollType Where CType='C'")
            CollType = rs.fields("CollType")
            AutoNum = rs.fields("AutoNum")

            rsH = db.OpenRemoteSet("Select * from CollectionH Where 1<>1")
            rsH.AddNew()
            CollNo = db.GetNextDocNo(AutoNum, CollDate, Branch)
            If CollNo = "" Then
                ErrMsg = "Generate Collection Number Failed !"
                Throw New Exception(ErrMsg)
            End If
            rsH.fields("CollNo") = CollNo
            rsH.fields("CollDate") = CollDate
            rsH.fields("CollType") = CollType
            rsH.fields("Branch") = Branch
            rsH.fields("SDate") = CollDate
            rsH.fields("EDate") = CollDate
            rsH.fields("PSDate") = CollDate
            rsH.fields("PEDate") = CollDate
            rsH.fields("CollStatus") = "R"
            rsH.fields("RepDate") = CollDate
            rsH.fields("Type") = "I"
            rsH.fields("CType") = "C"
            rsH.fields("ProfileF") = False
            rsH.fields("ID") = dbUser
            rsH.fields("Last_Opr") = dbUser
            rsH.fields("last_Date") = Today.Date
            rsH.fields("Last_Time") = Time()
            If Not rsH.Update Then Throw New Exception(db.ErrDesc)

            rsV = db.OpenRemoteSet("select * from CollectionV where 1<>1")
            RefNo = ""
            Ccy = ""
            For i = 0 To DataSet.Tables(0).Rows.Count - 1
                RefNo = tostr(DataSet.Tables(0).Rows(i)("DocNo"))
                Ccy = tostr(DataSet.Tables(0).Rows(i)("Currency"))
                If toval(DataSet.Tables(0).Rows(i)("DebtorF")) = -1 Then
                    Nominal = toval(DataSet.Tables(0).Rows(i)("Nominal"))
                Else
                    Nominal = -toval(DataSet.Tables(0).Rows(i)("Nominal"))
                End If
                rsD = db.OpenRemoteSet("Select * from CollectionD Where CollNo='" & CollNo & "' and RefNo='" & RefNo & "' and Currency='" & Ccy & "'")
                If rsD.Eof Then
                    rsD.AddNew()
                    rsD.fields("CollNo") = CollNo
                    rsD.fields("RefNo") = RefNo
                    rsD.fields("Currency") = Ccy
                    rsD.fields("AName") = Trim(tostr(DataSet.Tables(0).Rows(i)("AName"), 255))
                    rsD.fields("DueDate") = todate(DataSet.Tables(0).Rows(i)("DueDate"))
                Else
                    rsD.Edit()
                End If
                rsD.fields("Nominal") = toval(rsD.fields("Nominal")) + Nominal
                rsD.fields("Amount") = toval(rsD.fields("Amount")) + Nominal
                If Not rsD.Update Then Throw New Exception(db.ErrDesc)

                rsA = db.OpenRemoteSet("Select * from CollAmount Where CollNo='" & CollNo & "' and Currency='" & Ccy & "'")
                If rsA.Eof Then
                    rsA.AddNew()
                    rsA.fields("CollNo") = CollNo
                    rsA.fields("Currency") = Ccy
                Else
                    rsA.Edit()
                End If
                rsA.fields("TAmount") = toval(rsA.fields("TAmount")) + Nominal
                If Not rsA.Update Then Throw New Exception(db.ErrDesc)

                rsV.AddNew()
                rsV.fields("CollNo") = CollNo
                rsV.fields("RefNo") = RefNo
                rsV.fields("Currency") = Ccy
                rsV.fields("Voucher") = tostr(DataSet.Tables(0).Rows(i)("Voucher"))
                rsV.fields("Nominal") = Nominal
                rsV.fields("Amount") = Nominal
                If Not rsV.UpdateRow Then Throw New Exception(db.ErrDesc)

                If Not db.Execute("Update nVoucher Set CollNo='" & CollNo & "' Where Voucher='" & rsV.fields("Voucher") & "'") Then Throw New Exception(db.ErrDesc)
            Next

            ' Finally
            If Not rsV.Eof Then If Not rsV.Update Then Throw New Exception(db.ErrDesc)
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function EndorsePolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long
            Dim PDate As Date

            ANO = getInParamValue(InParam, 0)
            PDate = getInParamValue(InParam, 3)
            LANO = getInParamValue(InParam, 5)

            Dim Param(18) As Object
            Param(0) = getInParamValue(InParam, 1)       ' RegNo
            Param(1) = getInParamValue(InParam, 2)       ' Booking Date / ADate
            Param(2) = getInParamValue(InParam, 3)       ' Effective Date / PDate
            Param(3) = getInParamValue(InParam, 4)       ' Endorsement Method 'U' / 'R' : Reinstatement Claim
            Param(4) = getInParamValue(InParam, 5)       ' LANO
            Param(5) = getInParamValue(InParam, 6)       ' Remarks
            Param(6) = getInParamValue(InParam, 7)       ' CPCT
            Param(7) = getInParamValue(InParam, 8)       ' Memo
            Param(8) = getInParamValue(InParam, 9)       ' True : External Endorsement, False : Internal Endorsement
            Param(9) = getInParamValue(InParam, 10)      ' RefNo
            Param(10) = getInParamValue(InParam, 11)     ' RefDate
            Param(11) = getInParamValue(InParam, 12)     ' 0 : Default, 1 : No Add Refund, 2 : No Add Refund All
            Param(12) = getInParamValue(InParam, 13)     ' Wording Code
            Param(13) = getInParamValue(InParam, 14)     ' True : Endorse All
            Param(14) = getInParamValue(InParam, 15)     ' App Source : 'UW', 'RI','MKT'
            Param(15) = toval(getInParamValue(InParam, 16))     ' Adjust Coverage Period (True/False)
            Param(16) = getInParamValue(InParam, 17)     ' Reason Code
            Param(17) = toval(getInParamValue(InParam, 18))   ' Adjust Facultative Period (True/False)
            Param(18) = getInParamValue(InParam, 19)     ' Refund Type


            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not PreEndorseConditionsOK(ANO, PDate, Param(13), Param(14), Param(8), ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            Dim NewANO As Long
            If ANO = LANO Then
                Dim Result As Object = Nothing
                Result = Me.Direct_Call_Ext_Process("ProcessExt.NetV3.Proclib", "EndorseMaster", ANO, Param)
                If Result Is Nothing Then
                    Throw New Exception("Background Process is running !")
                Else
                    NewANO = toval(Result)
                    If NewANO = 0 Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            Else
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Endorsement", ANO, Param)
                If NewANO = 0 Then
                    Throw New Exception(Me.ErrDesc)
                End If
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = NewANO

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function Cancellation_Prorata_By_Refund(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim PDate As Date
            Dim Refund As Double
            Dim Prorata As Double

            ANO = getInParamValue(InParam, 0)
            PDate = getInParamValue(InParam, 1)
            Refund = getInParamValue(InParam, 2)
            Prorata = toval(getInParamValue(InParam, 3))

            Dim rs As RemoteSet
            Dim PremiumTotal As Double
            Dim PremiumInforced As Double
            Dim PremiumNotYetInforced As Double
            Dim NumCCY As Integer
            Dim CCY As String

            Dim TOC As String
            Dim Branch As String
            Dim BizMapID As String
            Dim SDate As Date
            Dim SQL As String = ""

            rs = db.OpenRemoteSet("Select TOC, Branch, BizMapID from Acceptance Where ANO=" & ANO)
            TOC = rs.fields("TOC")
            Branch = rs.fields("Branch")
            BizMapID = rs.fields("BizMapID")

            NumCCY = 0
            CCY = ""

            SQL = ""
            SQL = SQL & "Declare @AsAt as smalldatetime"
            SQL = SQL & " Declare @ANO as Int"

            SQL = SQL & " Select @ANO=" & ANO
            SQL = SQL & " Select @Asat='" & Format(PDate, "MM/DD/YYYY") & "'"

            SQL = SQL & " Select Max(P.Currency) as Currency, "
            SQL = SQL & " Count (Distinct P.Currency) as NumCCY,"
            SQL = SQL & " sum(R.Gross/P.OGross*P.Gross) as PremiumTotal, "
            SQL = SQL & " sum(R.Gross/P.OGross*P.Gross*P.Rate) as PremiumTotalInBased, "
            SQL = SQL & " sum(case When Not ((DateDiff(Day, R.EDate, @AsAt) >= 0 or DateDiff(Day, @AsAt, R.SDate) > 0) and datediff(day,@AsAt,A.SDate)<=0 and datediff(day,@AsAt,A.EDate)>=0) then R.Gross/P.OGross*P.Gross else 0 end) as PremiumInforced,"
            SQL = SQL & " sum(case When Not ((DateDiff(Day, R.EDate, @AsAt) >= 0 or DateDiff(Day, @AsAt, R.SDate) > 0) and datediff(day,@AsAt,A.SDate)<=0 and datediff(day,@AsAt,A.EDate)>=0) then R.Gross/P.OGross*P.Gross*P.Rate else 0 end) as PremiumInforcedInBased,"
            SQL = SQL & " min(case When (DateDiff(Day, @AsAt, R.SDate) > 0 and datediff(day,@AsAt,A.SDate)<=0 and datediff(day,@AsAt,A.EDate)>=0) then R.SDate else Null end) as SDate,"
            SQL = SQL & " sum(case When (DateDiff(Day, @AsAt, R.SDate) > 0 and datediff(day,@AsAt,A.SDate)<=0 and datediff(day,@AsAt,A.EDate)>=0) then R.Gross/P.OGross*P.Gross else 0 end) as PremiumNotYetInforced,"
            SQL = SQL & " sum(case When (DateDiff(Day, @AsAt, R.SDate) > 0 and datediff(day,@AsAt,A.SDate)<=0 and datediff(day,@AsAt,A.EDate)>=0) then R.Gross/P.OGross*P.Gross*P.Rate else 0 end) as PremiumNotYetInforcedInBased"
            SQL = SQL & " from RCalc R, PCalc P, Acceptance A "
            SQL = SQL & " Where P.ANO = R.ANO and A.ANO=R.ANO And P.Currency = R.Currency and R.ANO=@ANO"

            rs = db.OpenRemoteSet(SQL)
            If NumCCY = 0 Then NumCCY = toval(rs.fields("NumCCY"))
            If NumCCY <= 1 Then
                PremiumTotal = toval(rs.fields("PremiumTotal"))
                PremiumInforced = toval(rs.fields("PremiumInforced"))
                PremiumNotYetInforced = toval(rs.fields("PremiumNotYetInforced"))
            Else
                PremiumTotal = toval(rs.fields("PremiumTotalInBased"))
                PremiumInforced = toval(rs.fields("PremiumInforcedInBased"))
                PremiumNotYetInforced = toval(rs.fields("PremiumNotYetInforcedInBased"))
            End If
            If CCY = "" Then CCY = tostr(rs.fields("Currency"))
            SDate = todate(rs.fields("SDate"))

            ' Minimum Retained Premium
            Dim MRPremium As Double
            If NumCCY <= 1 Then
                rs = db.OpenRemoteSet("Select Top 1 B.MRPremium from BRRuleMin B, SysTab S Where B.TOC='" & TOC & "' and (B.Branch='<All>' or B.Branch='" & Branch & "') and (B.BizMapID='<All>' or B.BizMapID='" & BizMapID & "') and B.Currency='" & CCY & "' Order By B.MRPremium")
                MRPremium = toval(rs.fields("MRPremium"))
            Else
                rs = db.OpenRemoteSet("Select Top 1 B.MRPremium from BRRuleMin B, SysTab S Where B.TOC='" & TOC & "' and (B.Branch='<All>' or B.Branch='" & Branch & "') and (B.BizMapID='<All>' or B.BizMapID='" & BizMapID & "') and B.Currency=S.Currency Order By B.MRPremium")
                MRPremium = toval(rs.fields("MRPremium"))
            End If

            If PremiumInforced <> 0 Then
                If Prorata <> 0 AndAlso Refund = 0 Then
                    Refund = Prorata / 100 * PremiumInforced + PremiumNotYetInforced
                End If

                ' Compare with Minimum Retained Premium and Premium Total
                If PremiumTotal - MRPremium < Refund Then
                    Refund = PremiumTotal - MRPremium
                End If

                'If (Refund - PremiumNotYetInforced) < 0 Then
                '    Refund = PremiumNotYetInforced
                'End If

                Prorata = (Refund - PremiumNotYetInforced) / PremiumInforced * 100
            Else
                Refund = 0
                Prorata = 0
            End If

            ' Set Output Parameter
            ReDim OutParam(2)
            OutParam(0) = Prorata
            OutParam(1) = Refund
            OutParam(2) = IIf(Prorata = 0, SDate, PDate)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function CancelPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long
            Dim PDate As Date
            Dim OpenFlag As Boolean

            ANO = getInParamValue(InParam, 0)
            PDate = getInParamValue(InParam, 3)
            OpenFlag = getInParamValue(InParam, 4)
            LANO = getInParamValue(InParam, 5)


            Dim Param(17) As Object
            Param(0) = getInParamValue(InParam, 1)       ' RegNo
            Param(1) = getInParamValue(InParam, 2)       ' Booking Date / ADate
            Param(2) = getInParamValue(InParam, 3)       ' Effective Date / PDate
            Param(3) = getInParamValue(InParam, 4)       ' True : Open Flag
            Param(4) = getInParamValue(InParam, 5)       ' LANO
            Param(5) = getInParamValue(InParam, 6)       ' Remarks
            Param(6) = getInParamValue(InParam, 7)       ' CPCT
            Param(7) = getInParamValue(InParam, 8)       ' Memo
            Param(8) = getInParamValue(InParam, 9)       ' True : External Endorsement, False : Internal Endorsement
            Param(9) = getInParamValue(InParam, 10)      ' RefNo
            Param(10) = getInParamValue(InParam, 11)     ' RefDate
            Param(11) = getInParamValue(InParam, 12)     ' Branch
            Param(12) = getInParamValue(InParam, 13)     ' TOC
            Param(13) = getInParamValue(InParam, 14)     ' 0 : Default, 1 : No Add Refund, 2 : No Add Refund All
            Param(14) = getInParamValue(InParam, 15)     ' Wording Code
            Param(15) = getInParamValue(InParam, 16)     ' "A" : Cancel All
            Param(16) = getInParamValue(InParam, 17)     ' Reason Code
            Param(17) = getInParamValue(InParam, 18)     ' Refund Type

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not PreCancelConditionsOK(ANO, PDate, OpenFlag, (getInParamValue(InParam, 16) = "A"), getInParamValue(InParam, 13), getInParamValue(InParam, 14), getInParamValue(InParam, 9), ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            Dim NewANO As Long

            If ANO = LANO Then
                Dim Result As Object = Nothing
                Result = Me.Direct_Call_Ext_Process("ProcessExt.NetV3.Proclib", "CancelMaster", ANO, Param)
                If Result Is Nothing Then
                    Throw New Exception("Background Process is running !")
                Else
                    NewANO = toval(Result)
                    If NewANO = 0 Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            Else
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Cancellation", ANO, Param)
                If NewANO = 0 Then
                    Throw New Exception(Me.ErrDesc)
                End If
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = NewANO

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ClosePolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)

            Dim Param(3) As Object
            Param(0) = getInParamValue(InParam, 1)       ' Transaction Date
            Param(1) = getInParamValue(InParam, 2)       ' Remarks
            Param(2) = getInParamValue(InParam, 3)       ' Validation in Days
            Param(3) = getInParamValue(InParam, 4)       ' Reason

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            Dim rs As RemoteSet
            Dim LANO As Long
            Dim OpenFlag As Boolean
            rs = db.OpenRemoteSet("Select A.LANO, A.AType, A.AStatus, C.OpenFlag from Acceptance A, Cover C Where C.CNO=A.CNO and A.ANO=" & ANO)
            LANO = toval(rs.fields("LANO"))
            OpenFlag = toval(rs.fields("OpenFlag"))
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            ElseIf Not ((LANO = -1 Or ANO = LANO Or OpenFlag) And (tostr(rs.fields("AType")) = "N" And tostr(rs.fields("AStatus")) = "W")) Then
                ErrMsg = "Please Check Policy/Certificate Data Status !"
                Throw New Exception(ErrMsg)
            End If

            If Not PreCloseConditionsOk(ANO, LANO, OpenFlag, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "ClosePolicy", ANO, Param) Then Throw New Exception(Me.ErrDesc)

            ' Set Output Parameter

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function UndoClosePolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            Dim rs As RemoteSet
            Dim CNO As Long
            Dim LANO As Long
            Dim OpenFlag As Boolean
            rs = db.OpenRemoteSet("Select A.CNO, A.LANO, A.AStatus, C.CStatus, C.OpenFlag from Acceptance A, Cover C Where C.CNO=A.CNO and A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            ElseIf Not (rs.fields("AStatus") = "C" And rs.fields("CStatus") = "C") Then
                ErrMsg = "Please Check Policy Data Status !"
                Throw New Exception(ErrMsg)
            End If
            CNO = toval(rs.fields("CNO"))
            LANO = toval(rs.fields("LANO"))
            OpenFlag = toval(rs.fields("OpenFlag"))

            If Not PreUndoCloseConditionsOk(ANO, LANO, OpenFlag, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "UnCloseQuotation", CNO) Then Throw New Exception(Me.ErrDesc)

            ' Set Output Parameter

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function ExpirePolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long
            Dim CNO As Long
            Dim OpenFlag As Boolean
            Dim ExpDate As Date

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)
            CNO = getInParamValue(InParam, 2)
            OpenFlag = getInParamValue(InParam, 3)
            ExpDate = getInParamValue(InParam, 4)        ' Expiry Date

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not PreExpireConditionsOk(ANO, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If


            If ANO = LANO And Not OpenFlag Then
                If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "EXPIREMASTER", ANO, CNO, ExpDate) Then Throw New Exception(Me.ErrDesc)
            Else
                If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "EXPIRE", ANO, CNO, ExpDate) Then Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function UndoExpirePolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long
            Dim CNO As Long
            Dim OpenFlag As Boolean
            Dim ExpDate As Date

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)
            CNO = getInParamValue(InParam, 2)
            OpenFlag = getInParamValue(InParam, 3)
            ExpDate = getInParamValue(InParam, 4)        ' Expiry Date

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not PreUndoExpireConditionsOk(ANO, LANO, OpenFlag, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            If ANO = LANO And Not OpenFlag Then
                If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "UNDOEXPIREMASTER", ANO, CNO, ExpDate) Then Throw New Exception(Me.ErrDesc)
            Else
                If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "UNDOEXPIRE", ANO, CNO, ExpDate) Then Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ConfirmPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)

            Dim Param(3) As Object
            Param(0) = getInParamValue(InParam, 1)       ' CNO
            Param(1) = getInParamValue(InParam, 2)       ' Confirmed Date
            Param(2) = getInParamValue(InParam, 3)       ' Validation in Days
            Param(3) = getInParamValue(InParam, 4)       ' Comment

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not PreConfirmConditionsOK(ANO, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            Dim NewANO As Long
            NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "ConfirmQuotation", ANO, Param)
            If NewANO = 0 Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = NewANO

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function PreConfirmConditionsOK(ByVal ANO As Long, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim LANO As Long
        Dim OpenFlag As Boolean
        Dim PStatus As String
        Dim AType As String

        On Error GoTo errmsg

        rs = db.OpenRemoteSet("Select Cover.CNO, Acceptance.AType, Acceptance.LANO, Cover.OpenFlag, Acceptance.PStatus from Cover, Acceptance where Acceptance.CNO=Cover.CNO and Acceptance.ANO=" & ANO)
        OpenFlag = toval(rs.fields("OpenFlag"))
        LANO = toval(rs.fields("LANO"))
        PStatus = tostr(rs.fields("PStatus"))
        AType = tostr(rs.fields("AType"))

        If AType = "N" Then
            ErrMsg = "Confirm Only for Policy Data !"
            GoTo errmsg
        End If

        If ANO = LANO And Not OpenFlag Then ' Master
            rs = db.OpenRemoteSet("Select ANO, PStatus from Acceptance Where AStatus='W' and ANO<>LANO and LANO=" & LANO)
            Do While Not rs.Eof
                ' Do Premium Calculation If Necessary !
                If rs.fields("PStatus") = "I" Then
                    If Not Me.Direct_Call_Ext_Process("Process.NetV3.Proclib", "PREMIUM_CALC", toval(rs.fields("ANO"))) Then
                        ErrMsg = Me.ErrDesc
                        GoTo errmsg
                    End If
                End If
                If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "PRECONFIRMCONDITIONS", toval(rs.fields("ANO")), False) Then
                    ErrMsg = Me.ErrDesc
                    GoTo errmsg
                End If
                rs.MoveNext()
            Loop
        Else
            If ANO <> LANO Then
                If PStatus = "I" Then
                    If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "PREMIUM_CALC", ANO) Then
                        ErrMsg = Me.ErrDesc
                        GoTo errmsg
                    End If
                End If
                If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "PRECONFIRMCONDITIONS", ANO, False) Then
                    ErrMsg = Me.ErrDesc
                    GoTo errmsg
                End If
            End If
        End If

        PreConfirmConditionsOK = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy confirm : " & ErrMsg)
        PreConfirmConditionsOK = False
    End Function
    Private Function CheckRenewalData(ByVal ANO As Long, ByVal NewTOC As String, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        On Error GoTo errmsg


        rs = db.OpenRemoteSet("Select RegNo from Acceptance Where RANO=" & ANO)
        If Not rs.Eof Then
            ErrMsg = "Already Renewed ! Referred to The Policy Data. (Reg No. " & rs.fields("RegNo") & ")"
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select AType from Acceptance Where PANO=" & ANO, adOpenForwardonly)
        If Not rs.Eof Then
            If rs.fields("AType") <> "A" Then
                ErrMsg = "Can't Renew; While The Policy is being Endorsed !"
                GoTo errmsg
            End If
        End If

        Dim RegDate As Date
        Dim EDate As Date
        Dim RenewalDataF As Boolean
        Dim LANO As Long
        Dim OpenFlag As Boolean
        rs = db.OpenRemoteSet("Select TOC.RenewalDataF, NewTOC.NOINPUTF, TOPRO.NORENEWALF, Acceptance.LANO, Cover.OpenFlag, Acceptance.AdjustableF, Acceptance.RegDate, coalesce(RArrHeader.PPDate, Acceptance.EDate) as R_PPDate, coalesce(RArrHeader.NRDate, Acceptance.EDate) as NRDate, Acceptance.EDate from Acceptance Inner Join Cover on (Cover.CNO=Acceptance.CNO) Inner Join TOC on (TOC.TOC=Acceptance.TOC) Left Join TOPRO on (TOPRO.TOPRO=Acceptance.TOPRO) Left Join RArrHeader on (RArrHeader.ANO=Acceptance.ANO) Left Join TOC NewTOC on (NewTOC.TOC='" & NewTOC & "') Where Acceptance.ANO=" & ANO, adOpenForwardonly)
        If Not rs.Eof Then
            If toval(rs.fields("AdjustableF")) Then
                If Not db.GetPrivileges("RENEWONADJUSTABLEPOLICY") Then
                    ErrMsg = "Can't Renew; While The Policy is not yet Adjusted (Adjustable Policy) !"
                    GoTo errmsg
                End If
            End If

            If DateDiff(DateInterval.Day, rs.fields("NRDATE"), rs.fields("EDate")) > 0 Then
                ErrMsg = "Can't Renew; Policy Needs Treaty Reallocation, Please Review The Reinsurance Date on " & Format(rs.fields("NRDate"), "d") & " !"
                GoTo errmsg
            End If

            ' Clean Cut
            If DateDiff(DateInterval.Day, rs.fields("NRDate"), rs.fields("R_PPDate")) < 0 Then
                ErrMsg = "Can't Renew; Policy Needs Treaty Reallocation, Please Review The Reinsurance Date on " & Format(rs.fields("R_PPDate"), "d") & " !"
                GoTo errmsg
            End If

            If toval(rs.fields("NOINPUTF")) Then
                ErrMsg = "Can't Renew; Please Check No Input in Sub-Class Parameter Setting !"
                GoTo errmsg
            End If
            If toval(rs.fields("NORENEWALF")) Then
                ErrMsg = "Can't Renew; Please Check No Renewal in Product Parameter Setting !"
                GoTo errmsg
            End If
        End If
        EDate = rs.fields("EDate")
        RegDate = rs.fields("RegDate")
        RenewalDataF = rs.fields("RenewalDataF")
        LANO = rs.fields("LANO")
        OpenFlag = rs.fields("OpenFlag")

        rs = db.OpenRemoteSet("Select R.RenewalF, R.Action, R.PStatus from Renewal R Where R.ANO=" & ANO, adOpenForwardonly)
        If Not rs.Eof Then
            If (toval(rs.fields("RenewalF")) = 0 And toval(rs.fields("Action")) = 0) Or toval(rs.fields("Action")) = 3 Then
                ErrMsg = "Can't Renew; The Policy is Declined for Renewal !"
                GoTo errmsg
            End If
            If tostr(rs.fields("PStatus")) = "I" And RenewalDataF Then
                ErrMsg = "Can't Renew; The Renewal Data Has Changed; Drop and Generate Again !"
                GoTo errmsg
            End If
        End If

        ' Check Total Loss
        If ANO <> LANO Then
            rs = db.OpenQueryDef("QCLACC_6", adOpenForwardonly, ANO)
            If Not rs.Eof Then
                ErrMsg = "Already Has Claim Total Loss (ClaimNo : " & rs.fields("ClaimNo") & ") on " & Format(rs.fields("LossDate"), "d") & " !"
                GoTo errmsg
            End If
        End If

        ' Individual or Master Fleet or Detail Certificate
        If LANO = -1 Or (LANO <> -1 And ANO = LANO And Not OpenFlag) Or (ANO <> LANO And OpenFlag) Then
            If Not db.GetPrivileges("UNPAIDPREMIUM") Then
                If isNotPremiumPaid(ANO, Today.Date) Then
                    Dim RequestingUser As String = ""
                    If Not (CheckRequestingActivities("UNPAIDPREMIUM", "UND", ANO, RegDate, "", RequestingUser) AndAlso RequestingUser.ToUpper = dbUser.ToUpper) Then
                        ErrMsg = "Can't Renew The Policy which is Not Yet Paid !"
                        GoTo errmsg
                    End If
                End If
            End If
        End If

        CheckRenewalData = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Renewal Policy : " & ErrMsg)
        CheckRenewalData = False
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function RenewPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Dim ANO As Long
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim LANO As Long
            Dim CNO As Long
            Dim OpenFlag As Boolean
            Dim RenewalTo As Integer
            Dim ExpiredOnRenewal As Boolean
            Dim TOC As String

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)
            CNO = getInParamValue(InParam, 2)
            OpenFlag = getInParamValue(InParam, 3)
            TOC = getInParamValue(InParam, 11) ' TOC

            Dim Param(27) As Object
            Param(0) = ""           ' New Quotation No / CoverNo
            Param(1) = ""           ' New CI No. / RegNo
            Param(2) = getInParamValue(InParam, 4)   ' New Policy No
            Param(3) = getInParamValue(InParam, 5)   ' New Start Date
            Param(4) = getInParamValue(InParam, 6)   ' New Expiry Date
            Param(5) = getInParamValue(InParam, 7)   ' SType
            Param(6) = getInParamValue(InParam, 8)  ' Remarks
            Param(7) = getInParamValue(InParam, 9)  ' Memo
            Param(8) = ""
            Param(9) = ""
            Param(10) = ""
            Param(11) = getInParamValue(InParam, 10) ' Branch
            Param(12) = TOC
            Param(13) = getInParamValue(InParam, 12) ' ASource
            RenewalTo = toval(getInParamValue(InParam, 13)) ' 0 : Default, 1 : Individual
            ExpiredOnRenewal = toval(getInParamValue(InParam, 14)) ' True/False
            Param(14) = toval(getInParamValue(InParam, 15)) ' Apply Default CT
            Param(15) = toval(getInParamValue(InParam, 16)) ' Apply Decreasing Value Of Sum Insured
            Param(16) = toval(getInParamValue(InParam, 17))  'With Detail Flag
            Param(17) = toval(getInParamValue(InParam, 18))  'Apply Default Fee and Duty
            Param(18) = toval(getInParamValue(InParam, 19))  'Apply Default Period Term
            Param(19) = toval(getInParamValue(InParam, 20))  'Apply Default Object Valuation
            Param(20) = toval(getInParamValue(InParam, 21))  'Immediate Inforce
            Param(21) = toval(getInParamValue(InParam, 22))  'Auto Inforce Process Flag
            Param(22) = toval(getInParamValue(InParam, 23))  'Auto Print Flag
            Param(23) = tostr(getInParamValue(InParam, 24))  'Auto Print Version
            Param(24) = tostr(getInParamValue(InParam, 25))  'CertificateNo
            Param(25) = toval(getInParamValue(InParam, 26))  'Apply Without Busines Source
            Param(26) = toval(getInParamValue(InParam, 27))  'Apply Rate Table
            Param(27) = toval(getInParamValue(InParam, 28))  'Skip Detail with Total Loss

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If LANO = -1 Or OpenFlag Then
                If Not CheckRenewalData(ANO, TOC, ErrMsg) Then
                    Throw New Exception(ErrMsg)
                End If
            Else    ' Fleet
                If Not CheckRenewalData(ANO, TOC, ErrMsg) Then
                    Throw New Exception(ErrMsg)
                End If
                Dim rs As RemoteSet
                Dim SQL As String
                SQL = "Select ANO from Acceptance Where AStatus in ('I','C') and AType<>'C' and AType<>'A' and RenewableF=1 and ANO<>LANO and LANO=" & LANO & " "
                SQL = SQL & " and Not Exists (Select ANO from Renewal Where Renewal.ANO=Acceptance.ANO and ((Renewal.RenewalF=0 and Renewal.Action=0) or Renewal.Action=3)) "
                If toval(Param(27)) Then
                    SQL = SQL & " and Not Exists (Select Top 1 A.OANO from Acceptance A Inner Join Claim C1 on (C1.ANO=A.ANO) Inner Join CLLOSS C2 on (C2.CNO=C1.CNO) Where A.OANO=Acceptance.OANO and C1.CType<>'C' and (C1.CSTATUS IN ('P','T','A') OR C1.PCSTATUS IN ('A')) and not exists (Select pCNO from Claim Where Claim.pCNO=C1.CNO and Claim.CStatus<>'R') and C2.CTLF=1)"
                End If
                SQL = SQL & " Order By CertificateNo "
                rs = db.OpenRemoteSet(SQL)
                Do While Not rs.Eof
                    If Not CheckRenewalData(rs.fields("ANO"), TOC, ErrMsg) Then
                        Throw New Exception(ErrMsg)
                    End If
                    rs.MoveNext()
                Loop
            End If

            If ExpiredOnRenewal Then
                If PreExpireConditionsOk(ANO, ErrMsg) Then
                    If ANO = LANO And Not OpenFlag Then
                        If Not Me.Direct_Call_Ext_Process("Process.NetV3.Proclib", "EXPIREMASTER", ANO, CNO, Today.Date) Then
                            Throw New Exception(Me.ErrDesc)
                        End If
                    Else
                        If Not Me.Direct_Call_Ext_Process("Process.NetV3.Proclib", "EXPIRE", ANO, CNO, Today.Date) Then
                            Throw New Exception(Me.ErrDesc)
                        End If
                    End If
                Else
                    Throw New Exception(ErrMsg)
                End If
            End If

            Dim NewANO As Long
            If ANO = LANO Or LANO = -1 Then
                Dim Result As Object = Nothing
                Result = Me.Direct_Call_Ext_Process("ProcessExt.NetV3.Proclib", "Renewal", CNO, ANO, Param)
                If Result Is Nothing Then
                    Throw New Exception("Background Process is running !")
                Else
                    NewANO = toval(Result)
                    If NewANO = 0 Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            ElseIf OpenFlag Then
                If RenewalTo = 0 Then  ' Default
                    Dim rs As RemoteSet
                    rs = db.OpenRemoteSet("Select C.CoverNo, A.RegNo, A.PolicyNo, mA.CNO as M_CNO from Acceptance A Left Join Acceptance mA on (mA.ANO=A.LANO), Cover C where C.CNO=A.CNO and A.ANO=" & ANO)
                    Param(0) = rs.fields("CoverNo")
                    Param(1) = rs.fields("RegNo")
                    Param(2) = rs.fields("PolicyNo")
                    NewANO = Me.Direct_Call_Ext_Process("ProcessExt.NetV3.Proclib", "RenewalDetail", ANO, rs.fields("M_CNO"), LANO, Param)
                Else    ' Individual
                    Param(5) = "S"
                    Dim Result As Object = Nothing
                    Result = Me.Direct_Call_Ext_Process("ProcessExt.NetV3.Proclib", "Renewal", CNO, ANO, Param)
                    If Result Is Nothing Then
                        Throw New Exception("Background Process is running !")
                    Else
                        NewANO = toval(Result)
                    End If
                End If
                If NewANO = 0 Then
                    Throw New Exception(Me.ErrDesc)
                End If
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = NewANO

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            If db.CheckLockedObject("Policy" & ANO) Then db.ReleaseLockedObject("Policy" & ANO)
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ConvertPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long
            Dim CNO As Long
            Dim PolicyType As String
            Dim ISType As String
            Dim TOC As String

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)
            CNO = getInParamValue(InParam, 2)
            PolicyType = getInParamValue(InParam, 3) ' New PolicyType
            ISType = getInParamValue(InParam, 4)     ' New Insurance Source Type
            TOC = getInParamValue(InParam, 5)     ' New Sub Class of Business


            Dim rs As RemoteSet

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            rs = db.OpenRemoteSet("Select AType, TOC from Acceptance Where AStatus='W' and ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Please Check Policy Data Status !"
                Throw New Exception(ErrMsg)
            End If
            Dim AType As String
            Dim ASource As String
            AType = rs.fields("AType")

            If TOC = "" Then TOC = rs.fields("TOC")

            rs = db.OpenRemoteSet("Select ASource from Cover Where CNO=" & CNO)
            ASource = rs.fields("ASource")

            Dim SType As String

            rs = db.OpenRemoteSet("Select PolicyType.SType from Cover, PolicyType Where PolicyType.PolicyType=Cover.PolicyType and Cover.CNO=" & CNO)
            SType = rs.fields("SType")

            Dim NewANO As Long
            rs = db.OpenRemoteSet("Select SType from PolicyType Where PolicyType='" & PolicyType & "'")
            If tostr(rs.fields("SType")) = SType And (LANO = ANO Or LANO = -1) And (AType = "N" Or (AType = "E" And ASource = "MKT")) Then
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Convert_PolicyType", ANO, LANO, CNO, PolicyType, ISType, TOC)
            ElseIf SType = "S" And tostr(rs.fields("SType")) = "M" And (AType = "N" Or (AType = "E" And ASource = "MKT")) Then
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Convert_To_Master", ANO, LANO, CNO, PolicyType, ISType, TOC)
            ElseIf SType = "M" And LANO <> ANO And tostr(rs.fields("SType")) = "S" And (AType = "N") Then
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Convert_To_Individual", ANO, LANO, CNO, PolicyType, ISType, TOC)
            ElseIf SType = "M" And LANO = ANO And tostr(rs.fields("SType")) = "O" And (AType = "N" Or (AType = "E" And ASource = "MKT")) Then
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Convert_PolicyType", ANO, LANO, CNO, PolicyType, ISType, TOC)
            ElseIf SType = "O" And LANO = ANO And tostr(rs.fields("SType")) = "M" And (AType = "N" Or (AType = "E" And ASource = "MKT")) Then
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Convert_PolicyType", ANO, LANO, CNO, PolicyType, ISType, TOC)
            Else
                ErrMsg = "Conversion Not Available !"
                Throw New Exception(ErrMsg)
            End If
            If NewANO = 0 Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = NewANO

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function CopyPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long
            Dim CNO As Long
            Dim ISType As String

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)
            CNO = getInParamValue(InParam, 2)
            ISType = getInParamValue(InParam, 3) ' New Insurance Source Type

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            Dim NewANO As Long
            If ANO = LANO Then      ' Master Policy
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "COPYPOLICYMASTER", CNO, ANO, ISType)
            ElseIf LANO = -1 Then   ' Individual Policy
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "COPYPOLICY", CNO, ANO, ISType)
            Else                    ' Detail Policy
                NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "COPYCERTIFICATE", ANO, LANO, CNO)
            End If
            If NewANO = 0 Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = NewANO

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function ApplyMasterCoverage(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Dim ANO As Long
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim LANO As Long
            Dim sFilterCode As String

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)
            sFilterCode = getInParamValue(InParam, 2)  ' value : ('C01','C02')

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If ANO = LANO Then
                ' Execute the Process Here
                Dim Result As Object = Nothing
                Result = Me.Direct_Call_Ext_Process("ProcessExt.NetV3.Proclib", "APPLY_MASTER_RISKCOVERAGE", ANO, sFilterCode)
                If Result Is Nothing Then
                    Throw New Exception("Background Process is running !")
                Else
                    If Not Result Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            End If

            ' Set Output Parameter

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            If db.CheckLockedObject("Policy" & ANO) Then db.ReleaseLockedObject("Policy" & ANO)
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function ApplyMasterDeductible(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If ANO = LANO Then
                ' Execute the Process Here
                Dim Result As Object = Nothing
                Result = Me.Direct_Call_Ext_Process("ProcessExt.NetV3.Proclib", "APPLY_MASTER_DEDUCTIBLE", ANO)
                If Result Is Nothing Then
                    Throw New Exception("Background Process is running !")
                Else
                    If Not Result Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            End If

            ' Set Output Parameter

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ApplyMasterInterest(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Dim ANO As Long
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim LANO As Long

            ANO = getInParamValue(InParam, 0)
            LANO = getInParamValue(InParam, 1)

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If ANO = LANO Then
                ' Execute the Process Here
                Dim Result As Object = Nothing
                Result = Me.Direct_Call_Ext_Process("ProcessExt.NetV3.Proclib", "Apply_Master_Interest", ANO)
                If Result Is Nothing Then
                    Throw New Exception("Background Process is running !")
                Else
                    If Not Result Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            End If

            ' Set Output Parameter

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            If db.CheckLockedObject("Policy" & ANO) Then db.ReleaseLockedObject("Policy" & ANO)
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function AdjustPolicy(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long
            Dim PDate As Date
            Dim OpenFlag As Boolean

            ANO = getInParamValue(InParam, 0)
            PDate = getInParamValue(InParam, 3)
            OpenFlag = getInParamValue(InParam, 4)
            LANO = getInParamValue(InParam, 5)

            Dim Param(12) As Object
            Param(0) = getInParamValue(InParam, 1)       ' RegNo
            Param(1) = getInParamValue(InParam, 2)       ' Booking Date / ADate
            Param(2) = getInParamValue(InParam, 3)       ' Effective Date / PDate
            Param(3) = "U"              ' EMethod : "U"
            Param(4) = getInParamValue(InParam, 5)       ' LANO
            Param(5) = getInParamValue(InParam, 6)       ' Remarks
            Param(6) = getInParamValue(InParam, 7)       ' CPCT
            Param(7) = getInParamValue(InParam, 8)       ' Description
            Param(8) = getInParamValue(InParam, 9)       ' ""
            Param(9) = getInParamValue(InParam, 10)      ' ""
            Param(10) = getInParamValue(InParam, 11)     ' ""
            Param(11) = getInParamValue(InParam, 12)     ' Branch
            Param(12) = getInParamValue(InParam, 13)     ' TOC

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not PreAdjustConditionsOK(ANO, PDate, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            Dim NewANO As Long
            If ANO = LANO And Not OpenFlag Then
                Dim Result As Object = Nothing
                Result = Me.Direct_Call_Ext_Process("ProcessExt.NetV3.Proclib", "AdjustMaster", ANO, Param)
                If Result Is Nothing Then
                    Throw New Exception("Background Process is running !")
                Else
                    NewANO = toval(Result)
                    If NewANO = 0 Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            Else
                If ANO <> LANO Then
                    NewANO = Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Adjustment", ANO, Param)
                    If NewANO = 0 Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = NewANO

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function Convert_FLDID_Into_Array(ByVal FLDID As String) As String()
        Dim SFLDID() As String
        Dim Separator As String

        If InStr(FLDID, "|") <> 0 Then
            Separator = "|"
        ElseIf InStr(FLDID, "/") <> 0 Then
            Separator = "/"
        Else
            Separator = ""
        End If
        Dim i As Integer
        i = 0
        Do While Separator <> ""
            ReDim Preserve SFLDID(i)

            SFLDID(i) = Trim(MID(FLDID, 1, InStr(FLDID, Separator) - 1))

            FLDID = MID(FLDID, InStr(FLDID, Separator) + 1, Len(FLDID))
            If InStr(FLDID, "|") <> 0 Then
                Separator = "|"
            ElseIf InStr(FLDID, "/") <> 0 Then
                Separator = "/"
            Else
                Separator = ""
            End If
            i = i + 1
        Loop
        ReDim Preserve SFLDID(i)
        SFLDID(i) = FLDID
        Convert_FLDID_Into_Array = SFLDID
    End Function
    Private Function isDuplicateInsured(ByVal CoverNo As String, ByVal PolicyNo As String, ByVal CertificateNo As String, ByVal RefNo As String, ByVal AName As String, ByVal RegID As String, ByVal FLDID As String, ByVal ValueID As String, ByVal PDate As Date, ByVal PPDate As Date, Optional ByVal ANO As Long = 0, Optional ByVal ExclusiveF As Boolean = False) As Boolean
        Dim rs As RemoteSet
        Dim SQL As String
        Dim sPDate As String
        Dim sPPDate As String
        sPDate = Format(PDate, "MM/DD/YYYY")
        sPPDate = Format(PPDate, "MM/DD/YYYY")

        Dim AFLDID() As String
        Dim AVALUEID() As String
        Dim EachAccRegF As Boolean
        Dim UniqueRefNoF As Boolean
        Dim MaxNumDuplReg As Integer

        isDuplicateInsured = False

        If ANO > 0 Then
            Dim LANO As Long
            rs = db.OpenRemoteSet("Select A.LANO, TOC.EachAccRegF, TOC.UniqueRefNoF, TOC.MaxNumDuplReg from Acceptance A, TOC Where TOC.TOC=A.TOC and A.ANO=" & ANO)
            LANO = toval(rs.fields("LANO"))
            EachAccRegF = toval(rs.fields("EachAccRegF"))
            UniqueRefNoF = toval(rs.fields("UniqueRefNoF"))
            MaxNumDuplReg = toval(rs.fields("MaxNumDuplReg"))
            If ANO <> LANO Then
                If EachAccRegF Then
                    AFLDID = Convert_FLDID_Into_Array(FLDID)
                    AVALUEID = Convert_FLDID_Into_Array(ValueID)
                Else
                    ReDim AFLDID(0)
                    ReDim AVALUEID(0)
                    AFLDID(0) = FLDID
                    AVALUEID(0) = ValueID
                End If

                For i = 0 To AFLDID.GetUpperBound(0)
                    SQL = "Select Top " & tostr(MaxNumDuplReg + 1) & " C1.CoverNo as QuotationNo, A1.RegNo as CINo, A1.PolicyNo+case when A1.CertificateNo='' then '' else '-'+A1.CertificateNo end as PolicyNo, A1.AName as Insured, A1.SDate as Inception, A1.EDate as Expiry, A1.ANO   "
                    SQL = SQL & " from Cover C1, Acceptance A1, Cover C2, Acceptance A2, TOC    "
                    SQL = SQL & " Where A2.ANO = " & ANO & " And A2.ANO <> A2.LANO And C2.CNO = A2.CNO And TOC.TOC = A2.TOC "
                    SQL = SQL & " and A2.AType not in ('A','C') and A1.AType not in ('A','C') "
                    SQL = SQL & " and A1.ANO<>A1.LANO and A1.ANO<>A2.RANO and A1.OANO<>A2.OANO and C1.CNO=A1.CNO and Not (A1.AType='N' and A1.AStatus='C' and ((C1.CStatus='A' and C1.ASource='MKT') or C1.CStatus='C'))  "
                    SQL = SQL & " and A1.PDate<A1.PPDate "
                    SQL = SQL & " and ("
                    SQL = SQL & " (A1.pDate<='" & sPDate & "' and ((A1.PPDate>'" & sPDate & "' ) or (A1.PPDate>='" & sPDate & "' and A1.ExclusiveF=0 and A2.ExclusiveF=1))) Or "
                    SQL = SQL & " ('" & sPDate & "'<=A1.pDate and (('" & sPPDate & "'>A1.PDate ) or ('" & sPPDate & "'>=A1.PDate and A2.ExclusiveF=0 and A1.ExclusiveF=1)))"
                    SQL = SQL & " ) "

                    If EachAccRegF Then
                        SQL = SQL & " and (A1.REGID='" & RegID & "' and A1.FLDID<>'' and charindex('" & AFLDID(i) & "',A1.FLDID) <>0 and A1.ValueID<>'' and charindex('|" & AVALUEID(i) & "|','|'+A1.ValueID+'|') <>0) "
                    Else
                        SQL = SQL & " and (A1.REGID='" & RegID & "' and A1.FLDID<>'' and A1.FLDID='" & AFLDID(i) & "' and A1.ValueID<>'' and A1.ValueID='" & AVALUEID(i) & "') "
                    End If

                    rs = db.OpenRemoteSet(SQL)
                    isDuplicateInsured = (rs.RecordCount = MaxNumDuplReg + 1)
                    If isDuplicateInsured Then Exit For
                Next

                If Not isDuplicateInsured AndAlso UniqueRefNoF Then
                    SQL = "Select Top " & tostr(MaxNumDuplReg + 1) & " C1.CoverNo as QuotationNo, A1.RegNo as CINo, A1.PolicyNo+case when A1.CertificateNo='' then '' else '-'+A1.CertificateNo end as PolicyNo, A1.AName as Insured, A1.SDate as Inception, A1.EDate as Expiry, A1.ANO   "
                    SQL = SQL & " from Cover C1, Acceptance A1, Cover C2, Acceptance A2, TOC    "
                    SQL = SQL & " Where A2.ANO = " & ANO & " And A2.ANO <> A2.LANO And C2.CNO = A2.CNO And TOC.TOC = A2.TOC "
                    SQL = SQL & " and A2.AType not in ('A','C') and A1.AType not in ('A','C') "
                    SQL = SQL & " and A1.ANO<>A1.LANO and A1.ANO<>A2.RANO and A1.OANO<>A2.OANO and C1.CNO=A1.CNO and Not (A1.AType='N' and A1.AStatus='C' and ((C1.CStatus='A' and C1.ASource='MKT') or C1.CStatus='C'))  "
                    SQL = SQL & " and A1.PDate<A1.PPDate "
                    SQL = SQL & " and ("
                    SQL = SQL & " (A1.pDate<='" & sPDate & "' and ((A1.PPDate>'" & sPDate & "' ) or (A1.PPDate>='" & sPDate & "' and A1.ExclusiveF=0 and A2.ExclusiveF=1))) Or "
                    SQL = SQL & " ('" & sPDate & "'<=A1.pDate and (('" & sPPDate & "'>A1.PDate ) or ('" & sPPDate & "'>=A1.PDate and A2.ExclusiveF=0 and A1.ExclusiveF=1)))"
                    SQL = SQL & " ) "
                    SQL = SQL & " and (A1.TOC=A2.TOC and A1.RefNo=A2.RefNo)"

                    rs = db.OpenRemoteSet(SQL)
                    isDuplicateInsured = (rs.RecordCount = MaxNumDuplReg + 1)
                End If
            Else
                If EachAccRegF Then
                    AFLDID = Convert_FLDID_Into_Array(FLDID)
                    AVALUEID = Convert_FLDID_Into_Array(ValueID)
                Else
                    ReDim AFLDID(0)
                    ReDim AVALUEID(0)
                    AFLDID(0) = FLDID
                    AVALUEID(0) = ValueID
                End If

                For i = 0 To AFLDID.GetUpperBound(0)
                    SQL = "Select Top " & tostr(MaxNumDuplReg + 1) & " C1.CoverNo as QuotationNo, A1.RegNo as CINo, A1.PolicyNo+case when A1.CertificateNo='' then '' else '-'+A1.CertificateNo end as PolicyNo, A1.AName as Insured, A1.SDate as Inception, A1.EDate as Expiry, A1.ANO   "
                    SQL = SQL & " from Cover C1, Acceptance A1, Cover C2, Acceptance A2, TOC    "
                    SQL = SQL & " Where A2.ANO = " & ANO & " And A2.ANO = A2.LANO And C2.CNO = A2.CNO And TOC.TOC = A2.TOC "
                    SQL = SQL & " and A2.AType not in ('A','C') and A1.AType not in ('A','C') "
                    SQL = SQL & " and A1.ANO=A1.LANO and A1.ANO<>A2.RANO and A1.OANO<>A2.OANO and C1.CNO=A1.CNO and Not (A1.AType='N' and A1.AStatus='C' and ((C1.CStatus='A' and C1.ASource='MKT') or C1.CStatus='C'))  "
                    SQL = SQL & " and A1.PDate<A1.PPDate "
                    SQL = SQL & " and ("
                    SQL = SQL & " (A1.pDate<='" & sPDate & "' and ((A1.PPDate>'" & sPDate & "' ) or (A1.PPDate>='" & sPDate & "' and A1.ExclusiveF=0 and A2.ExclusiveF=1))) Or "
                    SQL = SQL & " ('" & sPDate & "'<=A1.pDate and (('" & sPPDate & "'>A1.PDate ) or ('" & sPPDate & "'>=A1.PDate and A2.ExclusiveF=0 and A1.ExclusiveF=1)))"
                    SQL = SQL & " ) "

                    If EachAccRegF Then
                        SQL = SQL & " and (A1.REGID='" & RegID & "' and A1.FLDID<>'' and charindex('" & AFLDID(i) & "',A1.FLDID) <>0 and A1.ValueID<>'' and charindex('|" & AVALUEID(i) & "|','|'+A1.ValueID+'|') <>0) "
                    Else
                        SQL = SQL & " and (A1.REGID='" & RegID & "' and A1.FLDID<>'' and A1.FLDID='" & AFLDID(i) & "' and A1.ValueID<>'' and A1.ValueID='" & AVALUEID(i) & "') "
                    End If

                    rs = db.OpenRemoteSet(SQL)
                    isDuplicateInsured = (rs.RecordCount = MaxNumDuplReg + 1)
                    If isDuplicateInsured Then Exit For
                Next

                If Not isDuplicateInsured AndAlso UniqueRefNoF Then
                    SQL = "Select Top " & tostr(MaxNumDuplReg + 1) & " C1.CoverNo as QuotationNo, A1.RegNo as CINo, A1.PolicyNo+case when A1.CertificateNo='' then '' else '-'+A1.CertificateNo end as PolicyNo, A1.AName as Insured, A1.SDate as Inception, A1.EDate as Expiry, A1.ANO   "
                    SQL = SQL & " from Cover C1, Acceptance A1, Cover C2, Acceptance A2, TOC    "
                    SQL = SQL & " Where A2.ANO = " & ANO & " And A2.ANO=A2.LANO And C2.CNO = A2.CNO And TOC.TOC = A2.TOC "
                    SQL = SQL & " and A2.AType not in ('A','C') and A1.AType not in ('A','C') "
                    SQL = SQL & " and A1.ANO=A1.LANO and A1.ANO<>A2.RANO and A1.OANO<>A2.OANO and C1.CNO=A1.CNO and Not (A1.AType='N' and A1.AStatus='C' and ((C1.CStatus='A' and C1.ASource='MKT') or C1.CStatus='C')) "
                    SQL = SQL & " and A1.PDate<A1.PPDate "
                    SQL = SQL & " and ("
                    SQL = SQL & " (A1.pDate<='" & sPDate & "' and ((A1.PPDate>'" & sPDate & "' ) or (A1.PPDate>='" & sPDate & "' and A1.ExclusiveF=0 and A2.ExclusiveF=1))) Or "
                    SQL = SQL & " ('" & sPDate & "'<=A1.pDate and (('" & sPPDate & "'>A1.PDate ) or ('" & sPPDate & "'>=A1.PDate and A2.ExclusiveF=0 and A1.ExclusiveF=1)))"
                    SQL = SQL & " ) "
                    SQL = SQL & " and (A1.TOC=A2.TOC and A1.RefNo=A2.RefNo) "
                    rs = db.OpenRemoteSet(SQL)
                    isDuplicateInsured = (rs.RecordCount = MaxNumDuplReg + 1)
                End If
            End If
        Else
            rs = db.OpenRemoteSet("Select TOP 1 TOC.EachAccRegF, TOC.UniqueRefNoF, TOC.MaxNumDuplReg from TOC Where REGID='" & RegID & "'")
            EachAccRegF = toval(rs.fields("EachAccRegF"))
            UniqueRefNoF = toval(rs.fields("UniqueRefNoF"))
            MaxNumDuplReg = toval(rs.fields("MaxNumDuplReg"))
            If EachAccRegF Then
                AFLDID = Convert_FLDID_Into_Array(FLDID)
                AVALUEID = Convert_FLDID_Into_Array(ValueID)
            Else
                ReDim AFLDID(0)
                ReDim AVALUEID(0)
                AFLDID(0) = FLDID
                AVALUEID(0) = ValueID
            End If

            For i = 0 To AFLDID.GetUpperBound(0)
                SQL = "Select Top " & tostr(MaxNumDuplReg + 1) & " C1.CoverNo as QuotationNo, A1.RegNo as CINo, A1.PolicyNo+case when A1.CertificateNo='' then '' else '-'+A1.CertificateNo end as PolicyNo, A1.AName as Insured, A1.SDate as Inception, A1.EDate as Expiry, A1.ANO  "
                SQL = SQL & " from Cover C1, Acceptance A1, TOC"
                SQL = SQL & " Where C1.CNO=A1.CNO and TOC.TOC=A1.TOC and A1.AType not in ('A','C') "
                SQL = SQL & " and Not (A1.AType='N' and A1.AStatus='C' and ((C1.CStatus='A' and C1.ASource='MKT') or C1.CStatus='C'))"
                SQL = SQL & " and (C1.CoverNo<>'" & CoverNo & "' or A1.PolicyNo<>'" & PolicyNo & "' or A1.CertificateNo<>'" & CertificateNo & "')"
                SQL = SQL & " and A1.PDate<A1.PPDate "
                SQL = SQL & " and ("
                SQL = SQL & " (A1.pDate<='" & sPDate & "' and ((A1.PPDate>'" & sPDate & "') or (A1.PPDate>='" & sPDate & "' and A1.ExclusiveF=0 and '" & IIf(ExclusiveF, 1, 0) & "' =1))) Or "
                SQL = SQL & " ('" & sPDate & "'<=A1.pDate and (('" & sPPDate & "'>A1.PDate) or ('" & sPPDate & "'>=A1.PDate and '" & IIf(ExclusiveF, 1, 0) & "' =0 and A1.ExclusiveF=1)))"
                SQL = SQL & " ) "

                If EachAccRegF Then
                    SQL = SQL & " and (A1.REGID='" & RegID & "' and A1.FLDID<>'' and charindex('" & AFLDID(i) & "',A1.FLDID) <>0 and A1.ValueID<>'' and charindex('|" & AVALUEID(i) & "|','|'+A1.ValueID+'|') <>0) "
                Else
                    SQL = SQL & " and (A1.REGID='" & RegID & "' and A1.FLDID<>'' and A1.FLDID='" & AFLDID(i) & "' and A1.ValueID<>'' and A1.ValueID='" & AVALUEID(i) & "') "
                End If

                rs = db.OpenRemoteSet(SQL)
                isDuplicateInsured = (rs.RecordCount = MaxNumDuplReg + 1)
                If isDuplicateInsured Then Exit For
            Next

            If Not isDuplicateInsured AndAlso UniqueRefNoF Then
                SQL = "Select Top " & tostr(MaxNumDuplReg + 1) & " C1.CoverNo as QuotationNo, A1.RegNo as CINo, A1.PolicyNo+case when A1.CertificateNo='' then '' else '-'+A1.CertificateNo end as PolicyNo, A1.AName as Insured, A1.SDate as Inception, A1.EDate as Expiry, A1.ANO  "
                SQL = SQL & " from Cover C1, Acceptance A1, TOC"
                SQL = SQL & " Where C1.CNO=A1.CNO and TOC.TOC=A1.TOC and A1.AType not in ('A','C') "
                SQL = SQL & " and Not (A1.AType='N' and A1.AStatus='C' and ((C1.CStatus='A' and C1.ASource='MKT') or C1.CStatus='C'))"
                SQL = SQL & " and (C1.CoverNo<>'" & CoverNo & "' or A1.PolicyNo<>'" & PolicyNo & "' or A1.CertificateNo<>'" & CertificateNo & "')"
                SQL = SQL & " and A1.PDate<A1.PPDate "
                SQL = SQL & " and ("
                SQL = SQL & " (A1.pDate<='" & sPDate & "' and ((A1.PPDate>'" & sPDate & "') or (A1.PPDate>='" & sPDate & "' and A1.ExclusiveF=0 and '" & IIf(ExclusiveF, 1, 0) & "' =1))) Or "
                SQL = SQL & " ('" & sPDate & "'<=A1.pDate and (('" & sPPDate & "'>A1.PDate) or ('" & sPPDate & "'>=A1.PDate and '" & IIf(ExclusiveF, 1, 0) & "' =0 and A1.ExclusiveF=1)))"
                SQL = SQL & " ) "
                SQL = SQL & " and (TOC.RegID='" & RegID & "' and A1.RefNo='" & RefNo & "')"
                rs = db.OpenRemoteSet(SQL)
                isDuplicateInsured = (rs.RecordCount = MaxNumDuplReg + 1)
            End If
        End If
    End Function


    <WebMethod(EnableSession:=False)> _
    Public Function GetListDuplicateInsured(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim CoverNo As String = getInParamValue(InParam, 0)
            Dim PolicyNo As String = getInParamValue(InParam, 1)
            Dim CertificateNo As String = getInParamValue(InParam, 2)
            Dim AName As String = getInParamValue(InParam, 3)
            Dim RegID As String = getInParamValue(InParam, 4)
            Dim FLDID As String = getInParamValue(InParam, 5)
            Dim ValueID As String = getInParamValue(InParam, 6)
            Dim PDate As Date = getInParamValue(InParam, 7)
            Dim PPDate As Date = getInParamValue(InParam, 8)
            Dim ANO As Long = toval(getInParamValue(InParam, 9))
            Dim RefNo As String = getInParamValue(InParam, 10)
            Dim ExclusiveF As Boolean = toval(getInParamValue(InParam, 11))

            Dim SQL As String
            Dim rs As RemoteSet
            Dim sPDate As String
            Dim sPPDate As String
            sPDate = Format(PDate, "MM/DD/YYYY")
            sPPDate = Format(PPDate, "MM/DD/YYYY")

            Dim AFLDID() As String
            Dim AVALUEID() As String
            Dim EachAccRegF As Boolean
            If ANO > 0 Then
                Dim LANO As Long
                rs = db.OpenRemoteSet("Select A.LANO, TOC.EachAccRegF from Acceptance A, TOC Where TOC.TOC=A.TOC and A.ANO=" & ANO)
                LANO = toval(rs.fields("LANO"))
                EachAccRegF = toval(rs.fields("EachAccRegF"))
                If ANO <> LANO Then
                    If EachAccRegF Then
                        AFLDID = Convert_FLDID_Into_Array(FLDID)
                        AVALUEID = Convert_FLDID_Into_Array(ValueID)
                    Else
                        ReDim AFLDID(0)
                        ReDim AVALUEID(0)
                        AFLDID(0) = FLDID
                        AVALUEID(0) = ValueID
                    End If

                    For i = 0 To AFLDID.GetUpperBound(0)
                        SQL = "Select C1.CoverNo as QuotationNo, A1.RegNo as CINo, A1.PolicyNo+case when A1.CertificateNo='' then '' else '-'+A1.CertificateNo end as PolicyNo, A1.AName as Insured, A1.SDate as Inception, A1.EDate as Expiry, A1.ANO   "
                        SQL = SQL & " from Cover C1, Acceptance A1, Cover C2, Acceptance A2, TOC    "
                        SQL = SQL & " Where A2.ANO = " & ANO & " And A2.ANO <> A2.LANO And C2.CNO = A2.CNO And TOC.TOC = A2.TOC "
                        SQL = SQL & " and A2.AType not in ('A','C') and A1.AType not in ('A','C') "
                        SQL = SQL & " and A1.ANO<>A1.LANO and A1.ANO<>A2.RANO and A1.OANO<>A2.OANO and C1.CNO=A1.CNO and Not (A1.AType='N' and A1.AStatus='C' and ((C1.CStatus='A' and C1.ASource='MKT') or C1.CStatus='C'))  "

                        SQL = SQL & " and A1.PDate<A1.PPDate "
                        SQL = SQL & " and ("
                        SQL = SQL & " (A1.pDate<='" & sPDate & "' and ((A1.PPDate>'" & sPDate & "') or (A1.PPDate>='" & sPDate & "' and A1.ExclusiveF=0 and A2.ExclusiveF=1))) or "
                        SQL = SQL & " ('" & sPDate & "'<=A1.PDate and (('" & sPPDate & "'>A1.PDate) or ('" & sPPDate & "'>=A1.PDate and A2.ExclusiveF=0 and A1.ExclusiveF=1)))"
                        SQL = SQL & " ) "

                        If EachAccRegF Then
                            SQL = SQL & " and ((A1.REGID='" & RegID & "' and A1.FLDID<>'' and charindex('" & AFLDID(i) & "',A1.FLDID) <>0 and A1.ValueID<>'' and charindex('|" & AVALUEID(i) & "|','|'+A1.ValueID+'|') <>0) "
                        Else
                            SQL = SQL & " and ((A1.REGID='" & RegID & "' and A1.FLDID<>'' and A1.FLDID='" & AFLDID(i) & "' and A1.ValueID<>'' and A1.ValueID='" & AVALUEID(i) & "') "
                        End If
                        SQL = SQL & " or (TOC.UniqueRefNoF=1 and A1.TOC=A2.TOC and A1.RefNo=A2.RefNo))"
                        rs = db.OpenRemoteSet(SQL)
                        Dataset = rs.GetDataSet
                        If Not rs.Eof Then Exit For
                    Next
                Else
                    If EachAccRegF Then
                        AFLDID = Convert_FLDID_Into_Array(FLDID)
                        AVALUEID = Convert_FLDID_Into_Array(ValueID)
                    Else
                        ReDim AFLDID(0)
                        ReDim AVALUEID(0)
                        AFLDID(0) = FLDID
                        AVALUEID(0) = ValueID
                    End If

                    For i = 0 To AFLDID.GetUpperBound(0)
                        SQL = "Select C1.CoverNo as QuotationNo, A1.RegNo as CINo, A1.PolicyNo+case when A1.CertificateNo='' then '' else '-'+A1.CertificateNo end as PolicyNo, A1.AName as Insured, A1.SDate as Inception, A1.EDate as Expiry, A1.ANO   "
                        SQL = SQL & " from Cover C1, Acceptance A1, Cover C2, Acceptance A2, TOC    "
                        SQL = SQL & " Where A2.ANO = " & ANO & " And A2.ANO = A2.LANO And C2.CNO = A2.CNO And TOC.TOC = A2.TOC "
                        SQL = SQL & " and A2.AType not in ('A','C') and A1.AType not in ('A','C') "
                        SQL = SQL & " and A1.ANO=A1.LANO and A1.ANO<>A2.RANO and A1.OANO<>A2.OANO and C1.CNO=A1.CNO and Not (A1.AType='N' and A1.AStatus='C' and ((C1.CStatus='A' and C1.ASource='MKT') or C1.CStatus='C'))  "

                        SQL = SQL & " and A1.PDate<A1.PPDate "
                        SQL = SQL & " and ("
                        SQL = SQL & " (A1.pDate<='" & sPDate & "' and ((A1.PPDate>'" & sPDate & "') or (A1.PPDate>='" & sPDate & "' and A1.ExclusiveF=0 and A2.ExclusiveF=1))) or "
                        SQL = SQL & " ('" & sPDate & "'<=A1.PDate and (('" & sPPDate & "'>A1.PDate) or ('" & sPPDate & "'>=A1.PDate and A2.ExclusiveF=0 and A1.ExclusiveF=1)))"
                        SQL = SQL & " ) "

                        If EachAccRegF Then
                            SQL = SQL & " and ((A1.REGID='" & RegID & "' and A1.FLDID<>'' and charindex('" & AFLDID(i) & "',A1.FLDID) <>0 and A1.ValueID<>'' and charindex('|" & AVALUEID(i) & "|','|'+A1.ValueID+'|') <>0) "
                        Else
                            SQL = SQL & " and ((A1.REGID='" & RegID & "' and A1.FLDID<>'' and A1.FLDID='" & AFLDID(i) & "' and A1.ValueID<>'' and A1.ValueID='" & AVALUEID(i) & "') "
                        End If
                        SQL = SQL & " or (TOC.UniqueRefNoF=1 and A1.TOC=A2.TOC and A1.RefNo=A2.RefNo))"
                        rs = db.OpenRemoteSet(SQL)
                        Dataset = rs.GetDataSet
                        If Not rs.Eof Then Exit For
                    Next
                End If

            Else
                rs = db.OpenRemoteSet("Select TOP 1 TOC.EachAccRegF from TOC Where REGID='" & RegID & "'")
                EachAccRegF = toval(rs.fields("EachAccRegF"))
                If EachAccRegF Then
                    AFLDID = Convert_FLDID_Into_Array(FLDID)
                    AVALUEID = Convert_FLDID_Into_Array(ValueID)
                Else
                    ReDim AFLDID(0)
                    ReDim AVALUEID(0)
                    AFLDID(0) = FLDID
                    AVALUEID(0) = ValueID
                End If

                For i = 0 To AFLDID.GetUpperBound(0)
                    SQL = "Select C1.CoverNo as QuotationNo, A1.RegNo as CINo, A1.PolicyNo+case when A1.CertificateNo='' then '' else '-'+A1.CertificateNo end as PolicyNo, A1.AName as Insured, A1.SDate as Inception, A1.EDate as Expiry, A1.ANO  "
                    SQL = SQL & " from Cover C1, Acceptance A1, TOC"
                    SQL = SQL & " Where C1.CNO=A1.CNO and TOC.TOC=A1.TOC and A1.AType not in ('A','C') "
                    SQL = SQL & " and Not (A1.AType='N' and A1.AStatus='C' and ((C1.CStatus='A' and C1.ASource='MKT') or C1.CStatus='C'))"
                    SQL = SQL & " and (C1.CoverNo<>'" & CoverNo & "' or A1.PolicyNo<>'" & PolicyNo & "' or A1.CertificateNo<>'" & CertificateNo & "')"
                    SQL = SQL & " and A1.PDate<A1.PPDate "
                    SQL = SQL & " and ("
                    SQL = SQL & " (A1.pDate<='" & sPDate & "' and ((A1.PPDate>'" & sPDate & "') or (A1.PPDate>='" & sPDate & "' and A1.ExclusiveF=0 and '" & IIf(ExclusiveF, 1, 0) & "' =1))) Or "
                    SQL = SQL & " ('" & sPDate & "'<=A1.pDate and (('" & sPPDate & "'>A1.PDate) or ('" & sPPDate & "'>=A1.PDate and '" & IIf(ExclusiveF, 1, 0) & "' =0 and A1.ExclusiveF=1)))"
                    SQL = SQL & " ) "

                    If EachAccRegF Then
                        SQL = SQL & " and ((A1.REGID='" & RegID & "' and A1.FLDID<>'' and charindex('" & AFLDID(i) & "',A1.FLDID) <>0 and A1.ValueID<>'' and charindex('|" & AVALUEID(i) & "|','|'+A1.ValueID+'|') <>0) "
                    Else
                        SQL = SQL & " and ((A1.REGID='" & RegID & "' and A1.FLDID<>'' and A1.FLDID='" & AFLDID(i) & "' and A1.ValueID<>'' and A1.ValueID='" & AVALUEID(i) & "') "
                    End If
                    SQL = SQL & " or (TOC.RegID='" & RegID & "' and TOC.UniqueRefNoF=1 and A1.RefNo='" & RefNo & "'))"
                    rs = db.OpenRemoteSet(SQL)
                    Dataset = rs.GetDataSet
                    If Not rs.Eof Then Exit For
                Next

            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    Private Function Get_UWYear(ByVal SDate As Date, ByVal EDate As Date, ByVal PDate As Date) As Integer
        Dim UWYear As Integer
        Dim PolicyYear As Integer

        If DateDiff(DateInterval.Day, PDate, EDate) < 0 Then
            PDate = EDate
        End If

        If DateDiff(DateInterval.Day, PDate, SDate) > 0 Then
            PDate = SDate
        End If

        PolicyYear = DateDiff(DateInterval.Year, SDate, PDate)
        If DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, PolicyYear, SDate), PDate) < 0 Then
            PolicyYear = PolicyYear - 1
        End If

        If DateDiff(DateInterval.Year, SDate, EDate) - 1 < PolicyYear And PolicyYear > 1 Then
            PolicyYear = DateDiff(DateInterval.Year, SDate, EDate) - 1
        End If

        UWYear = Year(SDate) + PolicyYear
        Get_UWYear = UWYear
    End Function
    Private Function PreInforceConditionsOK(ByVal ANO As Long, ByVal ADate As Date, ByVal PolicyNo As String, ByRef BackDatedF As Boolean, ByVal AdjustBookingDateF As Boolean, ByVal BucketF As Boolean, ByVal ReviewF As Boolean, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim Branch As String
        Dim TOC As String
        Dim OpenFlag As Boolean
        Dim PStatus As String
        Dim AStatus As String
        Dim CStatus As String
        Dim RANO As Long
        Dim PANO As Long
        Dim AType As String
        Dim pAType As String
        Dim EMethod As String
        Dim CNO As Long
        Dim ISType As String
        Dim SDate As Date
        Dim EDate As Date
        Dim PDate As Date
        Dim OANO As Long
        Dim SysUserID As String
        Dim LANO As Long
        Dim GracePeriod As Integer
        Dim ReqDate As Date
        Dim ReqTime As String
        Dim Installment As Integer
        Dim IDate_1 As Date

        On Error GoTo errmsg
        rs = db.OpenRemoteSet("Select coalesce(APayor.Installment, Payor.Installment) as Installment, coalesce(APayor.IDate_1, Payor.IDate_1) as IDate_1, coalesce(APayor.Grace, Payor.Grace) as Grace, Cover.CNO, Cover.ISType, Acceptance.LDate, Acceptance.LTime, Acceptance.LANO, Acceptance.SysUserID, Acceptance.OANO, Acceptance.TOC, Acceptance.RANO, Acceptance.PANO, Acceptance.SDate, Acceptance.EDate, Acceptance.PDate, Acceptance.Branch, Cover.OpenFlag, Acceptance.EMethod, Acceptance.AType, Acceptance.PStatus, Acceptance.AStatus, Cover.CStatus, pA.AType as pAType from Acceptance Left Join Acceptance pA on (pA.ANO=Acceptance.pANO) Left Join APayor on (APayor.ANO=Acceptance.ANO) Left Join Payor on (Payor.ANO=Acceptance.ANO) Inner Join Cover on (Cover.CNO=Acceptance.CNO) where Acceptance.ANO=" & ANO)
        Branch = tostr(rs.fields("Branch"))
        TOC = tostr(rs.fields("TOC"))
        OpenFlag = toval(rs.fields("OpenFlag"))
        PStatus = tostr(rs.fields("PStatus"))
        AStatus = tostr(rs.fields("AStatus"))
        CStatus = tostr(rs.fields("CStatus"))
        RANO = toval(rs.fields("RANO"))
        PANO = toval(rs.fields("PANO"))
        AType = tostr(rs.fields("AType"))
        pAType = tostr(rs.fields("pAType"))
        EMethod = tostr(rs.fields("EMethod"))
        CNO = toval(rs.fields("CNO"))
        ISType = tostr(rs.fields("ISType"))
        SDate = todate(rs.fields("SDate"))
        EDate = todate(rs.fields("EDate"))
        PDate = todate(rs.fields("PDate"))
        OANO = toval(rs.fields("OANO"))
        SysUserID = tostr(rs.fields("SysUserID"))
        LANO = toval(rs.fields("LANO"))
        GracePeriod = toval(rs.fields("Grace"))
        ReqDate = todate(rs.fields("LDate"))
        ReqTime = tostr(rs.fields("LTime"))
        Installment = toval(rs.fields("Installment"))
        IDate_1 = todate(rs.fields("IDate_1"))
        rs.CloseRemoteSet()

        If Not (CStatus = "A" And (AStatus = "W" Or AStatus = "T")) Then
            ErrMsg = "Please Check Policy Data Status !"
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select DueByEffectiveF from SysTab")
        Dim DueByEffectiveF As Boolean = toval(rs.fields("DueByEffectiveF"))
        Dim DueDate As Date

        If DueByEffectiveF Then
            DueDate = PDate
        Else
            DueDate = ADate
        End If

        If Not BucketF And Not ReviewF Then
            Dim EndDate As Date
            Dim LockedDate As Date
            Dim BLLocked As Integer
            EndDate = GetEndDate(Branch)
            LockedDate = GetLockedDate(Branch)
            BLLocked = GetUserBLLocked(db.Var_db_user)
            If DateDiff(DateInterval.Day, EndDate, ADate) <= 0 Or (DateDiff(DateInterval.Day, LockedDate, ADate) <= 0 And (DateDiff(DateInterval.Day, LockedDate, Today.Date) > BLLocked Or BLLocked = 0)) Then
                If BackDatedF Then
                    If DateDiff(DateInterval.Day, GetBackDatedDate(Branch), ADate) > 0 Then
                        ErrMsg = "Transactions for this period have been closed !"
                        GoTo errmsg
                    End If
                Else
                    If AdjustBookingDateF Then
                        If DateDiff(DateInterval.Day, EndDate, ADate) <= 0 Then
                            ADate = DateAdd(DateInterval.Day, 1, EndDate)
                        End If

                        If DateDiff(DateInterval.Day, LockedDate, ADate) <= 0 Then
                            ADate = DateAdd(DateInterval.Day, 1, LockedDate)
                        End If
                    Else
                        ErrMsg = "Transactions for this period have been closed !"
                        GoTo errmsg
                    End If
                End If
            Else
                BackDatedF = False
            End If
        End If

        Select Case ISType
            Case "C", "R"
                rs = db.OpenRemoteSet("Select CNO from ISource where CNO=" & CNO)
                If rs.Eof Then
                    ErrMsg = "Missing Insurance Source !"
                    GoTo errmsg
                End If
            Case "L"
                rs = db.OpenRemoteSet("Select CNO from Coinsurance where CNO=" & CNO)
                If rs.Eof Then
                    ErrMsg = "Missing Coinsurance Member !"
                    GoTo errmsg
                End If
        End Select

        If AType = "N" And Not BackDatedF Then
            If DateDiff(DateInterval.Day, SDate, PDate) <> 0 Then
                ErrMsg = "Inception date must be the same with Insurance Period !"
                GoTo errmsg
            End If
        Else
            If Not (AType = "C" And EMethod = "A") Then
                If DateDiff(DateInterval.Day, SDate, PDate) < 0 Then
                    ErrMsg = "Effective date can't be earlier than Insurance Period !"
                    GoTo errmsg
                End If
            End If
        End If

        If (AType = "N" Or (AType = "E" And pAType = "C")) And (ANO = LANO Or LANO = -1) And PolicyNo <> "" Then
            If toval(GetAppSetting("MaxLenghPolicyNo")) <> 0 AndAlso Len(PolicyNo) > toval(GetAppSetting("MaxLenghPolicyNo")) Then
                ErrMsg = "The supplied PolicyNo Exceeding Max Policy No Length (Max. " & tostr(GetAppSetting("MaxLenghPolicyNo")) & " char(s) length) !"
                GoTo errmsg
            Else
                rs = db.OpenRemoteSet("Select Top 1 ANO from Acceptance Where AType in ('N','E') and AStatus<>'C' and (LANO=-1 or LANO=ANO) and ANO<>" & RANO & " and ANO<>" & ANO & " and PolicyNo='" & PolicyNo & "'")
                If Not rs.Eof Then
                    ErrMsg = "Data Already Exists With The Same Policy No. !"
                    GoTo errmsg
                End If
            End If
        End If

        If AType <> "N" And Not BackDatedF And Not ReviewF Then
            If Not db.GetPrivileges("ALLOWBACKDATEDBOOKING") Then
                rs = db.OpenRemoteSet("Select Top 1 A2.ADate from Acceptance A1, Acceptance A2 Where A2.OANO=A1.OANO and A2.ANO<A1.ANO and A1.ANO=" & ANO & " and A2.ADate>'" & Format(ADate, "MM/DD/YYYY") & "'")
                If Not rs.Eof Then
                    ErrMsg = "Policy Booking Date Can't Be Earlier On : " & rs.fields("ADate") & " !"
                    GoTo errmsg
                End If
            End If
        End If

        If Not BackDatedF Then
            ' if Future Inception/Effective and Booking Date is earlier on Policy Year Effective (IFRS)
            'If DateDiff(DateInterval.Day, ADate, PDate) > 0 And Get_UWYear(SDate, EDate, ADate) < Get_UWYear(SDate, EDate, PDate) Then
            '    ErrMsg = "Policy Booking Date Can't Be Earlier On Effective Policy Year !"
            '    GoTo errmsg
            'End If
        End If

        If AType <> "A" And pAType <> "A" Then
            If ANO = LANO Then
                rs = db.OpenRemoteSet("Select C.ClaimNo, C.LossDate from Claim C, Acceptance A1, Acceptance A2, Acceptance A3 Where A3.ANO=" & ANO & " and A2.AStatus='W' and A2.LANO=A3.ANO and A1.OANO=A2.OANO and A1.ANO=C.ANO and C.LossDate>A2.PDate and C.CType<>'C' and C.CStatus in ('T','A')")
                If Not rs.Eof Then
                    ErrMsg = "Claim Already Registered No. " & rs.fields("ClaimNo") & " Loss Date On : " & rs.fields("LossDate") & " !"
                    GoTo errmsg
                End If
            Else
                rs = db.OpenRemoteSet("Select C.ClaimNo, C.LossDate from Claim C, Acceptance A1, Acceptance A2 Where A2.AStatus='W' and A2.ANO=" & ANO & " and A1.OANO=A2.OANO and A1.ANO=C.ANO and C.LossDate>A2.PDate and C.CType<>'C' and C.CStatus in ('T','A')")
                If Not rs.Eof Then
                    ErrMsg = "Claim Already Registered No. " & rs.fields("ClaimNo") & " Loss Date On : " & rs.fields("LossDate") & " !"
                    GoTo errmsg
                End If
            End If
        End If


        If ANO = LANO And PANO <> -1 Then
            rs = db.OpenRemoteSet("Select top 1 CertificateNo from Acceptance Where LANO=" & PANO & " and AStatus in ('W','T')")
            If Not rs.Eof Then
                ErrMsg = "Policy Has a Detail with Waiting Status (Ex. Certificate No. " & rs.fields("CertificateNo") & ") !"
                GoTo errmsg
            End If
        End If

        If EMethod <> "A" And AType <> "A" And pAType <> "A" Then
            rs = db.OpenRemoteSet("Select Top 1 A2.PDate from Acceptance A1, Acceptance A2 Where Not (A2.AType='C' and A2.EMethod='A') and A2.ANO=A1.PANO and A1.ANO=" & ANO & " and A2.PDate>'" & Format(PDate, "MM/DD/YYYY") & "'")
            If Not rs.Eof Then
                ErrMsg = "Policy Effective Date Can't Be Earlier On : " & rs.fields("PDate") & " !"
                GoTo errmsg
            End If

            If LANO <> ANO Then
                rs = db.OpenRemoteSet("Select R.PDate as R_PDate, R.PPDate as R_PPDate, R.RDate, R.NRDate from RArrHeader R Where R.ANO=" & ANO)
                If Not rs.Eof Then
                    ' Not Clean Cut
                    If (DateDiff(DateInterval.Day, rs.fields("R_PDate"), rs.fields("RDate")) = 0 And DateDiff(DateInterval.Day, rs.fields("R_PPDate"), rs.fields("NRDate")) = 0) And (DateDiff(DateInterval.Day, todate(rs.fields("RDate")), PDate) < 0 Or DateDiff(DateInterval.Day, todate(rs.fields("NRDate")), PDate) > 0) Then
                        ErrMsg = "Effective Date Should Be Between Treaty Period On " & Format(todate(rs.fields("RDate")), "d") & " - " & Format(todate(rs.fields("NRDate")), "d")
                        GoTo errmsg
                    End If

                    ' Clean Cut
                    If Not (DateDiff(DateInterval.Day, rs.fields("R_PDate"), rs.fields("RDate")) = 0 And DateDiff(DateInterval.Day, rs.fields("R_PPDate"), rs.fields("NRDate")) = 0) And DateDiff(DateInterval.Day, PDate, rs.fields("R_PPDate")) < 0 Then
                        ErrMsg = "Policy Needs Annual Treaty Reallocation, Please Review The Reinsurance Date on " & Format(rs.fields("R_PPDate"), "d") & " !"
                        GoTo errmsg
                    End If
                End If
            End If
        End If

        If Not ReviewF Then
            If Not db.GetPrivileges("ALLOWBACKDATEDTRX") AndAlso Not (AType = "C" And EMethod = "A") Then
                Dim MinBackDatedPolicy As Integer
                Dim MaxBackDatedPolicy As Integer
                GetMinMaxBackDatedPolicy(TOC, MinBackDatedPolicy, MaxBackDatedPolicy)
                If DateDiff(DateInterval.Day, PDate, ADate) > MaxBackDatedPolicy Or DateDiff(DateInterval.Day, PDate, ADate) < MinBackDatedPolicy Then
                    ErrMsg = "Policy Inception/Effective Date BackDate must be between " & tostr(MinBackDatedPolicy) & " - " & tostr(MaxBackDatedPolicy) & " Days !"
                    GoTo errmsg
                End If
            End If

            If GracePeriod <> 0 Then
                Dim ALLOWEXCEEDMINGRACEF As Boolean = db.GetPrivileges("ALLOWEXCEEDMINGRACE")
                Dim ALLOWEXCEEDMAXGRACEF As Boolean = db.GetPrivileges("ALLOWEXCEEDMAXGRACE")
                If (Not ALLOWEXCEEDMINGRACEF OrElse Not ALLOWEXCEEDMAXGRACEF) Then
                    Dim MinGracePeriodPolicy As Integer
                    Dim MaxGracePeriodPolicy As Integer
                    GetMinMaxGracePeriodPolicy(MinGracePeriodPolicy, MaxGracePeriodPolicy)

                    If MinGracePeriodPolicy <> 0 OrElse MaxGracePeriodPolicy <> 0 Then
                        If Not ALLOWEXCEEDMINGRACEF Then
                            If toval(GracePeriod) < MinGracePeriodPolicy Then
                                ErrMsg = "Policy Min Grace Period is " & MinGracePeriodPolicy & " day(s) !"
                                GoTo ErrMsg
                            End If
                        End If

                        If Not ALLOWEXCEEDMAXGRACEF Then
                            If toval(GracePeriod) > MaxGracePeriodPolicy Then
                                ErrMsg = "Policy Max Grace Period is " & MaxGracePeriodPolicy & " day(s) !"
                                GoTo ErrMsg
                            End If
                        End If
                    End If
                End If
            End If

            If Not db.GetPrivileges("ALLOWCHANGEFIRSTINSTALLMENTDUE") Then
                If Installment > 0 And DateDiff(DateInterval.Day, IDate_1, DateAdd(DateInterval.Day, GracePeriod, DueDate)) <> 0 Then
                    ErrMsg = "First Installment Due Not the Same With Due Date On : " & Format(DateAdd(DateInterval.Day, GracePeriod, DueDate), "d") & " !"
                    GoTo errmsg
                End If
            End If
        End If

        If Not BucketF Then
            If ANO = LANO Then ' Master
                If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "PREINFORCECONDITIONS", ANO, BackDatedF) Then
                    ErrMsg = Me.ErrDesc
                    GoTo errmsg
                End If
                If Not OpenFlag Then
                    rs = db.OpenRemoteSet("Select ANO, PStatus, AStatus, LDate, LTime from Acceptance Where ANO<>LANO and LANO=" & LANO)
                    If Not rs.Eof Then
                        Do While Not rs.Eof

                            If DateDiff(DateInterval.Day, ReqDate, todate(rs.fields("LDate"))) > 0 OrElse (DateDiff(DateInterval.Day, ReqDate, todate(rs.fields("LDate"))) = 0 And ReqTime < tostr(rs.fields("LTime"))) Then
                                ReqDate = todate(rs.fields("LDate"))
                                ReqTime = tostr(rs.fields("LTime"))
                            End If

                            ' Do Premium Calculation If Necessary !
                            If rs.fields("PStatus") = "I" Then
                                If Not Me.Direct_Call_Ext_Process("Process.NetV3.Proclib", "PREMIUM_CALC", toval(rs.fields("ANO"))) Then
                                    ErrMsg = Me.ErrDesc
                                    GoTo errmsg
                                End If
                                ReqDate = Today.Date
                                ReqTime = Time()
                            End If

                            If rs.fields("AStatus") = "W" Then
                                If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "PREINFORCECONDITIONS", toval(rs.fields("ANO")), BackDatedF) Then
                                    ErrMsg = Me.ErrDesc
                                    GoTo errmsg
                                End If
                            End If
                            rs.MoveNext()
                        Loop
                    Else
                        If AType = "N" Then
                            ErrMsg = "Master Policy Has No Detail Data !"
                            GoTo errmsg
                        End If
                    End If
                End If
            Else
                ' Do Premium Calculation If Necessary !
                If PStatus = "I" Then
                    If Not Me.Direct_Call_Process_Execute("Process.NetV3.Proclib", "PREMIUM_CALC", ANO) Then
                        ErrMsg = Me.ErrDesc
                        GoTo errmsg
                    End If
                    ReqDate = Today.Date
                    ReqTime = Time()
                End If
                If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "PREINFORCECONDITIONS", ANO, BackDatedF) Then
                    ErrMsg = Me.ErrDesc
                    GoTo errmsg
                End If
            End If
        End If

        If Not ReviewF Then
            If Not db.GetPrivileges("ALLOWPASSAPPLIMIT") Then
                rs = db.OpenQueryDef("QApproval_Production", adOpenKeyset, ANO)
                If Not rs.Eof Then
                    If AType = "C" AndAlso CBool(toval(GetAppSetting("LIMITOUTFORAPPROVALCANCELLATIONF"))) Then
                        If Not GetApprovalLimit(db.Var_db_user, tostr(rs.fields("TOC")), tostr(rs.fields("ISType")), "Out", toval(rs.fields("Amount")), tostr(rs.fields("Branch")), toval(rs.fields("ApprLimitPCT")), ANO, toval(rs.fields("TotalAmount"))) Then
                            Dim RequestingUser As String = ""
                            If Not (CheckRequestingActivities("ALLOWPASSAPPLIMIT", "UND", ANO, ReqDate, ReqTime, RequestingUser) AndAlso RequestingUser.ToUpper = dbUser.ToUpper) Then
                                ErrMsg = "No Limit Authorization To Approve !"
                                GoTo errmsg
                            End If
                        End If
                    Else
                        If Not GetApprovalLimit(db.Var_db_user, tostr(rs.fields("TOC")), tostr(rs.fields("ISType")), "In", toval(rs.fields("Amount")), tostr(rs.fields("Branch")), toval(rs.fields("ApprLimitPCT")), ANO, toval(rs.fields("TotalAmount"))) Then
                            Dim RequestingUser As String = ""
                            If Not (CheckRequestingActivities("ALLOWPASSAPPLIMIT", "UND", ANO, ReqDate, ReqTime, RequestingUser) AndAlso RequestingUser.ToUpper = dbUser.ToUpper) Then
                                ErrMsg = "No Limit Authorization To Approve !"
                                GoTo errmsg
                            End If
                        End If
                    End If
                End If
            End If
            If AType = "C" AndAlso Not db.GetPrivileges("ALLOWINFORCEONPOLICYCANCELLATION") Then
                ErrMsg = "No Privileges for Policy Cancellation !"
                GoTo errmsg
            End If
        End If

        ' Create Default Employee Contribution
        If AType = "N" Then
            rs = db.OpenRemoteSet("Select ECUType from ECUType Where DefaultF=1 and PolicyF=1")
            If Not rs.Eof Then
                Dim ECUtype As String
                ECUtype = tostr(rs.fields("ECUType"))
                rs = db.OpenRemoteSet("Select * from ACCEMPCONTH Where ANO=" & OANO)
                If rs.Eof Then
                    rs.AddNew()
                    rs.fields("ANO") = OANO
                    rs.fields("TOTALPCT") = 100
                    rs.fields("Remarks") = "Default Contribution"
                    rs.fields("Status") = "A"
                    rs.fields("UserID") = dbUser
                    rs.fields("CDate") = Today.Date
                    rs.fields("LUserID") = dbUser
                    rs.fields("LDate") = Today.Date
                    rs.fields("PUserID") = dbUser
                    rs.fields("PUserDate") = Today.Date
                    rs.fields("AUserID") = dbUser
                    rs.fields("AUserDate") = Today.Date
                    If Not rs.Update Then
                        ErrMsg = "Can't Create Default Quotation Employee Contribution !"
                        GoTo errmsg
                    End If

                    rs = db.OpenRemoteSet("Select * from ACCEMPCONTD Where 1<>1")
                    rs.AddNew()
                    rs.fields("ANO") = OANO
                    rs.fields("ID") = SysUserID
                    rs.fields("Remarks") = "Default Contribution"
                    rs.fields("PCT") = 100
                    rs.fields("ECUTYPE") = ECUtype
                    If Not rs.Update Then
                        ErrMsg = "Can't Create Default Quotation Employee Contribution !"
                        GoTo errmsg
                    End If
                End If
            End If
        End If

        PreInforceConditionsOK = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy Inforce : " & ErrMsg)
        PreInforceConditionsOK = False
    End Function
    Private Function PreUndoInforceConditionsOK(ByVal ANO As Long, ByVal LANO As Long, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim CNO As Long
        Dim PANO As Long
        Dim OpenFlag As Boolean
        Dim RegNo As String
        Dim AStatus As String
        Dim AType As String

        On Error GoTo errmsg

        rs = db.OpenRemoteSet("Select CNO, PANO, REGNO, AStatus, AType from Acceptance Where ANO=" & ANO)
        CNO = rs.fields("CNO")
        PANO = rs.fields("PANO")
        RegNo = rs.fields("RegNo")
        AStatus = rs.fields("AStatus")
        AType = rs.fields("AType")

        If AStatus = "T" Then
            PreUndoInforceConditionsOK = True
            Exit Function
        End If

        rs = db.OpenRemoteSet("Select OpenFlag from Cover Where CNO=" & CNO)
        OpenFlag = rs.fields("OpenFlag")

        If (Not OpenFlag AndAlso ANO <> LANO AndAlso LANO <> -1) Then
            ErrMsg = "Can't Undo Inforce Master Policy Certificate !"
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select AType from Acceptance Where PANO=" & ANO)
        If Not rs.Eof Then
            Select Case rs.fields("AType")
                Case "E"
                    ErrMsg = "Policy has a Endorsement Data, Please Review First !"
                Case "C"
                    ErrMsg = "Policy has a Cancellation Data, Please Review First !"
                Case "A"
                    ErrMsg = "Policy has an Adjustment Data, Please Review First !"
            End Select
            GoTo errmsg
        End If

        If AType <> "A" AndAlso AType <> "C" Then
            If ANO = LANO And Not OpenFlag Then
                ' Check Detail Acceptance for Conjunction Policies
                rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2 Inner Join Cover C on (C.CNO=A2.CNO) Left Join RArrHeader R on (R.ANO=A2.ANO) Where (R.Status='A') and A2.AType in ('N','E') and A2.AStatus Not In ('W','T') and A2.CANO=A1.ANO and A1.LANO=" & ANO)
                If Not rs.Eof Then
                    ErrMsg = "Must Undo-Inforce the Co. Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                    GoTo errmsg
                End If
            Else
                rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2 Inner Join Cover C on (C.CNO=A2.CNO) Left Join RArrHeader R on (R.ANO=A2.ANO) Where (R.Status='A') and A2.AType in ('N','E') and A2.AStatus Not In ('W','T') and A2.CANO=A1.ANO and A1.ANO=" & ANO)
                If Not rs.Eof Then
                    ErrMsg = "Must Undo-Inforce the Co. Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                    GoTo errmsg
                End If
            End If
        End If

        rs = db.OpenRemoteSet("Select Top 1 ANO from Acceptance Where AType='A' and AStatus='I' and PANO=" & PANO & " and ANO<>" & ANO)
        If Not rs.Eof Then
            ErrMsg = "Endorsed Policy has an Adjustment Data, Please Review First !"
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select Top 1 Claim.ClaimNo from Claim Where ANO=" & ANO)
        If Not rs.Eof Then
            ErrMsg = "Already Used By Claim No : " & rs.fields("ClaimNo")
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select Top 1 Claim.ClaimNo from Claim, Acceptance Where Acceptance.ANO=Claim.ANO and Acceptance.RegNo='" & RegNo & "' and Acceptance.LANO=" & ANO)
        If Not rs.Eof Then
            ErrMsg = "Already Used By Claim No : " & rs.fields("ClaimNo")
            GoTo errmsg
        End If

        If ANO = LANO And OpenFlag Then
            rs = db.OpenRemoteSet("Select Top 1 ANO from Acceptance Where ANO<>LANO and RegNo='" & RegNo & "' and LANO=" & LANO)
            If Not rs.Eof Then
                ErrMsg = "Already Has Detail Certificate !"
                GoTo errmsg
            End If
        End If

        PreUndoInforceConditionsOK = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy Undo Expire : " & ErrMsg)
        PreUndoInforceConditionsOK = False
    End Function
    Private Function PreEndorseConditionsOK(ByVal ANO As Long, ByVal NewPDate As Date, ByVal EndorseAllF As Boolean, ByVal AppSource As String, ByVal ExternalF As Boolean, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim AType As String
        Dim EMethod As String
        Dim SDate As Date
        Dim EDate As Date
        Dim PDate As Date
        Dim Remarks As String
        Dim LANO As Long
        Dim pAType As String
        Dim AStatus As String
        Dim MAStatus As String
        Dim OpenFlag As Boolean

        On Error GoTo errmsg

        ErrMsg = ""

        rs = db.OpenRemoteSet("Select A.AStatus, mA.AStatus as MAStatus, A.EMethod, A.AType, A.PDate, A.SDate, A.EDate, A.LANO, A.Remarks, pA.AType as pAType, C.OpenFlag from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) Left Join Acceptance mA on (mA.ANO=A.LANO and mA.ANO=mA.LANO) Left Join Acceptance pA on (pA.ANO=A.pANO) Where A.ANO=" & ANO)
        If rs.Eof Then
            ErrMsg = "Policy/Certificate Data Not Found !"
            GoTo errmsg
        End If
        EMethod = tostr(rs.fields("EMethod"))
        AType = tostr(rs.fields("AType"))
        PDate = todate(rs.fields("PDate"))
        SDate = todate(rs.fields("SDate"))
        EDate = todate(rs.fields("EDate"))
        LANO = toval(rs.fields("LANO"))
        Remarks = tostr(rs.fields("Remarks"))
        pAType = tostr(rs.fields("pAType"))
        AStatus = tostr(rs.fields("AStatus"))
        MAStatus = tostr(rs.fields("MAStatus"))
        OpenFlag = toval(rs.fields("OpenFlag"))

        If (LANO = ANO) AndAlso AType = "E" AndAlso AStatus = "W" Then
            PreEndorseConditionsOK = True
            Exit Function
        End If

        If AType = "A" Or pAType = "A" Then
            ErrMsg = "Endorsement not for Adjustment !"
            GoTo errmsg
        End If

        If AStatus <> "I" Then
            ErrMsg = "Please Check Policy Data Status !"
            GoTo errmsg
        End If

        If LANO <> -1 AndAlso ANO <> LANO AndAlso MAStatus <> "I" Then
            ErrMsg = "Please Check Master Policy Data Status !"
            GoTo errmsg
        End If

        If DateDiff(DateInterval.Day, PDate, NewPDate) < 0 And Not (AType = "C" And EMethod = "A") Then
            Select Case AType
                Case "N"
                    ErrMsg = "Effective Date for this endorsement can't be earlier than current policy."
                Case "E"
                    ErrMsg = "Effective Date for this endorsement can't be earlier than current endorsement."
                Case "C"
                    ErrMsg = "Effective Date for this endorsement can't be earlier than current cancellation."
            End Select
            GoTo errmsg
        End If

        If AType = "C" Then ' Endorsement Over Cancellation
            If Not db.GetPrivileges("EndorseCancellation") Then
                ErrMsg = "Not Authorized Endorse Over Cancellation !"
                GoTo errmsg
            End If
            If Remarks.ToUpper = "AUTO LAPSING" Then
                If Not db.GetPrivileges("ALLOWENDORSEALPOLICY") Then
                    ErrMsg = "Not Authorized Endorse Auto Lapse Policy !"
                    GoTo errmsg
                End If
            End If
        End If

        If AType <> "A" And EMethod <> "A" Then
            If DateDiff(DateInterval.Day, SDate, NewPDate) < 0 Then
                ErrMsg = "Effective Date can't be earlier than Start Date"
                GoTo errmsg
            End If

            If DateDiff(DateInterval.Day, EDate, NewPDate) > 0 Then
                ErrMsg = "Effective Date can't be greater than Expiry Date"
                GoTo errmsg
            End If
        End If

        If AType = "C" And EMethod <> "A" Then    ' Endorse Over Cancellation
            If DateDiff(DateInterval.Day, PDate, NewPDate) <> 0 Then
                ErrMsg = "Effective Date must be the Same with Cancellation Date"
                GoTo errmsg
            End If
        End If

        If ANO = LANO And Not OpenFlag AndAlso EndorseAllF Then
            ' Check Detail Acceptance for Conjunction Policies
            rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2, Cover C Where C.CNO=A2.CNO and A2.ANO=A1.CANO and A2.AType in ('C') and A1.AType in ('N','E','C') and A1.AStatus='I' and A1.LANO=" & ANO)
            If Not rs.Eof Then
                ErrMsg = "Must Endorse the Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                GoTo errmsg
            End If
        Else
            rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2, Cover C Where C.CNO=A2.CNO and A2.ANO=A1.CANO and A2.AType in ('C') and A1.AType in ('N','E','C') and A1.AStatus='I' and A1.ANO=" & ANO)
            If Not rs.Eof Then
                ErrMsg = "Must Endorse the Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                GoTo errmsg
            End If
        End If

        ' Check Existing Claim Incurred 
        If Not db.GetPrivileges("ALLOWENDORSEWITHCLAIMINCURRED") Then
            If EndorseAllF Then
                rs = db.OpenRemoteSet("Select C.ClaimNo, A1.CertificateNo, C.LossDate from Claim C, Acceptance A1, Acceptance A2 Where A1.AType in ('N','E','C') and A1.AStatus='I' and A1.ANO<>A1.LANO and A1.PPDate>='" & Format(NewPDate, "MM/DD/YYYY") & "' and A1.LANO=" & ANO & " and A2.OANO=A1.OANO and C.ANO=A2.ANO and C.CType<>'C' and C.CStatus<>'R' and (C.CStatus<>'C' or C.PCstatus='A') and not exists (Select pCNO from Claim Where Claim.pCNO=C.CNO and Claim.CStatus<>'R') order By C.LossDate Desc")
            Else
                rs = db.OpenRemoteSet("Select C.ClaimNo, A1.CertificateNo, C.LossDate from Claim C, Acceptance A1, Acceptance A2 Where A1.ANO=" & ANO & " and A2.OANO=A1.OANO and C.ANO=A2.ANO and C.CType<>'C' and C.CStatus<>'R' and (C.CStatus<>'C' or C.PCstatus='A') and not exists (Select pCNO from Claim Where Claim.pCNO=C.CNO and Claim.CStatus<>'R') order By C.LossDate Desc")
            End If
            If Not rs.Eof Then
                Do While Not rs.Eof
                    If IsDate(rs.fields("LossDate")) Then
                        If DateDiff(DateInterval.Day, rs.fields("LossDate"), NewPDate) < 0 Then
                            If rs.fields("CertificateNo") <> "" Then
                                ErrMsg = "Claim Already Incurred (ClaimNo : " & rs.fields("ClaimNo") & " in CertificateNo : " & rs.fields("CertificateNo") & ") on " & Format(rs.fields("LossDate"), "d") & " !" & vbCrLf
                            Else
                                ErrMsg = "Claim Already Incurred (ClaimNo : " & rs.fields("ClaimNo") & ") on " & Format(rs.fields("LossDate"), "d") & " !" & vbCrLf
                            End If
                        End If
                    End If
                    rs.MoveNext()
                Loop
                If ErrMsg <> "" Then GoTo errmsg
            End If
        End If

        ' Check Total Loss
        rs = db.OpenQueryDef("QCLACC_6", adOpenForwardonly, ANO)
        If Not rs.Eof Then
            ErrMsg = "Already Has Claim Total Loss (ClaimNo : " & rs.fields("ClaimNo") & ") on " & Format(rs.fields("LossDate"), "d") & " !"
            GoTo errmsg
        End If

        If EMethod <> "A" Then
            rs = db.OpenRemoteSet("Select R.PDate as R_PDate, R.PPDate as R_PPDate, R.RDate, R.NRDate from RArrHeader R Where R.ANO=" & ANO)
            If Not rs.Eof Then
                ' Not Clean Cut
                If (DateDiff(DateInterval.Day, rs.fields("R_PDate"), rs.fields("RDate")) = 0 And DateDiff(DateInterval.Day, rs.fields("R_PPDate"), rs.fields("NRDate")) = 0) And (DateDiff(DateInterval.Day, todate(rs.fields("RDate")), NewPDate) < 0 Or DateDiff(DateInterval.Day, todate(rs.fields("NRDate")), NewPDate) > 0) Then
                    ErrMsg = "Effective Date Should Be Between Treaty Period On " & Format(todate(rs.fields("RDate")), "d") & " - " & Format(todate(rs.fields("NRDate")), "d")
                    GoTo errmsg
                End If

                ' Clean Cut
                If Not (DateDiff(DateInterval.Day, rs.fields("R_PDate"), rs.fields("RDate")) = 0 And DateDiff(DateInterval.Day, rs.fields("R_PPDate"), rs.fields("NRDate")) = 0) And DateDiff(DateInterval.Day, NewPDate, rs.fields("R_PPDate")) < 0 Then
                    ErrMsg = "Policy Needs Annual Treaty Reallocation, Please Review The Reinsurance Date on " & Format(rs.fields("R_PPDate"), "d") & " !"
                    GoTo errmsg
                End If
            End If
        End If

        If CBool(toval(GetAppSetting("CheckUnPaidPremiumWhenEndorseF"))) Then
            If ExternalF Then
                If Not db.GetPrivileges("UNPAIDPREMIUM") Then
                    If Not isPremiumPaid(ANO, Today.Date) Then
                        Dim RequestingUser As String = ""
                        If Not (CheckRequestingActivities("UNPAIDPREMIUM", "UND", ANO, NewPDate, "", RequestingUser) AndAlso RequestingUser.ToUpper = dbUser.ToUpper) Then
                            ErrMsg = "Can't Endorse The Policy which is Not Yet Paid !"
                            GoTo errmsg
                        End If
                    End If
                End If
            End If
        End If

        If ANO = LANO Then
            rs = db.OpenRemoteSet("Select top 1 CertificateNo from Acceptance Where LANO=" & LANO & " and AStatus in ('W','T')")
            If Not rs.Eof Then
                ErrMsg = "Policy Has a Detail with Waiting Status (Ex. Certificate No. " & rs.fields("CertificateNo") & ") !"
                GoTo errmsg
            End If
        End If

        rs = db.OpenRemoteSet("Select A1.RegNo from Acceptance A1 Where ((AType in ('A') and AStatus<>'C') or AType not in ('A')) and PANO=" & ANO)
        If Not rs.Eof Then
            ErrMsg = "Already Endorsed (Cancelled) ! Referred to The Policy Data. (Reg No. " & rs.fields("RegNo") & ")"
            GoTo errmsg
        End If

        If AppSource = "MKT" Then
            If Not db.GetPrivileges("ALLOWENDORSEMKT") Then
                ErrMsg = "Not Allowed To Endorse The Marketing Data"
                GoTo errmsg
            End If
        End If

        PreEndorseConditionsOK = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy Endorse : " & ErrMsg)
        PreEndorseConditionsOK = False
    End Function
    Private Function PreCancelConditionsOK(ByVal ANO As Long, ByVal NewPDate As Date, ByVal OpenFlag As Boolean, ByVal CancelAllF As Boolean, ByVal TOC As String, ByVal NoRefundF As Integer, ByVal ExternalF As Boolean, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim AType As String
        Dim EMethod As String
        Dim SDate As Date
        Dim EDate As Date
        Dim PDate As Date
        Dim LANO As Long
        Dim OANO As Long
        Dim PolicyNo As String
        Dim CertificateNo As String
        Dim AStatus As String
        Dim MAStatus As String

        On Error GoTo errmsg

        rs = db.OpenRemoteSet("Select A.PolicyNo, A.CertificateNo, A.OANO, A.AStatus, mA.AStatus as MAStatus, A.EMethod, A.AType, A.PDate, A.SDate, A.EDate, A.LANO, A.Remarks, pA.AType as pAType from Acceptance A Left Join Acceptance mA on (mA.ANO=A.LANO and mA.ANO=mA.LANO) Left Join Acceptance pA on (pA.ANO=A.pANO) Where A.ANO=" & ANO)
        If rs.Eof Then
            ErrMsg = "Policy/Certificate Data Not Found !"
            GoTo errmsg
        End If
        EMethod = tostr(rs.fields("EMethod"))
        AType = tostr(rs.fields("AType"))
        PDate = todate(rs.fields("PDate"))
        SDate = todate(rs.fields("SDate"))
        EDate = todate(rs.fields("EDate"))
        LANO = toval(rs.fields("LANO"))
        OANO = toval(rs.fields("OANO"))
        PolicyNo = tostr(rs.fields("PolicyNo"))
        CertificateNo = tostr(rs.fields("CertificateNo"))
        AStatus = tostr(rs.fields("AStatus"))
        MAStatus = tostr(rs.fields("MAStatus"))

        If AType = "C" Then
            ErrMsg = "Cancellation not for Cancellation !"
            GoTo errmsg
        End If

        If AStatus <> "I" Then
            ErrMsg = "Please Check Policy Data Status !"
            GoTo errmsg
        End If

        If LANO <> -1 AndAlso ANO <> LANO AndAlso MAStatus <> "I" Then
            ErrMsg = "Please Check Master Policy Data Status !"
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select NOCANCELF from TOC Where TOC='" & TOC & "'")
        If toval(rs.fields("NOCANCELF")) Then
            If Not db.GetPrivileges("ALLOWCANCELDATA") Then
                ErrMsg = "Not Allowed to Cancel Data !"
                GoTo errmsg
            End If
        End If

        If Not CancelAllF Then
            If DateDiff(DateInterval.Day, SDate, NewPDate) < 0 Then
                ErrMsg = "Effective Date can't be earlier than Start Date"
                GoTo errmsg
            End If

            If DateDiff(DateInterval.Day, EDate, NewPDate) > 0 Then
                ErrMsg = "Effective Date can't be greater than Expiry Date"
                GoTo errmsg
            End If
            If DateDiff(DateInterval.Day, PDate, NewPDate) < 0 Then
                Select Case AType
                    Case "N"
                        ErrMsg = "Effective Date for this cancellation can't be earlier than current policy."
                    Case "E"
                        ErrMsg = "Effective Date for this cancellation can't be earlier than current endorsement."
                    Case "A"
                        ErrMsg = "Effective Date for this cancellation can't be earlier than current Adjustment."
                    Case Else
                        ErrMsg = "Effective Date for this cancellation can't be earlier than previous transaction."
                End Select
                GoTo errmsg
            End If
        End If

        If AType <> "A" Then

            If ANO = LANO And Not OpenFlag Then
                ' Check Detail Acceptance for Conjunction Policies
                rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2 Inner Join Cover C on (C.CNO=A2.CNO) Left Join RArrHeader R on (R.ANO=A2.ANO) Where (R.Status='A' or A2.AStatus='W') and A2.CANO=A1.ANO and A1.AType in ('N','E') and A1.LANO=" & ANO)
                If Not rs.Eof Then
                    ErrMsg = "Must Cancel the Co. Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                    GoTo errmsg
                End If
            Else
                rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2 Inner Join Cover C on (C.CNO=A2.CNO) Left Join RArrHeader R on (R.ANO=A2.ANO) Where (R.Status='A' or A2.AStatus='W') and A2.CANO=A1.ANO and A1.AType in ('N','E') and A1.ANO=" & ANO)
                If Not rs.Eof Then
                    ErrMsg = "Must Cancel the Co. Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                    GoTo errmsg
                End If
            End If

            ' Check Existing Claim Incurred
            If ANO = LANO And Not OpenFlag Then
                Dim rsDet As RemoteSet
                rsDet = db.OpenRemoteSet("Select ANO from Acceptance Where AType in ('N','E') and AStatus='I' and LANO=" & LANO)
                Do While Not rsDet.Eof
                    ' Check Existing Claim Incurred
                    rs = db.OpenRemoteSet("Select Top 1 C.LossDate, C.CStatus, C.pCStatus, C1.CTLF from Claim C Left Join CLLOSS C1 on (C1.CNO=C.CNO and C1.CTLF=1), Acceptance A1, Acceptance A2 Where A1.ANO=" & rsDet.fields("ANO") & " and A2.OANO=A1.OANO and C.ANO=A2.ANO and C.CType<>'C' and C.CStatus<>'R' and (C.CStatus<>'C' or C.PCstatus='A') and not exists (Select pCNO from Claim Where Claim.pCNO=C.CNO and Claim.CStatus<>'R') order By C.LossDate Desc")
                    If IsDate(rs.fields("LossDate")) Then
                        If DateDiff(DateInterval.Day, rs.fields("LossDate"), NewPDate) < 0 Then
                            ErrMsg = "Claim Already Incurred on " & Format(rs.fields("LossDate"), "d")
                            GoTo errmsg
                        ElseIf DateDiff(DateInterval.Day, rs.fields("LossDate"), NewPDate) = 0 Then
                            If Not ((rs.fields("CStatus") = "A" Or rs.fields("pCStatus") = "A") And toval(rs.fields("CTLF")) And (NoRefundF = 1 Or NoRefundF = 2 Or NoRefundF = 4)) Then
                                ErrMsg = "Claim Already Incurred on " & Format(rs.fields("LossDate"), "d")
                                GoTo errmsg
                            End If
                        End If
                    End If

                    ' Check Total Loss
                    If Not (NoRefundF = 1 Or NoRefundF = 2 Or NoRefundF = 4) Then
                        rs = db.OpenQueryDef("QCLACC_6", adOpenForwardonly, rsDet.fields("ANO"))
                        If Not rs.Eof Then
                            If (DateDiff(DateInterval.Month, todate(SDate), rs.fields("LossDate")) \ 12) >= (DateDiff(DateInterval.Month, todate(SDate), NewPDate) \ 12) Then
                                ErrMsg = "Already Has Claim Total Loss (ClaimNo : " & rs.fields("ClaimNo") & ") on " & Format(rs.fields("LossDate"), "d") & " !"
                                GoTo errmsg
                            End If
                        End If
                    End If

                    If Not CancelAllF Then
                        ' Not Check If Cancel All
                        rs = db.OpenRemoteSet("Select R.PDate as R_PDate, R.PPDate as R_PPDate, R.RDate, R.NRDate from RArrHeader R Where R.ANO=" & rsDet.fields("ANO"))
                        If Not rs.Eof Then
                            ' Not Clean Cut
                            If (DateDiff(DateInterval.Day, rs.fields("R_PDate"), rs.fields("RDate")) = 0 And DateDiff(DateInterval.Day, rs.fields("R_PPDate"), rs.fields("NRDate")) = 0) And (DateDiff(DateInterval.Day, todate(rs.fields("RDate")), NewPDate) < 0 Or DateDiff(DateInterval.Day, todate(rs.fields("NRDate")), NewPDate) > 0) Then
                                ErrMsg = "Effective Date Should Be Between Treaty Period On " & Format(todate(rs.fields("RDate")), "d") & " - " & Format(todate(rs.fields("NRDate")), "d")
                                GoTo errmsg
                            End If

                            ' Clean Cut
                            If Not (DateDiff(DateInterval.Day, rs.fields("R_PDate"), rs.fields("RDate")) = 0 And DateDiff(DateInterval.Day, rs.fields("R_PPDate"), rs.fields("NRDate")) = 0) And DateDiff(DateInterval.Day, NewPDate, rs.fields("R_PPDate")) < 0 Then
                                ErrMsg = "Policy Needs Annual Treaty Reallocation, Please Review The Reinsurance Date on " & Format(rs.fields("R_PPDate"), "d") & " !"
                                GoTo errmsg
                            End If
                        End If
                    End If
                    rsDet.MoveNext()
                Loop
            Else
                ' Check Existing Claim Incurred
                rs = db.OpenRemoteSet("Select Top 1 C.LossDate, C.CStatus, C.pCStatus, C1.CTLF from Claim C Left Join CLLOSS C1 on (C1.CNO=C.CNO and C1.CTLF=1), Acceptance A1, Acceptance A2 Where A1.ANO=" & ANO & " and A2.OANO=A1.OANO and C.ANO=A2.ANO and C.CType<>'C' and C.CStatus<>'R' and (C.CStatus<>'C' or C.PCstatus='A') and not exists (Select pCNO from Claim Where Claim.pCNO=C.CNO and Claim.CStatus<>'R') order By C.LossDate Desc")
                If IsDate(rs.fields("LossDate")) Then
                    If DateDiff(DateInterval.Day, rs.fields("LossDate"), NewPDate) < 0 Then
                        ErrMsg = "Claim Already Incurred on " & Format(rs.fields("LossDate"), "d")
                        GoTo errmsg
                    ElseIf DateDiff(DateInterval.Day, rs.fields("LossDate"), NewPDate) = 0 Then
                        If Not ((rs.fields("CStatus") = "A" Or rs.fields("pCStatus") = "A") And toval(rs.fields("CTLF")) And (NoRefundF = 1 Or NoRefundF = 2 Or NoRefundF = 4)) Then
                            ErrMsg = "Claim Already Incurred on " & Format(rs.fields("LossDate"), "d")
                            GoTo errmsg
                        End If
                    End If
                End If

                ' Check Total Loss
                If Not (NoRefundF = 1 Or NoRefundF = 2 Or NoRefundF = 4) Then
                    rs = db.OpenQueryDef("QCLACC_6", adOpenForwardonly, ANO)
                    If Not rs.Eof Then
                        If (DateDiff(DateInterval.Month, todate(SDate), rs.fields("LossDate")) \ 12) >= (DateDiff(DateInterval.Month, todate(SDate), NewPDate) \ 12) Then
                            ErrMsg = "Already Has Claim Total Loss (ClaimNo : " & rs.fields("ClaimNo") & ") on " & Format(rs.fields("LossDate"), "d") & " !"
                            GoTo errmsg
                        End If
                    End If
                End If

                If NoRefundF = 0 Then
                    Dim ClaimOnCancelF As Boolean
                    rs = db.OpenRemoteSet("Select ClaimOnCancelF from TOC Where TOC='" & TOC & "'")
                    ClaimOnCancelF = toval(rs.fields("ClaimOnCancelF"))
                    If ClaimOnCancelF Then
                        Dim Claim_Paid As Double
                        Dim Gross_Premium As Double
                        Dim OS_Claim As Double

                        rs = db.OpenQueryDef("CLAIM_PAID_BY_POLICY", adOpenForwardonly, PolicyNo, CertificateNo)
                        Claim_Paid = -toval(rs.fields("Claim_Paid"))
                        rs = db.OpenQueryDef("GROSS_PREMIUM_BY_POLICY", adOpenForwardonly, PolicyNo, CertificateNo)
                        Gross_Premium = toval(rs.fields("Gross_Premium"))
                        rs = db.OpenQueryDef("OS_CLAIM_BY_POLICY", adOpenForwardonly, PolicyNo, CertificateNo)
                        OS_Claim = toval(rs.fields("OS_Claim"))

                        If Gross_Premium <> 0 AndAlso (Claim_Paid + OS_Claim) <> 0 AndAlso (Claim_Paid + OS_Claim) > Gross_Premium Then
                            ErrMsg = "Amount Claim Exceeding Premium Amount !"
                            GoTo errmsg
                        End If
                    End If
                End If

                If Not CancelAllF Then
                    ' Not Check If Cancel All
                    rs = db.OpenRemoteSet("Select R.PDate as R_PDate, R.PPDate as R_PPDate, R.RDate, R.NRDate from RArrHeader R Where R.ANO=" & ANO)
                    If Not rs.Eof Then
                        ' Not Clean Cut
                        If (DateDiff(DateInterval.Day, rs.fields("R_PDate"), rs.fields("RDate")) = 0 And DateDiff(DateInterval.Day, rs.fields("R_PPDate"), rs.fields("NRDate")) = 0) And (DateDiff(DateInterval.Day, todate(rs.fields("RDate")), NewPDate) < 0 Or DateDiff(DateInterval.Day, todate(rs.fields("NRDate")), NewPDate) > 0) Then
                            ErrMsg = "Effective Date Should Be Between Treaty Period On " & Format(todate(rs.fields("RDate")), "d") & " - " & Format(todate(rs.fields("NRDate")), "d")
                            GoTo errmsg
                        End If

                        ' Clean Cut
                        If Not (DateDiff(DateInterval.Day, rs.fields("R_PDate"), rs.fields("RDate")) = 0 And DateDiff(DateInterval.Day, rs.fields("R_PPDate"), rs.fields("NRDate")) = 0) And DateDiff(DateInterval.Day, NewPDate, rs.fields("R_PPDate")) < 0 Then
                            ErrMsg = "Policy Needs Annual Treaty Reallocation, Please Review The Reinsurance Date on " & Format(rs.fields("R_PPDate"), "d") & " !"
                            GoTo errmsg
                        End If
                    End If
                End If
            End If

            If CancelAllF Then
                rs = db.OpenRemoteSet("Select Top 1 A1.RegNo from Acceptance A1 Where AType='A' and AStatus<>'C' and OANO=" & OANO)
                If Not rs.Eof Then
                    ErrMsg = "Policy Already Has Adjustment (Reg No. " & rs.fields("RegNo") & ")"
                    GoTo errmsg
                End If
            End If


            Dim SQL As String = ""
            SQL = SQL & " Declare @ANO Int"
            SQL = SQL & " Select @ANO=" & ANO
            SQL = SQL & " Select top 1 A1.RegNo, A1.PANO"
            SQL = SQL & " from Acceptance A1 "
            SQL = SQL & " Where ((A1.AType in ('A') and A1.AStatus<>'C') or A1.AType not in ('A')) and A1.pANO=@ANO"

            SQL = SQL & " Union all"

            SQL = SQL & " Select top 1 A1.RegNo, A1.PANO"
            SQL = SQL & " from Acceptance A1 "
            SQL = SQL & " Where not (A1.AType in ('C') and A1.EMethod='A' and A1.AStatus='I') and A1.RANO=@ANO"
            SQL = SQL & " and not exists (Select A2.pANO from acceptance A2 Where A2.pANO=A1.ANO)"

            rs = db.OpenRemoteSet(SQL)
            If Not rs.Eof Then
                If rs.fields("PANO") = ANO Then
                    ErrMsg = "Already Endorsed (Cancelled) ! Referred to The Policy Data. (Reg No. " & rs.fields("RegNo") & ")"
                Else
                    ErrMsg = "Already Renewed ! Referred to The Policy Data. (Reg No. " & rs.fields("RegNo") & ")"
                End If
                GoTo errmsg
            End If

            Dim CANCELONPAIDPREMIUMF As Boolean = db.GetPrivileges("CANCELONPAIDPREMIUM")
            Dim CANCELONUNPAIDPREMIUMF As Boolean = db.GetPrivileges("CANCELONUNPAIDPREMIUM")
            Dim isNotPremiumPaidF As Boolean = False

            If Not CANCELONPAIDPREMIUMF Or Not CANCELONUNPAIDPREMIUMF Then
                isNotPremiumPaidF = isNotPremiumPaid(ANO, Today.Date)
            End If

            If Not CANCELONPAIDPREMIUMF Then
                If Not isNotPremiumPaidF Then
                    ErrMsg = "Can't Cancel The Policy which is Already Paid !"
                    GoTo errmsg
                End If
            End If

            If Not CANCELONUNPAIDPREMIUMF Then
                If isNotPremiumPaidF Then
                    ErrMsg = "Can't Cancel The Policy which is Not Yet Paid !"
                    GoTo errmsg
                End If
            End If
        Else
            ' Adjustment Cancel
            rs = db.OpenRemoteSet("Select A1.RegNo, A1.PANO from Acceptance A1 Where A1.AType='C' and A1.PANO=" & ANO)
            If Not rs.Eof Then
                ErrMsg = "Already Endorsed (Cancelled) ! Referred to The Policy Data. (Reg No. " & rs.fields("RegNo") & ")"
                GoTo errmsg
            End If
        End If

        If ANO = LANO Then
            rs = db.OpenRemoteSet("Select top 1 CertificateNo from Acceptance Where LANO=" & LANO & " and AStatus in ('W','T')")
            If Not rs.Eof Then
                ErrMsg = "Policy Has a Detail with Waiting Status (Ex. Certificate No. " & rs.fields("CertificateNo") & ") !"
                GoTo errmsg
            End If
        End If

        PreCancelConditionsOK = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy Cancel : " & ErrMsg)
        PreCancelConditionsOK = False
    End Function

    Private Function PreCloseConditionsOk(ByVal ANO As Long, ByVal LANO As Long, ByVal OpenFlag As Boolean, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet

        rs = db.OpenRemoteSet("Select Top 1 RArrDet.Status from RArrDet Where Status<>'O' and ANO=" & ANO)
        If Not rs.Eof Then
            ErrMsg = "Having Confirmed Facultative Cession !"
            GoTo errmsg
        End If

        If ANO = LANO And Not OpenFlag Then
            ' Check Detail Acceptance for Conjunction Policies
            rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2 Inner Join Cover C on (C.CNO=A2.CNO) Left Join RArrHeader R on (R.ANO=A2.ANO) Where (R.Status='A' or A2.AStatus='W') and A2.CANO=A1.ANO and A1.AType in ('N','E') and A1.LANO=" & ANO)
            If Not rs.Eof Then
                ErrMsg = "Must Cancel the Co. Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                GoTo errmsg
            End If
        Else
            rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2 Inner Join Cover C on (C.CNO=A2.CNO) Left Join RArrHeader R on (R.ANO=A2.ANO) Where (R.Status='A' or A2.AStatus='W') and A2.CANO=A1.ANO and A1.AType in ('N','E') and A1.ANO=" & ANO)
            If Not rs.Eof Then
                ErrMsg = "Must Cancel the Co. Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                GoTo errmsg
            End If
        End If

        PreCloseConditionsOk = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy Close : " & ErrMsg)
        PreCloseConditionsOk = False
    End Function
    Private Function PreUndoCloseConditionsOk(ByVal ANO As Long, ByVal LANO As Long, ByVal OpenFlag As Boolean, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet

        On Error GoTo errmsg

        If ANO = LANO And Not OpenFlag Then
            ' Check Detail Acceptance for Conjunction Policies
            rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2, Cover C Where C.CStatus='C' and A2.AStatus='C' and C.CNO=A2.CNO and A2.ANO=A1.CANO and A1.AType in ('N','E') and A1.LANO=" & ANO)
            If Not rs.Eof Then
                ErrMsg = "Must Undo Close the Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                GoTo errmsg
            End If
        Else
            rs = db.OpenRemoteSet("Select Top 1 A2.ANO, C.CoverNo, A2.PolicyNo, A2.CertificateNo from Acceptance A1, Acceptance A2, Cover C Where C.CStatus='C' and A2.AStatus='C' and C.CNO=A2.CNO and A2.ANO=A1.CANO and A1.AType in ('N','E') and A1.ANO=" & ANO)
            If Not rs.Eof Then
                ErrMsg = "Must Undo Close The Quotation (" & rs.fields("CoverNo") & "), Policy (" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & ") !"
                GoTo errmsg
            End If
        End If

        PreUndoCloseConditionsOk = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy Close : " & ErrMsg)
        PreUndoCloseConditionsOk = False
    End Function
    Private Function PreExpireConditionsOk(ByVal ANO As Long, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim LANO As Long
        Dim OpenFlag As Boolean
        Dim OANO As Long
        Dim RANO As Long
        Dim sFLDID As String
        Dim sVALUEID As String
        Dim sVALUEDESC As String
        Dim AType As String
        Dim AStatus As String
        Dim TOCReallocationF As Boolean

        On Error GoTo errmsg

        rs = db.OpenRemoteSet("Select top 1 A1.PolicyNo, A1.CertificateNo, A1.AType, A1.AStatus, A1.EDate, A1.FLDID, A1.VALUEID, A1.VALUEDESC, A1.RANO, A1.OANO, A1.LANO, C.OpenFlag, R.NRDate, R.PPDate as R_PPDate, TOC.ReallocationF as TOCReallocationF from Acceptance A1 Left Join RArrheader R on (R.ANO=A1.ANO) Inner Join Cover C on (C.CNO=A1.CNO) Inner Join TOC on (TOC.TOC=A1.TOC) Where A1.ANO=" & ANO)
        If Not rs.Eof AndAlso toval(rs.fields("LANO")) <> ANO Then
            If DateDiff(DateInterval.Day, rs.fields("NRDATE"), rs.fields("EDate")) > 0 Then
                ErrMsg = "Can't Expire; The Policy : '" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & "' Needs Annual Treaty Reallocation, Please Review The Reinsurance Date on " & Format(rs.fields("NRDate"), "d") & " !"
                GoTo errmsg
            End If

            ' Clean Cut
            If DateDiff(DateInterval.Day, rs.fields("NRDate"), rs.fields("R_PPDate")) < 0 Then
                ErrMsg = "Can't Expire; The Policy : '" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & "' Needs Annual Treaty Reallocation, Please Review The Reinsurance Date on " & Format(rs.fields("R_PPDate"), "d") & " !"
                GoTo errmsg
            End If
        End If
        OANO = toval(rs.fields("OANO"))
        LANO = toval(rs.fields("LANO"))
        RANO = toval(rs.fields("RANO"))
        OpenFlag = toval(rs.fields("OpenFlag"))
        sFLDID = tostr(rs.fields("FLDID"))
        sVALUEID = tostr(rs.fields("VALUEID"))
        sVALUEDESC = tostr(rs.fields("VALUEDESC"))
        AType = tostr(rs.fields("AType"))
        AStatus = tostr(rs.fields("AStatus"))
        TOCReallocationF = toval(rs.fields("TOCReallocationF"))

        If AType = "A" Or AType = "C" Then
            ErrMsg = "Expiry Policy not for Adjustment or Cancellation !"
            GoTo errmsg
        End If

        If AStatus <> "I" Then
            ErrMsg = "Please Check Policy Data Status !"
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select top 1 A1.PolicyNo, A1.CertificateNo from Acceptance A1 Where PANO=" & ANO)
        If Not rs.Eof Then
            ErrMsg = "Can't Expire; While The Policy : '" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & "' is being Endorsed !"
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select top 1 A1.PolicyNo, A1.CertificateNo from Acceptance A1 Inner Join Acceptance Ra on (A1.OANO=Ra.OANO) Where A1.AType in ('N','E') and A1.AStatus='I' and Ra.ANO=" & RANO)
        If Not rs.Eof Then
            ErrMsg = "Can't Expire; While The Policy : '" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & "' is not yet Expired !"
            GoTo errmsg
        End If

        If TOCReallocationF Then
            rs = db.OpenRemoteSet("Select top 1 A1.PolicyNo, A1.CertificateNo from Acceptance A1, AccRegDet A2, RArrHeader R Where R.ANO=A1.ANO and R.Status='A' and A2.RegNo=R.QRegNo and A2.ReAllocF=1 and A1.ANO=" & ANO)
            If Not rs.Eof Then
                ErrMsg = "Can't Expire; The Policy : '" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & "' Needs Treaty Reallocation !"
                GoTo errmsg
            End If
        End If

        If LANO <> -1 And ANO <> LANO And OpenFlag Then
            Dim AFLDID As Object
            Dim AValueID As Object
            Dim AValueDesc As Object
            Dim i As Integer
            Dim FLDID As String
            Dim ValueID As String

            AFLDID = Convert_FLDID_Into_Array(SFLDID)
            AValueID = Convert_FLDID_Into_Array(SVALUEID)
            AValueDesc = Convert_FLDID_Into_Array(SVALUEDESC)

            For i = 0 To UBound(AValueID)
                If AValueID(i) <> AValueDesc(i) Then
                    FLDID = AFLDID(i)
                    ValueID = AValueID(i)
                    Exit For
                End If
            Next

            rs = db.OpenRemoteSet("Select Top 1 A.CheckInvoiceF from AInfoDet A Where A.ANO=" & LANO & " and A.FLDID='" & FLDID & "' and (A.VALUEID='<All>' or A.VALUEID='" & ValueID & "') and A.Status='A' Order By Case When A.VALUEID='<All>' Then 1 else 0 end")
            If toval(rs.fields("CheckInvoiceF")) Then
                rs = db.OpenRemoteSet("Select I.RefNo from ICover I Left join AccInvoice A on (A.INO=I.INO) Where coalesce(A.InvoiceNominal,0)=0 and I.ANO=" & ANO & "")
                If Not rs.Eof Then
                    ErrMsg = "Can't Expire; The Policy : '" & rs.fields("PolicyNo") & "-" & rs.fields("CertificateNo") & "', Ref No. : " & rs.fields("RefNo") & " needs Payment Invoice !"
                    GoTo errmsg
                End If
            End If
        End If

        PreExpireConditionsOk = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy Expire : " & ErrMsg)
        PreExpireConditionsOk = False
    End Function
    Private Function PreUndoExpireConditionsOk(ByVal ANO As Long, ByVal LANO As Long, ByVal OpenFlag As Boolean, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet

        On Error GoTo errmsg

        Dim RESTRICTUNDOEXPIRYONRENEWAL As Boolean = db.GetPrivileges("RESTRICTUNDOEXPIRYONRENEWAL")

        If ANO = LANO And Not OpenFlag Then
            Dim rsDet As RemoteSet
            ' Just Check one of details
            rsDet = db.OpenRemoteSet("Select Top 1 ANO from Acceptance Where AType Not in ('A','C') and ANO<>LANO and LANO=" & ANO)
            If Not rsDet.Eof Then
                rs = db.OpenRemoteSet("Select Top 1 Acceptance.NoRiskAccumulationF, Acceptance.AType, RArrHeader.Status from Acceptance, RArrHeader Where RArrHeader.ANO=Acceptance.ANO and Acceptance.AType<>'A' and Acceptance.RANO=" & rsDet.fields("ANO") & " Order by Acceptance.ANO Desc")
                If Not rs.Eof Then
                    If RESTRICTUNDOEXPIRYONRENEWAL Then
                        ErrMsg = "Can't Undo-Expire; While The Policy already Has a Renewal !"
                        GoTo errmsg
                    Else
                        If Not toval(rs.fields("NoRiskAccumulationF")) AndAlso rs.fields("AType") <> "C" AndAlso rs.fields("Status") <> "A" Then
                            ErrMsg = "Can't Undo-Expire; While The Renewed Policy is already closed !"
                            GoTo errmsg
                        End If
                    End If
                End If
            End If
        ElseIf ANO <> LANO Then
            rs = db.OpenRemoteSet("Select Top 1 Acceptance.NoRiskAccumulationF, Acceptance.AType, RArrHeader.Status from Acceptance, RArrHeader Where RArrHeader.ANO=Acceptance.ANO and Acceptance.AType<>'A' and Acceptance.RANO=" & ANO & " Order by Acceptance.ANO Desc")
            If Not rs.Eof Then
                If RESTRICTUNDOEXPIRYONRENEWAL Then
                    ErrMsg = "Can't Undo-Expire; While The Policy already Has a Renewal !"
                    GoTo errmsg
                Else
                    If Not toval(rs.fields("NoRiskAccumulationF")) AndAlso rs.fields("AType") <> "C" AndAlso rs.fields("Status") <> "A" Then
                        ErrMsg = "Can't Undo-Expire; While The Renewed Policy is already closed !"
                        GoTo errmsg
                    End If
                End If
            End If
        End If

        rs = db.OpenRemoteSet("Select Acceptance.ANO from Acceptance Where (AType not in ('A') or AStatus<>'C') and PANO=" & ANO)
        If Not rs.Eof Then
            ErrMsg = "Can't Undo-Expire; While The Policy is being Endorsed !"
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select Cover.CStatus, Cover.ASource, Acceptance.LANO, Acceptance.AType, mAcceptance.AType as MAType from Cover Inner Join Acceptance on (Cover.CNO=Acceptance.CNO) Left Join Acceptance MAcceptance on (MAcceptance.ANO=Acceptance.LANO) Where Acceptance.ANO=" & ANO)
        If rs.fields("CStatus") = "C" Then
            ErrMsg = "Can't Undo-Expire; While The Policy is Closed !"
            GoTo ErrMsg
        End If
        If rs.fields("ASource") = "MKT" And ((rs.fields("LANO") = ANO And rs.fields("MAType") = "N") Or (rs.fields("LANO") = -1 And rs.fields("AType") = "N")) Then
            ErrMsg = "Can't Undo-Expire; Confirmed Quotation Data !"
            GoTo ErrMsg
        End If

        PreUndoExpireConditionsOk = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy Undo Expire : " & ErrMsg)
        PreUndoExpireConditionsOk = False
    End Function
    Private Function PreAdjustConditionsOK(ByVal ANO As Long, ByVal NewPDate As Date, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim LANO As Long
        Dim OpenFlag As Boolean
        Dim SDate As Date
        Dim EDate As Date
        Dim PDate As Date
        Dim AdjustableF As Boolean
        Dim AStatus As String
        Dim AType As String

        On Error GoTo errmsg

        rs = db.OpenRemoteSet("Select Acceptance.AdjustableF, Acceptance.AStatus, Acceptance.AType, Acceptance.LANO, Acceptance.SDate, Acceptance.PDate, Acceptance.EDate, Cover.OpenFlag  from Cover,Acceptance where Cover.CNO=Acceptance.CNO and Acceptance.ANO=" & ANO)
        If rs.Eof Then
            ErrMsg = "Policy Data Not Found !"
            GoTo errmsg
        End If
        LANO = toval(rs.fields("LANO"))
        OpenFlag = toval(rs.fields("OpenFlag"))
        PDate = todate(rs.fields("PDate"))
        SDate = todate(rs.fields("SDate"))
        EDate = todate(rs.fields("EDate"))
        AdjustableF = toval(rs.fields("AdjustableF"))
        AStatus = tostr(rs.fields("AStatus"))
        AType = tostr(rs.fields("AType"))

        If Not (IIf(ANO = LANO And Not OpenFlag, True, AdjustableF) And (AStatus = "C" Or (AType = "C" And AStatus = "I"))) Then
            ErrMsg = "Please Check Policy Data Status !"
            GoTo errmsg
        End If

        If DateDiff(DateInterval.Day, SDate, NewPDate) < 0 Then
            ErrMsg = "Adjustment Date can't be earlier than Start Date"
            GoTo errmsg
        End If

        If ANO = LANO And Not OpenFlag Then
            rs = db.OpenRemoteSet("Select top 1 A1.LANO as ANO, A1.RegNo, A1.PolicyNo, A1.CertificateNo from Acceptance A1, Acceptance A2, AccAss A3 Where A3.ANO=A1.ANO and A3.AdjMethod='OVERLIMIT' and A1.AdjustableF=1 and A1.AType not in ('A','C') and A1.PDate<=A1.PPDate and A1.ANO=A2.PANO and A2.AdjustableF=1 and (A2.AStatus='C' or (A2.AType='C' and A2.AStatus='I')) and A2.ANO<>A2.LANO and A2.LANO=" & ANO)
        Else
            rs = db.OpenRemoteSet("Select top 1 A1.ANO as ANO, A1.RegNo, A1.PolicyNo, A1.CertificateNo from Acceptance A1, Acceptance A2, AccAss A3 Where A3.ANO=A1.ANO and A3.AdjMethod='OVERLIMIT' and A1.AdjustableF=1 and A1.AType not in ('A','C') and A1.PDate<=A1.PPDate and A1.ANO=A2.PANO and A2.ANO=" & ANO)
        End If
        If Not rs.Eof Then
            ErrMsg = "Found Data Not Yet Adjusted ! Referred To The Policy Data. (Reg No. " & rs.fields("RegNo") & ", Policy No. " & rs.fields("PolicyNo") & ", Certificate No. " & rs.fields("CertificateNo") & ")"
            GoTo errmsg
        End If

        rs = db.OpenRemoteSet("Select Top 1 A1.ANO, A1.RegNo from Acceptance A1 Where AType='A' and AStatus<>'C' and pANO=" & ANO)
        If Not rs.Eof Then
            ErrMsg = "Already Adjusted ! Referred to The Policy Data. (Reg No. " & rs.fields("RegNo") & ")"
            GoTo errmsg
        End If

        PreAdjustConditionsOK = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PNotes(ANO, "POLICY", "Policy Adjust : " & ErrMsg)
        PreAdjustConditionsOK = False
    End Function
    Private Function CheckRequestingActivities(ByVal FName As String, ByVal SType As String, ByVal AppNo As Object, ByVal PDate As Date, ByVal PTime As String, ByRef RequestingUser As String) As Boolean
        Dim SQL As String
        On Error GoTo errmsg

        SQL = "Select Top 1 UserID, RequestingUser "
        SQL = SQL & " from SysRequest"
        If SType = "UND" Or SType = "UND-MW" Then
            SQL = SQL & " Where Param1='SEND' and Param2 in ('UND','UND-MW') and Status='P' and ApprovedFName='" & FName & "' and ApprovedStatus='A' and ((LUserDate>'" & Format(PDate, "MM/DD/YYYY") & "') or (LUserDate='" & Format(PDate, "MM/DD/YYYY") & "' and LUserTime>='" & PTime & "'))"
        ElseIf SType = "CLM" Or SType = "CLM-MW" Then
            SQL = SQL & " Where Param1='SEND' and Param2 in ('CLM','CLM-MW') and Status='P' and ApprovedFName='" & FName & "' and ApprovedStatus='A' and ((LUserDate>'" & Format(PDate, "MM/DD/YYYY") & "') or (LUserDate='" & Format(PDate, "MM/DD/YYYY") & "' and LUserTime>='" & PTime & "'))"
        Else
            SQL = SQL & " Where Param1='SEND' and Param2='" & SType & "' and Status='P' and ApprovedFName='" & FName & "' and ApprovedStatus='A' and ((LUserDate>'" & Format(PDate, "MM/DD/YYYY") & "') or (LUserDate='" & Format(PDate, "MM/DD/YYYY") & "' and LUserTime>='" & PTime & "'))"
        End If
        Select Case SType
            Case "UND", "UND-MW", "IT", "MAR", "SVC", "POLH"
                SQL = SQL & " and Param3='" & AppNo & "'"
            Case "UND-OBJ"
                SQL = SQL & " and Param6='" & AppNo & "'"
            Case "CLM-EXP"
                SQL = SQL & " and Param6='" & AppNo & "'"
            Case "FA", "RQ"
                SQL = SQL & " and Param6='" & AppNo & "'"
            Case Else
                SQL = SQL & " and Param4='" & AppNo & "'"
        End Select
        SQL = SQL & " Order by ID Desc"

        Dim rs As RemoteSet
        rs = db.OpenRemoteSet(SQL)
        RequestingUser = tostr(rs.fields("RequestingUser"))
        If RequestingUser = "" Then
            RequestingUser = tostr(rs.fields("UserID"))
        End If
        CheckRequestingActivities = Not (rs.Eof)
        Exit Function
errmsg:
        CheckRequestingActivities = False
    End Function
    Private Function isAnyPremiumPaid(ByVal ANO As Long, ByVal PDate As Date) As Boolean
        Dim rs As RemoteSet
        Dim PremiumAmount As Double
        Dim PremiumPaid As Double
        rs = db.OpenQueryDef("QPremiumAmount", adOpenForwardonly, ANO, PDate)
        PremiumAmount = toval(rs.fields("PremiumAmount"))
        PremiumPaid = toval(rs.fields("PremiumPaid"))

        isAnyPremiumPaid = (PremiumPaid > 0)
    End Function
    Private Function isNotPremiumPaid(ByVal ANO As Long, ByVal PDate As Date) As Boolean
        Dim rs As RemoteSet
        Dim PremiumOutstanding As Double
        rs = db.OpenQueryDef("QPremiumAmount", adOpenForwardonly, ANO, PDate)
        PremiumOutstanding = toval(rs.fields("PremiumOutstanding"))
        isNotPremiumPaid = ((PremiumOutstanding) > toval(GetAppSetting("NominalPaidTolerance")))
    End Function
    Private Function isPremiumPaid(ByVal ANO As Long, ByVal PDate As Date) As Boolean
        Dim rs As RemoteSet
        Dim PremiumAmount As Double
        Dim PremiumPaid As Double
        rs = db.OpenQueryDef("QPremiumAmount", adOpenForwardonly, ANO, PDate)
        PremiumAmount = toval(rs.fields("PremiumAmount"))
        PremiumPaid = toval(rs.fields("PremiumPaid"))

        If (tostr(rs.fields("DueDate")) <> "" AndAlso IsDate(rs.fields("DueDate")) AndAlso DateDiff(DateInterval.Day, rs.fields("DueDate"), PDate) >= 0) Then
            isPremiumPaid = ((PremiumAmount - PremiumPaid) <= toval(GetAppSetting("NominalPaidTolerance")))
        Else
            isPremiumPaid = True
        End If
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPolicyStatus(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim CoverNo As String
            Dim PolicyNo As String
            Dim CertificateNo As String
            Dim RefNo As String
            Dim AName As String
            Dim RegID As String
            Dim FLDID As String
            Dim ValueID As String
            Dim PDate As Date
            Dim PPDate As Date
            Dim ANO As Long
            Dim ExclusiveF As Boolean

            CoverNo = getInParamValue(InParam, 0)
            PolicyNo = getInParamValue(InParam, 1)
            CertificateNo = getInParamValue(InParam, 2)
            RefNo = getInParamValue(InParam, 3)
            AName = getInParamValue(InParam, 4)
            RegID = getInParamValue(InParam, 5)
            FLDID = getInParamValue(InParam, 6)
            ValueID = getInParamValue(InParam, 7)
            PDate = getInParamValue(InParam, 8)
            PPDate = getInParamValue(InParam, 9)
            ANO = toval(getInParamValue(InParam, 10))
            ExclusiveF = toval(getInParamValue(InParam, 11))


            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = Not isDuplicateInsured(CoverNo, PolicyNo, CertificateNo, RefNo, AName, RegID, FLDID, ValueID, PDate, PPDate, ANO, ExclusiveF)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListAccountAssignment(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim LANO As Long
            Dim OpenFlag As Boolean

            ANO = getInParamValue(InParam, 0)    ' Acceptance No.
            LANO = getInParamValue(InParam, 1)      ' Master Acceptance
            OpenFlag = getInParamValue(InParam, 2)  ' Open Flag

            Dim rs As RemoteSet
            If ANO = LANO And Not OpenFlag Then
                rs = db.OpenQueryDef("View07a", adOpenKeyset, LANO)
            Else
                rs = db.OpenQueryDef("View07", adOpenKeyset, ANO)
            End If
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetListTC(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset1 As DataSet, ByRef Dataset2 As DataSet, ByRef Dataset3 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Section As String
            Dim TOPRO As String
            Dim DefaultF As Boolean
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)
            Section = getInParamValue(InParam, 1)
            TOPRO = getInParamValue(InParam, 2)
            DefaultF = getInParamValue(InParam, 3)

            Dim rs As RemoteSet

            ' Dataset1 : Clauses
            ' Dataset2 : Condition
            ' Dataset3 : Exclusion
            If Not DefaultF Then
                If TOPRO = "" Then
                    SQL = "select cast((Case When coalesce(A2.OrderNo,mA2.OrderNo) is null then 0 else 1 end) as bit) as ProcessFlag, "
                    SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                    SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                    SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, "
                    SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                    SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenCla G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join AccCla mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                    SQL = SQL & " Left Join AccCla A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                    SQL = SQL & " Left Join DefCla D on (D.TOC=TOC.TOC and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (coalesce(A2.OrderNo, '')<>'' or G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                Else
                    SQL = "select cast((Case When coalesce(A2.OrderNo,mA2.OrderNo) is null then 0 else 1 end) as bit) as ProcessFlag, "
                    SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                    SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                    SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, "
                    SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                    SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenCla G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join AccCla mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                    SQL = SQL & " Left Join AccCla A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                    SQL = SQL & " Left Join TOP_CLA D on (D.TOPRO='" & TOPRO & "' and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (coalesce(A2.OrderNo, '')<>'' or G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                End If
            Else
                If TOPRO = "" Then
                    SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenCla G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join DefCla A2 on (A2.TOC=TOC.TOC and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                Else
                    SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenCla G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join TOP_CLA A2 on (A2.TOPRO='" & TOPRO & "' and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                End If
            End If

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset1 = rs.GetDataSet


            If Not DefaultF Then
                If TOPRO = "" Then
                    SQL = "select cast((Case When coalesce(A2.OrderNo,mA2.OrderNo) is null then 0 else 1 end) as bit) as ProcessFlag, "
                    SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                    SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                    SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, "
                    SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                    SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenCon G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join AccCON mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                    SQL = SQL & " Left Join AccCON A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                    SQL = SQL & " Left Join DefCon D on (D.TOC=TOC.TOC and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (coalesce(A2.OrderNo, '')<>'' or G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                Else
                    SQL = "select cast((Case When coalesce(A2.OrderNo,mA2.OrderNo) is null then 0 else 1 end) as bit) as ProcessFlag, "
                    SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                    SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                    SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, "
                    SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                    SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenCon G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join AccCON mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                    SQL = SQL & " Left Join AccCON A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                    SQL = SQL & " Left Join TOP_CON D on (D.TOPRO='" & TOPRO & "' and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (coalesce(A2.OrderNo, '')<>'' or G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                End If
            Else
                If TOPRO = "" Then
                    SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenCon G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join DefCon A2 on (A2.TOC=TOC.TOC and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                Else
                    SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenCon G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join TOP_CON A2 on (A2.TOPRO='" & TOPRO & "' and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                End If
            End If

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset2 = rs.GetDataSet

            If Not DefaultF Then
                If TOPRO = "" Then
                    SQL = "select cast((Case When coalesce(A2.OrderNo,mA2.OrderNo) is null then 0 else 1 end) as bit) as ProcessFlag, "
                    SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                    SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                    SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, "
                    SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                    SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenEX G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join AccEX mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                    SQL = SQL & " Left Join AccEX A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                    SQL = SQL & " Left Join DefEX D on (D.TOC=TOC.TOC and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (coalesce(A2.OrderNo, '')<>'' or G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                Else
                    SQL = "select cast((Case When coalesce(A2.OrderNo,mA2.OrderNo) is null then 0 else 1 end) as bit) as ProcessFlag, "
                    SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                    SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                    SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                    SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, "
                    SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                    SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenEX G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join AccEX mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                    SQL = SQL & " Left Join AccEX A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                    SQL = SQL & " Left Join TOP_EX D on (D.TOPRO='" & TOPRO & "' and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (coalesce(A2.OrderNo, '')<>'' or G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                End If
            Else
                If TOPRO = "" Then
                    SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenEX G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join DefEX A2 on (A2.TOC=TOC.TOC and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                Else
                    SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, '' as Description, '' as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                    SQL = SQL & " from Acceptance A1 "
                    SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                    SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                    SQL = SQL & " Inner Join GenEX G on (G.COB=TOC.COB)"
                    SQL = SQL & " Left Join TOP_EX A2 on (A2.TOPRO='" & TOPRO & "' and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                    SQL = SQL & " Where A1.ANO=" & ANO
                    SQL = SQL & " and (G.AllowedF=1)"
                    SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                End If
            End If

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset3 = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveListTC(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset1 As DataSet, ByVal Dataset2 As DataSet, ByVal Dataset3 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            ' getInParamValue(InParam, 0) = ANO
            ' getInParamValue(InParam, 1) = Section
            ' getInParamValue(InParam, 2) = TOPRO
            ' getInParamValue(InParam, 3) = DefaultF
            ' getInParamValue(InParam, 4) = SaveAllF

            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            Dim j As Integer
            Const MaxCols = 8

            Dim bufferdata_1(Dataset1.Tables(0).Rows.Count - 1, MaxCols - 1) As Object
            j = 0
            For i = 0 To Dataset1.Tables(0).Rows.Count - 1
                If Dataset1.Tables(0).Rows(i)("ProcessFlag") = True Then
                    bufferdata_1(j, 0) = tostr(Dataset1.Tables(0).Rows(i)("OrderNo"))
                    bufferdata_1(j, 1) = tostr(Dataset1.Tables(0).Rows(i)("ShortDesc"))
                    bufferdata_1(j, 2) = tostr(Dataset1.Tables(0).Rows(i)("Description"))
                    bufferdata_1(j, 3) = tostr(Dataset1.Tables(0).Rows(i)("ShortDesc_1"))
                    bufferdata_1(j, 4) = tostr(Dataset1.Tables(0).Rows(i)("Description_1"))
                    bufferdata_1(j, 5) = toval(Dataset1.Tables(0).Rows(i)("DetailF"))
                    bufferdata_1(j, 6) = tostr(Dataset1.Tables(0).Rows(i)("Last_Opr"))
                    bufferdata_1(j, 7) = todate(Dataset1.Tables(0).Rows(i)("Last_Update"))
                    j = j + 1
                End If
            Next

            Dim bufferdata_2(Dataset2.Tables(0).Rows.Count - 1, MaxCols - 1) As Object
            j = 0
            For i = 0 To Dataset2.Tables(0).Rows.Count - 1
                If Dataset2.Tables(0).Rows(i)("ProcessFlag") = True Then
                    bufferdata_2(j, 0) = tostr(Dataset2.Tables(0).Rows(i)("OrderNo"))
                    bufferdata_2(j, 1) = tostr(Dataset2.Tables(0).Rows(i)("ShortDesc"))
                    bufferdata_2(j, 2) = tostr(Dataset2.Tables(0).Rows(i)("Description"))
                    bufferdata_2(j, 3) = tostr(Dataset2.Tables(0).Rows(i)("ShortDesc_1"))
                    bufferdata_2(j, 4) = tostr(Dataset2.Tables(0).Rows(i)("Description_1"))
                    bufferdata_2(j, 5) = toval(Dataset2.Tables(0).Rows(i)("DetailF"))
                    bufferdata_2(j, 6) = tostr(Dataset2.Tables(0).Rows(i)("Last_Opr"))
                    bufferdata_2(j, 7) = todate(Dataset2.Tables(0).Rows(i)("Last_Update"))
                    j = j + 1
                End If
            Next

            Dim bufferdata_3(Dataset3.Tables(0).Rows.Count - 1, MaxCols - 1) As Object
            j = 0
            For i = 0 To Dataset3.Tables(0).Rows.Count - 1
                If Dataset3.Tables(0).Rows(i)("ProcessFlag") = True Then
                    bufferdata_3(j, 0) = tostr(Dataset3.Tables(0).Rows(i)("OrderNo"))
                    bufferdata_3(j, 1) = tostr(Dataset3.Tables(0).Rows(i)("ShortDesc"))
                    bufferdata_3(j, 2) = tostr(Dataset3.Tables(0).Rows(i)("Description"))
                    bufferdata_3(j, 3) = tostr(Dataset3.Tables(0).Rows(i)("ShortDesc_1"))
                    bufferdata_3(j, 4) = tostr(Dataset3.Tables(0).Rows(i)("Description_1"))
                    bufferdata_3(j, 5) = toval(Dataset3.Tables(0).Rows(i)("DetailF"))
                    bufferdata_3(j, 6) = tostr(Dataset3.Tables(0).Rows(i)("Last_Opr"))
                    bufferdata_3(j, 7) = todate(Dataset3.Tables(0).Rows(i)("Last_Update"))
                    j = j + 1
                End If
            Next

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "SAVE_TC_DATA", InParam, TrimBuffer(bufferdata_1), TrimBuffer(bufferdata_2), TrimBuffer(bufferdata_3)) Then Throw New Exception(Me.ErrDesc)

            ' Set Output Parameter

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetTC(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Section As String
            Dim TOPRO As String
            Dim DefaultF As Boolean
            Dim TCIndex As Integer
            Dim OrderNo As String
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)
            Section = getInParamValue(InParam, 1)
            TOPRO = getInParamValue(InParam, 2)
            DefaultF = getInParamValue(InParam, 3)
            TCIndex = getInParamValue(InParam, 4)    ' 0 : Clauses, 1 : Conditions, 2 : Exclusions
            OrderNo = getInParamValue(InParam, 5)

            Dim rs As RemoteSet
            Select Case TCIndex
                Case 0
                    If Not DefaultF Then
                        If TOPRO = "" Then
                            SQL = "select "
                            SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                            SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                            SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Description,mA2.Description), G.Description) as Description, "
                            SQL = SQL & " coalesce(coalesce(A2.Description_1,mA2.Description_1), G.Description_1) as Description_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                            SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenCla G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join AccCla mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                            SQL = SQL & " Left Join AccCla A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                            SQL = SQL & " Left Join DefCla D on (D.TOC=TOC.TOC and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        Else
                            SQL = "select "
                            SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                            SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                            SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Description,mA2.Description), G.Description) as Description, "
                            SQL = SQL & " coalesce(coalesce(A2.Description_1,mA2.Description_1), G.Description_1) as Description_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                            SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenCla G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join AccCla mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                            SQL = SQL & " Left Join AccCla A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                            SQL = SQL & " Left Join TOP_CLA D on (D.TOPRO='" & TOPRO & "' and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        End If
                    Else
                        If TOPRO = "" Then
                            SQL = "select coalesce(A2.OrderNo, G.OrderNo) as OrderNo, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenCla G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join DefCla A2 on (A2.TOC=TOC.TOC and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        Else
                            SQL = "select coalesce(A2.OrderNo, G.OrderNo) as OrderNo, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenCla G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join TOP_CLA A2 on (A2.TOPRO='" & TOPRO & "' and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        End If
                    End If
                Case 1
                    If Not DefaultF Then
                        If TOPRO = "" Then
                            SQL = "select "
                            SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                            SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                            SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Description,mA2.Description), G.Description) as Description, "
                            SQL = SQL & " coalesce(coalesce(A2.Description_1,mA2.Description_1), G.Description_1) as Description_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                            SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenCon G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join AccCON mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                            SQL = SQL & " Left Join AccCON A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                            SQL = SQL & " Left Join DefCon D on (D.TOC=TOC.TOC and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        Else
                            SQL = "select "
                            SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                            SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                            SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Description,mA2.Description), G.Description) as Description, "
                            SQL = SQL & " coalesce(coalesce(A2.Description_1,mA2.Description_1), G.Description_1) as Description_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                            SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenCon G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join AccCON mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                            SQL = SQL & " Left Join AccCON A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                            SQL = SQL & " Left Join TOP_CON D on (D.TOPRO='" & TOPRO & "' and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        End If
                    Else
                        If TOPRO = "" Then
                            SQL = "select coalesce(A2.OrderNo, G.OrderNo) as OrderNo, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenCon G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join DefCon A2 on (A2.TOC=TOC.TOC and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        Else
                            SQL = "select coalesce(A2.OrderNo, G.OrderNo) as OrderNo, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenCon G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join TOP_CON A2 on (A2.TOPRO='" & TOPRO & "' and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        End If
                    End If
                Case 2
                    If Not DefaultF Then
                        If TOPRO = "" Then
                            SQL = "select "
                            SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                            SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                            SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Description,mA2.Description), G.Description) as Description, "
                            SQL = SQL & " coalesce(coalesce(A2.Description_1,mA2.Description_1), G.Description_1) as Description_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                            SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenEX G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join AccEX mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                            SQL = SQL & " Left Join AccEX A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                            SQL = SQL & " Left Join DefEX D on (D.TOC=TOC.TOC and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        Else
                            SQL = "select "
                            SQL = SQL & " coalesce(coalesce(A2.OrderNo,mA2.OrderNo), G.OrderNo) as OrderNo, "
                            SQL = SQL & " coalesce(A2.Last_Opr,mA2.Last_Opr) as Last_Opr, "
                            SQL = SQL & " cast(coalesce(coalesce(A2.Last_Update,mA2.Last_Update),convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc,mA2.ShortDesc), G.ShortDesc) as ShortDesc, "
                            SQL = SQL & " coalesce(coalesce(A2.ShortDesc_1,mA2.ShortDesc_1), G.ShortDesc_1) as ShortDesc_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Description,mA2.Description), G.Description) as Description, "
                            SQL = SQL & " coalesce(coalesce(A2.Description_1,mA2.Description_1), G.Description_1) as Description_1, "
                            SQL = SQL & " coalesce(coalesce(A2.Section,mA2.Section),'') as Section, "
                            SQL = SQL & " cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenEX G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join AccEX mA2 on (mA2.ANO=A1.LANO and C.OpenFlag=0 and mA2.OrderNo=G.OrderNo and mA2.Section='" & Section & "') "
                            SQL = SQL & " Left Join AccEX A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                            SQL = SQL & " Left Join TOP_EX D on (D.TOPRO='" & TOPRO & "' and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        End If
                    Else
                        If TOPRO = "" Then
                            SQL = "select coalesce(A2.OrderNo, G.OrderNo) as OrderNo, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenEX G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join DefEX A2 on (A2.TOC=TOC.TOC and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        Else
                            SQL = "select coalesce(A2.OrderNo, G.OrderNo) as OrderNo, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join GenEX G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join TOP_EX A2 on (A2.TOPRO='" & TOPRO & "' and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        End If
                    End If
            End Select

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPolicyIncompleteInfo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            Dim CNO As Long
            Dim LANO As Long
            Dim COCNO As String
            rs = db.OpenRemoteSet("Select CNO, LANO, COCNO from Acceptance Where ANO=" & ANO)
            CNO = toval(rs.fields("CNO"))
            LANO = toval(rs.fields("LANO"))
            COCNO = tostr(rs.fields("COCNO"))

            Dim ISType As String
            Dim OpenFlag As Boolean
            rs = db.OpenRemoteSet("Select ISType, OpenFlag from Cover Where CNO=" & CNO)
            ISType = tostr(rs.fields("ISType"))
            OpenFlag = toval(rs.fields("OpenFlag"))

            Dim Info As String = ""

            Select Case ISType
                Case "C", "R"
                    rs = db.OpenRemoteSet("Select Top 1 CNO from ISource where CNO=" & CNO, adOpenKeyset)
                    If rs.Eof Then Info = Info & "- No Insurance Source Data" & vbCrLf

                    If ISType = "C" Then
                        rs = db.OpenRemoteSet("Select Top 1 CNO from Coinsurance where CNO=" & CNO, adOpenKeyset)
                        If rs.Eof Then Info = Info & "- No Other Coinsurance Member Data" & vbCrLf
                    End If
                Case "L"
                    rs = db.OpenRemoteSet("Select Top 1 CNO from Coinsurance where CNO=" & CNO, adOpenKeyset)
                    If rs.Eof Then Info = Info & "- No Coinsurance Member Data" & vbCrLf
            End Select

            rs = db.OpenRemoteSet("Select Top 1 RCover.ANO from RCover where ANO=" & ANO, adOpenKeyset)
            If rs.Eof Then Info = Info & "- No Risk Coverage Data" & vbCrLf

            If COCNO = "" Then
                rs = db.OpenRemoteSet("Select Top 1 RCover.ANO from RCover, RateTab where RateTab.Code=RCover.Code and RateTab.CompulsoryF=1 and RCover.ANO=" & ANO, adOpenKeyset)
                If Not rs.Eof Then Info = Info & "- No COCNO for Compulsory Cover" & vbCrLf
            End If

            rs = db.OpenRemoteSet("Select Top 1 ICover.ANO from ICover where ANO=" & ANO, adOpenKeyset)
            If rs.Eof Then Info = Info & "- No Interest Insured Data" & vbCrLf

            rs = db.OpenRemoteSet("Select Top 1 AccDed.ANO from AccDed where ANO=" & ANO, adOpenKeyset)
            If rs.Eof Then Info = Info & "- No Deductible Data" & vbCrLf

            rs = db.OpenRemoteSet("Select Top 1 AccFD.ANO from AccFD where ANO=" & ANO, adOpenKeyset)
            If rs.Eof Then Info = Info & "- No Fee & Duty Data" & vbCrLf

            rs = db.OpenRemoteSet("Select Top 1 AccCla.ANO from AccCla where ANO=" & ANO, adOpenKeyset)
            If rs.Eof Then Info = Info & "- No Terms and Conditions Data" & vbCrLf

            rs = db.OpenRemoteSet("Select Top 1 AccBS.ANO from AccBS where ANO=" & ANO, adOpenKeyset)
            If rs.Eof Then Info = Info & "- No Agent/Broker Data" & vbCrLf

            If LANO <> -1 And Not OpenFlag Then
                rs = db.OpenRemoteSet("Select Top 1 Acceptance.ANO from Acceptance where ANO<>LANO and LANO=" & LANO, adOpenKeyset)
                If rs.Eof Then Info = Info & "- No Detail Acceptance Data" & vbCrLf
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = Info

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListBeneficiaries(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet

            SQL = "SELECT Cast(0 as bit) as ProcessFlag, Beneficiaries.*  "
            SQL = SQL & " FROM Beneficiaries, Acceptance"
            SQL = SQL & " WHERE Beneficiaries.CNO = Acceptance.CNO And Acceptance.ANO =" & ANO

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetBeneficiaries(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim Name As String
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenQueryDef("QBeneficiaries", adOpenKeyset, CNO, ID, Name)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveBeneficiaries(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim ID As String
            Dim Name As String
            Dim PID As String
            Dim PName As String

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)
            ' Former Value
            PID = getInParamValue(InParam, 3)
            PName = getInParamValue(InParam, 4)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim EditMode As Integer
            Dim CNO As Long
            rs = db.OpenRemoteSet("Select A.AStatus, A.CNO from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * from Beneficiaries Where CNO=" & CNO & " and ID='" & PID & "' and Name='" & PName & "'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("CNO") = CNO
            rs.fields("ID") = ID
            rs.fields("Name") = Name
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropBeneficiaries(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim Name As String
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * from Beneficiaries Where CNO=" & CNO & " and ID='" & ID & "' and Name='" & Name & "'")
            If rs.Eof Then
                ErrMsg = "No Beneficiaries Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListInParty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet

            SQL = "SELECT Cast(0 as bit) as ProcessFlag, InParty.*  "
            SQL = SQL & " FROM InParty, Acceptance"
            SQL = SQL & " WHERE InParty.CNO = Acceptance.CNO And Acceptance.ANO =" & ANO
            SQL = SQL & " ORDER BY INPARTY.IPNO"

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetInParty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim Name As String
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenQueryDef("QInParty", adOpenKeyset, CNO, ID, Name)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveInParty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim ID As String
            Dim Name As String
            Dim PID As String
            Dim PName As String

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)
            ' Former Value
            PID = getInParamValue(InParam, 3)
            PName = getInParamValue(InParam, 4)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim EditMode As Integer
            Dim CNO As Long
            rs = db.OpenRemoteSet("Select A.AStatus, A.CNO from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * from InParty Where CNO=" & CNO & " and ID='" & PID & "' and Name='" & PName & "'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("CNO") = CNO
            rs.fields("ID") = ID
            rs.fields("Name") = Name
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropInParty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim Name As String
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * from InParty Where CNO=" & CNO & " and ID='" & ID & "' and Name='" & Name & "'")
            If rs.Eof Then
                ErrMsg = "No Interested Party Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function


    <WebMethod(EnableSession:=False)> _
    Public Function GetListAlternateBeneficiaries(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet

            SQL = "SELECT Cast(0 as bit) as ProcessFlag, Beneficiaries.*  "
            SQL = SQL & " FROM ABeneficiaries Beneficiaries"
            SQL = SQL & " WHERE Beneficiaries.ANO =" & ANO

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetAlternateBeneficiaries(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim Name As String

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenQueryDef("QABeneficiaries", adOpenKeyset, ANO, ID, Name)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveAlternateBeneficiaries(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim ID As String
            Dim Name As String
            Dim PID As String
            Dim PName As String

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)
            ' Former Value
            PID = getInParamValue(InParam, 3)
            PName = getInParamValue(InParam, 4)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            rs = db.OpenRemoteSet("Select * from ABeneficiaries Where ANO=" & ANO & " and ID='" & PID & "' and Name='" & PName & "'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("ANO") = ANO
            rs.fields("ID") = ID
            rs.fields("Name") = Name
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropAlternateBeneficiaries(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim Name As String
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from ABeneficiaries Where ANO=" & ANO & " and ID='" & ID & "' and Name='" & Name & "'")
            If rs.Eof Then
                ErrMsg = "No Beneficiaries Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetListAlternateInParty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet

            SQL = "SELECT Cast(0 as bit) as ProcessFlag, InParty.*  "
            SQL = SQL & " FROM AInParty InParty"
            SQL = SQL & " WHERE InParty.ANO =" & ANO
            SQL = SQL & " ORDER BY INPARTY.IPNO"

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetAlternateInParty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim Name As String

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenQueryDef("QAInParty", adOpenKeyset, ANO, ID, Name)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveAlternateInParty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim ID As String
            Dim Name As String
            Dim PID As String
            Dim PName As String

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)
            ' Former Value
            PID = getInParamValue(InParam, 3)
            PName = getInParamValue(InParam, 4)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            rs = db.OpenRemoteSet("Select * from AInParty InParty Where ANO=" & ANO & " and ID='" & PID & "' and Name='" & PName & "'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("ANO") = ANO
            rs.fields("ID") = ID
            rs.fields("Name") = Name
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropAlternateInParty(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim Name As String

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            Name = getInParamValue(InParam, 2)

            Dim rs As RemoteSet

            rs = db.OpenRemoteSet("Select * from AInParty InParty Where ANO=" & ANO & " and ID='" & ID & "' and Name='" & Name & "'")
            If rs.Eof Then
                ErrMsg = "No Interested Party Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetListSurveyor(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet

            SQL = "SELECT Cast(0 as bit) as ProcessFlag, Surveyor.*, SysUser.Name  "
            SQL = SQL & " FROM Acceptance, Surveyor, SysUser"
            SQL = SQL & " WHERE Surveyor.CNO = Acceptance.CNO And SysUser.ID = Surveyor.ID and Acceptance.ANO=" & ANO

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetSurveyor(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim SDate As Date
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            SDate = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenQueryDef("QSurveyor", adOpenKeyset, CNO, ID, SDate)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveSurveyor(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim ID As String
            Dim SDate As Date
            Dim PID As String
            Dim PSDate As Date

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            SDate = getInParamValue(InParam, 2)
            ' Former Value
            PID = getInParamValue(InParam, 3)
            PSDate = getInParamValue(InParam, 4)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim EditMode As Integer
            Dim CNO As Long
            rs = db.OpenRemoteSet("Select A.AStatus, A.CNO from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * from Surveyor Where CNO=" & CNO & " and ID='" & PID & "' and SDate='" & Format(PSDate, "MM/DD/YYYY") & "'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("CNO") = CNO
            rs.fields("ID") = ID
            rs.fields("SDate") = SDate
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropSurveyor(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim SDate As Date
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            SDate = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * from Surveyor Where CNO=" & CNO & " and ID='" & ID & "' and SDate='" & Format(SDate, "MM/DD/YYYY") & "'")
            If rs.Eof Then
                ErrMsg = "No Surveyor Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function


    <WebMethod(EnableSession:=False)> _
    Public Function GetObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FLDID As String
            Dim ValueID As String
            Dim SeqNo As Integer
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            SeqNo = toval(getInParamValue(InParam, 3))

            Dim rs As RemoteSet

            SQL = "Select A.*, coalesce(G.DESCRIPTION,A.Remarks) as Description"
            SQL = SQL & " FROM AINFODET A"
            SQL = SQL & " Left Join GENDTAB G on (G.FLDID=A.FLDID and G.VALUEID=A.VALUEID)"
            SQL = SQL & " WHERE A.ANO=" & ANO & " and A.FLDID='" & FLDID & "' and A.VALUEID='" & ValueID & "' and A.SeqNo=" & SeqNo & ""

            rs = db.OpenRemoteSet(SQL)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FLDID As String
            Dim ValueID As String
            Dim SeqNo As Integer

            ANO = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            SeqNo = toval(getInParamValue(InParam, 3))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim PolicyType As String
            Dim ISType As String
            Dim TOC As String
            Dim Branch As String
            rs = db.OpenRemoteSet("Select A.AStatus, A.AType, A.TOC, A.Branch, C.PolicyType, C.ISType from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If Not (rs.fields("AStatus") <> "C" AndAlso (rs.fields("AType") = "N" Or rs.fields("AType") = "E")) Then
                    Throw New Exception("Policy Already Closed !")
                End If
            End If
            TOC = rs.fields("TOC")
            Branch = rs.fields("Branch")
            PolicyType = rs.fields("PolicyType")
            ISType = rs.fields("ISType")

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from AInfoDet Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID = '" & ValueID & "' and SeqNo=" & SeqNo)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
                If Not (toval(rs.fields("AdmNo")) = -1 AndAlso (rs.fields("Status") = "N" Or (rs.fields("Status") <> "C" And Not (db.GetPrivileges("FORCEAPPROVEOBJECTREGISTER") And db.GetUserType <> "SUPERVISOR")))) Then
                    Throw New Exception("Please Check Object Register Status !")
                End If
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("ANO") = ANO
            rs.fields("FLDID") = FLDID
            rs.fields("ValueID") = ValueID
            rs.fields("SeqNo") = SeqNo

            Dim DocNo As String = rs.fields("DocNo")
            If db.GetPrivileges("FORCEAPPROVEOBJECTREGISTER") And db.GetUserType <> "SUPERVISOR" Then
                rs.fields("Status") = "N"
            Else
                If DocNo = "" Then
                    DocNo = GetDocumentNumbering("KLP", PolicyType, ISType, todate(rs.fields("DocDate")), Branch, TOC, "")
                    If DocNo = "" Then
                        Throw New Exception("Generate Document No. Failed ")
                    End If
                    rs.fields("DocNo") = DocNo
                End If

                rs.fields("Status") = "A"
                rs.fields("AUserID") = dbUser
                rs.fields("AUserDate") = Today.Date
                rs.fields("AUserTime") = Time()

                If Not db.Execute("Update AInfoDet Set Status='C', CUserID='" & dbUser & "', CUserDate='" & Format(Today.Date, "MM/DD/YYYY") & "', CUserTime='" & Time() & "' from AInfoDet Where Status='A' and ANO=" & ANO & " and FLDID='" & FLDID & "' and ValueID='" & ValueID & "' and SeqNo<" & SeqNo) Then
                    Throw New Exception(db.ErrDesc)
                End If
            End If
            rs.fields("AdmNo") = -1
            rs.fields("Last_Opr") = dbUser
            rs.fields("Last_Update") = Today.Date
            rs.fields("Last_Time") = Time()
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            ReDim OutParam(0)
            OutParam(0) = docno

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FLDID As String
            Dim ValueID As String
            Dim SeqNo As Integer

            ANO = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            SeqNo = toval(getInParamValue(InParam, 3))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus, AType from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If Not (rs.fields("AStatus") <> "C" AndAlso (rs.fields("AType") = "N" Or rs.fields("AType") = "E")) Then
                    Throw New Exception("Policy Already Closed !")
                End If
            End If

            rs = db.OpenRemoteSet("Select * from AInfoDet Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID='" & ValueID & "' and SeqNo=" & SeqNo)
            If rs.Eof Then
                ErrMsg = "No Object Register Data !"
                Throw New Exception(ErrMsg)
            End If

            If toval(rs.fields("AdmNo")) <> -1 Then
                Throw New Exception("Please Check Object Register Posting Status !")
            Else
                If db.GetPrivileges("FORCEAPPROVEOBJECTREGISTER") And db.GetUserType <> "SUPERVISOR" Then
                    If Not (rs.fields("Status") = "N") Then
                        Throw New Exception("Please Check Object Register Status !")
                    End If
                End If
            End If
            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            rs = db.OpenRemoteSet("Select ImageID from AInfoDetPIC Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID='" & ValueID & "'")
            If Not rs.Eof Then
                Do While Not rs.Eof
                    If toval(rs.fields("ImageID")) <> 0 Then
                        Dim wsc As New localhost6.WSCommon(GetAppSetting("CareWSImage"))
                        wsc.Timeout = System.Threading.Timeout.Infinite
                        Dim Param() As Object = {ANO, FLDID, ValueID, toval(rs.fields("ImageID"))}
                        If Not wsc.DropObjectRegisterDocument(dbUser, dbPassword, Param, ErrMsg) Then Throw New Exception(ErrMsg)
                    End If
                    rs.MoveNext()
                Loop
            End If

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetRevisionObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FLDID As String
            Dim ValueID As String
            Dim SeqNo As Integer

            ANO = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            SeqNo = toval(getInParamValue(InParam, 3))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus, A.AType from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If Not (rs.fields("AStatus") <> "C" AndAlso (rs.fields("AType") = "N" Or rs.fields("AType") = "E")) Then
                    Throw New Exception("Policy Already Closed !")
                End If
            End If

            Dim SQL As String
            SQL = "Select A.*, coalesce(G.DESCRIPTION,A.Remarks) as Description"
            SQL = SQL & " FROM AINFODET A"
            SQL = SQL & " Left Join GENDTAB G on (G.FLDID=A.FLDID and G.VALUEID=A.VALUEID)"
            SQL = SQL & " WHERE A.ANO=" & ANO & " and A.FLDID='" & FLDID & "' and A.VALUEID='" & ValueID & "' and A.SeqNo=" & SeqNo & ""

            rs = db.OpenRemoteSet(SQL)
            If rs.Eof Then
                ErrMsg = "No Object Register Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not (rs.fields("Status") = "A") Then
                Throw New Exception("Please Check Object Register Status !")
            End If
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            fbuffer(GetFieldIndex("SeqNo", findex)) = SetFieldValue(toval(rs.fields("SeqNo")) + 1)
            fbuffer(GetFieldIndex("Fee", findex)) = SetFieldValue(0)
            fbuffer(GetFieldIndex("Fee_1", findex)) = SetFieldValue(0)
            fbuffer(GetFieldIndex("Fee_2", findex)) = SetFieldValue(0)
            fbuffer(GetFieldIndex("Status", findex)) = SetFieldValue("N")

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function Search_FLDMAP_In_ObjInfo(ByVal ANO As Long, ByVal TOC As String, ByVal ParamArray FLDMAP() As String) As Object
        Dim i As Integer
        Dim j As Integer
        Dim rsTOC As RemoteSet
        Dim rsInfo As RemoteSet
        Dim FLDID As String

        Dim NItem As Integer = FLDMAP.GetUpperBound(0) + 1
        Dim NFound As Integer = 0
        Dim Value(NItem - 1) As Object

        rsTOC = db.OpenQueryDef("QTOC_GENHTAB", adOpenForwardonly, TOC)
        rsInfo = db.OpenRemoteSet("Select * from AInfo where ANO=" & ANO)
        For i = 1 To 15
            FLDID = rsTOC.fields("FLDID" & i)
            If FLDID <> "" Then
                For j = 0 To NItem - 1
                    If UCase(rsTOC.fields("FLDMAP" & i)) = UCase(FLDMAP(j)) AndAlso tostr(Value(j)) = "" Then
                        Value(j) = rsInfo.fields("VALUEID" & i)
                        If tostr(Value(j)) = "" Then
                            Value(j) = rsInfo.fields("VALUEDESC" & i)
                        End If
                        NFound = NFound + 1
                    End If
                Next
                If NFound = NItem Then Exit For
            Else
                Exit For
            End If
        Next

        If NFound <> NItem Then
            rsTOC = db.OpenQueryDef("QTOC_EXT_GENHTAB", adOpenForwardonly, TOC)
            rsInfo = db.OpenRemoteSet("Select * from AInfo_Ext where ANO=" & ANO)
            For i = 16 To 31
                FLDID = rsTOC.fields("FLDID" & i)
                If FLDID <> "" Then
                    For j = 0 To NItem - 1
                        If UCase(rsTOC.fields("FLDMAP" & i)) = UCase(FLDMAP(j)) AndAlso tostr(Value(j)) = "" Then
                            Value(j) = rsInfo.fields("VALUEID" & i)
                            If tostr(Value(j)) = "" Then
                                Value(j) = rsInfo.fields("VALUEDESC" & i)
                            End If
                            NFound = NFound + 1
                        End If
                    Next
                    If NFound = NItem Then Exit For
                Else
                    Exit For
                End If
            Next
        End If

        Search_FLDMAP_In_ObjInfo = Value
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetFLDMAP_In_ObjInfo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim TOC As String
            Dim FLDMAP As String

            ANO = getInParamValue(InParam, 0)
            TOC = getInParamValue(InParam, 1)
            FLDMAP = getInParamValue(InParam, 2)

            Dim ValueParam As Object

            ValueParam = Search_FLDMAP_In_ObjInfo(ANO, TOC, FLDMAP)

            ReDim OutParam(0)
            OutParam(0) = ValueParam(0)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetObjectRegisterNotes(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FLDID As String
            Dim ValueID As String
            Dim SeqNo As Integer

            ANO = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            SeqNo = toval(getInParamValue(InParam, 3))

            Dim rs As RemoteSet

            rs = db.OpenRemoteSet("Select Notes from AInfoDet Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID = '" & ValueID & "' and SeqNo=" & SeqNo)

            ReDim OutParam(0)
            OutParam(0) = tostr(rs.fields("Notes"))

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveObjectRegisterNotes(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FLDID As String
            Dim ValueID As String
            Dim SeqNo As Integer
            Dim Notes As String

            ANO = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            SeqNo = toval(getInParamValue(InParam, 3))
            Notes = tostr(getInParamValue(InParam, 4))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim OANO As Long
            rs = db.OpenRemoteSet("Select A.OANO, An.Notes from Acceptance A, AInfoDet An  Where A.ANO=An.ANO and An.ANO=" & ANO & " and An.FLDID='" & FLDID & "' and An.VALUEID = '" & ValueID & "' and An.SeqNo=" & SeqNo)
            If rs.Eof Then
                Throw New Exception("Object Register Not Found !")
            End If

            OANO = rs.fields("OANO")
            If rs.fields("Notes") <> Notes Then

                Save_Log_Object("AINFODET", "ANO", ANO, "FLDID", FLDID, "VALUEID", ValueID, "SEQNO", SeqNo, "Notes", rs.fields("Notes"), Notes, "Update")
                If Not db.Execute("Update AInfoDet Set Notes='" & tostr(Notes) & "' Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID = '" & ValueID & "' and SeqNo=" & SeqNo) Then Throw New Exception(db.ErrDesc)

                rs = db.OpenRemoteSet("Select * from PNOTES Where 1<>1")
                rs.AddNew()
                rs.fields("ANO") = OANO
                rs.fields("SystemF") = False
                rs.fields("Date") = Today.Date
                rs.fields("Time") = Time()
                rs.fields("UserID") = dbUser
                rs.fields("Category") = "POLICY"
                rs.fields("Notes") = Notes
                If Not rs.Update Then Throw New Exception(rs.ErrDesc)
            End If

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    Private Function GetReinsuranceRate(ByVal TOC As String, ByVal CCY As String, ByVal SDate As Date, Optional ByVal PHolder As String = "", Optional ByVal RDate As Date = #1/1/1900#) As Double
        Dim rs As RemoteSet
        Dim CCYTo As String
        Dim BaseCCY As String
        Dim Rate As Double

        If CDate(RDate) = #1/1/1900# Then
            RDate = SDate
        End If
        Rate = 1
        rs = db.OpenRemoteSet("Select Currency from SysTab", adOpenForwardonly)
        BaseCCY = rs.fields("Currency")
        If CCY.ToUpper <> BaseCCY.ToUpper Then
            CCYTo = CCY
            Do
                CCY = CCYTo
                rs = db.OpenQueryDef("GetReinsuranceRate", adOpenForwardonly, TOC, CCY, SDate, PHolder, RDate)
                CCYTo = tostr(rs.fields("CCYTO"))
                Rate = Rate * toval(rs.fields("Rate"))
            Loop While CCYTo.ToUpper <> CCY.ToUpper And CCYTo.ToUpper <> "" And CCYTo.ToUpper <> BaseCCY.ToUpper
        End If
        GetReinsuranceRate = Rate
    End Function
    Private Function GetClaimReinsuranceRate(ByVal TOC As String, ByVal CCY As String, ByVal SDate As Date, Optional ByVal PHolder As String = "", Optional ByVal RDate As Date = #1/1/1900#) As Double
        Dim rs As RemoteSet
        Dim CCYTo As String
        Dim BaseCCY As String
        Dim Rate As Double

        If CDate(RDate) = #1/1/1900# Then
            RDate = SDate
        End If
        Rate = 1
        rs = db.OpenRemoteSet("Select Currency from SysTab", adOpenForwardonly)
        BaseCCY = rs.fields("Currency")
        If CCY.ToUpper <> BaseCCY.ToUpper Then
            CCYTo = CCY
            Do
                CCY = CCYTo
                rs = db.OpenQueryDef("GetReinsuranceRate", adOpenForwardonly, TOC, CCY, SDate, PHolder, RDate)
                CCYTo = tostr(rs.fields("CCYTO"))
                Rate = Rate * toval(rs.fields("CRate"))
            Loop While CCYTo.ToUpper <> CCY.ToUpper And CCYTo.ToUpper <> "" And CCYTo.ToUpper <> BaseCCY.ToUpper
        End If
        GetClaimReinsuranceRate = Rate
    End Function
    <WebMethod(EnableSession:=False)> _
   Public Function ApproveObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FLDID As String
            Dim ValueID As String
            Dim SeqNo As Integer
            Dim AuthorizedDate As Date
            Dim AuthorizedBy As String

            ANO = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            SeqNo = toval(getInParamValue(InParam, 3))
            AuthorizedDate = todate(getInParamValue(InParam, 4))
            AuthorizedBy = tostr(getInParamValue(InParam, 5))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim TOC As String
            Dim Branch As String
            Dim ISType As String
            Dim PolicyType As String
            rs = db.OpenRemoteSet("Select A.AStatus, A.AType, A.TOC, A.Branch, C.PolicyType, C.ISType from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If Not (rs.fields("AStatus") <> "C" AndAlso (rs.fields("AType") = "N" Or rs.fields("AType") = "E")) Then
                    Throw New Exception("Policy Already Closed !")
                End If
            End If
            TOC = rs.fields("TOC")
            Branch = rs.fields("Branch")
            ISType = rs.fields("ISType")
            PolicyType = rs.fields("PolicyType")

            rs = db.OpenRemoteSet("Select * from AInfoDet Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID='" & ValueID & "' and SeqNo=" & SeqNo)
            If rs.Eof Then
                ErrMsg = "No Object Register Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not (rs.fields("Status") = "N") Then
                Throw New Exception("Please Check Object Register Status !")
            End If
            Dim ReqDate As Date = todate(rs.fields("Last_Update"))
            Dim ReqTime As String = tostr(rs.fields("Last_Time"))
            Dim AdmNo As Long = toval(rs.fields("AdmNo"))
            Dim ObjectID As String = tostr(ANO) & "-" & FLDID & "-" & ValueID

            Dim Limit As Double
            Limit = rs.fields("Limit")
            If Limit <> 0 Then
                If Not db.GetPrivileges("ALLOWPASSAPPLIMIT") Then
                    Limit = Limit * GetReinsuranceRate(TOC, rs.fields("Currency"), todate(rs.fields("DocDate")))
                    If Not GetApprovalLimit(dbUser, TOC, ISType, "In", Limit, Branch, 0, ANO, 0) Then
                        Dim RequestingUser As String = ""
                        If Not (CheckRequestingActivities("ALLOWPASSAPPLIMIT", "UND-OBJ", ObjectID, ReqDate, ReqTime, RequestingUser) AndAlso RequestingUser.ToUpper = dbUser.ToUpper) Then
                            Throw New Exception("No Limit Authorization !")
                        End If
                    End If
                End If
            End If

            rs.Edit()

            Dim DocNo As String = rs.fields("DocNo")
            If DocNo = "" Then
                DocNo = GetDocumentNumbering("KLP", PolicyType, ISType, todate(rs.fields("DocDate")), Branch, TOC, "")
                If DocNo = "" Then
                    Throw New Exception("Generate Document No. Failed ")
                End If
                rs.fields("DocNo") = DocNo
            End If

            rs.fields("Status") = "A"
            rs.fields("AUserID") = dbUser
            rs.fields("AUserDate") = Today.Date
            rs.fields("AUserTime") = Time()
            rs.fields("AuthorizedBy") = AuthorizedBy
            rs.fields("AuthorizedDate") = AuthorizedDate
            If Not rs.Update Then
                Throw New Exception(db.ErrDesc)
            End If

            ' Already Posting to admlink
            If AdmNo <> -1 Then
                If Not db.Execute("Update Admlink Set DocNo='" & DocNo & "', RefNo='" & DocNo & "' from Admlink, nVoucher Where nVoucher.Voucher=Admlink.Voucher and Admlink.AdmNo=" & AdmNo) Then
                    Throw New Exception(db.ErrDesc)
                End If

                If Not db.Execute("Update nVoucher Set DocNo='" & DocNo & "', RefNo='" & DocNo & "' from Admlink, nVoucher Where nVoucher.Voucher=Admlink.Voucher and Admlink.AdmNo=" & AdmNo) Then
                    Throw New Exception(db.ErrDesc)
                End If
            End If

            If Not db.Execute("Update AInfoDet Set Status='C', CUserID='" & dbUser & "', CUserDate='" & Format(Today.Date, "MM/DD/YYYY") & "', CUserTime='" & Time() & "' from AInfoDet Where Status='A' and ANO=" & ANO & " and FLDID='" & FLDID & "' and ValueID='" & ValueID & "' and SeqNo<" & SeqNo) Then
                Throw New Exception(db.ErrDesc)
            End If

            ' Auto Done Task if ActionList ActionEndF=1
            rs = db.OpenRemoteSet("Select S.ID from SysRequest S Left Join SysRequest pS on (pS.ID=s.PID) Inner Join ActionList A on (A.ActionCode=(case When s.ActionCode='' and coalesce(pS.ActionCode, '')<>'' then  coalesce(pS.ActionCode, '') else s.ActionCode end) and A.ActionEndF=1 and (A.RCondition=10 or A.ApprovedFName='ALLOWPASSAPPLIMIT')) Where S.Status='' and S.Param2 in ('UND-OBJ') and S.Param6='" & ObjectID & "'")
            If Not rs.Eof Then
                Dim ws As New WSCommon
                Dim IParam(7) As Object
                Dim OParam() As Object = Nothing
                IParam(0) = rs.fields("ID")
                IParam(5) = "A"
                IParam(7) = True
                If Not ws.DoneSMSTask(db, dbUser, dbPassword, IParam, OParam, ErrMsg) Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
   Public Function DeclineObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FLDID As String
            Dim ValueID As String
            Dim SeqNo As Integer

            ANO = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            SeqNo = toval(getInParamValue(InParam, 3))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim TOC As String
            Dim Branch As String
            Dim ISType As String
            rs = db.OpenRemoteSet("Select A.AStatus, A.AType, A.TOC, A.Branch, C.ISType from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If Not (rs.fields("AStatus") <> "C" AndAlso (rs.fields("AType") = "N" Or rs.fields("AType") = "E")) Then
                    Throw New Exception("Policy Already Closed !")
                End If
            End If
            TOC = rs.fields("TOC")
            Branch = rs.fields("Branch")
            ISType = rs.fields("ISType")

            rs = db.OpenRemoteSet("Select * from AInfoDet Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID='" & ValueID & "' and SeqNo=" & SeqNo)
            If rs.Eof Then
                ErrMsg = "No Object Register Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not (rs.fields("Status") = "N") Then
                Throw New Exception("Please Check Object Register Status !")
            End If

            Dim Limit As Double
            Limit = rs.fields("Limit")
            If Limit <> 0 Then
                Limit = Limit * GetReinsuranceRate(TOC, rs.fields("Currency"), todate(rs.fields("DocDate")))
                If Not GetApprovalLimit(dbUser, TOC, ISType, "In", Limit, Branch, 0, ANO, 0) Then
                    Throw New Exception("No Limit Authorization !")
                End If
            End If

            rs.Edit()
            rs.fields("Status") = "C"
            rs.fields("CUserID") = dbUser
            rs.fields("CUserDate") = Today.Date
            rs.fields("CUserTime") = Time()

            If Not rs.Update Then
                Throw New Exception(db.ErrDesc)
            End If

            If Not db.Execute("Update AInfoDet Set Status='C', CUserID='" & dbUser & "', CUserDate='" & Format(Today.Date, "MM/DD/YYYY") & "', CUserTime='" & Time() & "' from AInfoDet Where Status='A' and ANO=" & ANO & " and FLDID='" & FLDID & "' and ValueID='" & ValueID & "' and SeqNo<" & SeqNo) Then
                Throw New Exception(db.ErrDesc)
            End If

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Sub Get_BankAccount(ByVal ID As String, ByVal Currency As String, ByRef Bank As String, ByRef BankBranch As String, ByRef AccountNo As String, ByRef BankAccount As String, ByRef BankName As String)
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select Bank, BankBranch, AccountNo, BankAccount, BankName from vProAcc ProAcc Where ID='" & ID & "' and Currency='" & Currency & "'")
        Bank = tostr(rs.fields("Bank"))
        BankBranch = tostr(rs.fields("BankBranch"))
        AccountNo = tostr(rs.fields("AccountNo"))
        BankAccount = tostr(rs.fields("BankAccount"))
        BankName = tostr(rs.fields("BankName"))
    End Sub
    <WebMethod(EnableSession:=False)> _
    Public Function PostingObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean

        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim FLDID As String
            Dim ValueID As String
            Dim SeqNo As Integer

            ANO = getInParamValue(InParam, 0)
            FLDID = getInParamValue(InParam, 1)
            ValueID = getInParamValue(InParam, 2)
            SeqNo = toval(getInParamValue(InParam, 3))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet

            Dim br_Currency As String
            rs = db.OpenRemoteSet("SELECT Currency FROM SYSTAB")
            br_Currency = rs.fields("Currency")

            Dim TOC As String
            Dim Branch As String
            Dim ID As String
            Dim pPolicyNo As string
            rs = db.OpenRemoteSet("Select A.AStatus, A.AType, A.TOC, A.Branch, A.PolicyNo, Payor.ID from Acceptance A Inner Join Payor on (Payor.ANO=A.ANO) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If Not (rs.fields("AStatus") <> "C" AndAlso (rs.fields("AType") = "N" Or rs.fields("AType") = "E")) Then
                    Throw New Exception("Policy Already Closed !")
                End If
            End If
            TOC = rs.fields("TOC")
            Branch = rs.fields("Branch")
            ID = rs.fields("ID")
            pPolicyNo = rs.fields("PolicyNo")


            rs = db.OpenRemoteSet("Select * from AInfoDet Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID='" & ValueID & "' and SeqNo=" & SeqNo)
            If rs.Eof Then
                ErrMsg = "No Object Register Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not ((rs.fields("Status") = "N" OrElse rs.fields("Status") = "A") AndAlso rs.fields("AdmNo") = -1) Then
                Throw New Exception("Please Check Object Register Status !")
            End If

            Dim PID As String
            Dim PCCY As String = br_Currency
            Dim PCode As String
            Dim PRefNo As String
            Dim PDocNo As String
            Dim PADate As Date
            Dim pBillDate As String
            Dim pSubject As String
            Dim pBranch As String
            Dim PType As String

            PID = ID
            PCode = TOC
            PType = "OT"
            PRefNo = rs.fields("DocNo")
            PDocNo = rs.fields("DocNo")
            PADate = Today.Date
            pBillDate = Today.Date
            pSubject = "Others"
            pBranch = Branch

            Dim rsAdm As RemoteSet
            Dim AdmNo As Long

            rsAdm = db.OpenQueryDef("QAdmLink", adOpenKeyset, PID, PCode, PType, PCCY, PRefNo, PADate, pBillDate, PDocNo, pSubject, pBranch)
            If Not rsAdm.Eof Then
                rsAdm.Edit()
            Else
                rsAdm.AddNew()
                AdmNo = db.GetIdentity("Admlink") + 1
                rsAdm.fields("AdmNo") = AdmNo
                rsAdm.fields("ID") = PID
                rsAdm.fields("CODE") = PCode
                rsAdm.fields("TYPE") = PType
                rsAdm.fields("Currency") = PCCY
                rsAdm.fields("Subject") = pSubject
                rsAdm.fields("Branch") = pBranch
                rsAdm.fields("RefNo") = PRefNo
                rsAdm.fields("DocNo") = PDocNo
                rsAdm.fields("A_PolicyNo") = pPolicyNo
                rsAdm.fields("Date") = PADate
                rsAdm.fields("BillDate") = pBillDate
                Dim BankAccount As String = ""
                Dim BankName As String = ""
                Dim Bank As String = ""
                Dim BankBranch As String = ""
                Dim AccountNo As String = ""
                Get_BankAccount(PID, PCCY, Bank, BankBranch, AccountNo, BankAccount, BankName)
                rsAdm.fields("Bank") = Bank
                rsAdm.fields("BankBranch") = BankBranch
                rsAdm.fields("AccountNo") = AccountNo
                rsAdm.fields("BankAccount") = BankAccount
                rsAdm.fields("BankName") = BankName
            End If
            rsAdm.fields("CT") = ""
            rsAdm.fields("Source") = PID
            rsAdm.fields("AgentID") = PID
            rsAdm.fields("BrokerID") = ""
            rsAdm.fields("CPAID") = ""
            rsAdm.fields("BatchNo") = ""
            rsAdm.fields("WPC") = 0
            rsAdm.fields("Extra_1") = "-"
            rsAdm.fields("Remarks") = rs.fields("Remarks")
            rsAdm.fields("Statement") = 0
            rsAdm.fields("VoucherF") = 0
            rsAdm.fields("Amount_3") = toval(rsAdm.fields("Amount_3")) + (toval(rs.fields("Fee")) + toval(rs.fields("Fee_1")) + toval(rs.fields("Fee_2")))
            rsAdm.fields("LUserID") = dbUser
            rsAdm.fields("LDate") = Today.Date
            rsAdm.fields("LTime") = Time()

            ' Rounding Amount
            Dim Amount As Double
            Dim TAmount As Double

            TAmount = toval(rsAdm.fields("Amount_1")) + toval(rsAdm.fields("Amount_2")) + toval(rsAdm.fields("Amount_3")) + toval(rsAdm.fields("Amount_4")) + toval(rsAdm.fields("Amount_5")) + toval(rsAdm.fields("Amount_6")) + toval(rsAdm.fields("Amount_7")) + toval(rsAdm.fields("Amount_8")) + toval(rsAdm.fields("Amount_9")) + toval(rsAdm.fields("Amount_10")) + toval(rsAdm.fields("Amount_18")) + toval(rsAdm.fields("Amount_19")) + toval(rsAdm.fields("Amount_20"))
            Amount = RoundCCY(TAmount, PCCY) - TAmount - toval(rsAdm.fields("Amount_11"))
            rsAdm.fields("Amount_11") = toval(rsAdm.fields("Amount_11")) + Amount
            rsAdm.fields("Amount") = TAmount + toval(rsAdm.fields("Amount_11")) + toval(rsAdm.fields("Amount_13")) + toval(rsAdm.fields("Amount_14"))
            If toval(rsAdm.fields("Amount")) = 0 Then
                rsAdm.fields("WriteF") = 1
            Else
                rsAdm.fields("WriteF") = 0
            End If
            If Not rsAdm.Update Then Throw New Exception(rsAdm.ErrDesc)

            rs.Edit()
            rs.fields("AdmNo") = AdmNo
            If rs.Update Then
                If rsAdm.fields("WriteF") = False Then
                    If Me.Direct_Call_Process_Execute("Process.NetV3.ProcLib", "CREATE_VOUCHER", AdmNo) = "" Then
                        Throw New Exception(Me.ErrDesc)
                    End If
                End If
            Else
                Throw New Exception(rs.ErrDesc)
            End If

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenQueryDef("View17L", adOpenKeyset, ANO)

            Dataset = rs.GetDataSet

            Dim NewRow() As DataRow = Nothing
            Dim Datarow As DataRow
            Dim i As Integer
            Dim k As Integer
            Dim l As Integer
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Datarow = Dataset.Tables(0).Rows(i)

                If Datarow("Status") = "A" Then
                    k = 0
                    If NewRow Is Nothing Then
                        l = -1
                    Else
                        l = NewRow.GetUpperBound(0)
                    End If

                    For k = 0 To l
                        If NewRow(k)("Currency") = Datarow("Currency") Then
                            Exit For
                        End If
                    Next
                    If k > l Then
                        ReDim Preserve NewRow(k)
                        NewRow(k) = Dataset.Tables(0).NewRow
                        NewRow(k)("Remarks") = "Total"
                        NewRow(k)("Currency") = Datarow("Currency")
                        NewRow(k)("Limit") = 0
                    End If
                    NewRow(k)("Limit") = toval(NewRow(k)("Limit")) + toval(Datarow("Limit"))
                End If
            Next

            If Not NewRow Is Nothing Then
                For k = 0 To NewRow.GetUpperBound(0)
                    Dataset.Tables(0).Rows.Add(NewRow(k))
                Next
            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ShowObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)   ' Certificate ANO
    
            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select LANO, FLDID, VALUEID, VALUEDESC from Acceptance Where ANO=" & ANO)

            Dim AFLDID() As String
            Dim AValueID() As String
            Dim AValueDesc() As String
            AFLDID = Split(rs.fields("FLDID"), "|")
            AValueID = Split(rs.fields("VALUEID"), "|")
            AValueDesc = Split(rs.fields("VALUEDESC"), "|")
            Dim LANO As Long = rs.fields("LANO")

            Dim FLDID As String = ""
            Dim ValueID As String = ""
            Dim i As Integer
            Dim j As Integer
            For i = 0 To UBound(AValueID)
                If AValueID(i) <> AValueDesc(i) Then
                    FLDID = AFLDID(i)
                    ValueID = AValueID(i)
                    Exit For
                End If
            Next

            Dim SQL As String
            SQL = "Select A.*, coalesce(G.DESCRIPTION,A.Remarks) as Description"
            SQL = SQL & " FROM AINFODET A"
            SQL = SQL & " Left Join GENDTAB G on (G.FLDID=A.FLDID and G.VALUEID=A.VALUEID)"
            SQL = SQL & " WHERE A.ANO=" & LANO & " and A.FLDID='" & FLDID & "' and A.VALUEID='" & ValueID & "' and A.Status='A'"

            rs = db.OpenRemoteSet(SQL)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function VerifyAinfoDet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            Dim i As Long
            Dim NumVerify As Long = 0
            Dim ANO As Long = toval(getInParamValue(InParam, 0))
            Dim FLDID As String = tostr(getInParamValue(InParam, 1))
            Dim skipPolicyCheckF As Boolean = toval(getInParamValue(InParam, 2))

            db.Begintrans()

            Dim rs As RemoteSet
            Dim PolicyType As String
            Dim ISType As String
            Dim TOC As String
            Dim Branch As String
            Dim SDate As Date
            Dim EDate As Date
            rs = db.OpenRemoteSet("Select A.AStatus, A.SDate, A.EDate, A.AType, A.TOC, A.Branch, C.PolicyType, C.ISType from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If Not (rs.fields("AStatus") <> "C" AndAlso (rs.fields("AType") = "N" Or rs.fields("AType") = "E")) Then
                    Throw New Exception("Policy Already Closed !")
                End If
            End If
            TOC = rs.fields("TOC")
            Branch = rs.fields("Branch")
            PolicyType = rs.fields("PolicyType")
            ISType = rs.fields("ISType")
            SDate = rs.fields("SDate")
            EDate = rs.fields("EDate")

            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If (Dataset.Tables(0).Rows(i)("ProcessFlag") = True) Then
                    Try

                        Dim ValueID As String

                        StrMsg = ""

                        ValueID = tostr(GetItemValue("ID", Dataset, i))

                        If ValueID = "" Then
                            StrMsg = StrMsg & "Invalid Object Register ID !" & vbCrLf
                        End If

                        If toval(GetItemValue("MAXEXTPMT", Dataset, i)) < 0 Then
                            StrMsg = StrMsg & "Invalid Max. Extend Payment !" & vbCrLf
                        End If

                        If toval(GetItemValue("TERMOFPMT", Dataset, i)) < 0 Then
                            StrMsg = StrMsg & "Invalid Term OF Payment  (1) !" & vbCrLf
                        End If

                        If toval(GetItemValue("TERMOFPMT_EXT", Dataset, i)) < 0 Then
                            StrMsg = StrMsg & "Invalid Term OF Payment  (2) !" & vbCrLf
                        End If

                        If tostr(GetItemValue("TermofPMTType", Dataset, i)) = "" Then
                            StrMsg = StrMsg & "Invalid Term OF Payment  Type (1) !" & vbCrLf
                        End If

                        If tostr(GetItemValue("TermofPMTType", Dataset, i)) <> "" Then
                            rs = db.OpenRemoteSet("Select TERMOFPMTTYPE from TERMOFPMTTYPE WHERE TERMOFPMTTYPE='" & tostr(GetItemValue("TERMOFPMTTYPE", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Invalid Term OF Payment  Type (1) !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("TermofPMTType_Ext", Dataset, i)) <> "" Then
                            rs = db.OpenRemoteSet("Select TERMOFPMTTYPE from TERMOFPMTTYPE WHERE TERMOFPMTTYPE='" & tostr(GetItemValue("TermofPMTType_Ext", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Invalid Term OF Payment  Type (2) !" & vbCrLf
                            End If
                        End If

                        If toval(GetItemValue("CWP", Dataset, i)) < 0 Then
                            StrMsg = StrMsg & "Invalid Claim Waiting Period !" & vbCrLf
                        End If

                        If toval(GetItemValue("Limit", Dataset, i)) < 0 Then
                            StrMsg = StrMsg & "Invalid Nominal Limit !" & vbCrLf
                        End If

                        If toval(GetItemValue("Limit", Dataset, i)) <> 0 AndAlso tostr(GetItemValue("Currency", Dataset, i)) = "" Then
                            StrMsg = StrMsg & "Invalid Currency !" & vbCrLf
                        End If

                        If tostr(GetItemValue("Currency", Dataset, i)) <> "" Then
                            rs = db.OpenRemoteSet("Select Currency from Currency WHERE ALLOWEDF=1 and Currency='" & tostr(GetItemValue("Currency", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Invalid Currency !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("VPStart", Dataset, i)) = "" Or Not IsDate(GetItemValue("VPStart", Dataset, i)) Then
                            StrMsg = StrMsg & "Invalid Validity Start Date !" & vbCrLf
                        End If

                        If tostr(GetItemValue("VPEnd", Dataset, i)) = "" Or Not IsDate(GetItemValue("VPEnd", Dataset, i)) Then
                            StrMsg = StrMsg & "Invalid Validity End Date !" & vbCrLf
                        End If

                        If DateDiff(DateInterval.Day, todate(GetItemValue("VPStart", Dataset, i)), todate(GetItemValue("VPEnd", Dataset, i))) <= 0 Then
                            StrMsg = StrMsg & "Invalid Validity Period !" & vbCrLf
                        End If

                        If DateDiff(DateInterval.Day, todate(GetItemValue("VPStart", Dataset, i)), todate(SDate)) < 0 Then
                            StrMsg = StrMsg & "Invalid Validity Period !" & vbCrLf
                        End If

                        If DateDiff(DateInterval.Day, todate(GetItemValue("VPEnd", Dataset, i)), todate(EDate)) < 0 Then
                            StrMsg = StrMsg & "Invalid Validity Period !" & vbCrLf
                        End If

                        If StrMsg = "" Then
                            Dataset.Tables(0).Rows(i)("Status") = "Verified"
                        Else
                            Dataset.Tables(0).Rows(i)("Status") = "failed "
                        End If

                    Catch ex As Exception
                        StrMsg = StrMsg & ex.Message
                        Dataset.Tables(0).Rows(i)("Status") = "failed "
                    End Try
                    Dataset.Tables(0).Rows(i)("Message") = StrMsg
                    NumVerify = NumVerify + 1
                End If
            Next

            ReDim OutParam(1)
            OutParam(0) = NumVerify
            OutParam(1) = ""

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function UploadAinfoDet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            Dim ProcessOpt As Integer = InParam(0)
            Dim i As Long
            Dim numProcess As Long = 0

            Dim ANO As Long = toval(getInParamValue(InParam, 1))
            Dim FLDID As String = tostr(getInParamValue(InParam, 2))

            db.Begintrans()

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim PolicyType As String
            Dim ISType As String
            Dim TOC As String
            Dim Branch As String
            Dim SDate As Date
            Dim EDate As Date
            rs = db.OpenRemoteSet("Select A.AStatus, A.SDate, A.EDate, A.AType, A.TOC, A.Branch, C.PolicyType, C.ISType from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If Not (rs.fields("AStatus") <> "C" AndAlso (rs.fields("AType") = "N" Or rs.fields("AType") = "E")) Then
                    Throw New Exception("Policy Already Closed !")
                End If
            End If
            TOC = rs.fields("TOC")
            Branch = rs.fields("Branch")
            PolicyType = rs.fields("PolicyType")
            ISType = rs.fields("ISType")
            SDate = rs.fields("SDate")
            EDate = rs.fields("EDate")


            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                StrMsg = ""
                If (Dataset.Tables(0).Rows(i)("ProcessFlag") = True) And (Dataset.Tables(0).Rows(i)("Status") = "Verified" Or ProcessOpt <> 0) Then
                    If Dataset.Tables(0).Rows(i)("Status") <> "Verified" And ProcessOpt = 2 Then
                        ErrMsg = "Record No. " & i + 1 & " is not Verified Yet !"
                        Throw New Exception(ErrMsg)
                    End If
                    Try

                        Dim ValueID As String
                        Dim SeqNo As Integer
                        Dim EditMode As Integer

                        StrMsg = ""

                        ValueID = tostr(GetItemValue("ID", Dataset, i))

                        rs = db.OpenRemoteSet("Select * from AInfoDet Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID = '" & ValueID & "' Order by SeqNo Desc")
                        SeqNo = toval(rs.fields("SeqNo"))
                        If rs.Eof Then
                            rs.AddNew()
                        ElseIf rs.fields("Status") = "A" And db.GetPrivileges("FORCEAPPROVEOBJECTREGISTER") And db.GetUserType <> "SUPERVISOR" Then
                            ' Revision
                            rs.AddNew()
                            SeqNo = SeqNo + 1
                        Else
                            rs.Edit()
                            If Not (rs.fields("Status") = "N" Or (rs.fields("Status") <> "C" And toval(rs.fields("AdmNo")) = -1 And Not (db.GetPrivileges("FORCEAPPROVEOBJECTREGISTER") And db.GetUserType <> "SUPERVISOR"))) Then
                                Throw New Exception("Please Check Object Register Status !")
                            End If
                        End If
                        EditMode = rs.editmode


                        rs.fields("Remarks") = tostr(GetItemValue("Remarks", Dataset, i))
                        rs.fields("MaxExtPMT") = tostr(GetItemValue("MaxExtPMT", Dataset, i))
                        rs.fields("TermOfPMT") = toval(GetItemValue("TermOfPMT", Dataset, i))
                        rs.fields("TermofPMTType") = tostr(GetItemValue("TermofPMTType", Dataset, i))
                        rs.fields("TermOfPMT_Ext") = toval(GetItemValue("TermOfPMT_Ext", Dataset, i))
                        rs.fields("TermofPMTType_Ext") = tostr(GetItemValue("TermofPMTType_Ext", Dataset, i))
                        rs.fields("Currency") = tostr(GetItemValue("Currency", Dataset, i))
                        rs.fields("Limit") = toval(GetItemValue("Limit", Dataset, i))
                        rs.fields("VPStart") = todate(GetItemValue("VPStart", Dataset, i))
                        rs.fields("VPEnd") = todate(GetItemValue("VPEnd", Dataset, i))
                        rs.fields("DocNo") = tostr(GetItemValue("DocNo", Dataset, i))
                        rs.fields("DocDate") = tostr(GetItemValue("DocDate", Dataset, i))
                        rs.fields("Fee") = toval(GetItemValue("Fee", Dataset, i))
                        rs.fields("Fee_1") = toval(GetItemValue("Fee_1", Dataset, i))
                        rs.fields("Fee_2") = toval(GetItemValue("Fee_2", Dataset, i))
                        rs.fields("Notes") = tostr(GetItemValue("Notes", Dataset, i))
                        rs.fields("CWP") = toval(GetItemValue("CWP", Dataset, i))
                        rs.fields("ContractType") = tostr(GetItemValue("ContractType", Dataset, i))
                        rs.fields("ContractDescription") = tostr(GetItemValue("ContractDescription", Dataset, i))
                        rs.fields("Commodity") = tostr(GetItemValue("Commodity", Dataset, i))
                        rs.fields("CoShare") = toval(GetItemValue("CoShare", Dataset, i))
                        rs.fields("Rate") = toval(GetItemValue("Rate", Dataset, i))
                        rs.fields("CheckInvoiceF") = toval(GetItemValue("CheckInvoiceF", Dataset, i))
                        rs.fields("AllowedF") = True
                        rs.fields("ANO") = ANO
                        rs.fields("FLDID") = FLDID
                        rs.fields("ValueID") = ValueID
                        rs.fields("SeqNo") = SeqNo

                        Dim DocNo As String = rs.fields("DocNo")
                        If db.GetPrivileges("FORCEAPPROVEOBJECTREGISTER") And db.GetUserType <> "SUPERVISOR" Then
                            rs.fields("Status") = "N"
                        Else
                            If DocNo = "" Then
                                DocNo = GetDocumentNumbering("KLP", PolicyType, ISType, todate(rs.fields("DocDate")), Branch, TOC, "")
                                If DocNo = "" Then
                                    Throw New Exception("Generate Document No. Failed ")
                                End If
                                rs.fields("DocNo") = DocNo
                            End If

                            rs.fields("Status") = "A"
                            rs.fields("AUserID") = dbUser
                            rs.fields("AUserDate") = Today.Date
                            rs.fields("AUserTime") = Time()

                            If Not db.Execute("Update AInfoDet Set Status='C', CUserID='" & dbUser & "', CUserDate='" & Format(Today.Date, "MM/DD/YYYY") & "', CUserTime='" & Time() & "' from AInfoDet Where Status='A' and ANO=" & ANO & " and FLDID='" & FLDID & "' and ValueID='" & ValueID & "' and SeqNo<" & SeqNo) Then
                                Throw New Exception(db.ErrDesc)
                            End If
                        End If
                        rs.fields("AdmNo") = -1
                        rs.fields("Last_Opr") = dbUser
                        rs.fields("Last_Update") = Today.Date
                        rs.fields("Last_Time") = Time()
                        If Not rs.Update Then Throw New Exception(db.ErrDesc)

                        If StrMsg = "" Then
                            Dataset.Tables(0).Rows(i)("Status") = "Processed"
                        Else
                            Dataset.Tables(0).Rows(i)("Status") = "failed "
                        End If
                    Catch ex As Exception
                        StrMsg = ex.Message
                        Dataset.Tables(0).Rows(i)("Status") = "failed "
                    End Try
                    Dataset.Tables(0).Rows(i)("Message") = StrMsg
                    numProcess = numProcess + 1
                End If

            Next

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            ReDim OutParam(1)
            OutParam(0) = numProcess

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetUnderwritingInfo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * from UINFO Where CNO=" & CNO)

            findex = rs.fields_index
            fbuffer = rs.GetRecord


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveUnderwritingInfo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus, A.CNO from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If
            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * from UINFO Where CNO=" & CNO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("CNO") = CNO
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter   
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListCoinsurance(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet

            SQL = "SELECT Cast(0 as bit) as ProcessFlag, PROFILE.ID, PROFILE.Salutation + ' ' + PROFILE.NAME as Name, COINSURANCE.DOCNO, COINSURANCE.SHARE, COINSURANCE.COMM, COINSURANCE.HF, COINSURANCE.JP, COINSURANCE.RCOMM, COINSURANCE.Ujrohreins, COINSURANCE.UjrohCedant, COINSURANCE.PCOLLECTIONF, COINSURANCE.CCOLLECTIONF, Coinsurance.NotApplyDiscF, Coinsurance.CR, Coinsurance.NetF, Coinsurance.NetHF, Coinsurance.NetOutGoF, Coinsurance.WPC  "
            SQL = SQL & " FROM COINSURANCE, Profile, ACCEPTANCE"
            SQL = SQL & " WHERE COINSURANCE.ID=PROFILE.ID AND COINSURANCE.CNO=ACCEPTANCE.CNO AND ACCEPTANCE.ANO=" & ANO
            SQL = SQL & " UNION ALL"
            SQL = SQL & " SELECT Cast(0 as bit) as ProcessFlag, '' AS ID, SYSTAB.COMPANY AS NAME, ''  AS DOCNO, ISOURCE.SHARE, ISOURCE.FEE AS COMM, ISOURCE.HF, 0 as JP, 0 as RCOMM, ISOURCE.UjrohReins, ISOURCE.UjrohCedant, Cast(0 as bit) as PCOLLECTIONF, Cast(0 as bit) as CCOLLECTIONF, Cast(0 as bit) as NotApplyDiscF, ISource.CR, Cast(0 as bit) as NetF, Cast(0 as bit) as NetHF, Cast(0 as bit) as NetOutGoF, ISource.WPC"
            SQL = SQL & " FROM ISOURCE, SYSTAB, COVER, ACCEPTANCE"
            SQL = SQL & " WHERE ISOURCE.CNO=COVER.CNO AND COVER.ISTYPE='C' AND COVER.CNO=ACCEPTANCE.CNO AND ACCEPTANCE.ANO=" & ANO

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetCoinsurance(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef DataSet As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenQueryDef("QCoinsurance_1", adOpenKeyset, CNO, ID)
            findex = rs.fields_index
            fbuffer = rs.GetRecord



            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveCoinsurance(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim ID As String
            Dim PID As String

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            ' Former Value
            PID = getInParamValue(InParam, 2)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim CNO As Long
            Dim rs As RemoteSet
            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select Top 1 Cover.CNO, Cover.OCNO, Acceptance.AStatus, Acceptance.TOC, Acceptance.Branch, Acceptance.CT, Cover.CoverDate, Cover.PolicyType, Cover.ISType from Cover, Acceptance Where Cover.CNO=Acceptance.CNO and Acceptance.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            Dim TOC As String
            Dim Branch As String
            Dim CT As String
            Dim PolicyType As String
            Dim ISType As String
            Dim CoverDate As Date
            Dim DocNO As String
            Dim OCNO As Long
            CNO = rs.fields("CNO")
            OCNO = rs.fields("OCNO")
            CoverDate = rs.fields("CoverDate")
            Branch = rs.fields("Branch")
            CT = rs.fields("CT")
            TOC = rs.fields("TOC")
            PolicyType = rs.fields("PolicyType")
            ISType = rs.fields("ISType")

            rs = db.OpenRemoteSet("Select * from Coinsurance Where CNO=" & CNO & " and ID='" & PID & "'")
            If rs.Eof Then
                rs.AddNew()
                DocNO = GetFieldValue(fbuffer(GetFieldIndex("DocNo", findex)))
                If DocNO = "" Then
                    DocNO = GetDocumentNumbering("CS", PolicyType, ISType, CoverDate, Branch, TOC, CT)
                    If DocNO = "" Then Throw New Exception("Generate Coinsurance Number Failed !")
                End If
            Else
                rs.Edit()
                DocNO = rs.fields("DocNo")
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("DocNo") = DocNO
            rs.fields("CNO") = CNO
            rs.fields("ID") = ID
            If rs.editmode = AdEditAdd Then
                rs.fields("CDate") = Today.Date
                rs.fields("UserID") = dbUser
            End If
            rs.fields("LDate") = Today.Date
            rs.fields("LUserID") = dbUser
            If Not rs.Update Then Throw New Exception(db.ErrDesc)



            If Not rs.Eof Then
                If Not rs.Update Then Throw New Exception(db.ErrDesc)
            End If

            If Not db.Execute("Update Acceptance Set PStatus='I' from Acceptance Where Acceptance.AStatus='W' and Acceptance.ANO=" & ANO & "") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            ' Detail Acceptance
            If Not db.Execute("Update Acceptance Set PStatus='I' from Acceptance Where Acceptance.AStatus='W' and Acceptance.ANO<>Acceptance.LANO and Acceptance.LANO=" & ANO & "") Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = DocNO

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function DropCoinsurance(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * From Coinsurance Where CNO=" & CNO & " and ID='" & ID & "'")
            If rs.Eof Then
                ErrMsg = "No Coinsurance Data !"
                Throw New Exception(ErrMsg)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Update Acceptance Set PStatus='I' from Acceptance Where Acceptance.AStatus='W' and Acceptance.ANO=" & ANO & "") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            ' Detail Acceptance
            If Not db.Execute("Update Acceptance Set PStatus='I' from Acceptance Where Acceptance.AStatus='W' and Acceptance.ANO<>Acceptance.LANO and Acceptance.LANO=" & ANO & "") Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveListCoinsurance(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            ' getInParamValue(InParam, 0) = ANO
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet

            rs = db.OpenRemoteSet("Select Top 1 Cover.CNO, Acceptance.AStatus, Acceptance.TOC, Acceptance.Branch, Acceptance.CT, Cover.CoverDate, Cover.PolicyType, Cover.ISType from Cover, Acceptance Where Cover.CNO=Acceptance.CNO and Acceptance.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If
            Dim CNO As Long
            Dim TOC As String
            Dim Branch As String
            Dim CT As String
            Dim PolicyType As String
            Dim ISType As String
            Dim CoverDate As Date

            CNO = rs.fields("CNO")
            CoverDate = rs.fields("CoverDate")
            Branch = rs.fields("Branch")
            CT = rs.fields("CT")
            TOC = rs.fields("TOC")
            PolicyType = rs.fields("PolicyType")
            ISType = rs.fields("ISType")

            If Not db.Execute("Select * into #Coinsurance from Coinsurance Where 1<>1 and CNO=" & CNO) Then Throw New Exception(db.ErrDesc)
            rs = db.OpenRemoteSet("Select * from #Coinsurance Where 1<>1")

            Dim i As Integer
            Dim DataRow As DataRow
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                DataRow = Dataset.Tables(0).Rows(i)
                If tostr(DataRow("ID")) <> "" AndAlso toval(DataRow("Share")) <> 0 Then
                    rs.AddNew()
                    rs.fields("CNO") = CNO
                    rs.fields("ID") = DataRow("ID")
                    rs.fields("DocNO") = DataRow("DocNO")
                    If rs.fields("DocNo") = "" Then
                        rs.fields("DocNo") = GetDocumentNumbering("CS", PolicyType, ISType, CoverDate, Branch, TOC, CT)
                        If rs.fields("DocNo") = "" Then Throw New Exception("Generate Coinsurance Number Failed !")
                    End If
                    rs.fields("Share") = toval(DataRow("Share"))
                    rs.fields("Comm") = toval(DataRow("Comm"))
                    rs.fields("HF") = toval(DataRow("HF"))
                    rs.fields("JP") = toval(DataRow("JP"))
                    rs.fields("RComm") = toval(DataRow("RComm"))
                    rs.fields("CR") = toval(DataRow("CR"))
                    rs.fields("UjrohReins") = toval(DataRow("UjrohReins"))
                    rs.fields("UjrohCedant") = toval(DataRow("UjrohCedant"))
                    rs.fields("PCOLLECTIONF") = toval(DataRow("PCOLLECTIONF"))
                    rs.fields("CCOLLECTIONF") = toval(DataRow("CCOLLECTIONF"))
                    rs.fields("NotApplyDiscF") = toval(DataRow("NotApplyDiscF"))
                    rs.fields("NetF") = toval(DataRow("NetF"))
                    rs.fields("NetHF") = toval(DataRow("NetHF"))
                    rs.fields("NetOutGoF") = toval(DataRow("NetOutGoF"))
                    rs.fields("WPC") = toval(DataRow("WPC"))

                    If toval(rs.fields("NetOutGoF")) Then
                        rs.fields("NetHF") = False
                    ElseIf toval(rs.fields("NetHF")) Then
                        rs.fields("NetOutGoF") = False
                    End If
                    rs.fields("CDate") = Today.Date
                    rs.fields("UserID") = dbUser
                    rs.fields("LDate") = Today.Date
                    rs.fields("LUserID") = dbUser
                    If Not rs.UpdateRow Then Throw New Exception(rs.ErrDesc)
                End If
            Next
            If Not rs.Eof Then
                If Not rs.Update Then Throw New Exception(rs.ErrDesc)
            End If

            If Not db.Execute("Update C1 Set UserID=C2.UserID, CDate=C2.CDate from #Coinsurance C1, Coinsurance C2 Where C2.CNO=C1.CNO and C2.ID=C1.ID and C2.CNO=" & CNO) Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Delete from Coinsurance Where CNO=" & CNO) Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Insert into Coinsurance Select * from #Coinsurance") Then Throw New Exception(db.ErrDesc)
          
            If Not db.Execute("Drop Table #Coinsurance") Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetInsuranceSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef DataSet1 As DataSet, ByRef DataSet2 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select CNO from Acceptance Where ANO=" & ANO)
            CNO = rs.fields("CNO")

            rs = db.OpenQueryDef("QISource", adOpenKeyset, CNO)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            rs = db.OpenRemoteSet("Select ISourceDet.*, RateTab.Description from ISourceDet, RateTab Where RateTab.Code=ISourceDet.Code and ISourceDet.CNO=" & CNO)
            DataSet1 = rs.GetDataSet

            rs = db.OpenRemoteSet("Select ISSecurity.*, Profile.Name from ISSecurity, Profile Where Profile.ID=ISSecurity.ID and ISSecurity.CNO=" & CNO)
            DataSet2 = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveInsuranceSource(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset1 As DataSet, ByVal Dataset2 As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim OverWritePayorF As Boolean
            Dim OverWriteDAF As Boolean

            ANO = getInParamValue(InParam, 0)
            OverWritePayorF = getInParamValue(InParam, 1)
            OverWriteDAF = getInParamValue(InParam, 2)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim CNO As Long
            Dim ID As String
            Dim rs As RemoteSet
            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select A.AStatus, A.CNO from Acceptance A Where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            CNO = rs.fields("CNO")
            rs = db.OpenRemoteSet("Select * from ISource Where CNO=" & CNO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("CNO") = CNO
            If Not rs.Update Then Throw New Exception(db.ErrDesc)
            ID = rs.fields("ID")

            If Not db.Execute("Delete from ISourceDet Where CNO=" & CNO) Then Throw New Exception(db.ErrDesc)
            rs = db.OpenRemoteSet("Select * from ISourceDet Where 1<>1")
            For i = 0 To Dataset1.Tables(0).Rows.Count - 1
                If tostr(Dataset1.Tables(0).Rows(i)("Code")) <> "" Then
                    rs.AddNew()
                    rs.fields("CNO") = CNO
                    rs.fields("Code") = tostr(Dataset1.Tables(0).Rows(i)("Code"))
                    rs.fields("Fee") = toval(Dataset1.Tables(0).Rows(i)("Fee"))
                    rs.fields("HF") = toval(Dataset1.Tables(0).Rows(i)("HF"))
                    If Not rs.UpdateRow Then Throw New Exception(db.ErrDesc)
                End If
            Next

            If Not rs.Eof Then
                If Not rs.Update Then Throw New Exception(db.ErrDesc)
            End If

            If Not db.Execute("Delete from ISSecurity Where CNO=" & CNO) Then Throw New Exception(db.ErrDesc)
            rs = db.OpenRemoteSet("Select * from ISSecurity Where 1<>1")
            For i = 0 To Dataset2.Tables(0).Rows.Count - 1
                If tostr(Dataset2.Tables(0).Rows(i)("ID")) <> "" Then
                    rs.AddNew()
                    rs.fields("CNO") = CNO
                    rs.fields("ID") = tostr(Dataset2.Tables(0).Rows(i)("ID"))
                    rs.fields("Share") = toval(Dataset2.Tables(0).Rows(i)("Share"))
                    If Not rs.UpdateRow Then Throw New Exception(db.ErrDesc)
                End If
            Next

            If Not rs.Eof Then
                If Not rs.Update Then Throw New Exception(db.ErrDesc)
            End If

            If Not db.Execute("Update Acceptance Set PStatus='I' from Acceptance Where Acceptance.AStatus='W' and Acceptance.ANO=" & ANO & "") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            ' Detail Acceptance
            If Not db.Execute("Update Acceptance Set PStatus='I' from Acceptance Where Acceptance.AStatus='W' and Acceptance.ANO<>Acceptance.LANO and Acceptance.LANO=" & ANO & "") Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter

            If Not Create_Payor(ANO, ID, , , OverWritePayorF) Then Throw New Exception(db.ErrDesc)
            If Not Create_AccDA(ANO, ID, OverWriteDAF) Then Throw New Exception(db.ErrDesc)

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetDeclaredValue(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim INO As Long
            Dim DPeriod As Integer  ' 
            INO = getInParamValue(InParam, 0)
            DPeriod = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            Dim ANO As Long
            Dim TOI As String
            Dim Remark As String
            Dim RINO As Long
            Dim SDate As Date
            Dim EDate As Date
            Dim DDate As Date
            Dim MaxAmount As Double
            Dim Method As String

            rs = db.OpenQueryDef("QICover", adOpenForwardonly, INO)
            ANO = toval(rs.fields("ANO"))
            TOI = tostr(rs.fields("TOI"))
            Remark = tostr(rs.fields("Remark"))
            RINO = toval(rs.fields("RINO"))
            MaxAmount = toval(rs.fields("Amount"))

            rs = db.OpenRemoteSet("Select AdjMethod from TOI Where TOI='" & TOI & "'")
            Method = tostr(rs.fields("AdjMethod"))

            rs = db.OpenRemoteSet("Select PDate, Case When AStatus='W' Then PPDate else EDate end as PPDate from Acceptance Where ANO=" & ANO)
            SDate = todate(rs.fields("PDate"))
            EDate = todate(rs.fields("PPDate"))

            DDate = Last_date_of_period(DateAdd(DateInterval.Month, DPeriod - 1, SDate), 1)
            If DateDiff(DateInterval.Day, SDate, DDate) < 0 Then DDate = SDate
            If DateDiff(DateInterval.Day, EDate, DDate) > 0 Then DDate = EDate

            If Method = "AVERAGE" Then
                Dim SQL As String
                SQL = "Select top 1 A2.PDate, ICover.SI"
                SQL = SQL & " from Acceptance A1, Acceptance a2, ICover"
                SQL = SQL & " Where A1.ANO=" & ANO & " and A2.OANO=A1.OANO and A2.ANO<=A1.PANO and A2.ANO<=A1.ANO and A2.AType not in ('A','C')"
                SQL = SQL & " and ICover.ANO=A2.ANO and ICover.Adjustable=1 and ICover.TOI='" & TOI & "' and (ICover.Remark Like '" & Remark & "' and ICover.RINO=" & RINO & ")"
                SQL = SQL & " and '" & Format(DDate, "MM/DD/YYYY") & "' between A2.PDate and A2.PPDate"
                SQL = SQL & " order By A2.PDate, A2.ANO Desc"
                rs = db.OpenRemoteSet(SQL)
                MaxAmount = toval(rs.fields("SI"))
            End If

            rs = db.OpenQueryDef("QICoverDet", adOpenKeyset, INO, DPeriod)
            Dataset = rs.GetDataSet

            ' Set Output Parameter
            ReDim OutParam(1)
            OutParam(0) = MaxAmount
            OutParam(1) = DDate


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveDeclaredValue(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim INO As Long
            Dim DPeriod As Integer
            Dim MaxAmount As Double
            INO = getInParamValue(InParam, 0)
            DPeriod = getInParamValue(InParam, 1)
            MaxAmount = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            Dim ANO As Long
            Dim TOI As String
            Dim Method As String
            Dim SDate As Date
            Dim EDate As Date
            Dim AType As String

            rs = db.OpenRemoteSet("Select ANO, TOI from ICover Where INO=" & INO)
            ANO = toval(rs.fields("ANO"))
            TOI = tostr(rs.fields("TOI"))

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            rs = db.OpenRemoteSet("Select AStatus, PDate, PPDate, AType from Acceptance Where ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            SDate = todate(rs.fields("PDate"))
            EDate = todate(rs.fields("PPDate"))
            AType = tostr(rs.fields("AType"))

            rs = db.OpenRemoteSet("Select AdjMethod from TOI Where TOI='" & TOI & "'")
            Method = tostr(rs.fields("AdjMethod"))

            If Not db.Execute("Delete from ICoverDet Where INO=" & INO & " and DPeriod=" & DPeriod) Then Throw New Exception(db.ErrDesc)

            Dim i As Integer
            Dim j As Integer

            j = 0
            rs = db.OpenRemoteSet("Select * from ICoverDet Where 1<>1")
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If tostr(Dataset.Tables(0).Rows(i)("Description")) <> "" Then
                    rs.AddNew()
                    rs.fields("INO") = INO
                    rs.fields("DPeriod") = DPeriod
                    rs.fields("OrderNo") = j + 1
                    rs.fields("Description") = tostr(Dataset.Tables(0).Rows(i)("Description"))
                    rs.fields("Qty") = toval(Dataset.Tables(0).Rows(i)("Qty"))
                    rs.fields("Unit") = tostr(Dataset.Tables(0).Rows(i)("Unit"))
                    rs.fields("UValue") = toval(Dataset.Tables(0).Rows(i)("UValue"))
                    rs.fields("Amount") = toval(Dataset.Tables(0).Rows(i)("Amount"))
                    rs.fields("ANO") = ANO
                    If Not rs.UpdateRow Then Throw New Exception(db.ErrDesc)
                    j = j + 1
                End If
            Next

            If Not rs.Eof Then
                If Not rs.Update Then Throw New Exception(db.ErrDesc)
            End If

            Dim Limit As Double
            Dim Amount As Double
            Dim rsI As RemoteSet

            If Method = "OVERLIMIT" Then
                Limit = MaxAmount
            Else
                Limit = 0
            End If

            If Method = "AVERAGE" AndAlso tostr(GetAppSetting("AdjustableStockCalculationMethod")) = "" Then
                Dim SQL As String
                Dim rsA As RemoteSet
                Dim pDate As Date
                Dim pPDate As Date

                SQL = "Select ICover.SI, b.pDate, b.PPDate"
                SQL = SQL & " from acceptance a"
                SQL = SQL & " Inner Join acceptance b on (b.OANO=a.OANO and b.ANO<a.ANO and b.pDate<b.PPdate)"
                SQL = SQL & " Cross Apply (Select sum(SI) as SI from ICover Where ICover.ANO=b.ANO and ICover.Adjustable=1 and ICover.TOI='" & TOI & "') ICover"
                SQL = SQL & " where a.ANO=" & ANO
                SQL = SQL & " and b.ANO>coalesce((Select c.ANO from Acceptance c Left join acceptance d on (d.pANO=c.ANO and d.AType='C') Where coalesce(d.ANO,0)=0 and c.AType='A' and c.OANO=a.OANO and c.ANO<a.ANO),0)"
                SQL = SQL & " Order By b.pDate, b.PPDate"

                Dim SI As Double = 0
                rsA = db.OpenRemoteSet(SQL)
                If Not rsA.Eof Then
                    i = 1
                    pDate = rsA.fields("PDate")
                    Dim NumMonth As Integer = 0
                    Do While Not rsA.Eof

                        Do While True
                            SI = toval(rsA.fields("SI"))
                            pPDate = rsA.fields("pPDate")
                            rsA.MoveNext()
                            If rsA.Eof OrElse SI <> toval(rsA.fields("SI")) Then Exit Do
                        Loop

                        j = DateDiff(DateInterval.Month, SDate, pPDate)
                        If DateDiff(DateInterval.Day, DateAdd(DateInterval.Month, j, SDate), pPDate) > 0 Then
                            j = j + 1
                        End If
                        j = j - NumMonth
                        If j > 0 Then
                            Dim Amt As Double
                            Amt = 0
                            Dim k As Integer
                            For k = 1 To j
                                rs = db.OpenQueryDef("QICoverDet", adOpenForwardonly, INO, i)
                                If Not rs.Eof Then
                                    Do While Not rs.Eof
                                        If toval(rs.fields("UValue")) > Limit Then
                                            Amt = Amt + (toval(rs.fields("UValue")) - Limit) * toval(rs.fields("Qty"))
                                        End If
                                        rs.MoveNext()
                                    Loop
                                Else
                                    Amt = Amt + SI
                                End If
                                i = i + 1
                            Next k
                            Amount = Amount + Amt / j * DateDiff(DateInterval.Day, pDate, pPDate) / DateDiff(DateInterval.Day, SDate, EDate)
                        End If
                        pDate = pPDate
                        NumMonth = NumMonth + j
                    Loop
                End If
            Else
                Amount = 0
                j = DateDiff(DateInterval.Month, SDate, EDate)
                If DateDiff(DateInterval.Day, DateAdd(DateInterval.Month, j, SDate), EDate) > 0 Then
                    j = j + 1
                End If
                For i = 1 To j
                    rs = db.OpenQueryDef("QICoverDet", adOpenForwardonly, INO, i)
                    If Not rs.Eof Then
                        Do While Not rs.Eof
                            If toval(rs.fields("UValue")) > Limit Then
                                Amount = Amount + (toval(rs.fields("UValue")) - Limit) * toval(rs.fields("Qty"))
                            End If
                            rs.MoveNext()
                        Loop
                    Else
                        If Method = "TOTAL" Then
                            Amount = Amount + toval(MaxAmount) / j
                        ElseIf Method = "AVERAGE" Then
                            Amount = Amount + toval(MaxAmount)
                        Else    ' Overlimit
                            ' Nothing
                        End If
                    End If
                Next i
                If Method = "AVERAGE" Then Amount = Amount / j
            End If

            rsI = db.OpenQueryDef("QICover", adOpenKeyset, INO)
            If rsI.Eof Then Throw New Exception(db.ErrDesc)
            rsI.Edit()
            If Method = "TOTAL" Then
                rsI.fields("SI") = Amount
            ElseIf Method = "AVERAGE" Then
                rsI.fields("SI") = Amount
            Else    ' OverLimit
                rs = db.OpenQueryDef("QICoverExt", adOpenForwardonly, INO)
                If tostr(rs.fields("FLDTAG1")) = "Maximum Carrying" Then
                    rsI.fields("SI") = Limit * toval(rs.fields("FLDVALUE1")) + Amount
                Else
                    rsI.fields("SI") = Limit + Amount
                End If
            End If
            rsI.fields("SI_OC") = rsI.fields("SI") / rsI.fields("Rate_OC")
            If Not rsI.Update Then Throw New Exception(db.ErrDesc)

            If Not db.Execute("Update Acceptance Set PStatus='I' Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected = 0 Then
                ErrMsg = "Policy Already Processed !"
                Throw New Exception(ErrMsg)
            End If

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SearchSection(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Section As String
            Dim ShortDesc As String
            Dim Description As String
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)
            Section = getInParamValue(InParam, 1)
            ShortDesc = getInParamValue(InParam, 2)
            Description = getInParamValue(InParam, 3)

            Dim rs As RemoteSet
            SQL = "Select Cast(0 as bit) as ProcessFlag, AccSection.* "
            SQL = SQL & " from AccSection "
            SQL = SQL & " Where ANO=" & ANO
            If Section <> "" Then
                SQL = SQL & " and Section Like '" & Section & "%'"
            End If
            If ShortDesc <> "" Then
                SQL = SQL & " and ShortDesc Like '" & ShortDesc & "%'"
            End If
            If Description <> "" Then
                SQL = SQL & " and Description Like '" & Description & "%'"
            End If

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetSection(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Section As String

            ANO = getInParamValue(InParam, 0)
            Section = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccSection Where ANO=" & ANO & " and Section='" & Section & "'")
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveSection(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Section As String

            ANO = getInParamValue(InParam, 0)
            Section = getInParamValue(InParam, 1)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A Where ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from AccSection Where ANO=" & ANO & " and Section='" & Section & "'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            ' Check Status
            If EditMode = AdEditInProgress Then
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            If EditMode = AdEditAdd Then
                rs.fields("ANO") = ANO
                rs.fields("Section") = Section
            End If
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    Private Function SaveSMS(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            ' Get Input Parameter
            Dim ANO As Long
            Dim RequestName As String
            Dim RequestMessage As String
            Dim ActionCode As String
            Dim ApprovedFName As String
            Dim DueDate As Date
            Dim Param2 As String

            ANO = getInParamValue(InParam, 0)
            RequestName = getInParamValue(InParam, 1)
            RequestMessage = getInParamValue(InParam, 2)
            ActionCode = getInParamValue(InParam, 3)
            ApprovedFName = getInParamValue(InParam, 4)
            DueDate = todate(getInParamValue(InParam, 5))
            Param2 = tostr(getInParamValue(InParam, 6))

            If Param2 = "" Then Param2 = "UND"

            Dim EmailF As Boolean = CBool(toval(GetAppSetting("EmailTaskF")))
            Dim EmailBody As String = ""
            Dim RCondition As Integer = 0

            Dim rs As RemoteSet
            If ActionCode <> "" Then
                rs = db.OpenRemoteSet("Select RCondition from ActionList Where ActionCode='" & ActionCode & "'")
                RCondition = toval(rs.fields("RCondition"))
            End If

            Dim rsAcc As RemoteSet
            rsAcc = db.OpenRemoteSet("Select ANO, OANO, RegNo, PolicyNo, CertificateNo, AName, RequestF, RequestMessage, RequestName from Acceptance Where ANO=" & ANO)
            If rsAcc.Eof Then
                ErrMsg = "No Data Available !"
                Throw New Exception(ErrMsg)
            End If

            rs = db.OpenRemoteSet("Select sum(Case When S.Param5<>S.UserID Then 1 else 0 end) as NumToOthers, sum(case When S.UserID<>'" & dbUser & "' and S.Param5=S.UserID then 1 else 0 end) as NumToMeOthers from SysRequest S, Acceptance A Where S.Param1='SEND' and S.Param2 in ('UND','UND-MW','IT') and A.ANO=S.Param3 and S.Status='' and A.OANO=" & rsAcc.fields("OANO"))
            If UCase(RequestName) = UCase(dbUser) Then
                If toval(rs.fields("NumToMeOthers")) <> 0 Then
                    Throw New Exception("Already in Message Request !")
                End If
            Else
                If toval(rs.fields("NumToOthers")) <> 0 Then
                    Throw New Exception("Already in Message Request !")
                End If
            End If

            EmailBody = "C/I No. : " & rsAcc.fields("RegNo") & " " & ", Policy No. : " & rsAcc.fields("PolicyNo") & ", Certificate No. : " & rsAcc.fields("certificateNo") & ", Insured : " & rsAcc.fields("AName")
            rsAcc.Edit()
            rsAcc.fields("RequestF") = True
            rsAcc.fields("RequestMessage") = RequestMessage
            rsAcc.fields("RequestName") = RequestName
            If rsAcc.fields("RequestName") = "" Then
                Dim rsA As RemoteSet
                Select Case RCondition
                    Case 1, 6, 7
                        rsA = db.OpenRemoteSet("Select Top 1 ID from fn_GetPolicyActionUser('" & ActionCode & "'," & ANO & ") S Where S.Type='U' order by LastRequestDate")
                        If Not rsA.Eof Then
                            RequestName = rsA.fields("ID")
                        Else
                            Throw New Exception("No Limit Authorization To Approve !")
                        End If
                End Select
                rsAcc.fields("RequestName") = RequestName
            End If
            If Not rsAcc.Update Then Throw New Exception(db.ErrDesc)

            If Not IPCRequest("CARE", RequestMessage, "SEND", Param2, ANO, 0, RequestName, "", dbUser, Today.Date, Time(), ActionCode, ApprovedFName, DueDate, -1, dbUser, ErrMsg, EmailF, EmailBody) Then Throw New Exception(ErrMsg)
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveMessage(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            If Not SaveSMS(dbUser, dbPassword, InParam, OutParam, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    Private Function SaveSMSObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            ' Get Input Parameter
            Dim ObjectID As String
            Dim RequestName As String
            Dim RequestMessage As String
            Dim ActionCode As String
            Dim ApprovedFName As String
            Dim DueDate As Date
            Dim Param2 As String

            ObjectID = getInParamValue(InParam, 0)   ' Dim ObjectID As String = tostr(ANO) & "-" & FLDID & "-" & ValueID
            RequestName = getInParamValue(InParam, 1)
            RequestMessage = getInParamValue(InParam, 2)
            ActionCode = getInParamValue(InParam, 3)
            ApprovedFName = getInParamValue(InParam, 4)
            DueDate = todate(getInParamValue(InParam, 5))
            Param2 = tostr(getInParamValue(InParam, 6))

            If Param2 = "" Then Param2 = "UND-OBJ"

            Dim EmailF As Boolean = CBool(toval(GetAppSetting("EmailTaskF")))
            Dim EmailBody As String = ""
            Dim RCondition As Integer = 0

            Dim rs As RemoteSet
            If ActionCode <> "" Then
                rs = db.OpenRemoteSet("Select RCondition from ActionList Where ActionCode='" & ActionCode & "'")
                RCondition = toval(rs.fields("RCondition"))
            End If

            rs = db.OpenRemoteSet("Select sum(Case When Param5<>UserID Then 1 else 0 end) as NumToOthers, sum(case When UserID<>'" & dbUser & "' and Param5=UserID then 1 else 0 end) as NumToMeOthers from SysRequest Where Param1='SEND' and Param2 in ('UND-OBJ') and Param6='" & ObjectID & "' and Status=''")
            If UCase(RequestName) = UCase(dbUser) Then
                If toval(rs.fields("NumToMeOthers")) <> 0 Then
                    Throw New Exception("Already in Message Request !")
                End If
            Else
                If toval(rs.fields("NumToOthers")) <> 0 Then
                    Throw New Exception("Already in Message Request !")
                End If
            End If

            rs = db.OpenRemoteSet("Select A1.RegNo, A1.PolicyNo, A1.AName, G.Description, A2.Status from Acceptance A1, AInfoDet A2, GenDTab G where A2.ANO=A1.ANO and G.FLDID=A2.FLDID and G.VALUEID=A2.VALUEID and cast(A2.ANO as varchar(10))+'-'+A2.FLDID+'-'+A2.VALUEID='" & ObjectID & "' Order By A2.SeqNo Desc")
            If Not rs.Eof Then
                EmailBody = "C/I No. : " & rs.fields("RegNo") & " " & ", Policy No. : " & rs.fields("PolicyNo") & ", Insured : " & rs.fields("AName") & ", Description : " & rs.fields("Description")
                rs = db.OpenRemoteSet("Select Top 1 * from AInfoDet A2 Where cast(A2.ANO as varchar(10))+'-'+A2.FLDID+'-'+A2.VALUEID='" & ObjectID & "' Order by A2.SeqNo Desc")
                rs.Edit()
                rs.fields("RequestF") = True
                rs.fields("RequestMessage") = RequestMessage
                rs.fields("RequestName") = RequestName
                If rs.fields("RequestName") = "" Then
                    Dim rsA As RemoteSet
                    Select Case RCondition
                        Case 10
                            rsA = db.OpenRemoteSet("Select Top 1 ID from fn_GetObjectRegisterActionUser('" & ActionCode & "','" & ObjectID & "') S Where S.Type='U' order by LastRequestDate")
                            If Not rsA.Eof Then
                                RequestName = rsA.fields("ID")
                            Else
                                Throw New Exception("No Limit Authorization To Approve !")
                            End If
                    End Select
                    rs.fields("RequestName") = RequestName
                End If
                If Not rs.Update Then Throw New Exception(db.ErrDesc)
            Else
                ErrMsg = "No Data Available !"
                Throw New Exception(ErrMsg)
            End If
            rs.CloseRemoteSet()

            If Not IPCRequest("CARE", RequestMessage, "SEND", Param2, 0, 0, RequestName, ObjectID, dbUser, Today.Date, Time(), ActionCode, ApprovedFName, DueDate, -1, dbUser, ErrMsg, EmailF, EmailBody) Then Throw New Exception(ErrMsg)
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveMessageObjectRegister(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            If Not SaveSMSObjectRegister(dbUser, dbPassword, InParam, OutParam, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SubmitRequestForPolicyApproval(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim RegNo As String = getInParamValue(InParam, 0)
            Dim CertificateNo As String = getInParamValue(InParam, 1)
            Dim Amount As Double = toval(getInParamValue(InParam, 2))

            Dim ANO As Long
            Dim RequestName As String = ""
            Dim RequestMessage As String = "Policy Request for Approval .."
            Dim ActionCode As String
            Dim ApprovedFName As String = ""
            Dim DueDate As Date = Today.Date

            Dim rs As RemoteSet
            Dim TOC As String
            Dim Branch As String

            rs = db.OpenRemoteSet("Select A1.LANO, A1.ANO, A1.TOC, A1.Branch, A1.AStatus, A1.pStatus, C.OpenFlag from Acceptance A1 Inner Join Cover C on (C.CNO=A1.CNO) Where A1.RegNo='" & RegNo & "' and A1.CERTIFICATENO='" & CertificateNo & "' and not exists (Select pANO from Acceptance A2 WHere A2.pANO=A1.ANO) Order By A1.ANO Desc")
            If rs.Eof Then
                Throw New Exception("No Policy/Certificate Data !")
            ElseIf rs.fields("AStatus") <> "W" Then
                Throw New Exception("Please Check Policy/Certificate Status !")
            End If
            ANO = toval(rs.fields("ANO"))
            TOC = tostr(rs.fields("TOC"))
            Branch = tostr(rs.fields("Branch"))

            If rs.fields("pStatus") = "I" Then  ' In-Consistent
                If rs.fields("LANO") = rs.fields("ANO") And Not toval(rs.fields("OpenFlag")) Then   ' Fleet
                    If Not Me.Direct_Call_Ext_Process("Process.NetV3.Proclib", "Premium_Calc_All", ANO) Then Throw New Exception(Me.ErrDesc)
                Else
                    If Not Me.Direct_Call_Ext_Process("Process.NetV3.Proclib", "Premium_Calc", ANO) Then Throw New Exception(Me.ErrDesc)
                End If
            End If

            Dim SQL As String

            SQL = ""
            SQL = SQL & " Select A1.ActionCode, A1.ApprovedFName "
            SQL = SQL & " from ActionList A1"
            SQL = SQL & " Where A1.ActionGroup='UND' and A1.ActionStartF=1 and A1.AllowedF=1 and A1.RCondition=1"
            SQL = SQL & " and exists "
            SQL = SQL & " ("
            SQL = SQL & " Select S2.Type, * "
            SQL = SQL & " from ActionUser S1, SysUser S2, SysUserTOC S3, SysUserBranch S4"
            SQL = SQL & " Where S1.ActionCode=A1.ActionCode and (S2.ID = S1.ActionUser Or S2.Type = S1.ActionUser) and S2.AllowedF=1 and cast(convert(varchar(10), getdate(),1) as smalldatetime)<=S2.Expiry and S2.LockedF=0"
            SQL = SQL & " and S3.ID=S2.ID and S3.TOC='" & TOC & "'"
            SQL = SQL & " and S4.ID=S2.ID and S4.AllowedF=1 and S4.Branch='" & Branch & "'"
            SQL = SQL & " )"

            rs = db.OpenRemoteSet(SQL)
            If rs.Eof Then
                Throw New Exception("No Action List Available for Action Group : 'UND', Action Task Condition : Check Policy Approval !")
            End If

            ActionCode = tostr(rs.fields("ActionCode"))
            ApprovedFName = tostr(rs.fields("ApprovedFName"))

            rs = db.OpenRemoteSet("Select Top 1 ID from fn_GetPolicyActionUser('" & ActionCode & "'," & ANO & ") S Where S.Type='U' order by LastRequestDate")
            If Not rs.Eof Then
                RequestName = rs.fields("ID")
            Else
                Throw New Exception("No Limit Authorization To Process !")
            End If

            ReDim InParam(5)

            InParam(0) = ANO
            InParam(1) = RequestName
            InParam(2) = RequestMessage
            InParam(3) = ActionCode
            InParam(4) = ApprovedFName
            InParam(5) = DueDate

            If Not SaveSMS(dbUser, dbPassword, InParam, OutParam, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DropSection(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Section As String

            ANO = getInParamValue(InParam, 0)
            Section = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccSection Where ANO=" & ANO & " and Section='" & Section & "'")
            If rs.Eof Then
                ErrMsg = "No Section Data !"
                Throw New Exception(ErrMsg)
            End If
            If Not db.Execute("Delete from ICover where ANO=" & ANO & " and Section='" & Section & "'") Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Delete from AccDed where ANO=" & ANO & " and Section='" & Section & "'") Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Delete from AccCla where ANO=" & ANO & " and Section='" & Section & "'") Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Delete from AccCON where ANO=" & ANO & " and Section='" & Section & "'") Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Delete from AccEx  where ANO=" & ANO & " and Section='" & Section & "'") Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Delete from FACCla where ANO=" & ANO & " and Section='" & Section & "'") Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Delete from FACCon where ANO=" & ANO & " and Section='" & Section & "'") Then Throw New Exception(db.ErrDesc)
            If Not db.Execute("Delete from FACEx  where ANO=" & ANO & " and Section='" & Section & "'") Then Throw New Exception(db.ErrDesc)
            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListPolicyServicing(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim PolicyType As String
            Dim TOC As String
            Dim Segment As String
            Dim PID As String
            Dim PolicyNo As String
            Dim Priority As String
            Dim COB As String
            Dim RefNo As String
            Dim ServiceNo As String
            Dim SQL As String

            PolicyType = getInParamValue(InParam, 0)
            TOC = getInParamValue(InParam, 1)
            Segment = getInParamValue(InParam, 2)
            PID = getInParamValue(InParam, 3)
            Policyno = getInParamValue(InParam, 4)
            Priority = getInParamValue(InParam, 5)
            COB = getInParamValue(InParam, 6)
            RefNo = getInParamValue(InParam, 7)
            ServiceNo = getInParamValue(InParam, 8)
            Dim rs As RemoteSet

            SQL = "Select Cast(0 as bit) as ProcessFlag, S.* from Service S, TOC Where TOC.TOC=S.TOC and S.ASource='UW' and S.SType='New' and S.SStatus='P' and S.PolicyType Like '" & PolicyType & "%' and S.TOC Like '" & TOC & "%' and S.Segment Like '" & Segment & "%' and S.PID Like '" & PID & "%' and S.PolicyNo Like '" & PolicyNo & "%' and S.Priority Like '" & Priority & "%' and TOC.COB Like '" & COB & "%' and S.RefNo Like '" & RefNo & "%' and S.ServiceNo Like '" & ServiceNo & "%'"
            If db.GetUserType <> "SUPERVISOR" Then
                SQL = SQL & " and Exists (Select TOC from SysUserTOC Where SysUserTOC.TOC=S.TOC and ID='" & dbUser & "')"
                SQL = SQL & " and Exists (Select Branch from SysUserBranch Where SysUserBranch.Branch=S.Branch and ID='" & dbUser & "')"
            End If

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetBusinessResult(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim CalculateF As Boolean

            ANO = getInParamValue(InParam, 0)
            CalculateF = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from BusinessResult Where ANO=" & ANO)
            If rs.Eof Or CalculateF Then
                If Not Check_Calculation(ANO, ErrMsg) Then Throw New Exception(ErrMsg)
                If Not Me.Direct_Call_Process_Execute("PROCESS.NETV3.PROCLIB", "BusinessResult_Calculation", ANO) Then Throw New Exception(Me.ErrDesc)
                rs = db.OpenRemoteSet("Select * from BusinessResult Where ANO=" & ANO)
            End If
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveBusinessResult(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A Where ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            rs = db.OpenRemoteSet("Select * from BusinessResult Where ANO=" & ANO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("ANO") = ANO
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListDocument(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)    ' Original Acceptance ANO

            Dim rs As RemoteSet
            rs = db.OpenQueryDef("QACCDOCLIST", adOpenKeyset, ANO)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    Public Function SaveListDoc(ByVal db As ClassLib, ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            Try
                Me.db = db
                Me.dbUser = db.Var_db_user
                Me.dbUserType = db.Var_db_User_Type
                Me.dbName = db.Var_Db_Name
                Me.dbServer = db.Var_Db_Server
                Me.dbTrusted = db.Get_DbTrusted
            Catch ex As Exception
            End Try

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)    ' Original Acceptance ANO

            ' Delete First
            Dim SQL As String
            SQL = "Delete ACCDOCLIST "
            SQL = SQL & " from ACCDOCLIST"
            SQL = SQL & " LEFT JOIN DOCTYPE ON (DOCTYPE.DOCTYPE=ACCDOCLIST.DOCTYPE) "
            SQL = SQL & " LEFT JOIN ACCDOCLISTPIC on (ACCDOCLISTPIC.ANO=ACCDOCLIST.ANO AND ACCDOCLISTPIC.DOCTYPE=ACCDOCLIST.DOCTYPE)"
            SQL = SQL & " where coalesce(DOCTYPE.COLLECTIONF,0)=0 AND coalesce(ACCDOCLISTPIC.ANO,0)=0 AND ACCDOCLIST.ANO=" & ANO

            If Not db.Execute(SQL) Then Throw New Exception(db.ErrDesc)
            Dim rs As RemoteSet

            Dim DocType As String

            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If (Dataset.Tables(0).Rows(i)("ProcessFlag") = True Or Dataset.Tables(0).Rows(i)("CollectionF") = True Or Dataset.Tables(0).Rows(i)("OptionalF") = True Or Dataset.Tables(0).Rows(i)("RequiredF") = True) Then
                    DocType = tostr(Dataset.Tables(0).Rows(i)("DOCTYPE"))
                    rs = db.OpenRemoteSet("Select * from ACCDOCLIST Where ANO=" & ANO & " and DocType='" & DocType & "'")
                    If rs.Eof Then
                        rs.AddNew()
                    Else
                        rs.Edit()
                    End If
                    rs.fields("ANO") = ANO
                    rs.fields("Description") = tostr(Dataset.Tables(0).Rows(i)("Description"))
                    rs.fields("DocType") = DocType
                    rs.fields("RequiredF") = toval(Dataset.Tables(0).Rows(i)("RequiredF"))
                    If toval(rs.fields("RequiredF")) Then
                        rs.fields("RequiredBy") = tostr(Dataset.Tables(0).Rows(i)("RequiredBy"))
                        rs.fields("RequiredDate") = todate(Dataset.Tables(0).Rows(i)("RequiredDate"))
                    Else
                        rs.fields("RequiredBy") = ""
                        rs.fields("RequiredDate") = ""
                    End If
                    rs.fields("availablef") = toval(Dataset.Tables(0).Rows(i)("ProcessFlag"))
                    rs.fields("RecDate") = todate(Dataset.Tables(0).Rows(i)("RecDate"))
                    rs.fields("Remarks") = tostr(Dataset.Tables(0).Rows(i)("Remarks"))
                    rs.fields("Last_Opr") = dbUser
                    rs.fields("Last_Update") = Today.Date
                    If Not rs.Update Then Throw New Exception(db.ErrDesc)
                End If
            Next
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveListDocument(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)    ' Original Acceptance ANO

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            If Not SaveListDoc(db, dbUser, dbPassword, InParam, Dataset, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListDocStatus(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Purpose As String

            ANO = getInParamValue(InParam, 0)    ' Original Acceptance No. (First ANO)
            Purpose = getInParamValue(InParam, 1)    ' Document Purpose 'POLICY' or 'UW'

            Dim rs As RemoteSet
            rs = db.OpenQueryDef("QACCDOCSTAT", adOpenKeyset, ANO, Purpose)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveListDocStatus(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)    ' Original Acceptance No. (First ANO)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim StatID As String

            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If Dataset.Tables(0).Rows(i)("ProcessFlag") = True And Dataset.Tables(0).Rows(i)("StatID") <> "" Then
                    StatID = Dataset.Tables(0).Rows(i)("StatID")
                    rs = db.OpenRemoteSet("Select * from AccDocStat Where ANO=" & ANO & " and StatID='" & StatID & "'")
                    If rs.Eof Then
                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("StatID") = StatID
                    Else
                        rs.Edit()
                    End If
                    rs.fields("Remarks") = tostr(Dataset.Tables(0).Rows(i)("Remarks"))

                    If toval(Dataset.Tables(0).Rows(i)("CheckF")) Then
                        rs.fields("ActionBy") = tostr(Dataset.Tables(0).Rows(i)("ActionBy"))
                        rs.fields("ActionDate") = todate(Dataset.Tables(0).Rows(i)("ActionDate"))
                        rs.fields("ActionTime") = tostr(Dataset.Tables(0).Rows(i)("ActionTime"))
                        rs.fields("AcceptedBy") = tostr(Dataset.Tables(0).Rows(i)("AcceptedBy"))
                        rs.fields("AcceptedDate") = todate(Dataset.Tables(0).Rows(i)("AcceptedDate"))
                        rs.fields("AcceptedTime") = tostr(Dataset.Tables(0).Rows(i)("AcceptedTime"))
                    Else
                        rs.fields("ActionBy") = ""
                        rs.fields("ActionDate") = ""
                        rs.fields("ActionTime") = ""
                        rs.fields("AcceptedBy") = ""
                        rs.fields("AcceptedDate") = ""
                        rs.fields("AcceptedTime") = ""
                    End If

                    If toval(Dataset.Tables(0).Rows(i)("CheckF")) <> toval(rs.fields("CheckF")) Then
                        rs.fields("CheckF") = toval(Dataset.Tables(0).Rows(i)("CheckF"))
                        rs.fields("Last_Opr") = dbUser
                        rs.fields("Last_Update") = Today.Date
                    End If
                    If Not rs.Update Then Throw New Exception(rs.ErrDesc)
                End If
            Next

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
   Public Function GetListCNC(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)    ' Original Acceptance No. (First ANO)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select Cast(0 as bit) as ProcessFlag, * from ACCCNC Where ANO=" & ANO)
            Dataset = rs.GetDataSet


            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
        Public Function OpenListCNC(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)    ' Original Acceptance No. (First ANO)

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If Dataset.Tables(0).Rows(i)("ProcessFlag") = True AndAlso Dataset.Tables(0).Rows(i)("CNCNO") <> "" Then
                    If Not db.Execute("Update AccCNC Set CNCStatus='O', UserID='" & dbUser & "', PDate='" & Format(Today.Date, "MM/DD/YYYY") & "',PTime='" & Time() & "' Where ANO=" & ANO & " and CNCNO='" & Dataset.Tables(0).Rows(i)("CNCNO") & "'") Then Throw New Exception(db.ErrDesc)
                End If
            Next

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
        Public Function DeclineListCNC(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)    ' Original Acceptance No. (First ANO)

            If Not db.GetLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)

            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If Dataset.Tables(0).Rows(i)("ProcessFlag") = True AndAlso Dataset.Tables(0).Rows(i)("CNCNO") <> "" Then
                    If Not db.Execute("Update AccCNC Set CNCStatus='C', UserID='" & dbUser & "', PDate='" & Format(Today.Date, "MM/DD/YYYY") & "',PTime='" & Time() & "' Where ANO=" & ANO & " and CNCNO='" & Dataset.Tables(0).Rows(i)("CNCNO") & "'") Then Throw New Exception(db.ErrDesc)
                End If
            Next

            If Not db.ReleaseLockedObject("Policy" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveListCNC(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)    ' Original Acceptance No. (First ANO)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet

            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If (Dataset.Tables(0).Rows(i)("ProcessFlag") = True Or Dataset.Tables(0).Rows(i)("CNCSTATUS") = "O") And Dataset.Tables(0).Rows(i)("CNCNO") <> "" Then
                    rs = db.OpenRemoteSet("Select * from AccCNC Where ANO=" & ANO & " and CNCNO='" & Dataset.Tables(0).Rows(i)("CNCNO") & "'")
                    If rs.Eof Then
                        rs.AddNew()
                        rs.fields("ANO") = ANO
                        rs.fields("CNCNO") = Dataset.Tables(0).Rows(i)("CNCNO")
                    Else
                        rs.Edit()
                    End If
                    rs.fields("CNCTYPE") = Dataset.Tables(0).Rows(i)("CNCTYPE")
                    rs.fields("CNCISSUEDBY") = Dataset.Tables(0).Rows(i)("CNCISSUEDBY")
                    rs.fields("CNCISSUEDDATE") = todate(Dataset.Tables(0).Rows(i)("CNCISSUEDDATE"))
                    rs.fields("CNCPLACE") = Dataset.Tables(0).Rows(i)("CNCPLACE")
                    rs.fields("UserID") = dbUser
                    rs.fields("PDate") = Today.Date
                    rs.fields("PTime") = Time()
                    rs.fields("CNCSTATUS") = "A"
                    If Not rs.Update Then Throw New Exception(rs.ErrDesc)
                End If
            Next

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function PreAddCollateralOk(ByVal OANO As Long, ByVal CollateralNo As String, ByVal AStatus As String, ByVal AType As String, ByVal COLLATERALTYPE As String, ByVal NOMINAL As Double, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim SQL As String

        On Error GoTo errmsg

        rs = db.OpenRemoteSet("Select CashDepositF from COLLATERALTYPE Where COLLATERALTYPE='" & COLLATERALTYPE & "'")
        If (toval(rs.fields("CashDepositF")) AndAlso AStatus <> "W") Or (AType <> "N" And AType <> "E") Then
            ErrMsg = "Can't Add Collateral When Policy already Processed !"
            GoTo errmsg
        End If

        SQL = "Select C.COLLATERALNO, MAX(C.NOMINAL) as NOMINAL, SUM(A.NOMINAL) AS ALLOCATED"
        SQL = SQL & " from COLLATERAL C"
        SQL = SQL & " OUTER APPLY (SELECT A1.NOMINAL from ACCCOLLATERAL A1, ACCEPTANCE A2 Where A1.COLLATERALNO=C.COLLATERALNO and A2.OANO=A1.ANO and A2.OANO<>" & OANO & " and A2.AStatus in ('W','T','I') and A2.AType in ('N','E')) A"
        SQL = SQL & " Where C.COLLATERALNO='" & CollateralNo & "'"
        SQL = SQL & " Group By C.COLLATERALNO"
        rs = db.OpenRemoteSet(SQL)
        If toval(rs.fields("NOMINAL")) <> 0 AndAlso toval(rs.fields("NOMINAL")) < toval(rs.fields("ALLOCATED")) + NOMINAL Then
            ErrMsg = "Exceeding Collateral Nominal !"
            GoTo errmsg
        End If

        PreAddCollateralOk = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PreAddCollateralOk = False
    End Function
    Private Function PreEditCollateralOk(ByVal OANO As Long, ByVal CollateralNo As String, ByVal AStatus As String, ByVal AType As String, ByVal COLLATERALTYPE As String, ByVal NOMINAL As Double, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim SQL As String

        On Error GoTo errmsg

        rs = db.OpenRemoteSet("Select CashDepositF from COLLATERALTYPE Where COLLATERALTYPE='" & COLLATERALTYPE & "'")
        If (toval(rs.fields("CashDepositF")) AndAlso AStatus <> "W") Or (AType <> "N" And AType <> "E") Then
            ErrMsg = "Can't Edit Collateral When Policy already Processed !"
            GoTo errmsg
        End If

        SQL = "Select C.COLLATERALNO, MAX(C.NOMINAL) as NOMINAL, SUM(A.NOMINAL) AS ALLOCATED"
        SQL = SQL & " from COLLATERAL C"
        SQL = SQL & " OUTER APPLY (SELECT A1.NOMINAL from ACCCOLLATERAL A1, ACCEPTANCE A2 Where A1.COLLATERALNO=C.COLLATERALNO and A2.OANO=A1.ANO and A2.OANO<>" & OANO & " and A2.AStatus in ('W','T','I') and A2.AType in ('N','E')) A"
        SQL = SQL & " Where C.COLLATERALNO='" & CollateralNo & "'"
        SQL = SQL & " Group By C.COLLATERALNO"
        rs = db.OpenRemoteSet(SQL)
        If toval(rs.fields("NOMINAL")) <> 0 AndAlso toval(rs.fields("NOMINAL")) < toval(rs.fields("ALLOCATED")) + NOMINAL Then
            ErrMsg = "Exceeding Collateral Nominal !"
            GoTo errmsg
        End If

        SQL = "SELECT C2.ClaimNo, sum(C1.AMOUNT) AS AMOUNT "
        SQL = SQL & " FROM CLSUBDET C1, CLAIM C2, CLSUB C3, ACCEPTANCE A1 "
        SQL = SQL & " WHERE A1.OANO=" & OANO & " and C1.Item='" & CollateralNo & "' and A1.ANO=C2.ANO AND C1.CNO=C2.CNO AND C3.CNO=C1.CNO AND C3.ID=C1.ID AND C3.CURRENCY=C1.CURRENCY AND C3.REFNO=C1.REFNO AND (C2.CSTATUS<>'C' or C2.PCSTATUS='A') and not exists (Select pCNO from Claim Where Claim.pCNO=C2.CNO and Claim.CStatus<>'R') AND C3.STATUS='P' "
        SQL = SQL & " GROUP BY C2.ClaimNo"

        rs = db.OpenRemoteSet(SQL)
        If Not rs.Eof Then
            ErrMsg = "Please Check Claim Subrogation, Claim No : " & rs.fields("ClaimNo")
            GoTo errmsg
        End If

        PreEditCollateralOk = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PreEditCollateralOk = False
    End Function
    Private Function PreDropCollateralOk(ByVal OANO As Long, ByVal CollateralNo As String, ByVal AStatus As String, ByVal AType As String, ByVal COLLATERALTYPE As String, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim SQL As String

        On Error GoTo errmsg

        rs = db.OpenRemoteSet("Select CashDepositF from COLLATERALTYPE Where COLLATERALTYPE='" & COLLATERALTYPE & "'")
        If (toval(rs.fields("CashDepositF")) AndAlso AStatus <> "W") Or (AType <> "N" And AType <> "E") Then
            ErrMsg = "Can't Drop Collateral When Policy already Processed !"
            GoTo errmsg
        End If

        SQL = "SELECT C2.ClaimNo, sum(C1.AMOUNT) AS AMOUNT "
        SQL = SQL & " FROM CLSUBDET C1, CLAIM C2, CLSUB C3, ACCEPTANCE A1 "
        SQL = SQL & " WHERE A1.OANO=" & OANO & " and C1.Item='" & CollateralNo & "' and A1.ANO=C2.ANO AND C1.CNO=C2.CNO AND C3.CNO=C1.CNO AND C3.ID=C1.ID AND C3.CURRENCY=C1.CURRENCY AND C3.REFNO=C1.REFNO AND (C2.CSTATUS<>'C' or C2.PCSTATUS='A') and not exists (Select pCNO from Claim Where Claim.pCNO=C2.CNO and Claim.CStatus<>'R') AND C3.STATUS='P' "
        SQL = SQL & " GROUP BY C2.ClaimNo"

        rs = db.OpenRemoteSet(SQL)
        If Not rs.Eof Then
            ErrMsg = "Please Check Claim Subrogation, Claim No : " & rs.fields("ClaimNo")
            GoTo errmsg
        End If

        PreDropCollateralOk = True
        Exit Function
errmsg:
        ErrMsg = StrError(ErrMsg)
        PreDropCollateralOk = False
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListCollateral(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)   ' Acceptance ANO

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select OANO from Acceptance Where ANO=" & ANO)
            Dim OANO As Long = rs.fields("OANO")

            rs = db.OpenQueryDef("QACCCOLLATERALLIST", adOpenKeyset, OANO)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetCollateral(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CollateralNo As String
            ANO = getInParamValue(InParam, 0)   ' Acceptance ANO 
            CollateralNo = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select OANO from Acceptance Where ANO=" & ANO)
            Dim OANO As Long = rs.fields("OANO")

            rs = db.OpenRemoteSet("Select * from AccCollateral Where ANO=" & OANO & " and CollateralNo='" & CollateralNo & "'")
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function UpdateCollateralUsedInProfile(ByVal rsSrc As RemoteSet, ByVal CollateralNo As String) As Boolean
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select * from Collateral Where CollateralNo='" & CollateralNo & "'")
        If Not rs.Eof Then
            rs.Edit()
            rs.fields("ExpDate") = rsSrc.fields("ExpDate")
            rs.fields("RecDate") = rsSrc.fields("RecDate")
            rs.fields("RECEIVEDBY") = rsSrc.fields("RECEIVEDBY")
            rs.fields("RECEIVEDBYID_NO") = rsSrc.fields("RECEIVEDBYID_NO")
            rs.fields("RECEIVEDBYTITLE") = rsSrc.fields("RECEIVEDBYTITLE")
            rs.fields("RECEIVEDFROM") = rsSrc.fields("RECEIVEDFROM")
            rs.fields("RECEIVEDFROMTITLE") = rsSrc.fields("RECEIVEDFROMTITLE")
            rs.fields("RELDATE") = rsSrc.fields("RELDATE")
            rs.fields("RELEASEDBY") = rsSrc.fields("RELEASEDBY")
            rs.fields("RELEASEDBYID_NO") = rsSrc.fields("RELEASEDBYID_NO")
            rs.fields("RELEASEDBYTITLE") = rsSrc.fields("RELEASEDBYTITLE")
            rs.fields("RELEASEDTO") = rsSrc.fields("RELEASEDTO")
            rs.fields("RELEASEDTOTITLE") = rsSrc.fields("RELEASEDTOTITLE")
            rs.fields("Status") = rsSrc.fields("Status")
            rs.fields("Attrib_1") = rsSrc.fields("Attrib_1")
            rs.fields("Attrib_2") = rsSrc.fields("Attrib_2")
            rs.fields("Attrib_3") = rsSrc.fields("Attrib_3")
            rs.fields("Attrib_4") = rsSrc.fields("Attrib_4")
            rs.fields("Attrib_5") = rsSrc.fields("Attrib_5")
            rs.fields("Attrib_6") = rsSrc.fields("Attrib_6")
            rs.fields("Attrib_7") = rsSrc.fields("Attrib_7")
            rs.fields("Attrib_8") = rsSrc.fields("Attrib_8")
            rs.fields("Attrib_9") = rsSrc.fields("Attrib_9")
            rs.fields("Attrib_10") = rsSrc.fields("Attrib_10")
            rs.fields("LAST_OPR") = rsSrc.fields("LAST_OPR")
            rs.fields("LAST_UPDATE") = rsSrc.fields("LAST_UPDATE")
            If Not rs.Update Then GoTo errmsg
        End If
        UpdateCollateralUsedInProfile = True
        Exit Function
errmsg:
        UpdateCollateralUsedInProfile = False
    End Function
    Private Function UpdateCollateralUsedInPolicy(ByVal rsSrc As RemoteSet, ByVal CollateralNo As String, ByVal OANO As String) As Boolean
        Dim rs As Remoteset
        rs = db.OpenRemoteSet("Select * from AccCollateral Where ANO<>" & OANO & " and CollateralNo='" & CollateralNo & "'")
        If Not rs.Eof Then
            rs.Edit()
            rs.fields("ExpDate") = rsSrc.fields("ExpDate")
            rs.fields("RecDate") = rsSrc.fields("RecDate")
            rs.fields("RECEIVEDBY") = rsSrc.fields("RECEIVEDBY")
            rs.fields("RECEIVEDBYID_NO") = rsSrc.fields("RECEIVEDBYID_NO")
            rs.fields("RECEIVEDBYTITLE") = rsSrc.fields("RECEIVEDBYTITLE")
            rs.fields("RECEIVEDFROM") = rsSrc.fields("RECEIVEDFROM")
            rs.fields("RECEIVEDFROMTITLE") = rsSrc.fields("RECEIVEDFROMTITLE")
            rs.fields("RELDATE") = rsSrc.fields("RELDATE")
            rs.fields("RELEASEDBY") = rsSrc.fields("RELEASEDBY")
            rs.fields("RELEASEDBYID_NO") = rsSrc.fields("RELEASEDBYID_NO")
            rs.fields("RELEASEDBYTITLE") = rsSrc.fields("RELEASEDBYTITLE")
            rs.fields("RELEASEDTO") = rsSrc.fields("RELEASEDTO")
            rs.fields("RELEASEDTOTITLE") = rsSrc.fields("RELEASEDTOTITLE")
            rs.fields("Status") = rsSrc.fields("Status")
            rs.fields("Attrib_1") = rsSrc.fields("Attrib_1")
            rs.fields("Attrib_2") = rsSrc.fields("Attrib_2")
            rs.fields("Attrib_3") = rsSrc.fields("Attrib_3")
            rs.fields("Attrib_4") = rsSrc.fields("Attrib_4")
            rs.fields("Attrib_5") = rsSrc.fields("Attrib_5")
            rs.fields("Attrib_6") = rsSrc.fields("Attrib_6")
            rs.fields("Attrib_7") = rsSrc.fields("Attrib_7")
            rs.fields("Attrib_8") = rsSrc.fields("Attrib_8")
            rs.fields("Attrib_9") = rsSrc.fields("Attrib_9")
            rs.fields("Attrib_10") = rsSrc.fields("Attrib_10")
            rs.fields("LAST_OPR") = rsSrc.fields("LAST_OPR")
            rs.fields("LAST_UPDATE") = rsSrc.fields("LAST_UPDATE")
            If Not rs.Update Then GoTo errmsg
        End If
        UpdateCollateralUsedInPolicy = True
        Exit Function
errmsg:
        UpdateCollateralUsedInPolicy = False
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveCollateral(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CollateralNo As String
            ANO = getInParamValue(InParam, 0)   ' Acceptance ANO
            CollateralNo = getInParamValue(InParam, 1)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select OANO, AStatus, AType from Acceptance Where ANO=" & ANO)
            Dim OANO As Long = rs.fields("OANO")
            Dim AStatus As String = rs.fields("AStatus")
            Dim AType As String = rs.fields("AType")

            Dim COLLATERALTYPE As String = GetFieldValue(fbuffer(GetFieldIndex("COLLATERALTYPE", findex)))
            Dim Nominal As Double = toval(GetFieldValue(fbuffer(GetFieldIndex("Nominal", findex))))

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from AccCollateral Where ANO=" & OANO & " and CollateralNo='" & CollateralNo & "'")
            If rs.Eof Then
                If Not PreAddCollateralOk(OANO, CollateralNo, AStatus, AType, COLLATERALTYPE, Nominal, ErrMsg) Then
                    Throw New Exception(ErrMsg)
                End If
                rs.AddNew()
            Else
                If Not PreEditCollateralOk(OANO, CollateralNo, AStatus, AType, COLLATERALTYPE, Nominal, ErrMsg) Then
                    Throw New Exception(ErrMsg)
                End If
                rs.Edit()
            End If
            EditMode = rs.editmode

            ' Check Status
            If EditMode = AdEditInProgress Then
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            If EditMode = AdEditAdd Then
                rs.fields("ANO") = OANO
                rs.fields("CollateralNo") = CollateralNo
            End If
            rs.fields("LAST_OPR") = dbUser
            rs.fields("LAST_UPDATE") = Today.Date
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            If Not UpdateCollateralUsedInPolicy(rs, CollateralNo, OANO) Then
                Throw New Exception("Update Collateral Used in Policy Failed !")
            End If

            If Not UpdateCollateralUsedInProfile(rs, CollateralNo) Then
                Throw New Exception("Update Collateral in Profile Failed !")
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = CollateralNo

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DropCollateral(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CollateralNo As String
            ANO = getInParamValue(InParam, 0)   ' Acceptance ANO
            CollateralNo = getInParamValue(InParam, 1)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If


            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select OANO, AStatus, AType from Acceptance Where ANO=" & ANO)
            Dim OANO As Long = rs.fields("OANO")
            Dim AStatus As String = rs.fields("AStatus")
            Dim AType As String = rs.fields("AType")

            rs = db.OpenRemoteSet("Select * from AccCollateral Where ANO=" & OANO & " and CollateralNo='" & CollateralNo & "'")
            If rs.Eof Then
                ErrMsg = "No Collateral Data !"
                Throw New Exception(ErrMsg)
            Else
                If Not PreDropCollateralOk(OANO, CollateralNo, AStatus, AType, rs.fields("COLLATERALTYPE"), ErrMsg) Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetListClaimCollateral(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CollateralNo As String
            ANO = getInParamValue(InParam, 0)   ' Acceptance ANO
            CollateralNo = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select OANO from Acceptance Where ANO=" & ANO)
            Dim OANO As Long = rs.fields("OANO")

            Dim SQL As String
            SQL = "SELECT C2.ClaimNo, C2.LossDate, cast(C2.Description as varchar(512)) as Description, C2.ClaimantInfo, cast(C2.Remarks as varchar(512)) as Remarks, sum(C1.AMOUNT) AS AMOUNT "
            SQL = SQL & " FROM CLSUBDET C1, CLAIM C2, CLSUB C3, ACCEPTANCE A1 "
            SQL = SQL & " WHERE A1.OANO=" & OANO & " and C1.Item='" & CollateralNo & "' and A1.ANO=C2.ANO AND C1.CNO=C2.CNO AND C3.CNO=C1.CNO AND C3.ID=C1.ID AND C3.CURRENCY=C1.CURRENCY AND C3.REFNO=C1.REFNO"
            SQL = SQL & " AND C2.CType<>'C' AND (C2.CStatus<>'C' or C2.PCStatus='A') and not exists (Select pCNO from Claim Where Claim.pCNO=C2.CNO and Claim.CStatus<>'R') "
            SQL = SQL & " AND C3.STATUS='P' "
            SQL = SQL & " GROUP BY C2.ClaimNo, C2.LossDate, cast(C2.Description as varchar(512)), C2.ClaimantInfo, cast(C2.Remarks as varchar(512))"

            rs = db.OpenRemoteSet(sql, adOpenKeyset)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetListCompliance(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)   ' Current Acceptance ANO 

            Dim rs As RemoteSet
            Dim OANO As Long
            rs = db.OpenRemoteSet("Select A.OANO from Acceptance A Where A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))

            SQL = "Select A.CID, C.DESCRIPTION, A.Status, A.PUserID, A.PUserDate, A.AUserID, A.AUserDate  from ACCCOMPSHEETH A, COMPLIANCE C Where C.CID=A.CID and A.ANO=" & OANO

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function PrepareCompSheet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CID As String
            ANO = getInParamValue(InParam, 0)   ' Current Acceptance ANO 
            CID = getInParamValue(InParam, 1)   ' Compliance ID

            Dim rs As RemoteSet
            Dim SQL As String = ""
            Dim LANO As Long

            Dim OANO As Long
            rs = db.OpenRemoteSet("Select A.OANO from Acceptance A Where A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))

            rs = db.OpenRemoteSet("Select * from AccCompSheetH Where ANO=" & OANO & " and CID='" & CID & "'")
            If rs.Eof Then
                rs = db.OpenRemoteSet("Select A.ANO, A.LANO, C.OpenFlag from Acceptance A, Cover C Where C.CNO=A.CNO and A.ANO=" & ANO)
                LANO = rs.fields("LANO")
                If rs.fields("ANO") = rs.fields("LANO") And Not rs.fields("OpenFlag") Then  ' Master Fleet
                    SQL = "Select mA.ANO, C1.CID, Max(A.PStatus) as PStatus, Sum(P.OurShare*P.Rate) as SI, MAX(C2.MinTSI) as MinTSI, MAX(C2.MaxTSI) as MaxTSI, Max(C1.PassGrade) as PassGrade"
                    SQL = SQL & " from Acceptance A, Acceptance mA, COVER C, PCALC P, COMPLIANCE C1, COMPLIANCET C2"
                    SQL = SQL & " Where mA.ANO=" & ANO & " and mA.ANO=A.LANO and c.CNO = mA.CNO "
                    SQL = SQL & " And P.ANO = a.ANO "
                    SQL = SQL & " And C2.TOC = A.TOC And ((C2.BType in ('','D') and C.ISType in ('C','I','L')) or (C2.BType in ('','I') and C.ISType in ('R','T'))) And C1.CID = C2.CID And C1.MustApplyF=1 and C1.AllowedF = 1 And C1.Effective<=A.PDate"
                    SQL = SQL & " Group By mA.ANO, C1.CID"
                    SQL = SQL & " Having ((Max(C2.MinTSI) = 0) Or (Max(C2.MinTSI) <> 0 And Max(C2.MinTSI) <= Sum(P.OurShare * P.Rate))) And ((Max(C2.MaxTSI) = 0) Or (Max(C2.MaxTSI) <> 0 And Max(C2.MaxTSI) >= Sum(P.OurShare * P.Rate)))"
                ElseIf rs.fields("ANO") = rs.fields("LANO") And rs.fields("OpenFlag") Then
                    SQL = "Select A.ANO, C1.CID, Max(A.PStatus) as PStatus, Sum(I1.SI*coalesce(I2.Share,100)/100) as SI, MAX(C2.MinTSI) as MinTSI, MAX(C2.MaxTSI) as MaxTSI, Max(C1.PassGrade) as PassGrade"
                    SQL = SQL & " from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) Left Join ICover I1 on (I1.ANO=A.ANO and I1.Category='M') Left Join ISource I2 on (I2.CNO=C.CNO), COMPLIANCE C1, COMPLIANCET C2"
                    SQL = SQL & " Where A.ANO=" & ANO & " and c.CNO = a.CNO "
                    SQL = SQL & " And C2.TOC = A.TOC And ((C2.BType in ('','D') and C.ISType in ('C','I','L')) or (C2.BType in ('','I') and C.ISType in ('R','T'))) And C1.CID = C2.CID And C1.MustApplyF=1 and C1.AllowedF = 1 And C1.Effective<=A.PDate"
                    SQL = SQL & " Group By A.ANO, C1.CID"
                    SQL = SQL & " Having ((Max(C2.MinTSI) = 0) Or (Max(C2.MinTSI) <> 0 And Max(C2.MinTSI) <= Sum(I1.SI*coalesce(I2.Share,100)/100))) And ((Max(C2.MaxTSI) = 0) Or (Max(C2.MaxTSI) <> 0 And Max(C2.MaxTSI) >= Sum(I1.SI*coalesce(I2.Share,100)/100)))"
                Else
                    SQL = "Select A.ANO, C1.CID, Max(A.PStatus) as PStatus, Sum(P.OurShare*P.Rate) as SI, MAX(C2.MinTSI) as MinTSI, MAX(C2.MaxTSI) as MaxTSI, Max(C1.PassGrade) as PassGrade"
                    SQL = SQL & " from Acceptance A, COVER C, PCALC P, COMPLIANCE C1, COMPLIANCET C2"
                    SQL = SQL & " Where A.ANO=" & ANO & " and c.CNO = a.CNO "
                    SQL = SQL & " And P.ANO = a.ANO "
                    SQL = SQL & " And C2.TOC = A.TOC And ((C2.BType in ('','D') and C.ISType in ('C','I','L')) or (C2.BType in ('','I') and C.ISType in ('R','T'))) And C1.CID = C2.CID And C1.MustApplyF=1 and C1.AllowedF = 1 And C1.Effective<=A.PDate"
                    SQL = SQL & " Group By A.ANO, C1.CID"
                    SQL = SQL & " Having ((Max(C2.MinTSI) = 0) Or (Max(C2.MinTSI) <> 0 And Max(C2.MinTSI) <= Sum(P.OurShare * P.Rate))) And ((Max(C2.MaxTSI) = 0) Or (Max(C2.MaxTSI) <> 0 And Max(C2.MaxTSI) >= Sum(P.OurShare * P.Rate)))"
                End If

                rs = db.OpenRemoteSet(SQL)
                If rs.Eof Then
                    Throw New Exception("Compliance Sheet Not Available !")
                ElseIf rs.fields("PStatus") = "I" And ANO <> LANO Then
                    Throw New Exception("Please do the Premium Calculation First !")
                End If

                Dim PassGrade As Integer = toval(rs.fields("PassGrade"))
                CID = tostr(rs.fields("CID"))
                rs = db.OpenRemoteSet("Select * from AccCompSheetH Where ANO=" & OANO & " and CID='" & CID & "'")
                If Not rs.Eof Then
                    Throw New Exception("Compliance Sheet Data Already Prepared !")
                End If
                rs.AddNew()
                rs.fields("ANO") = OANO
                rs.fields("CID") = CID
                rs.fields("Status") = "R"
                rs.fields("CDate") = Today.Date
                rs.fields("UserID") = dbUser
                rs.fields("LDate") = Today.Date
                rs.fields("LUserID") = dbUser
                rs.fields("PassGrade") = PassGrade
                If Not rs.Update Then Throw New Exception(rs.ErrDesc)
            End If


            If Not db.Execute("Delete from AccCompSheetD Where ANO=" & OANO & " and CID='" & CID & "'") Then Throw New Exception(db.ErrDesc)

            SQL = "Insert into AccCompSheetD(ANO, CID, RiskNo, RiskDesc, SID, SIDDESC, Description, Value)"
            SQL = SQL & " Select " & OANO & " as ANO, C1.CID, C1.RiskNo, C1.Description as RiskDesc, S.SID, S.Description as SIDDesc, C2.SDescription, SV.Value"
            SQL = SQL & " from ComplianceR C1, ComplianceS C2 Left Join ScoringV SV on (SV.SID=C2.SID and SV.Description=C2.SDescription), Scoring S"
            SQL = SQL & " Where C2.OptionalF=0 and C1.CID='" & CID & "' and C1.CID=C2.CID and C1.RiskNo=C2.RiskNo and S.SID=C2.SID"

            If Not db.Execute(SQL) Then Throw New Exception(db.ErrDesc)

            ReDim OutParam(0)
            OutParam(0) = CID

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetCompSheet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CID As String
            ANO = getInParamValue(InParam, 0)   ' Current Acceptance ANO 
            CID = getInParamValue(InParam, 1)   ' Compliance Sheet ID

            Dim rs As RemoteSet
            Dim OANO As Long
            Dim LANO As Long
            Dim SQL As String
            rs = db.OpenRemoteSet("Select A.OANO, A.ANO, A.LANO, C.OpenFlag from Acceptance A, Cover C Where C.CNO=A.CNO and A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))
            LANO = toval(rs.fields("LANO"))
            If rs.fields("ANO") = rs.fields("LANO") And Not rs.fields("OpenFlag") Then  ' Master Fleet
                SQL = "Select C.CoverNo, mA.ANO, mA.OANO, mA.RegNo, mA.PolicyNo, mA.CertificateNo, C.PHolder, Max(A1.Indicator) as Indicator, Max(A1.Status) as Status, Max(A1.Score) as Score, Max(A1.PassGrade) as PassGrade, Max(P.Name) as PHolderName, SUM(Pcalc.Ourshare*Pcalc.Rate) as OurShare,"
                SQL = SQL & " Max(A1.UserID) as UserID, Max(A1.CDate) as CDate,"
                SQL = SQL & " Max(A1.LUserID) as LUserID, Max(A1.LDate) as LDate,"
                SQL = SQL & " Max(A1.PUserID) as PUserID, Max(A1.PUserDate) as PUserDate,"
                SQL = SQL & " Max(A1.AUserID) as AUserID, Max(A1.AUserDate) as AUserDate,"
                SQL = SQL & " Max(A1.CUserID) as CUserID, Max(A1.CloseDate) as CloseDate,"
                SQL = SQL & " Max(A1.Comment) as Comment"
                SQL = SQL & " from acceptance A Inner join Acceptance mA on (mA.ANO=A.LANO) Inner join cover C on (C.CNO=mA.CNO) Inner join Profile P on (P.ID=C.PHolder) Left Join PCalc on (PCalc.ANO=A.ANO) Left Join AccCompSheetH A1 on (A1.ANO=mA.OANO and A1.CID='" & CID & "') Left Join Compliance C1 on (C1.CID=A1.CID)"
                SQL = SQL & " Where mA.ANO=" & ANO
                SQL = SQL & " Group By C.CoverNo, mA.ANO, mA.OANO, mA.RegNo, mA.PolicyNo, mA.CertificateNo, C.PHolder"
            ElseIf rs.fields("ANO") = rs.fields("LANO") And rs.fields("OpenFlag") Then
                SQL = "Select C.CoverNo, A.ANO, A.OANO, A.RegNo, A.PolicyNo, A.CertificateNo, C.PHolder, Max(A1.Indicator) as Indicator, Max(A1.Status) as Status, Max(A1.Score) as Score, Max(A1.PassGrade) as PassGrade, Max(P.Name) as PHolderName, SUM(I1.SI*coalesce(I2.Share,100)/100) as OurShare,"
                SQL = SQL & " Max(A1.UserID) as UserID, Max(A1.CDate) as CDate,"
                SQL = SQL & " Max(A1.LUserID) as LUserID, Max(A1.LDate) as LDate,"
                SQL = SQL & " Max(A1.PUserID) as PUserID, Max(A1.PUserDate) as PUserDate,"
                SQL = SQL & " Max(A1.AUserID) as AUserID, Max(A1.AUserDate) as AUserDate,"
                SQL = SQL & " Max(A1.CUserID) as CUserID, Max(A1.CloseDate) as CloseDate,"
                SQL = SQL & " Max(A1.Comment) as Comment"
                SQL = SQL & " from acceptance A Inner Join cover C on (C.CNO=A.CNO) Inner join Profile P on (P.ID=C.PHolder) Left Join ICover I1 on (I1.ANO=A.ANO and I1.Category='M') Left Join ISource I2 on (I2.CNO=C.CNO) Left Join AccCompSheetH A1 on (A1.ANO=A.OANO and A1.CID='" & CID & "') Left Join Compliance C1 on (C1.CID=A1.CID)"
                SQL = SQL & " Where A.ANO=" & ANO
                SQL = SQL & " Group By C.CoverNo, A.ANO, A.OANO, A.RegNo, A.PolicyNo, A.CertificateNo, C.PHolder"
            Else
                SQL = "Select C.CoverNo, A.ANO, A.OANO, A.RegNo, A.PolicyNo, A.CertificateNo, C.PHolder, Max(A1.Indicator) as Indicator, Max(A1.Status) as Status, Max(A1.Score) as Score, Max(A1.PassGrade) as PassGrade, Max(P.Name) as PHolderName, SUM(Pcalc.Ourshare*Pcalc.Rate) as OurShare,"
                SQL = SQL & " Max(A1.UserID) as UserID, Max(A1.CDate) as CDate,"
                SQL = SQL & " Max(A1.LUserID) as LUserID, Max(A1.LDate) as LDate,"
                SQL = SQL & " Max(A1.PUserID) as PUserID, Max(A1.PUserDate) as PUserDate,"
                SQL = SQL & " Max(A1.AUserID) as AUserID, Max(A1.AUserDate) as AUserDate,"
                SQL = SQL & " Max(A1.CUserID) as CUserID, Max(A1.CloseDate) as CloseDate,"
                SQL = SQL & " Max(A1.Comment) as Comment"
                SQL = SQL & " from acceptance A Inner Join cover C on (C.CNO=A.CNO) Inner join Profile P on (P.ID=C.PHolder) Left Join PCalc on (PCalc.ANO=A.ANO) Left Join AccCompSheetH A1 on (A1.ANO=A.OANO and A1.CID='" & CID & "') Left Join Compliance C1 on (C1.CID=A1.CID)"
                SQL = SQL & " Where A.ANO=" & ANO
                SQL = SQL & " Group By C.CoverNo, A.ANO, A.OANO, A.RegNo, A.PolicyNo, A.CertificateNo, C.PHolder"
            End If

            rs = db.OpenRemoteSet(SQL)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            SQL = "Select A.CID, A.RiskNo, A.RiskDesc, A.SID, A.SIDDesc, A.Remarks, A.Description, A.Value, C.PassGrade, C.OptionalF"
            SQL = SQL & " from AccCompSheetD A, ComplianceS C"
            SQL = SQL & " Where A.ANO=" & OANO & " and A.CID='" & CID & "'"
            SQL = SQL & " and C.CID=A.CID and C.RiskNo=A.RiskNo and C.SID=A.SID"
            SQL = SQL & " Order By A.CID, A.RiskNo, A.SID"

            rs = db.OpenRemoteSet(SQL)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function UpdateCompSheet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CID As String
            Dim Comment As String
            ANO = getInParamValue(InParam, 0)   ' Current Acceptance ANO 
            CID = getInParamValue(InParam, 1)   ' Compliance Sheet ID
            Comment = getInParamValue(InParam, 2)   ' Compliance Comment

            Dim rs As RemoteSet
            Dim OANO As Long
            rs = db.OpenRemoteSet("Select A.OANO from Acceptance A Where A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))

            rs = db.OpenRemoteSet("Select * from AccCompSheetH Where ANO=" & OANO & " and CID='" & CID & "'")
            If rs.Eof Then
                Throw New Exception("No Compliance Sheet Data !")
            ElseIf rs.fields("Status") <> "R" Then
                Throw New Exception("Compliance Sheet Data Already Processed !")
            End If

            Dim PassGrade As Integer
            Dim IPassGradeF As Boolean
            rs = db.OpenRemoteSet("Select IPassGradeF, PassGrade from Compliance C Where C.CID='" & CID & "'")
            IPassGradeF = toval(rs.fields("IPassGradeF"))
            If IPassGradeF Then
                PassGrade = toval(rs.fields("PassGrade"))
            Else
                PassGrade = toval(rs.fields("PassGrade"))
            End If

            Const MaxCols = 10
            Dim Bufferdata(Dataset.Tables(0).Rows.Count - 1, MaxCols - 1) As Object

            Dim i As Integer
            Dim Score As Integer = 0
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Bufferdata(i, 0) = tostr(Dataset.Tables(0).Rows(i)("CID"))
                Bufferdata(i, 1) = tostr(Dataset.Tables(0).Rows(i)("RiskNo"))
                Bufferdata(i, 2) = tostr(Dataset.Tables(0).Rows(i)("RiskDesc"))
                Bufferdata(i, 3) = tostr(Dataset.Tables(0).Rows(i)("SID"))
                Bufferdata(i, 4) = tostr(Dataset.Tables(0).Rows(i)("SIDDesc"))
                Bufferdata(i, 5) = tostr(Dataset.Tables(0).Rows(i)("Remarks"))
                Bufferdata(i, 6) = tostr(Dataset.Tables(0).Rows(i)("Description"))
                Bufferdata(i, 7) = tostr(Dataset.Tables(0).Rows(i)("Value"))
                Bufferdata(i, 8) = tostr(Dataset.Tables(0).Rows(i)("PassGrade"))
                Bufferdata(i, 9) = toval(Dataset.Tables(0).Rows(i)("OptionalF"))
                If tostr(Bufferdata(i, 0)) <> "" Then
                    Score = Score + toval(Bufferdata(i, 7))
                End If
                If toval(Bufferdata(i, 9)) Then ' if Optional add passGrade
                    PassGrade = PassGrade + toval(Bufferdata(i, 8))
                End If
            Next

            InParam(0) = OANO
            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Save_Compliance_Sheet", InParam, TrimBuffer(Bufferdata)) Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            ReDim OutParam(1)
            OutParam(0) = Score
            OutParam(1) = PassGrade

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DropCompSheet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CID As String
            ANO = getInParamValue(InParam, 0)   ' Current Acceptance ANO 
            CID = getInParamValue(InParam, 1)   ' Compliance Sheet ID

            Dim rs As RemoteSet
            Dim OANO As Long
            rs = db.OpenRemoteSet("Select A.OANO from Acceptance A Where A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))

            rs = db.OpenRemoteSet("Select * from AccCompSheetH Where ANO=" & OANO & " and CID='" & CID & "'")
            If rs.Eof Then
                Throw New Exception("No Compliance Sheet Data !")
            ElseIf rs.fields("Status") <> "R" Then
                Throw New Exception("Compliance Sheet Data Already Processed !")
            End If
            If Not rs.Delete Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function ReviewCompSheet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CID As String
            ANO = getInParamValue(InParam, 0)   ' Current Acceptance ANO 
            CID = getInParamValue(InParam, 1)   ' Compliance Sheet ID

            Dim rs As RemoteSet
            Dim OANO As Long
            rs = db.OpenRemoteSet("Select A.OANO from Acceptance A Where A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))

            rs = db.OpenRemoteSet("Select * from AccCompSheetH Where ANO=" & OANO & " and CID='" & CID & "'")
            If rs.Eof Then
                Throw New Exception("No Compliance Sheet Data !")
            ElseIf rs.fields("Status") <> "R" Then
                Throw New Exception("Compliance Sheet Data Already Processed !")
            End If
            rs.Edit()
            rs.fields("Status") = "P"
            rs.fields("PUserDate") = Today.Date
            rs.fields("PUserID") = dbUser
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ApproveCompSheet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CID As String
            ANO = getInParamValue(InParam, 0)   ' Current Acceptance ANO 
            CID = getInParamValue(InParam, 1)   ' Compliance Sheet ID

            Dim rs As RemoteSet
            Dim OANO As Long
            rs = db.OpenRemoteSet("Select A.OANO from Acceptance A Where A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))

            rs = db.OpenRemoteSet("Select * from AccCompSheetH Where Status='P' and ANO=" & OANO & " and CID='" & CID & "'")
            If rs.Eof Then
                Throw New Exception("No Compliance Sheet Data !")
            Else
                If Not Check_Compliance_User(rs.fields("CID"), dbUser, rs.fields("INDICATOR")) Then
                    Throw New Exception("User Not Authorized for Approval ! (Please Check Compliance User) ")
                End If

                If Not (toval(rs.fields("Score")) >= toval(rs.fields("PassGrade"))) Then
                    Throw New Exception("Compliance Risk Scoring Below Minimum Passing Grade !")
                End If
                rs.Edit()
            End If
            rs.fields("Status") = "A"
            rs.fields("AUserDate") = Today.Date
            rs.fields("AUserID") = dbUser
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DeclineCompSheet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CID As String
            ANO = getInParamValue(InParam, 0)   ' Current Acceptance ANO 
            CID = getInParamValue(InParam, 1)   ' Compliance Sheet ID

            Dim rs As RemoteSet
            Dim OANO As Long
            rs = db.OpenRemoteSet("Select A.OANO from Acceptance A Where A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))

            rs = db.OpenRemoteSet("Select * from AccCompSheetH Where Status='P' and ANO=" & OANO & " and CID='" & CID & "'")
            If rs.Eof Then
                Throw New Exception("No Compliance Sheet Data !")
            Else
                If Not Check_Compliance_User(rs.fields("CID"), dbUser, rs.fields("INDICATOR")) Then
                    Throw New Exception("User Not Authorized for Declining ! (Please Check Compliance User) ")
                End If
                rs.Edit()
            End If
            rs.fields("Status") = "C"
            rs.fields("CloseDate") = Today.Date
            rs.fields("CUserID") = dbUser
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ReDoCompSheet(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim CID As String
            ANO = getInParamValue(InParam, 0)   ' Current Acceptance ANO 
            CID = getInParamValue(InParam, 1)   ' Compliance Sheet ID

            Dim rs As RemoteSet
            Dim OANO As Long
            rs = db.OpenRemoteSet("Select A.OANO from Acceptance A Where A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))

            rs = db.OpenRemoteSet("Select * from AccCompSheetH Where Status in ('P','A','C') and ANO=" & OANO & " and CID='" & CID & "'")
            If rs.Eof Then
                Throw New Exception("No Compliance Sheet Data !")
            Else
                If Not Check_Compliance_User(rs.fields("CID"), dbUser, rs.fields("INDICATOR")) Then
                    Throw New Exception("User Not Authorized for Redo ! (Please Check Compliance User) ")
                End If
                rs.Edit()
            End If
            rs.fields("Status") = "R"
            rs.fields("AUserID") = ""
            rs.fields("PUserID") = ""
            rs.fields("CUserID") = ""
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function PrepareEmpContribution(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)   ' Original Acceptance ANO (OANO)

            Dim rs As RemoteSet

            rs = db.OpenRemoteSet("Select * from ACCEMPCONTH Where ANO=" & ANO)
            If Not rs.Eof Then
                Throw New Exception("Employee Contribution Already Prepared !")
            End If
            rs.AddNew()
            rs.fields("ANO") = ANO
            rs.fields("Status") = "R"
            rs.fields("TotalPCT") = 100
            rs.fields("CDate") = Today.Date
            rs.fields("UserID") = dbUser
            rs.fields("LDate") = Today.Date
            rs.fields("LUserID") = dbUser
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            If Not db.Execute("Delete from ACCEMPCONTD Where ANO=" & ANO) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetEmpContribution(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)   ' Acceptance ANO

            Dim rs As RemoteSet
            Dim OANO As Long
            Dim LANO As Long
            Dim SQL As String
            rs = db.OpenRemoteSet("Select A.OANO, A.ANO, A.LANO, C.OpenFlag from Acceptance A, Cover C Where C.CNO=A.CNO and A.ANO=" & ANO)
            OANO = toval(rs.fields("OANO"))
            LANO = toval(rs.fields("LANO"))
            If rs.fields("ANO") = rs.fields("LANO") And Not rs.fields("OpenFlag") Then  ' Master Fleet
                SQL = "Select C.CoverNo, mA.ANO, mA.OANO, mA.RegNo, mA.PolicyNo, mA.CertificateNo, C.PHolder, Max(A1.Remarks) as Remarks, Max(A1.Status) as Status, Max(A1.TotalPCT) as TotalPCT, Max(P.Name) as PHolderName, SUM(Pcalc.Ourshare*Pcalc.Rate) as OurShare,"
                SQL = SQL & " Max(A1.UserID) as UserID, Max(A1.CDate) as CDate,"
                SQL = SQL & " Max(A1.LUserID) as LUserID, Max(A1.LDate) as LDate,"
                SQL = SQL & " Max(A1.PUserID) as PUserID, Max(A1.PUserDate) as PUserDate,"
                SQL = SQL & " Max(A1.AUserID) as AUserID, Max(A1.AUserDate) as AUserDate,"
                SQL = SQL & " Max(A1.CUserID) as CUserID, Max(A1.CloseDate) as CloseDate"
                SQL = SQL & " from acceptance A Inner join Acceptance mA on (mA.ANO=A.LANO) Inner join cover C on (C.CNO=mA.CNO) Inner join Profile P on (P.ID=C.PHolder) Left Join PCalc on (PCalc.ANO=A.ANO) Left Join ACCEMPCONTH A1 on (A1.ANO=mA.OANO) "
                SQL = SQL & " Where mA.ANO=" & ANO
                SQL = SQL & " Group By C.CoverNo, mA.ANO, mA.OANO, mA.RegNo, mA.PolicyNo, mA.CertificateNo, C.PHolder"
            ElseIf rs.fields("ANO") = rs.fields("LANO") And rs.fields("OpenFlag") Then
                SQL = "Select C.CoverNo, A.ANO, A.OANO, A.RegNo, A.PolicyNo, A.CertificateNo, C.PHolder, Max(A1.Remarks) as Remarks, Max(A1.Status) as Status, Max(A1.TotalPCT) as TotalPCT, Max(P.Name) as PHolderName, SUM(I1.SI*coalesce(I2.Share,100)/100) as OurShare,"
                SQL = SQL & " Max(A1.UserID) as UserID, Max(A1.CDate) as CDate,"
                SQL = SQL & " Max(A1.LUserID) as LUserID, Max(A1.LDate) as LDate,"
                SQL = SQL & " Max(A1.PUserID) as PUserID, Max(A1.PUserDate) as PUserDate,"
                SQL = SQL & " Max(A1.AUserID) as AUserID, Max(A1.AUserDate) as AUserDate,"
                SQL = SQL & " Max(A1.CUserID) as CUserID, Max(A1.CloseDate) as CloseDate"
                SQL = SQL & " from acceptance A Inner Join cover C on (C.CNO=A.CNO) Inner join Profile P on (P.ID=C.PHolder) Left Join ICover I1 on (I1.ANO=A.ANO and I1.Category='M') Left Join ISource I2 on (I2.CNO=C.CNO) Left Join ACCEMPCONTH A1 on (A1.ANO=A.OANO) "
                SQL = SQL & " Where A.ANO=" & ANO
                SQL = SQL & " Group By C.CoverNo, A.ANO, A.OANO, A.RegNo, A.PolicyNo, A.CertificateNo, C.PHolder"
            Else
                SQL = "Select C.CoverNo, A.ANO, A.OANO, A.RegNo, A.PolicyNo, A.CertificateNo, C.PHolder, Max(A1.Remarks) as Remarks, Max(A1.Status) as Status, Max(A1.TotalPCT) as TotalPCT, Max(P.Name) as PHolderName, SUM(Pcalc.Ourshare*Pcalc.Rate) as OurShare,"
                SQL = SQL & " Max(A1.UserID) as UserID, Max(A1.CDate) as CDate,"
                SQL = SQL & " Max(A1.LUserID) as LUserID, Max(A1.LDate) as LDate,"
                SQL = SQL & " Max(A1.PUserID) as PUserID, Max(A1.PUserDate) as PUserDate,"
                SQL = SQL & " Max(A1.AUserID) as AUserID, Max(A1.AUserDate) as AUserDate,"
                SQL = SQL & " Max(A1.CUserID) as CUserID, Max(A1.CloseDate) as CloseDate"
                SQL = SQL & " from acceptance A Inner Join cover C on (C.CNO=A.CNO) Inner join Profile P on (P.ID=C.PHolder) Left Join PCalc on (PCalc.ANO=A.ANO) Left Join ACCEMPCONTH A1 on (A1.ANO=A.OANO)"
                SQL = SQL & " Where A.ANO=" & ANO
                SQL = SQL & " Group By C.CoverNo, A.ANO, A.OANO, A.RegNo, A.PolicyNo, A.CertificateNo, C.PHolder"
            End If

            rs = db.OpenRemoteSet(SQL)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            SQL = "Select * from ACCEMPCONTD Where ANO=" & OANO & " Order By PCT"

            rs = db.OpenRemoteSet(SQL)
            Dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function UpdateEmpContribution(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim Remarks As String
            ANO = getInParamValue(InParam, 0)   ' Original Acceptance ANO (OANO)
            Remarks = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from ACCEMPCONTH Where ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("No Employee Contribution Data !")
            ElseIf rs.fields("Status") <> "R" Then
                Throw New Exception("Employee Contribution Data Already Processed !")
            End If

            Const MaxCols = 4
            Dim Bufferdata(Dataset.Tables(0).Rows.Count - 1, MaxCols - 1) As Object

            Dim i As Integer
            Dim TotalPCT As Integer = 0
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Bufferdata(i, 0) = tostr(Dataset.Tables(0).Rows(i)("ID"))
                Bufferdata(i, 1) = tostr(Dataset.Tables(0).Rows(i)("Remarks"))
                Bufferdata(i, 2) = toval(Dataset.Tables(0).Rows(i)("PCT"))
                Bufferdata(i, 3) = tostr(Dataset.Tables(0).Rows(i)("ECUTYPE"))
                If tostr(Bufferdata(i, 0)) <> "" Then
                    TotalPCT = TotalPCT + toval(Bufferdata(i, 2))
                End If
            Next

            If Not Me.Direct_Call_Process_Execute("ProcessExt.NetV3.Proclib", "Save_Employee_Contribution", InParam, TrimBuffer(Bufferdata)) Then
                Throw New Exception(Me.ErrDesc)
            End If

            ' Set Output Parameter
            ReDim OutParam(0)
            OutParam(0) = TotalPCT

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DropEmpContribution(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)   ' Original Acceptance ANO (OANO)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from ACCEMPCONTH Where ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("No Employee Contribution Data !")
            ElseIf rs.fields("Status") <> "R" Then
                Throw New Exception("Employee Contribution Data Already Processed !")
            End If
            If Not rs.Delete Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ReviewEmpContribution(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)   ' Original Acceptance ANO (OANO)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from ACCEMPCONTH Where ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("No Employee Contribution Data !")
            ElseIf rs.fields("Status") <> "R" Then
                Throw New Exception("Employee Contribution Data Already Processed !")
            End If
            rs.Edit()
            rs.fields("Status") = "P"
            rs.fields("PUserDate") = Today.Date
            rs.fields("PUserID") = dbUser
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ApproveEmpContribution(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)   ' Original Acceptance ANO (OANO)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from ACCEMPCONTH Where Status='P' and ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("No Employee Contribution Data !")
            Else
                If rs.fields("TotalPCT") <> 100 Then
                    Throw New Exception("Total Employee Contribution must be 100 !")
                End If
                rs.Edit()
            End If
            rs.fields("Status") = "A"
            rs.fields("AUserDate") = Today.Date
            rs.fields("AUserID") = dbUser
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DeclineEmpContribution(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)   ' Original Acceptance ANO (OANO)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from ACCEMPCONTH Where Status='P' and ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("No Employee Contribution Data !")
            Else
                rs.Edit()
            End If
            rs.fields("Status") = "C"
            rs.fields("CloseDate") = Today.Date
            rs.fields("CUserID") = dbUser
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ReDoEmpContribution(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)   ' Original Acceptance ANO (OANO)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from ACCEMPCONTH Where Status in ('P','A','C') and ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("No Employee Contribution Data !")
            Else
                rs.Edit()
            End If
            rs.fields("Status") = "R"
            rs.fields("AUserID") = ""
            rs.fields("PUserID") = ""
            rs.fields("CUserID") = ""
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPolicyRefNo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)
            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select ANO, RefNo, SecurityNo, COCNO From ACCEPTANCE Where ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SavePolicyRefNo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select ANO, RefNo, SecurityNo, COCNO from ACCEPTANCE Where ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("Policy Data Not Found !")
            End If

            ' Check Duplicate COC NO
            Dim PCOCNO As String
            Dim COCNO As String
            PCOCNO = rs.fields("COCNO")
            COCNO = GetFieldValue(fbuffer(GetFieldIndex("COCNO", findex)))
            If PCOCNO <> COCNO And COCNO <> "" Then
                Dim rsA As RemoteSet
                rsA = db.OpenRemoteSet("Select top 1 A1.RegNo, A1.PolicyNo, A1.CertificateNo from Acceptance A1, Acceptance A2 Where A1.OANO<>A2.OANO and A2.ANO=" & ANO & " and A1.COCNO='" & COCNO & "'")
                If Not rsA.Eof Then
                    Throw New Exception("COC NO already Used By C/I No : " & rsA.fields("RegNo") & ", Policy No : " & rsA.fields("PolicyNo") & ", Certificate No. : " & rsA.fields("CertificateNo") & " !")
                End If
            End If

            rs.Edit()
            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function


    <WebMethod(EnableSession:=False)> _
    Public Function GetPolicyLocalVAT(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)
            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.ANO, A.AStatus, A.AType, A.CNO, A.TOC, A.Branch, A.AID, A.BVATPCT, A.TVATPCT, A.OverwriteBVATF, A.OverwriteTVATF, A.LUserID, A.LDate, A.LTime from Acceptance A Where A.ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SavePolicyLocalVAT(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.ANO, A.AStatus, A.AType, A.CNO, A.TOC, A.Branch, A.AID, A.BVATPCT, A.TVATPCT, A.OverwriteBVATF, A.OverwriteTVATF, A.LUserID, A.LDate, A.LTime from Acceptance A Where A.ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("Policy Data Not Found !")
            ElseIf Not (rs.fields("AStatus") = "W" And (rs.fields("AType") = "N" Or rs.fields("AType") = "E")) Then
                Throw New Exception("Please Check Policy Data Status !")
            End If

            rs.Edit()
            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("LUserID") = dbUser
            rs.fields("LDate") = Today.Date
            rs.fields("LTime") = Time()
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPolicyRelaxation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)
            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.ExclusiveF, A.SDate, A.STime, A.EDate, A.ETime, A.ExtRSDate, A.ExtREDate, A.ExtRPeriod, C.PPeriod, C.Prorata from Acceptance A, Cover C Where C.CNO=A.CNO and A.ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SavePolicyRelaxation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from ACCEPTANCE Where ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("Policy Data Not Found !")
            ElseIf rs.fields("AStatus") <> "W" Then
                Throw New Exception("Policy Already Processed !")
            End If

            Dim PTerm As Integer
            Dim ExtRPeriod As Integer
            PTerm = 12

            rs.Edit()
            rs.fields("ExtRSdate") = todate(GetFieldValue(fbuffer(GetFieldIndex("ExtRSDate", findex))))
            rs.fields("ExtREdate") = todate(GetFieldValue(fbuffer(GetFieldIndex("ExtREdate", findex))))

            If Not toval(rs.fields("ExclusiveF")) And rs.fields("STime") < rs.fields("ETime") Then
                ExtRPeriod = Premium_Calc_Prorata(rs.fields("ExtRSdate"), DateAdd("d", 1, rs.fields("ExtREdate")), toval(GetFieldValue(fbuffer(GetFieldIndex("Prorata", findex)))), PTerm, toval(GetFieldValue(fbuffer(GetFieldIndex("PPeriod", findex)))), "") * PTerm
            Else
                ExtRPeriod = Premium_Calc_Prorata(rs.fields("ExtRSdate"), rs.fields("ExtREdate"), toval(GetFieldValue(fbuffer(GetFieldIndex("Prorata", findex)))), PTerm, toval(GetFieldValue(fbuffer(GetFieldIndex("PPeriod", findex)))), "") * PTerm
            End If

            rs.fields("ExtRPeriod") = ExtRPeriod
            rs.fields("PStatus") = "I"
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function RemovePolicyRelaxation(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select AStatus from ACCEPTANCE Where ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("Policy Data Not Found !")
            ElseIf rs.fields("AStatus") <> "W" Then
                Throw New Exception("Policy Already Processed !")
            End If

            If Not db.Execute("Update Acceptance Set ExtRSDate=Null, ExtREDate=Null,ExtRPeriod=0, pStatus='I' from Acceptance Where AStatus='W' and ANO=" & ANO) Then Throw New Exception(db.ErrDesc)

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetExtAInfo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim rs As RemoteSet
            rs = db.OpenQueryDef("QAINFO_EXT", adOpenKeyset, ANO)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveExtAInfo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus from Acceptance A Where ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            Else
                If rs.fields("AStatus") <> "W" Then
                    Throw New Exception("Policy Already Processed !")
                End If
            End If

            rs = db.OpenRemoteSet("Select * from AINFO_Ext Where ANO=" & ANO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("ANO") = ANO
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter   
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
       Public Function GetPolicyNotification(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim AType As String
            Dim CCY As String
            Dim DocType As String

            ANO = getInParamValue(InParam, 0)    ' Claim No
            ID = getInParamValue(InParam, 1)
            AType = getInParamValue(InParam, 2)
            CCY = getInParamValue(InParam, 3)
            DocType = getInParamValue(InParam, 4)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from PNOTIFY where ANO=" & ANO & " and ID='" & ID & "' and Type='" & AType & "' and Currency_OC='" & CCY & "' and Document='" & DocType & "'")
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SavePolicyNotification(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim ID As String
            Dim AType As String
            Dim CCY As String
            Dim DocType As String

            ANO = getInParamValue(InParam, 0)    ' Claim No
            ID = getInParamValue(InParam, 1)
            AType = getInParamValue(InParam, 2)
            CCY = getInParamValue(InParam, 3)
            DocType = getInParamValue(InParam, 4)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from PNotify where ANO=" & ANO & " and ID='" & ID & "' and [Type]='" & AType & "' and Currency_OC='" & CCY & "' and Document='" & DocType & "'")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If

            EditMode = rs.editmode

            If EditMode = AdEditAdd Then
                rs.fields("ANO") = ANO
                rs.fields("ID") = ID
                rs.fields("Type") = AType
                rs.fields("Currency_OC") = CCY
                rs.fields("Document") = DocType
                rs.fields("Date") = GetFieldValue(fbuffer(GetFieldIndex("Date", findex)))
                rs.fields("Time") = GetFieldValue(fbuffer(GetFieldIndex("Time", findex)))
                rs.fields("Remarks") = GetFieldValue(fbuffer(GetFieldIndex("Remarks", findex)))
                rs.fields("UserID") = dbUser
            Else
                If rs.fields("Date") <> GetFieldValue(fbuffer(GetFieldIndex("Date", findex))) Then
                    Save_Log_Object("PNotify", "ANO", ANO, "ID", ID, "TYPE", AType, "Document", DocType, "Date", rs.fields("Date"), GetFieldValue(fbuffer(GetFieldIndex("Date", findex))), "Update")
                    rs.fields("Date") = GetFieldValue(fbuffer(GetFieldIndex("Date", findex)))
                End If

                If rs.fields("Time") <> GetFieldValue(fbuffer(GetFieldIndex("Time", findex))) Then
                    Save_Log_Object("PNotify", "ANO", ANO, "ID", ID, "TYPE", AType, "Document", DocType, "Time", rs.fields("Time"), GetFieldValue(fbuffer(GetFieldIndex("Time", findex))), "Update")
                    rs.fields("Time") = GetFieldValue(fbuffer(GetFieldIndex("Time", findex)))
                End If

                If rs.fields("Remarks") <> GetFieldValue(fbuffer(GetFieldIndex("Remarks", findex))) Then
                    Save_Log_Object("PNotify", "ANO", ANO, "ID", ID, "TYPE", AType, "Document", DocType, "Remarks", rs.fields("Remarks"), GetFieldValue(fbuffer(GetFieldIndex("Remarks", findex))), "Update")
                    rs.fields("Remarks") = GetFieldValue(fbuffer(GetFieldIndex("Remarks", findex)))
                End If

                If rs.fields("UserID") <> dbUser Then
                    Save_Log_Object("PNotify", "ANO", ANO, "ID", ID, "TYPE", AType, "Document", DocType, "UserID", rs.fields("UserID"), dbUser, "Update")
                    rs.fields("UserID") = dbUser
                End If
            End If

            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SearchPolicyNotes(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long = getInParamValue(InParam, 0)   ' Acceptance No Original ANO (ONO)
            Dim Category As String = getInParamValue(InParam, 1)  ' Policy Notes Category (QUOTATION/POLICY/OTHERS) - Refer to PNCategory
            Dim PDate As String = getInParamValue(InParam, 2)  ' Input as at, Format "MM/DD/YYYY"
            Dim Matters As String = getInParamValue(InParam, 3)  ' Matters 
            Dim sType As String = getInParamValue(InParam, 4)  ' ALL / USER / SYSTEM
            Dim Status As String = getInParamValue(InParam, 5)  ' Refer to PNStatus
            Dim IncludeRenewedPolicyF As Boolean = toval(getInParamValue(InParam, 6))  ' IncludeRenewedPolicyF (True : Include Renewed Policy, False : Not)

            Dim SQL As String = ""

            SQL = "Select P.* "
            SQL = SQL & " from PNotes P"
            SQL = SQL & " Where Exists ("
            SQL = SQL & " Select Top 1 A.OANO from Acceptance A Where A.OANO=P.ANO and A.OANO=" & ANO
            If IncludeRenewedPolicyF Then
                SQL = SQL & " Union All"
                SQL = SQL & " Select Top 1 A.OANO from Acceptance A, Acceptance RA Where A.OANO=P.ANO and RA.RANO=A.ANO and RA.OANO=" & ANO
            End If
            SQL = SQL & " )"
            If Category <> "" Then
                SQL = SQL & " and P.Category='" & Category & "'"
            End If

            If Matters <> "" Then
                SQL = SQL & " and P.Matters='" & Matters & "'"
            End If

            If PDate <> "" Then
                SQL = SQL & " and P.Date<='" & PDate & "'"
            End If

            Select Case sType.ToUpper
                Case "ALL"
                Case "USER"
                    SQL = SQL & " and P.SystemF=0"
                Case "SYSTEM"
                    SQL = SQL & " and P.SystemF=1"
            End Select

            If Status <> "" Then
                SQL = SQL & " and P.Status='" & Status & "'"
            End If

            SQL = SQL & " Order by PID Desc"


            Dim rs As RemoteSet
            Dim C_MaxRecords As Long
            C_MaxRecords = toval(GetAppSetting("CareMaxRecords"))
            rs = db.OpenRemoteSet(SQL, , , , C_MaxRecords)
            Dataset = rs.GetDataSet

            If Dataset Is Nothing Then Throw New Exception(ErrMsg)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SavePolicyNotes(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim PID As Long
            Dim FLDID As String
            Dim VALUEID As String

            ANO = getInParamValue(InParam, 0)   ' Acceptance No Original ANO (ONO)
            PID = getInParamValue(InParam, 1)    ' Policy Notes PID
            FLDID = getInParamValue(InParam, 2)    ' from Object Register
            VALUEID = getInParamValue(InParam, 3)    ' from Object Register

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from PNOTES Where PID=" & PID)
            If rs.Eof Then
                rs.AddNew()
            Else
                Dim rsS As RemoteSet
                rsS = db.OpenRemoteSet("Select Top 1 PID, Date, UserID from PNOTES Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID='" & VALUEID & "' Order by PID Desc")
                If Not rsS.Eof Then
                    If rsS.fields("PID") > PID Then
                        Throw New Exception("Can't Update The Notes Which is not the Last One !")
                    ElseIf rsS.fields("UserID") <> dbUser Then
                        Throw New Exception("Can't Update The Notes Which is not belong to the User !")
                    ElseIf DateDiff(DateInterval.Day, rsS.fields("Date"), Today.Date) > 0 Then
                        Throw New Exception("Can't Update The Notes Which already Lapse !")
                    End If
                End If
                rs.Edit()
                If toval(rs.fields("SystemF")) Then
                    Throw New Exception("Can't Update System Generated !")
                End If
            End If

            EditMode = rs.editmode

            If EditMode = AdEditAdd Then
                rs.fields("ANO") = ANO
            End If
            rs.fields("SystemF") = False
            rs.fields("Date") = Today.Date
            rs.fields("Time") = Time()
            rs.fields("UserID") = dbUser
            rs.fields("Category") = GetFieldValue(fbuffer(GetFieldIndex("Category", findex)))
            rs.fields("Matters") = GetFieldValue(fbuffer(GetFieldIndex("Matters", findex)))
            rs.fields("Notes") = GetFieldValue(fbuffer(GetFieldIndex("Notes", findex)))
            rs.fields("Status") = GetFieldValue(fbuffer(GetFieldIndex("Status", findex)))
            rs.fields("FLDID") = FLDID
            rs.fields("VALUEID") = VALUEID
            For i = 1 To 10
                rs.fields("LBLNotes_" & i) = GetFieldValue(fbuffer(GetFieldIndex("LBLNotes_" & i, findex)))
                rs.fields("Notes_" & i) = GetFieldValue(fbuffer(GetFieldIndex("Notes_" & i, findex)))
            Next
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            If EditMode = AdEditAdd Then
                rs = db.OpenRemoteSet("Select Top 1 PID from PNOTES Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID='" & VALUEID & "' Order by PID Desc")
            End If
            ReDim OutParam(0)
            OutParam(0) = toval(rs.fields("PID"))

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPolicyNotes(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim PID As Long
            Dim FLDID As String
            Dim VALUEID As String

            ANO = getInParamValue(InParam, 0)   ' Acceptance No Original ANO (ONO)
            PID = getInParamValue(InParam, 1)    ' Policy Notes PID
            FLDID = getInParamValue(InParam, 2)    ' from Object Register
            VALUEID = getInParamValue(InParam, 3)    ' from Object Register

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from PNOTES where PID=" & PID)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DropPolicyNotes(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim PID As Long
            Dim FLDID As String
            Dim VALUEID As String

            ANO = getInParamValue(InParam, 0)   ' Acceptance No Original ANO (ONO)
            PID = getInParamValue(InParam, 1)    ' Policy Notes PID
            FLDID = getInParamValue(InParam, 2)    ' from Object Register
            VALUEID = getInParamValue(InParam, 3)    ' from Object Register

            If Not db.GetLockedObject("POLICY" & ANO) Then Throw New Exception(db.ErrDesc)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from PNOTES where PID=" & PID)
            If rs.Eof Then
                ErrMsg = "No Policy Notes Data !"
                Throw New Exception(ErrMsg)
            Else
                Dim rsS As RemoteSet
                rsS = db.OpenRemoteSet("Select Top 1 PID, Date, UserID from PNOTES Where ANO=" & ANO & " and FLDID='" & FLDID & "' and VALUEID='" & VALUEID & "' Order by PID Desc")
                If Not rsS.Eof Then
                    If rsS.fields("PID") > PID Then
                        Throw New Exception("Can't Delete The Notes Which is not the Last One !")
                    ElseIf rsS.fields("UserID") <> dbUser Then
                        Throw New Exception("Can't Delete The Notes Which is not belong to the User !")
                    ElseIf DateDiff(DateInterval.Day, rsS.fields("Date"), Today.Date) > 0 Then
                        Throw New Exception("Can't Delete The Notes Which already Lapse !")
                    End If
                End If

                If toval(rs.fields("SystemF")) Then
                    Throw New Exception("Can't Delete System Generated !")
                End If
            End If
            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

            If Not db.ReleaseLockedObject("POLICY" & ANO) Then Throw New Exception(db.ErrDesc)
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function UpdateBankAccount(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim AdmNo As Long
            Dim ProAccF As Boolean
            Dim AccountNo As String
            Dim Bank As String
            Dim BankBranch As String
            Dim BankAccount As String
            Dim BankName As String
            Dim RefType As String
            Dim BankChargesF As Boolean


            AdmNo = getInParamValue(InParam, 0)
            ProAccF = getInParamValue(InParam, 1)    ' Update Profile Account Flag (False : No, True : Yes)
            AccountNo = getInParamValue(InParam, 2)
            Bank = getInParamValue(InParam, 3)
            BankBranch = getInParamValue(InParam, 4)
            BankAccount = getInParamValue(InParam, 5)
            BankName = getInParamValue(InParam, 6)
            RefType = getInParamValue(InParam, 7)
            BankChargesF = getInParamValue(InParam, 8)

            Dim datamaster As RemoteSet
            datamaster = db.OpenRemoteSet("Select * from Admlink Where AdmNo=" & AdmNo)
            If datamaster.Eof Then
                ErrMsg = "No Transaction Found !"
                Throw New Exception(ErrMsg)
            End If
            datamaster.Edit()
            datamaster.fields("AccountNo") = AccountNo
            datamaster.fields("Bank") = Bank
            datamaster.fields("BankBranch") = BankBranch
            datamaster.fields("BankAccount") = BankAccount
            datamaster.fields("BankName") = BankName
            datamaster.fields("RefType") = RefType
            datamaster.fields("BankChargesF") = BankChargesF
            If Not datamaster.Update Then
                Throw New Exception(db.ErrDesc)
            End If

            If Not db.Execute("Update nVoucher Set AccountNo=A.AccountNo, Bank=A.Bank, BankBranch=A.BankBranch, BankAccount=A.BankAccount, BankName=A.BankName, RefType=A.RefType from nVoucher, Admlink A Where nVoucher.Voucher=A.Voucher and A.AdmNo=" & AdmNo) Then
                Throw New Exception(db.ErrDesc)
            End If
            If ProAccF Then
                Dim rs As RemoteSet

                rs = db.OpenRemoteSet("Select ProAcc.ProAcc, ProAcc.Bank, ProAcc.BankBranch, ProAcc.AccountNo, ProAcc.BankAccount, ProAcc.BankName from ProAcc WHere Exists (Select nVoucher.ProAcc from nVoucher, Admlink Where ProAcc.ProAcc=nVoucher.ProAcc and Admlink.Voucher=nVoucher.VOucher and Admlink.AdmNo=" & AdmNo & ")")
                If Not rs.Eof Then
                    rs.Edit()

                    If rs.fields("Bank") <> datamaster.fields("Bank") Then
                        Save_Log_Object("PROACC", "PROACC", rs.fields("ProAcc"), "", "", "", "", "", "", "Bank", rs.fields("Bank"), datamaster.fields("Bank"), "Update")
                        rs.fields("Bank") = datamaster.fields("Bank")
                    End If

                    If rs.fields("BankBranch") <> datamaster.fields("BankBranch") Then
                        Save_Log_Object("PROACC", "PROACC", rs.fields("ProAcc"), "", "", "", "", "", "", "BankBranch", rs.fields("BankBranch"), datamaster.fields("BankBranch"), "Update")
                        rs.fields("BankBranch") = datamaster.fields("BankBranch")
                    End If

                    If DecryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), rs.fields("AccountNo"), 1) <> datamaster.fields("AccountNo") Then
                        Save_Log_Object("PROACC", "PROACC", rs.fields("ProAcc"), "", "", "", "", "", "", "AccountNo", DecryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), rs.fields("AccountNo"), 1), datamaster.fields("AccountNo"), "Update")
                        rs.fields("AccountNo") = EncryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), datamaster.fields("AccountNo"))
                    End If

                    If DecryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), rs.fields("BankAccount"), 1) <> datamaster.fields("BankAccount") Then
                        Save_Log_Object("PROACC", "PROACC", rs.fields("ProAcc"), "", "", "", "", "", "", "BankAccount", DecryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), rs.fields("BankAccount"), 1), datamaster.fields("BankAccount"), "Update")
                        rs.fields("BankAccount") = EncryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), datamaster.fields("BankAccount"))
                    End If

                    If rs.fields("BankName") <> datamaster.fields("BankName") Then
                        Save_Log_Object("PROACC", "PROACC", rs.fields("ProAcc"), "", "", "", "", "", "", "BankName", rs.fields("BankName"), datamaster.fields("BankName"), "Update")
                        rs.fields("BankName") = datamaster.fields("BankName")
                    End If
                    rs.fields("Last_Update") = Today.Date
                    rs.fields("Last_Opr") = dbUser
                    If Not rs.Update Then Throw New Exception(db.ErrDesc)
                End If
            End If

            ' Set Output Parameter

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function Check_Compliance_User(ByVal CID As String, ByVal ID As String, ByVal Indicator As String) As Boolean
        Dim rs As Remoteset
        Dim SQL As String
        SQL = "Select C1.ID"
        SQL = SQL & " from ComplianceU  C1"
        SQL = SQL & " Left Join ComplianceUI C2 on (C2.CID=C1.CID and C2.ID=C1.ID and C2.INDICATOR='" & Indicator & "')"
        SQL = SQL & " Where C1.CID='" & CID & "' and C1.ID='" & ID & "'"
        SQL = SQL & " and (Coalesce(C2.INDICATOR,'')='" & Indicator & "' or Not Exists (Select Top 1 C.CID, C.ID From ComplianceUI C Where C.CID=C1.CID and C.ID=C1.ID))"
        rs = db.OpenRemoteSet(SQL)
        Check_Compliance_User = (Not rs.Eof)
    End Function

    Private Function Create_Payor(ByVal ANO As Long, ByVal ID As String, Optional ByVal WPC As Integer = 0, Optional ByVal Grace As Integer = 0, Optional ByVal OverWriteF As Boolean = True) As Boolean
        Try
            Dim rs As RemoteSet
            Dim rsProfile As RemoteSet

            rsProfile = db.OpenRemoteSet("Select ID, Address_1, Address_2, Address_3, Disc, WPC, Grace from Profile Where ID='" & ID & "'")
            If rsProfile.Eof Then
                Throw New Exception("No Profile Found (" & ID & ") !")
            End If

            rs = db.OpenQueryDef("QPayor", adOpenKeyset, ANO)
            If Not rs.Eof Then
                rs.Edit()
            Else
                rs.AddNew()
                rs.fields("ANO") = ANO
                rs.fields("DefaultPayorF") = True
            End If
            If toval(rs.fields("DefaultPayorF")) Then
                ' Only Apply for new payor or DefaultPayorF=1
                rs.fields("Discount") = toval(rsProfile.fields("Disc"))
                rs.fields("WPC") = IIf(WPC = 0, toval(rsProfile.fields("WPC")), WPC)
                rs.fields("Grace") = IIf(Grace = 0, toval(rsProfile.fields("Grace")), Grace)
                rs.fields("Period") = 1 ' Monthly
            End If

            If rs.editmode = AdEditAdd Or OverWriteF Then
                rs.fields("ID") = tostr(rsProfile.fields("ID"))
                rs.fields("Address_1") = tostr(rsProfile.fields("Address_1"))
                rs.fields("Address_2") = tostr(rsProfile.fields("Address_2"))
                rs.fields("Address_3") = tostr(rsProfile.fields("Address_3"))
            End If
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)
            Create_Payor = True
        Catch ex As Exception
            StrMsg = ex.Message
            Create_Payor = False
        End Try
    End Function
    Private Function Create_AccDA(ByVal ANO As Long, ByVal ID As String, Optional ByVal OverWriteF As Boolean = True) As Boolean
        Dim rs As RemoteSet
        Dim Attention As String
        Dim Address As String
        Dim Phone_1 As String
        Dim Phone_2 As String
        Dim Fax_1 As String
        Dim Fax_2 As String
        Dim EMail As String
        Dim Mobile_1 As String
        Dim Mobile_2 As String
        Dim rsProfile As RemoteSet

        On Error GoTo ErrMsg

        rsProfile = db.OpenRemoteSet("Select top 1 * from PAlias Where ID='" & ID & "' and Type='DELIVERY'")
        If Not rsProfile.Eof Then
            Attention = tostr(rsProfile.fields("Name"))
            Address = tostr(rsProfile.fields("Address"))
            Phone_1 = tostr(rsProfile.fields("Phone_1"))
            Phone_2 = tostr(rsProfile.fields("Phone_2"))
            Fax_1 = tostr(rsProfile.fields("Fax_1"))
            Fax_2 = tostr(rsProfile.fields("Fax_2"))
            EMail = tostr(rsProfile.fields("EMail"))
            Mobile_1 = tostr(rsProfile.fields("Mobile_1"))
            Mobile_2 = tostr(rsProfile.fields("Mobile_2"))
        Else
            rsProfile = db.OpenRemoteSet("Select P.Salutation, P.Name, P.Address_1, P.Address_2, P.Address_3, coalesce(City.Description, P.City) as City, P.ZipCode, P.Phone_1, P.Phone_2, P.Fax_1, P.Fax_2, P.Email, P.Mobile_1, P.Mobile_2, coalesce(Province.Description, P.Province) as Province from vProfile P Left Join City on (City.City=P.City) Left Join Province on (Province.Province=P.Province) Where P.ID='" & ID & "'")
            If rsProfile.Eof Then GoTo ErrMsg
            Attention = Trim(tostr(rsProfile.fields("Salutation")) + " " + tostr(rsProfile.fields("Name")))
            Address = tostr(rsProfile.fields("Address_1"))
            If tostr(rsProfile.fields("Address_2")) <> "" Then Address = Address & vbCrLf & tostr(rsProfile.fields("Address_2"))
            If tostr(rsProfile.fields("Address_3")) <> "" Then Address = Address & vbCrLf & tostr(rsProfile.fields("Address_3"))
            If tostr(rsProfile.fields("City")) <> "" Then Address = Address & vbCrLf & tostr(rsProfile.fields("City"))
            If tostr(rsProfile.fields("Province")) <> "" Then Address = Address & vbCrLf & tostr(rsProfile.fields("Province"))
            If tostr(rsProfile.fields("ZipCode")) <> "" Then Address = Address & ", " & tostr(rsProfile.fields("ZipCode"))
            Phone_1 = tostr(rsProfile.fields("Phone_1"))
            Phone_2 = tostr(rsProfile.fields("Phone_2"))
            Fax_1 = tostr(rsProfile.fields("Fax_1"))
            Fax_2 = tostr(rsProfile.fields("Fax_2"))
            EMail = tostr(rsProfile.fields("EMail"))
            Mobile_1 = tostr(rsProfile.fields("Mobile_1"))
            Mobile_2 = tostr(rsProfile.fields("Mobile_2"))
        End If

        rs = db.OpenQueryDef("QAccDA", adOpenKeyset, ANO)
        If Not rs.Eof Then
            rs.Edit()
        Else
            rs.AddNew()
            rs.fields("ANO") = ANO
        End If
        If rs.editmode = adEditAdd Or OverWriteF Then
            rs.fields("Attention") = Attention
            rs.fields("Address") = Address
            rs.fields("Phone_1") = Phone_1
            rs.fields("Phone_2") = Phone_2
            rs.fields("Fax_1") = Fax_1
            rs.fields("Fax_2") = Fax_2
            rs.fields("EMail") = EMail
            rs.fields("Mobile_1") = Mobile_1
            rs.fields("Mobile_2") = Mobile_2
        End If
        If Not rs.Update Then GoTo ErrMsg
        Create_AccDA = True
        Exit Function
ErrMsg:
        Create_AccDA = False
    End Function
    Private Function Create_AccCA(ByVal ANO As Long, ByVal ID As String, Optional ByVal OverWriteF As Boolean = True) As Boolean
        Dim rs As RemoteSet
        Dim Attention As String
        Dim Address As String
        Dim Phone_1 As String
        Dim Phone_2 As String
        Dim Fax_1 As String
        Dim Fax_2 As String
        Dim EMail As String
        Dim Mobile_1 As String
        Dim Mobile_2 As String
        Dim rsProfile As RemoteSet

        On Error GoTo ErrMsg

        rsProfile = db.OpenRemoteSet("Select top 1 * from PAlias Where ID='" & ID & "' and Type='CORRESPONDENCE'")
        If Not rsProfile.Eof Then
            Attention = tostr(rsProfile.fields("Name"))
            Address = tostr(rsProfile.fields("Address"))
            Phone_1 = tostr(rsProfile.fields("Phone_1"))
            Phone_2 = tostr(rsProfile.fields("Phone_2"))
            Fax_1 = tostr(rsProfile.fields("Fax_1"))
            Fax_2 = tostr(rsProfile.fields("Fax_2"))
            EMail = tostr(rsProfile.fields("EMail"))
            Mobile_1 = tostr(rsProfile.fields("Mobile_1"))
            Mobile_2 = tostr(rsProfile.fields("Mobile_2"))
        Else
            rsProfile = db.OpenRemoteSet("Select P.Salutation, P.Name, P.Address_1, P.Address_2, P.Address_3, coalesce(City.Description, P.City) as City, P.ZipCode, P.Phone_1, P.Phone_2, P.Fax_1, P.Fax_2, P.Email, P.Mobile_1, P.Mobile_2, coalesce(Province.Description, P.Province) as Province from vProfile P Left Join City on (City.City=P.City) Left Join Province on (Province.Province=P.Province) Where P.ID='" & ID & "'")
            If rsProfile.Eof Then GoTo ErrMsg
            Attention = Trim(tostr(rsProfile.fields("Salutation")) + " " + tostr(rsProfile.fields("Name")))
            Address = tostr(rsProfile.fields("Address_1"))
            If tostr(rsProfile.fields("Address_2")) <> "" Then Address = Address & vbCrLf & tostr(rsProfile.fields("Address_2"))
            If tostr(rsProfile.fields("Address_3")) <> "" Then Address = Address & vbCrLf & tostr(rsProfile.fields("Address_3"))
            If tostr(rsProfile.fields("City")) <> "" Then Address = Address & vbCrLf & tostr(rsProfile.fields("City"))
            If tostr(rsProfile.fields("Province")) <> "" Then Address = Address & vbCrLf & tostr(rsProfile.fields("Province"))
            If tostr(rsProfile.fields("ZipCode")) <> "" Then Address = Address & ", " & tostr(rsProfile.fields("ZipCode"))
            Phone_1 = tostr(rsProfile.fields("Phone_1"))
            Phone_2 = tostr(rsProfile.fields("Phone_2"))
            Fax_1 = tostr(rsProfile.fields("Fax_1"))
            Fax_2 = tostr(rsProfile.fields("Fax_2"))
            EMail = tostr(rsProfile.fields("EMail"))
            Mobile_1 = tostr(rsProfile.fields("Mobile_1"))
            Mobile_2 = tostr(rsProfile.fields("Mobile_2"))
        End If

        rs = db.OpenQueryDef("QAccCA", adOpenKeyset, ANO)
        If Not rs.Eof Then
            rs.Edit()
        Else
            rs.AddNew()
            rs.fields("ANO") = ANO
        End If
        If rs.editmode = AdEditAdd Or OverWriteF Then
            rs.fields("Attention") = Left(Attention, 255)
            rs.fields("Address") = Left(Address, 255)
            rs.fields("Phone_1") = Phone_1
            rs.fields("Phone_2") = Phone_2
            rs.fields("Fax_1") = Fax_1
            rs.fields("Fax_2") = Fax_2
            rs.fields("EMail") = EMail
            rs.fields("Mobile_1") = Mobile_1
            rs.fields("Mobile_2") = Mobile_2
        End If
        If Not rs.Update Then GoTo ErrMsg
        Create_AccCA = True
        Exit Function
ErrMsg:
        StrMsg = StrError(StrMsg)
        Create_AccCA = False
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetClaimPaymentSchedule(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long

            ANO = getInParamValue(InParam, 0)    ' Original ANO

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * From ClaimPaymentSchedule Where ANO=" & ANO)
            findex = rs.fields_index()
            fbuffer = rs.GetRecord()
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function SaveClaimPaymentSchedule(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long
            ANO = getInParamValue(InParam, 0) ' Original ANO

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim EditMode As Integer

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select ANO from ACCEPTANCE Where ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("Policy Data Not Found !")
            End If

            rs = db.OpenRemoteSet("Select * from ClaimPaymentSchedule Where ANO=" & ANO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If

            EditMode = rs.editmode

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                If EditMode = AdEditInProgress Then
                    If rs.fields(findex(i)) <> GetFieldValue(fbuffer(i)) Then
                        Save_Log_Object("ClaimPaymentSchedule", "ANO", ANO, "", "", "", "", "", "", findex(i), rs.fields(findex(i)), GetFieldValue(fbuffer(i)), "Update")
                    End If
                End If
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            If EditMode = AdEditAdd Then
                rs.fields("ANO") = ANO
            End If
            rs.fields("Last_Opr") = dbUser
            rs.fields("Last_Update") = Today.Date
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function Get_RateCon(ByVal ANO As Long, ByVal Code As String, ByVal TSI As Double, Optional ByVal Unit As String = "C", Optional ByRef Loading As Double = 0, Optional ByRef TExcess As Double = 0, Optional ByRef PCTLOSS As Double = 0, Optional ByRef PCALC As String = "", Optional ByRef HoldF As Boolean = False, Optional ByRef HoldMessage As String = "", Optional ByVal pSDate As Object = Nothing, Optional ByRef CID As String = "", Optional ByRef SRate As Double = 0, Optional ByRef RateConF As Boolean = True, Optional ByRef ScalePCT As Double = 0) As Double
        Dim rs As RemoteSet
        Dim RateLoadingF As Boolean
        Dim Rate As Double
        Dim SDate As Date
        If Unit = "" Then Unit = "C"
        If TSI = 0 Then TSI = 1
        If IsDate(pSDate) Then
            SDate = pSDate
        Else
            rs = db.OpenRemoteSet("Select SDate from Acceptance Where ANO=" & ANO)
            SDate = rs.fields("SDate")
        End If
        rs = db.OpenRemoteSet("Select Top 1 * from fn_RateCon(" & ANO & ",'" & Code & "'," & tovalSF(TSI) & ",'" & Format(SDate, "MM/DD/YYYY") & "') Where AllowedF=1")
        RateConF = (Not rs.Eof)
        Code = tostr(rs.fields("Code"))
        RateLoadingF = toval(rs.fields("RateLoadingF"))
        Rate = toval(rs.fields("Rate"))
        Loading = toval(rs.fields("Loading")) / 100
        TExcess = toval(rs.fields("TExcess"))
        PCTLOSS = toval(rs.fields("PCTLOSS"))
        PCALC = tostr(rs.fields("PCALC"))
        If TSI <> 0 And Unit = "F" Then
            ' To Flat Rate
            Rate = Round(TSI * Rate / 100, 2)
        ElseIf Unit = "M" Then
            ' To Mile
            Rate = Rate * 10
        End If
        If toval(rs.fields("ProgressiveF")) AndAlso TSI > toval(rs.fields("MaxTSI")) Then Rate = 0
        If Loading = 0 Then Loading = 1
        If RateLoadingF OrElse PCALC = "S" Then
            Loading = Loading * Get_RateLoading(ANO, Code, SDate, TSI, HoldF, HoldMessage, CID, , SRate, ScalePCT)
            If TSI <> 0 And Unit = "F" Then
                ' To Flat Rate
                SRate = Round(TSI * SRate / 100, 2)
            ElseIf Unit = "M" Then
                ' To Mile
                SRate = SRate * 10
            End If
        End If
        Get_RateCon = Rate
    End Function
    Private Function Get_RateLoading(ByVal ANO As Long, ByVal Code As String, ByVal PDate As Date, Optional ByVal TSI As Double = 0, Optional ByRef HoldF As Boolean = False, Optional ByRef HoldMessage As String = "", Optional ByRef CID As String = "", Optional ByRef AggregateTSI As Double = 0, Optional ByRef SRate As Double = 0, Optional ByRef ScalePCT As Double = 0) As Double
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select * from Get_RateLoading(" & ANO & ",'" & Code & "','" & Format(todate(PDate), "MM/DD/YYYY") & "'," & tovalSF(TSI) & ") RateCon")
        HoldF = toval(rs.fields("HoldF"))
        HoldMessage = tostr(rs.fields("HoldMessage"))
        CID = tostr(rs.fields("CID"))
        AggregateTSI = toval(rs.fields("AggregateTSI"))
        SRate = toval(rs.fields("SRate"))
        ScalePCT = toval(rs.fields("ScalePCT"))
        Get_RateLoading = toval(rs.fields("Loading"))
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function Get_Short_Period_Calc(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter

            Dim ANO As Long
            Dim TOC As String
            Dim SDate As Date
            Dim EDate As Date
            Dim PDate As Date

            ANO = getInParamValue(InParam, 0)
            TOC = getInParamValue(InParam, 1)
            SDate = getInParamValue(InParam, 2)
            EDate = getInParamValue(InParam, 3)
            PDate = getInParamValue(InParam, 4)

            Dim ValueID As String = Search_VALUEID_For_PScale(ANO, TOC, "TypeOfTrip")

            ReDim OutParam(0)
            OutParam(0) = Short_Period_Calc(TOC, SDate, EDate, PDate, ValueID)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Private Function Search_VALUEID_For_PScale(ByVal ANO As Long, ByVal TOC As String, ByVal ParamArray FLDMAP() As Object) As String
        Dim i As Integer
        Dim j As Integer
        Dim rsTOC As RemoteSet
        Dim rsInfo As RemoteSet
        Dim FLDID As String

        Dim NItem As Integer
        NItem = UBound(FLDMAP)

        rsTOC = db.OpenQueryDef("QTOC_GENHTAB", adOpenForwardonly, TOC)
        rsInfo = db.OpenRemoteSet("Select * from AInfo where ANO=" & ANO)
        For i = 1 To 15
            FLDID = rsTOC.fields("FLDID" & i)
            If FLDID <> "" Then
                For j = 0 To NItem
                    If UCase(rsTOC.fields("FLDMAP" & i)) = UCase(FLDMAP(j)) Then
                        Search_VALUEID_For_PScale = rsInfo.fields("VALUEID" & i)
                        Exit Function
                    End If
                Next
            Else
                Exit For
            End If
        Next

        rsTOC = db.OpenQueryDef("QTOC_EXT_GENHTAB", adOpenForwardonly, TOC)
        rsInfo = db.OpenRemoteSet("Select * from AInfo_Ext where ANO=" & ANO)
        For i = 16 To 31
            FLDID = rsTOC.fields("FLDID" & i)
            If FLDID <> "" Then
                For j = 0 To NItem
                    If UCase(rsTOC.fields("FLDMAP" & i)) = UCase(FLDMAP(j)) Then
                        Search_VALUEID_For_PScale = rsInfo.fields("VALUEID" & i)
                        Exit Function
                    End If
                Next
            Else
                Exit For
            End If
        Next
        Search_VALUEID_For_PScale = ""
    End Function
    Private Function Short_Period_Prorata(ByVal TOC As String, ByVal SDate As Date, ByVal EDate As Date, ByVal Basis As String, ByVal ValueID As String) As Double
        Dim rsP As Remoteset
        Dim NOP As Double
        Dim Diff As Integer

        'M
        If Basis = "M" Then
            NOP = DateDiff(DateInterval.Month, SDate, EDate)

            Diff = DateDiff(DateInterval.Day, DateAdd(DateInterval.Month, NOP, SDate), EDate)
            NOP = NOP + Diff / 30
            rsP = db.OpenQueryDef("QPREMIUM_SCALE", adOpenForwardOnly, TOC, "M", NOP, ValueID)
            If Not rsP.Eof Then
                Short_Period_Prorata = rsP.fields("PCT")
            Else
                Short_Period_Prorata = 100
            End If
        Else
            NOP = DateDiff(DateInterval.Day, SDate, EDate)
            rsP = db.OpenQueryDef("QPREMIUM_SCALE", adOpenForwardOnly, TOC, "D", NOP, ValueID)
            If Not rsP.Eof Then
                Short_Period_Prorata = rsP.fields("PCT")
            Else
                Short_Period_Prorata = 100
            End If
        End If
    End Function
    Private Function Short_Period_Calc(ByVal TOC As String, ByVal SDate As Date, ByVal EDate As Date, ByVal PDate As Date, Optional ByVal ValueID As String = "") As Double
        If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, -1, EDate), DateAdd(DateInterval.Month, 1, SDate)) > 0 Then
            ' Short Period Policy
            Short_Period_Calc = Short_Period_Prorata(TOC, SDate, PDate, "D", ValueID)
        Else
            Short_Period_Calc = Short_Period_Prorata(TOC, SDate, PDate, "M", ValueID)
        End If
    End Function
    Private Function GetPremiumRate(ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            ' Get Input Parameter
            Dim ANO As Long
            Dim Code As String
            Dim Unit As String
            Dim SDate As Date
            Dim PCTLoss As Double

            ANO = getInParamValue(InParam, 0)
            Code = getInParamValue(InParam, 1)
            Unit = getInParamValue(InParam, 2)
            SDate = getInParamValue(InParam, 3)
            PCTLoss = toval(getInParamValue(InParam, 4))
            If PCTLoss = 0 Then PCTLoss = 100

            Dim SQL As String
            Dim SIBasedID As Integer
            Dim SI As Double
            Dim Rate As Double
            Dim Loading As Double
            Dim TExcess As Double
            Dim SRate As Double
            Dim PCALC As String = ""
            Dim RateConF As Boolean = True
            Dim ScalePCT As Double

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select SIBasedID from RateTab Where Code='" & Code & "'")
            SIBasedID = toval(rs.fields("SIBasedID"))

            SQL = "Select I.ANO, I.Currency, Max(A.TOC) as TOC, Max(A.SDate) as PSDate, Max(C.PHolder) as PHolder, sum(I.Amount) as Amount, sum(I.SI) as SI "
            SQL = SQL & " From ICover I, Acceptance A, Cover C"
            SQL = SQL & " Where A.ANO=I.ANO and C.CNO=A.CNO and I.ANO=" & ANO & " and Exists (Select P.TOI from RateCov R, PClass P Where R.Code='" & Code & "' and P.TOR=R.TOR and P.TOI=I.TOI)"
            SQL = SQL & " Group by I.ANO, I.Currency"
            rs = db.OpenRemoteSet(SQL)
            SI = 0
            Do While Not rs.Eof
                If SIBasedID = 0 Then
                    SI = SI + toval(rs.fields("SI")) * PCTLoss / 100 * GetReinsuranceRate(rs.fields("TOC"), rs.fields("Currency"), rs.fields("PSDate"), rs.fields("PHolder"))
                Else
                    SI = SI + toval(rs.fields("Amount")) * PCTLoss / 100 * GetReinsuranceRate(rs.fields("TOC"), rs.fields("Currency"), rs.fields("PSDate"), rs.fields("PHolder"))
                End If
                rs.MoveNext()
            Loop
            Rate = Get_RateCon(ANO, Code, SI, Unit, Loading, TExcess, PCTLoss, PCALC, , , SDate, , SRate, RateConF, ScalePCT)

            If Not RateConF Then
                Throw New Exception("No Rate Condition !")
            End If

            ' Set Output Parameter
            ReDim OutParam(6)
            OutParam(0) = Rate
            OutParam(1) = Loading
            OutParam(2) = PCTLoss
            OutParam(3) = SRate
            OutParam(4) = TExcess
            OutParam(5) = PCALC
            OutParam(6) = ScalePCT
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPremiumRate(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            'Dim ANO As Long
            'Dim Code As String
            'Dim Unit As String
            'Dim SDate As Date
            'Dim PCTLoss As Double

            'ANO = getInParamValue(InParam, 0)
            'Code = getInParamValue(InParam, 1)
            'Unit = getInParamValue(InParam, 2)
            'SDate = getInParamValue(InParam, 3)
            'PCTLoss = toval(getInParamValue(InParam, 4))

            If Not GetPremiumRate(InParam, OutParam, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            ' Set Output Parameter
            'ReDim OutParam(5)
            'OutParam(0) = Rate()
            'OutParam(1) = Loading
            'OutParam(2) = PCTLoss
            'OutParam(3) = SRate
            'OutParam(4) = TExcess
            'OutParam(5) = PCALC

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function GetInvoice(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim INO As Long
            ANO = getInParamValue(InParam, 0)   ' Acceptance ANO 
            INO = getInParamValue(InParam, 1)   ' Interest No

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from AccInvoice Where ANO=" & ANO & " and INO=" & INO)
            findex = rs.fields_index
            fbuffer = rs.GetRecord

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function SaveInvoice(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim INO As Long
            ANO = getInParamValue(InParam, 0)   ' Acceptance ANO 
            INO = getInParamValue(InParam, 1)   ' Interest No

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select ANO from Acceptance Where pANO=" & ANO & " and AType not in ('A','C')")
            If Not rs.Eof Then
                Throw New Exception("Already Has An Endorsement !")
            End If

            Dim EditMode As Integer
            rs = db.OpenRemoteSet("Select * from AccInvoice Where ANO=" & ANO & " and INO=" & INO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            EditMode = rs.editmode

            ' Check Status
            If EditMode = AdEditInProgress Then
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                If EditMode = AdEditInProgress Then
                    If Not Save_Log_Object("AccInvoice", "ANO", ANO, "INO", INO, "", "", "", "", rs.fields_Name(i), rs.fields(findex(i)), GetFieldValue(fbuffer(i)), "Update") Then Throw New Exception("Can't Save to Log !")
                End If
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            If EditMode = AdEditAdd Then
                rs.fields("ANO") = ANO
                rs.fields("INO") = INO
            End If
            rs.fields("Last_Opr") = dbUser
            rs.fields("Last_Update") = Today.Date
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter
            ' Nothing

            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function DropInvoice(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim ANO As Long
            Dim INO As Long
            ANO = getInParamValue(InParam, 0)   ' Acceptance ANO 
            INO = getInParamValue(InParam, 1)   ' Interest No

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select ANO from Acceptance Where pANO=" & ANO & " and AType not in ('A','C')")
            If Not rs.Eof Then
                Throw New Exception("Already Has An Endorsement !")
            End If

            rs = db.OpenRemoteSet("Select * from AccInvoice Where ANO=" & ANO & " and INO=" & INO)
            If rs.Eof Then
                ErrMsg = "No Invoice Data !"
                Throw New Exception(ErrMsg)
            Else
            End If

            If Not Save_Log_Object("AccInvoice", "ANO", ANO, "INO", INO, "InvoiceNo", rs.fields("InvoiceNo"), "", "", "", "", "", "Delete") Then Throw New Exception("Can't Save to Log !")

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function VerifyAccInvoice(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            Dim i As Long
            Dim NumVerify As Long = 0
            Dim skipPolicyCheckF As Boolean = toval(getInParamValue(InParam, 1))
            db.Begintrans()
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If (Dataset.Tables(0).Rows(i)("ProcessFlag") = True) Then
                    Try
                        Dim rs As RemoteSet
                        Dim PolicyNo As String
                        Dim CertificateNo As String
                        Dim InvoiceNo As String
                        Dim InvoiceDate As String
                        Dim InvoiceNominal As Double
                        Dim InvoiceRemarks As String

                        StrMsg = ""

                        PolicyNo = tostr(GetItemValue("PolicyNo", Dataset, i))
                        CertificateNo = tostr(GetItemValue("CertificateNo", Dataset, i))
                        InvoiceNo = tostr(GetItemValue("InvoiceNo", Dataset, i))
                        InvoiceDate = todate(GetItemValue("InvoiceDate", Dataset, i))
                        InvoiceNominal = toval(GetItemValue("InvoiceNominal", Dataset, i))
                        InvoiceRemarks = tostr(GetItemValue("InvoiceRemarks", Dataset, i))

                        Dim ANO As Long
                        Dim INO As Long
                        rs = db.OpenRemoteSet("Select Top 1 A.ANO from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) Where A.AType not in ('A','C') and A.LANO<>-1 and A.ANO<>A.LANO and A.PolicyNo='" & PolicyNo & "' and A.CertificateNo='" & CertificateNo & "' Order By A.ANO Desc")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Please Check Policy No. " & PolicyNo & " and Certificate No. " & CertificateNo & " !" & vbCrLf
                        End If

                        ANO = rs.fields("ANO")

                        If InvoiceNo = "" Then
                            StrMsg = StrMsg & "No Invoice No !" & vbCrLf
                        End If

                        rs = db.OpenRemoteSet("Select INO from ICover Where ANO=" & rs.fields("ANO") & " and RefNo='" & InvoiceNo & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Please Check interest Ref No. " & InvoiceNo & ", Policy No. " & PolicyNo & " and Certificate No. " & CertificateNo & " !" & vbCrLf
                        End If
                        INO = rs.fields("INO")

                        If tostr(GetItemValue("InvoiceDate", Dataset, i)) = "" Or Not IsDate(GetItemValue("InvoiceDate", Dataset, i)) Then
                            StrMsg = StrMsg & "Incorrect Invoice Date !" & vbCrLf
                        End If

                        If StrMsg = "" Then
                            Dataset.Tables(0).Rows(i)("Status") = "Verified"
                        Else
                            Dataset.Tables(0).Rows(i)("Status") = "failed "
                        End If

                    Catch ex As Exception
                        StrMsg = StrMsg & ex.Message
                        Dataset.Tables(0).Rows(i)("Status") = "failed "
                    End Try
                    Dataset.Tables(0).Rows(i)("Message") = StrMsg
                    NumVerify = NumVerify + 1
                End If
            Next


            ReDim OutParam(1)
            OutParam(0) = NumVerify
            OutParam(1) = ""

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function UploadAccInvoice(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            Dim ProcessOpt As Integer = InParam(0)
            Dim i As Long
            Dim numProcess As Long = 0

            db.Begintrans()

            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                StrMsg = ""
                If (Dataset.Tables(0).Rows(i)("ProcessFlag") = True) And (Dataset.Tables(0).Rows(i)("Status") = "Verified" Or ProcessOpt <> 0) Then
                    If Dataset.Tables(0).Rows(i)("Status") <> "Verified" And ProcessOpt = 2 Then
                        ErrMsg = "Record No. " & i + 1 & " is not Verified Yet !"
                        Throw New Exception(ErrMsg)
                    End If
                    Try
                        Dim rs As RemoteSet
                        Dim PolicyNo As String
                        Dim CertificateNo As String
                        Dim InvoiceNo As String
                        Dim InvoiceDate As String
                        Dim InvoiceNominal As Double
                        Dim InvoiceRemarks As String

                        StrMsg = ""

                        PolicyNo = tostr(GetItemValue("PolicyNo", Dataset, i))
                        CertificateNo = tostr(GetItemValue("CertificateNo", Dataset, i))
                        InvoiceNo = tostr(GetItemValue("InvoiceNo", Dataset, i))
                        InvoiceDate = todate(GetItemValue("InvoiceDate", Dataset, i))
                        InvoiceNominal = toval(GetItemValue("InvoiceNominal", Dataset, i))
                        InvoiceRemarks = tostr(GetItemValue("InvoiceRemarks", Dataset, i))

                        Dim ANO As Long
                        Dim INO As Long
                        rs = db.OpenRemoteSet("Select Top 1 A.ANO from Acceptance A Inner Join Cover C on (C.CNO=A.CNO) Where A.AType not in ('A','C') and A.LANO<>-1 and A.ANO<>A.LANO and A.PolicyNo='" & PolicyNo & "' and A.CertificateNo='" & CertificateNo & "' Order By A.ANO Desc")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Please Check Policy No. " & PolicyNo & " and Certificate No. " & CertificateNo & " !" & vbCrLf
                        End If

                        ANO = rs.fields("ANO")

                        If InvoiceNo = "" Then
                            StrMsg = StrMsg & "No Invoice No !" & vbCrLf
                        End If

                        rs = db.OpenRemoteSet("Select INO from ICover Where ANO=" & rs.fields("ANO") & " and RefNo='" & InvoiceNo & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Please Check interest Ref No. " & InvoiceNo & ", Policy No. " & PolicyNo & " and Certificate No. " & CertificateNo & " !" & vbCrLf
                        End If
                        INO = rs.fields("INO")

                        rs = db.OpenRemoteSet("Select * from AccInvoice Where ANO=" & ANO & " and INO=" & INO)
                        If rs.Eof Then
                            rs.AddNew()
                            rs.fields("ANO") = ANO
                            rs.fields("INO") = INO
                        Else
                            rs.Edit()
                        End If
                        rs.fields("InvoiceNo") = InvoiceNo
                        rs.fields("InvoiceDate") = InvoiceDate
                        rs.fields("InvoiceNominal") = InvoiceNominal
                        rs.fields("InvoiceRemarks") = InvoiceRemarks
                        rs.fields("Last_Opr") = dbUser
                        rs.fields("Last_Update") = Today.Date
                        If Not rs.Update Then Throw New Exception(db.ErrDesc)

                        If StrMsg = "" Then
                            Dataset.Tables(0).Rows(i)("Status") = "Processed"
                        Else
                            Dataset.Tables(0).Rows(i)("Status") = "failed "
                        End If
                    Catch ex As Exception
                        StrMsg = ex.Message
                        Dataset.Tables(0).Rows(i)("Status") = "failed "
                    End Try
                    Dataset.Tables(0).Rows(i)("Message") = StrMsg
                    numProcess = numProcess + 1
                End If

            Next

            ReDim OutParam(1)
            OutParam(0) = numProcess

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    Private Function GetLockedDate(ByVal Branch As String) As Date
        Dim LockedDate As Date
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select LockedDate from SysTab", adOpenForwardonly)
        LockedDate = todate(rs.fields("LockedDate"))
        rs = db.OpenRemoteSet("Select LockedDate from Branch Where Branch='" & Branch & "' and LockedDate>'" & Format(LockedDate, "MM/DD/YYYY") & "'")
        If Not rs.Eof Then
            LockedDate = todate(rs.fields("LockedDate"))
        End If
        GetLockedDate = LockedDate
    End Function
    Private Sub GetMinMaxBackDatedPolicy(ByVal TOC As String, ByRef MinBackDatedPolicy As Integer, ByRef MaxBackDatedPolicy As Integer)
        Dim rs As RemoteSet

        rs = db.OpenRemoteSet("Select MinBackDatedPolicy, MaxBackDatedPolicy from SysTab")
        MinBackDatedPolicy = toval(rs.fields("MinBackDatedPolicy"))
        MaxBackDatedPolicy = toval(rs.fields("MaxBackDatedPolicy"))
        rs = db.OpenRemoteSet("Select MinBackDatedPolicy, MaxBackDatedPolicy from TOC Where TOC='" & TOC & "'")
        If toval(rs.fields("MinBackDatedPolicy")) <> 0 Then
            MinBackDatedPolicy = toval(rs.fields("MinBackDatedPolicy"))
        End If
        If toval(rs.fields("MaxBackDatedPolicy")) <> 0 Then
            MaxBackDatedPolicy = toval(rs.fields("MaxBackDatedPolicy"))
        End If
    End Sub
    Private Sub GetMinMaxGracePeriodPolicy(ByRef MinGracePeriodPolicy As Integer, ByRef MaxGracePeriodPolicy As Integer)
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select MinGracePeriodPolicy, MaxGracePeriodPolicy from SysTab")
        MinGracePeriodPolicy = toval(rs.fields("MinGracePeriodPolicy"))
        MaxGracePeriodPolicy = toval(rs.fields("MaxGracePeriodPolicy"))
    End Sub
    Private Function GetBackDatedDate(ByVal Branch As String) As Date
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select BackDatedDate from Branch Where Branch='" & Branch & "'", adOpenForwardonly)
        GetBackDatedDate = todate(rs.fields("BackDatedDate"))
    End Function
    Private Function GetUserBLLocked(ByVal ID As String) As Integer
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select BLLocked from SysUser Where ID='" & ID & "'")
        GetUserBLLocked = toval(rs.fields("BLLocked"))
    End Function

    Private Function GetEndDate(ByVal Branch As String) As Date
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select End_Date from Branch Where Branch='" & Branch & "'", adOpenForwardonly)
        GetEndDate = todate(rs.fields("End_Date"))
    End Function
    Private Function GetApprovalLimit(ByVal UserID As String, ByVal TOC As String, ByVal ISType As String, ByVal InOut As String, ByVal Amount As Double, Optional ByVal Branch As String = "", Optional ByVal ApprLimitPct As Double = 100, Optional ByVal ANO As Long = -1, Optional ByVal TotalAmount As Double = 0) As Boolean
        Dim rs As RemoteSet

        Dim UType As String
        Dim BType As String
        Dim LimitIn As Double
        Dim LimitOut As Double
        Dim TLimitIn As Double = 0

        On Error GoTo errmsg


        If ISType = "I" Or ISType = "L" Or ISType = "C" Then
            BType = "DIRECT"
        Else
            BType = "INDIRECT"
        End If

        rs = db.OpenRemoteSet("select S1.Type, Case When S1.LimitIn<>0 then S1.LimitIn else S2.LimitIn end as LimitIn, Case When S1.LimitOut<>0 then S1.LimitOut else S2.LimitOut end as LimitOut from sysuser S1, SysUserType S2 where S2.Type=S1.Type and S1.ID='" & UserID & "'")
        If Not rs.Eof Then
            UType = rs.fields("Type")
            If UType = "SUPERVISOR" Then
                Return True
            Else
                LimitIn = toval(rs.fields("LimitIn"))
                LimitOut = toval(rs.fields("LimitOut"))
                If TOC <> "" Then
                    rs = db.OpenRemoteSet("Select Top 1 Effective, LimitIn, LimitOut from SysUserTypeTOC Where Type='" & UType & "' and BType='" & BType & "' and TOC='" & TOC & "' and Effective<='" & Format(Today.Date, "MM/DD/YYYY") & "' Order By Effective Desc")
                    If Not rs.Eof Then
                        If toval(rs.fields("LimitIn")) <> 0 Then LimitIn = toval(rs.fields("LimitIn"))
                        If toval(rs.fields("LimitOut")) <> 0 Then LimitOut = toval(rs.fields("LimitOut"))

                        If ANO <> -1 Then
                            rs = db.OpenQueryDef("QGetApprovalLimitByTOL", adOpenKeyset, ANO, UType, TOC, BType, rs.fields("Effective"))
                            If Not rs.Eof Then
                                ApprLimitPct = toval(rs.fields("Limit"))
                            End If
                        End If
                    End If

                    rs = db.OpenRemoteSet("Select Top 1 S.LimitIn, S.TotalLimitIn, Case When '" & InOut & "'='Out_1' and TOC.AppLimitOutOption='LimitOut_1' Then S.LimitOut_1 Else S.LimitOut End as LimitOut from SysUserTOC S Inner Join TOC on (TOC.TOC=S.TOC) Where S.ID='" & UserID & "' and S.TOC='" & TOC & "' and S.Effective<='" & Format(Today.Date, "MM/DD/YYYY") & "' Order By S.Effective Desc")
                    If toval(rs.fields("LimitIn")) <> 0 Then LimitIn = toval(rs.fields("LimitIn"))
                    If toval(rs.fields("LimitOut")) <> 0 Then LimitOut = toval(rs.fields("LimitOut"))
                    If toval(rs.fields("TotalLimitIn")) <> 0 Then TLimitIn = toval(rs.fields("TotalLimitIn"))

                    If Branch <> "" Then
                        rs = db.OpenRemoteSet("Select Top 1 LimitIn, LimitOut from SysUserTOCBranch Where ID='" & UserID & "' and TOC='" & TOC & "' and Branch='" & Branch & "' and Effective<='" & Format(Today.Date, "MM/DD/YYYY") & "' Order By Effective Desc")
                        If toval(rs.fields("LimitIn")) <> 0 Then LimitIn = toval(rs.fields("LimitIn"))
                        If toval(rs.fields("LimitOut")) <> 0 Then LimitOut = toval(rs.fields("LimitOut"))
                    End If
                End If
                Select Case InOut
                    Case "In"
                        'If Amount <= LimitIn Or LimitIn = 0 Then Return True
                        If TotalAmount = 0 Then
                            If Amount <= LimitIn * ApprLimitPct / 100 Then Return True
                        Else
                            If Amount <= LimitIn * ApprLimitPct / 100 AndAlso TotalAmount <= TLimitIn * ApprLimitPct / 100 Then Return True
                        End If
                    Case "Out", "Out_1"
                        'If Amount <= LimitOut Or LimitOut = 0 Then Return True
                        If Amount <= LimitOut * ApprLimitPct / 100 Then Return True
                End Select
            End If
            ' Get User Delegation
            rs = db.OpenRemoteSet("Select top 1 ID from SysUserDelegation Where PID='" & UserID & "' and status='R' and cast(convert(varchar(10), getdate(),1) as smalldatetime) between effective and expiry")
            If Not rs.Eof AndAlso rs.fields("ID") <> dbUser Then
                Return GetApprovalLimit(rs.fields("ID"), TOC, ISType, InOut, Amount, Branch, ApprLimitPct, ANO, TotalAmount)
            End If
            Return False
        Else
            Return False
        End If
        Exit Function
errmsg:
        Return False
    End Function

    Private Function Get_FLDTAG(ByVal FLDID As String) As String
        Dim rs As RemoteSet
        rs = db.OpenQueryDef("QGenHTab", adOpenKeyset, FLDID)
        Get_FLDTAG = tostr(rs.fields("FLDTAG"))
    End Function
    Public Function Get_Accumulation(ByVal RegID As String, ByVal FLDID As String, ByVal ValueID As String, ByVal RType As String, ByVal RYear As Integer) As Double
        Dim rs As RemoteSet

        rs = db.OpenQueryDef("QAccumulation", adOpenForwardonly, RegID, FLDID, ValueID, RType, RYear)
        Get_Accumulation = toval(rs.fields("Amount"))
    End Function
    Private Function Compare_DueDate_Payor_and_Facultative(ByVal ANO As Long, ByVal LANO As Long, ByVal MinLapseFacDue As Integer, ByVal BookingDate As Date, ByVal Effective As Date, Optional ByVal FID As String = "", Optional ByRef ErrMsg As String = "") As Boolean
        Dim rsP As RemoteSet
        Dim rsF As RemoteSet
        Dim PDueDate As Date
        Dim FDueDate As Date
        Dim FacDueByEffectiveF As Boolean
        Dim DueByEffectiveF As Boolean
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select FacDueByEffectiveF, DueByEffectiveF from SysTab")
        FacDueByEffectiveF = toval(rs.fields("FacDueByEffectiveF"))
        DueByEffectiveF = toval(rs.fields("DueByEffectiveF"))

        ErrMsg = ""
        If ANO <> LANO And LANO <> -1 Then
            rsP = db.OpenRemoteSet("Select wpc, grace, installment, idate_1, idate_2, idate_3, idate_4, idate_5, idate_6, idate_7, idate_8, idate_9, idate_10, idate_11, idate_12, idate_13, idate_14, idate_15, idate_16, idate_17, idate_18, idate_19, idate_20, idate_21, idate_22, idate_23, idate_24 from APayor Where ANO=" & ANO)
            If rsP.Eof Then
                rsP = db.OpenRemoteSet("Select wpc, grace, installment, idate_1, idate_2, idate_3, idate_4, idate_5, idate_6, idate_7, idate_8, idate_9, idate_10, idate_11, idate_12, idate_13, idate_14, idate_15, idate_16, idate_17, idate_18, idate_19, idate_20, idate_21, idate_22, idate_23, idate_24 from Payor Where ANO=" & LANO)
            End If
        Else
            rsP = db.OpenRemoteSet("Select wpc, grace, installment, idate_1, idate_2, idate_3, idate_4, idate_5, idate_6, idate_7, idate_8, idate_9, idate_10, idate_11, idate_12, idate_13, idate_14, idate_15, idate_16, idate_17, idate_18, idate_19, idate_20, idate_21, idate_22, idate_23, idate_24 from Payor Where ANO=" & ANO)
        End If
        If FID <> "" Then
            rsF = db.OpenRemoteSet("Select wpc, grace, installment, idate_1, idate_2, idate_3, idate_4, idate_5, idate_6, idate_7, idate_8, idate_9, idate_10, idate_11, idate_12, idate_13, idate_14, idate_15, idate_16, idate_17, idate_18, idate_19, idate_20, idate_21, idate_22, idate_23, idate_24, R.SDate from RArrDet R, FPart F Where F.ANO=R.ANO and F.FID=R.FID and R.AType<>'C' and R.ANO=" & ANO & " and R.FID='" & FID & "'")
        Else
            rsF = db.OpenRemoteSet("Select wpc, grace, installment, idate_1, idate_2, idate_3, idate_4, idate_5, idate_6, idate_7, idate_8, idate_9, idate_10, idate_11, idate_12, idate_13, idate_14, idate_15, idate_16, idate_17, idate_18, idate_19, idate_20, idate_21, idate_22, idate_23, idate_24, R.SDate from RArrDet R, FPart F Where F.ANO=R.ANO and F.FID=R.FID and R.AType<>'C' and R.ANO=" & ANO & "")
        End If
        Dim i As Integer
        Dim fInstallment As Integer
        Dim pInstallment As Integer
        Dim Installment As Integer
        Dim fPCT As Integer = 0
        Dim pPCT As Integer = 0
        Dim FacDueDate As Date
        Dim PolDueDate As Date
        Do While Not rsF.Eof
            If FacDueByEffectiveF Then
                FacDueDate = IIf(todate(rsF.fields("SDate")) < Effective, Effective, todate(rsF.fields("SDate")))
            Else
                If DueByEffectiveF Then
                    FacDueDate = Effective
                Else
                    FacDueDate = BookingDate
                End If
            End If

            If DueByEffectiveF Then
                PolDueDate = Effective
            Else
                PolDueDate = BookingDate
            End If

            fInstallment = rsF.fields("Installment")
            pInstallment = rsP.fields("Installment")

            If fInstallment <> pInstallment AndAlso Not db.GetPrivileges("ALLOWFACINSTDIFFERSWITHPAYORINST") Then
                ErrMsg = "Facultative installment differs with payor installment (Fac Slip No. " & rsF.fields("SlipNO") & ") !"
                Return False
            End If

            If pInstallment <> 0 Then
                If fInstallment <> 0 Then
                    If pInstallment <> fInstallment Then
                        ' Get First Installment
                        PDueDate = rsP.fields("IDate_" & 1)
                        FDueDate = rsF.fields("IDate_" & 1)
                        If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), FDueDate) < 0 Then
                            If MinLapseFacDue = 0 Then
                                ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Payor Due on " & Format(PDueDate, "dd-mmm-yy") & " !"
                            Else
                                ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Date on " & Format(DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), "dd-mmm-yy") & " !"
                            End If
                            Return False
                        End If
                        ' Get Last Installment
                        PDueDate = rsP.fields("IDate_" & pInstallment)
                        FDueDate = rsF.fields("IDate_" & fInstallment)
                        If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), FDueDate) < 0 Then
                            If MinLapseFacDue = 0 Then
                                ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Payor Due on " & Format(PDueDate, "dd-mmm-yy") & " !"
                            Else
                                ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Date on " & Format(DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), "dd-mmm-yy") & " !"
                            End If
                            Return False
                        End If
                    Else
                        Installment = fInstallment
                        For i = 1 To Installment
                            'Loop for Each Installment
                            PDueDate = rsP.fields("IDate_" & i)
                            FDueDate = rsF.fields("IDate_" & i)
                            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), FDueDate) < 0 Then
                                If MinLapseFacDue = 0 Then
                                    ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Payor Due on " & Format(PDueDate, "dd-mmm-yy") & " !"
                                Else
                                    ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Date on " & Format(DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), "dd-mmm-yy") & " !"
                                End If
                                Return False
                            End If
                        Next
                    End If
                Else
                    ' Get Last Installment
                    PDueDate = rsP.fields("IDate_" & pInstallment)
                    FDueDate = DateAdd(DateInterval.Day, rsF.fields("Grace") + rsF.fields("WPC"), FacDueDate)

                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), FDueDate) < 0 Then
                        If MinLapseFacDue = 0 Then
                            ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Payor Due on " & Format(PDueDate, "dd-mmm-yy") & " !"
                        Else
                            ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Date on " & Format(DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), "dd-mmm-yy") & " !"
                        End If
                        Return False
                    End If
                End If
            Else
                If fInstallment <> 0 Then
                    PDueDate = DateAdd(DateInterval.Day, rsP.fields("Grace") + rsP.fields("WPC"), PolDueDate)
                    ' Get First Installment
                    FDueDate = rsF.fields("IDate_" & 1)
                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), FDueDate) < 0 Then
                        If MinLapseFacDue = 0 Then
                            ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Payor Due on " & Format(PDueDate, "dd-mmm-yy") & " !"
                        Else
                            ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Date on " & Format(DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), "dd-mmm-yy") & " !"
                        End If
                        Return False
                    End If
                Else
                    PDueDate = DateAdd(DateInterval.Day, rsP.fields("Grace") + rsP.fields("WPC"), PolDueDate)
                    FDueDate = DateAdd(DateInterval.Day, rsF.fields("Grace") + rsF.fields("WPC"), FacDueDate)

                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), FDueDate) < 0 Then
                        If MinLapseFacDue = 0 Then
                            ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Payor Due on " & Format(PDueDate, "dd-mmm-yy") & " !"
                        Else
                            ErrMsg = "Incorrect Payment Due To Reinsurer (Fac Slip No. " & rsF.fields("SlipNO") & ") on " & Format(FDueDate, "dd-mmm-yy") & ", must be greater than or equal to Date on " & Format(DateAdd(DateInterval.Day, MinLapseFacDue, PDueDate), "dd-mmm-yy") & " !"
                        End If
                        Return False
                    End If
                End If
            End If

            rsF.MoveNext()
        Loop
        Return True
    End Function

    Private Function Insert_Policy_Auto_Print(ByVal ANO As Long, Optional ByVal Version As String = "", Optional ByVal RepName As String = "") As Boolean
        Try
            Dim rs As RemoteSet
            Dim rsA As RemoteSet
            rsA = db.OpenRemoteSet("Select Top 1 * from APDBucket Where AppID=" & ANO & " and AppType='POLICY' and Status in ('R','P','T')")
            If rsA.Eof Then
                rs = db.OpenRemoteSet("Select * from Get_Policy_APRule(" & ANO & ") R Where R.Version Like '" & Version & "%'")
                Do While Not rs.Eof
                    rsA.AddNew()
                    rsA.fields("AppID") = ANO
                    rsA.fields("AppType") = "POLICY"   ' Policy
                    rsA.fields("RuleID") = rs.fields("RuleID")
                    rsA.fields("Class") = rs.fields("Class")
                    rsA.fields("RepName") = IIf(rs.fields("RepName") = "", RepName, rs.fields("RepName"))
                    rsA.fields("ExeName") = rs.fields("ExeName")
                    rsA.fields("Version") = IIf(Version = "", rs.fields("Version"), Version)
                    rsA.fields("eStampDutyF") = toval(rs.fields("ApplyeStampDutyF"))
                    rsA.fields("NeedEmailF") = toval(rs.fields("NeedEmailF"))
                    rsA.fields("NeedPrintF") = toval(rs.fields("NeedPrintF"))
                    rsA.fields("NeedeStampDutyF") = toval(rs.fields("NeedeStampDutyF"))
                    rsA.fields("PrintLogoF") = toval(rs.fields("PrintLogoF"))
                    rsA.fields("EMailTo") = tostr(rs.fields("EMailTo"))
                    rsA.fields("NeedeStampDutyF") = toval(rs.fields("NeedeStampDutyF"))
                    rsA.fields("ShowPremiumCalcF") = toval(rs.fields("ShowPremiumCalcF"))
                    rsA.fields("ShowDiscountF") = toval(rs.fields("ShowDiscountF"))
                    rsA.fields("NeedESignF") = toval(rs.fields("NeedESignF"))
                    rsA.fields("Last_Opr") = dbUser
                    rsA.fields("Last_Update") = Today.Date
                    rsA.fields("Status") = "R"
                    If Not rsA.UpdateRow Then Throw New Exception(db.ErrDesc)
                    rs.MoveNext()
                Loop
                If Not rsA.Eof Then
                    If Not rsA.Update Then Throw New Exception(db.ErrDesc)
                End If
                Insert_Policy_Auto_Print = True
            End If
        Catch ex As Exception
            StrMsg = ex.Message
            Insert_Policy_Auto_Print = False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function LoadClauseFromExcel(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam As Object(), ByRef OutParam As Object(), ByRef dsExcel As DataSet, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()


            ' Get Input Parameter
            Dim ANO As Long
            Dim Section As String
            Dim TOPRO As String
            Dim DefaultF As Boolean
            Dim Upload As String
            Dim SQL As String

            ANO = getInParamValue(InParam, 0)
            Section = getInParamValue(InParam, 1)
            TOPRO = getInParamValue(InParam, 2)
            DefaultF = getInParamValue(InParam, 3)
            Upload = getInParamValue(InParam, 4)    ' CLA/CON/EX

            If Upload = "" Then Upload = "CLA"
            Dim GenTable As String
            Dim AccTable As String
            Dim DefTable As String
            Dim TOPTable As String

            Select Case Upload
                Case "CLA"
                    GenTable = "GenCla"
                    AccTable = "AccCla"
                    DefTable = "DefCla"
                    TOPTable = "TOP_CLA"
                Case "CON"
                    GenTable = "GenCON"
                    AccTable = "AccCON"
                    DefTable = "DefCON"
                    TOPTable = "TOP_CON"
                Case "EX"
                    GenTable = "GenEx"
                    AccTable = "AccEx"
                    DefTable = "DefEx"
                    TOPTable = "TOP_Ex"
                Case Else
                    GenTable = "GenCla"
                    AccTable = "AccCla"
                    DefTable = "DefCla"
                    TOPTable = "TOP_CLA"
            End Select

            Dim OrderNo As String
            Dim ShortDesc As String
            Dim ShortDesc_1 As String
            Dim Description As String
            Dim Description_1 As String

            For i = 0 To dsExcel.Tables(0).Rows.Count - 1
                OrderNo = tostr(GetItemValue("OrderNo", dsExcel, i))
                ShortDesc = tostr(GetItemValue("ShortDesc", dsExcel, i))
                ShortDesc_1 = tostr(GetItemValue("ShortDesc_1", dsExcel, i))
                Description = tostr(GetItemValue("Description", dsExcel, i))
                Description_1 = tostr(GetItemValue("Description_1", dsExcel, i))

                If OrderNo <> "" Then
                    If Not DefaultF Then
                        If TOPRO = "" Then
                            SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, A2.Last_Opr, cast(coalesce(A2.Last_update,convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join " & GenTable & " G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join " & AccTable & " A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                            SQL = SQL & " Left Join " & DefTable & " D on (D.TOC=TOC.TOC and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " and (coalesce(A2.OrderNo, '')<>'' or G.AllowedF=1)"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        Else
                            SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, A2.Last_Opr, cast(coalesce(A2.Last_update,convert(varchar(10), getdate(),1)) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(D.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join " & GenTable & " G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join " & AccTable & " A2 on (A2.ANO=A1.ANO and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "') "
                            SQL = SQL & " Left Join " & TOPTable & " D on (D.TOPRO='" & TOPRO & "' and D.OrderNo=G.OrderNo and D.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " and (coalesce(A2.OrderNo, '')<>'' or G.AllowedF=1)"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        End If
                    Else
                        If TOPRO = "" Then
                            SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join " & GenTable & " G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join " & DefTable & " A2 on (A2.TOC=TOC.TOC and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " and (G.AllowedF=1)"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        Else
                            SQL = "select cast((Case When A2.OrderNo is null then 0 else 1 end) as bit) as ProcessFlag, '' as Last_Opr, cast(convert(varchar(10), getdate(),1) as smalldatetime) as Last_Update, coalesce(A2.OrderNo, G.OrderNo) as OrderNo, coalesce(A2.ShortDesc, G.ShortDesc) as ShortDesc, coalesce(A2.ShortDesc_1, G.ShortDesc_1) as ShortDesc_1, coalesce(A2.Description, G.Description) as Description, coalesce(A2.Description_1, G.Description_1) as Description_1, coalesce(A2.Section,'') as Section, cast(coalesce(A2.CompulsoryF,0) as Bit) as CompulsoryF, Cast(0 as bit) as DetailF, G.AllowedF"
                            SQL = SQL & " from Acceptance A1 "
                            SQL = SQL & " Inner Join Cover C on (C.CNO=A1.CNO)"
                            SQL = SQL & " Inner Join TOC on (TOC.TOC=A1.TOC)"
                            SQL = SQL & " Inner Join " & GenTable & " G on (G.COB=TOC.COB)"
                            SQL = SQL & " Left Join " & TOPTable & " A2 on (A2.TOPRO='" & TOPRO & "' and A2.OrderNo=G.OrderNo and A2.Section='" & Section & "')"
                            SQL = SQL & " Where A1.ANO=" & ANO & " and G.OrderNo='" & OrderNo & "'"
                            SQL = SQL & " and (G.AllowedF=1)"
                            SQL = SQL & " Order By G.StatPop Desc, G.OrderNo"
                        End If
                    End If
                    Dim rs As RemoteSet
                    rs = db.OpenRemoteSet(SQL)
                    Dim ds As DataSet = rs.GetDataSet
                    If Dataset Is Nothing Then
                        Dataset = ds.Clone
                    End If

                    If ds.Tables(0).Rows.Count <> 0 Then
                        ds.Tables(0).Rows(0)("ProcessFlag") = True
                        ds.Tables(0).Rows(0)("DetailF") = True

                        If ShortDesc <> "" Then
                            ds.Tables(0).Rows(0)("shortdesc") = ShortDesc
                        End If

                        If ShortDesc_1 <> "" Then
                            ds.Tables(0).Rows(0)("ShortDesc_1") = ShortDesc_1
                        End If

                        If Description <> "" Then
                            ds.Tables(0).Rows(0)("Description") = Description
                        End If

                        If Description_1 <> "" Then
                            ds.Tables(0).Rows(0)("Description_1") = Description_1
                        End If

                        ds.AcceptChanges()
                        Dataset.Merge(ds)
                    End If
                End If
            Next

            ' Set Output Parameter

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetPayorVA(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()


            If Not GetPayorDataVA(db, InParam, OutParam, ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If

            db.Committrans()

            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = db.ErrDesc
            db.RollBack()
            CloseConnection()
            If ErrMsg = "" Then ErrMsg = ex.Message
            Return False
        End Try
    End Function
    Public Function GetPayorDataVA(ByVal db As ClassLib, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            Try
                Me.db = db
                Me.dbUser = db.Var_db_user
                Me.dbUserType = db.Var_db_User_Type
                Me.dbName = db.Var_Db_Name
                Me.dbServer = db.Var_Db_Server
                Me.dbTrusted = db.Get_DbTrusted
            Catch ex As Exception
            End Try

            ' Get Input Parameter
            Dim ANO As String
            Dim Bank As String
            Dim TBName As String
            ANO = getInParamValue(InParam, 0)
            Bank = getInParamValue(InParam, 1)      ' Bank Code Validation
            TBName = getInParamValue(InParam, 2)     ' PAYOR or APAYOE
            If TBName = "" Then TBName = "Payor"

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select ANO, ID, RefType, Bank, AccountNo, BankName from " & TBName & " Where ANO=" & ANO)
            If rs.Eof Then
                Throw New Exception("No Payor Data !")
            ElseIf rs.fields("RefType") = "VIRTUAL" AndAlso rs.fields("Bank") <> "" AndAlso rs.fields("AccountNo") <> "" Then
                Throw New Exception("VA Account Already Assigned !")
            End If

            ' Get Virtual Account if, RefType not yet Virtual
            Dim SQL As String
            Dim rsV As RemoteSet
            Dim BankName As String
            Dim AccountNo As String
            Dim VID As Long

            SQL = "Select Top 1 P.ID, A.Branch, Pc.Currency, A.TOC, C.SEGMENT, P.CGroup, P.SCGroup"
            SQL = SQL & " from Acceptance A, Cover C, PCalc Pc, Profile P"
            SQL = SQL & " Where C.CNO=A.CNO and Pc.ANO=A.ANO and P.ID='" & rs.fields("ID") & "' and A.ANO=" & ANO
            rsV = db.OpenRemoteSet(SQL)
            If rsV.Eof Then
                Throw New Exception("No Premium Information !")
            End If

            Dim PID As String = rsV.fields("ID")
            Dim CGroup As String = rsV.fields("CGroup")
            Dim SCGroup As String = rsV.fields("SCGroup")
            Dim Segment As String = rsV.fields("Segment")
            Dim Branch As String = rsV.fields("Branch")
            Dim TOC As String = rsV.fields("TOC")
            Dim pCCY As String = rsV.fields("Currency")

            Dim rsH As RemoteSet

            SQL = "Select Top 1 V1.*"
            SQL = SQL & " from VAccount V1 "
            SQL = SQL & " Left Join VSCGroup V7 on (V7.VID=V1.VID and (V7.SCGroup='" & CGroup & "' or V7.SCGroup='" & SCGroup & "')) "
            SQL = SQL & " Inner Join VTOC V3 on (V3.VID=V1.VID and V3.TOC='" & TOC & "')"
            SQL = SQL & " Inner Join VSegment V5 on (V5.VID=V1.VID and V5.Segment='" & Segment & "')"
            SQL = SQL & " Inner Join VBranch V6 on (V6.VID=V1.VID and V6.Branch='" & Branch & "')"
            SQL = SQL & " Where (V1.ApplyVSCGroupF=0 or coalesce(V7.SCGroup,'')<>'') and (V1.XApplyVSCGroupF=0 or coalesce(V7.SCGroup,'')='') and (V1.Currency='' or V1.Currency='" & pCCY & "') "
            SQL = SQL & " and V1.AutoGeneratedF=1 and V1.RestrictedF=0 and V1.UsedByCollectionF=0"
            If Bank <> "" Then
                SQL = SQL & " and V1.Bank='" & Bank & "'"
            End If

            rsH = db.OpenRemoteSet(SQL)
            If rsH.Eof Then
                Throw New Exception("No VA Account for Collection Available (Company Group : " & CGroup & ", Sub Company Group : " & SCGroup & ", Sub Class : " & TOC & ", Segment : " & Segment & ", Branch : " & Branch & ", Currency : " & pCCY & ", Bank : " & Bank & ") !")
            End If

            VID = toval(rsH.fields("VID"))

            If Not db.GetLockedObject("VACCOUNT" & VID) Then Throw New Exception(db.ErrDesc)

            Bank = tostr(rsH.fields("Bank"))
            BankName = tostr(rsH.fields("BankName"))

            If tostr(rsH.fields("AutoNum")) <> "" Then
                AccountNo = db.GetNextDocNo(tostr(rsH.fields("AutoNum")), Today.Date, Branch)
                If AccountNo = "" Then
                    StrMsg = "Generate Virtual Account Number Failed !"
                    Throw New Exception(StrMsg)
                End If
                AccountNo = rsH.fields("BIN") & AccountNo
            Else
                AccountNo = rsH.fields("BIN") & Format(toval(rsH.fields("SeqNo")) + 1, rsH.fields("sformat"))
                If Not db.Execute("Update VAccount Set SeqNo=" & toval(rsH.fields("SeqNo")) + 1 & " Where VID='" & VID & "'") Then Throw New Exception(db.ErrDesc)
            End If

            If Not db.ReleaseLockedObject("VACCOUNT" & VID) Then Throw New Exception(db.ErrDesc)

            Dim rsD As RemoteSet
            rsD = db.OpenRemoteSet("Select * from VAccountD Where VID=" & VID & " and AccountNo='" & AccountNo & "'")
            If rsD.Eof Then
                rsD.AddNew()
            Else
                rsD.Edit()
            End If
            rsD.fields("VID") = VID
            rsD.fields("AccountNo") = AccountNo
            rsD.fields("AllocatedF") = True
            rsD.fields("Last_Update") = Today.Date
            rsD.fields("Last_Opr") = dbUser
            If Not rsD.Update Then Throw New Exception(rsD.ErrDesc)

            rs.Edit()
            rs.fields("RefType") = "VIRTUAL"
            rs.fields("Bank") = Bank
            rs.fields("BankName") = BankName
            rs.fields("AccountNo") = AccountNo
            If Not rs.Update Then Throw New Exception(rs.ErrDesc)

            ' Set Output Parameter
            ReDim OutParam(4)
            OutParam(0) = "VIRTUAL"
            OutParam(1) = Bank
            OutParam(2) = BankName
            OutParam(3) = AccountNo
            OutParam(4) = VID

            If LockedKey <> "" Then If Not db.ReleaseLockedObject(LockedKey) Then Throw New Exception(db.ErrDesc)

            GetPayorDataVA = True
        Catch ex As Exception
            ErrMsg = ex.Message
            GetPayorDataVA = False
        End Try
    End Function
    Private Function TrimBuffer(ByVal Bufferdata As Object, Optional ByVal CheckedColumn As Integer = 0) As Object
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim Rows As Integer
        Dim Cols As Integer
        Dim buffer As Object = Nothing
        If Not Bufferdata Is Nothing Then
            Rows = UBound(Bufferdata, 1) + 1
            Cols = UBound(Bufferdata, 2) + 1

            For i = Rows - 1 To 0 Step -1
                If Bufferdata(i, CheckedColumn) <> "" Then Exit For
            Next i
            ReDim buffer(i, Cols - 1)
            For j = 0 To i
                For k = 0 To Cols - 1
                    buffer(j, k) = Bufferdata(j, k)
                Next k
            Next j
        End If
        TrimBuffer = buffer
    End Function
    Private Function Save_Log_Object(ByVal LogName As String, ByVal LogKeyID As String, ByVal LogID As String, ByVal LogKeyIDExt1 As String, ByVal LogIDExt1 As String, ByVal LogKeyIDExt2 As String, ByVal LogIDExt2 As String, ByVal LogKeyIDExt3 As String, ByVal LogIDExt3 As String, ByVal ColName As String, ByVal ColValueOld As Object, ByVal ColValueNew As Object, ByVal Action As String) As Boolean
        Dim rs As RemoteSet
        On Error GoTo ErrMessage
        If LogName <> "" And (ColValueOld <> ColValueNew Or Action <> "Update") Then
            rs = db.OpenRemoteSet("Select * from SysLogObject Where 1<>1")
            rs.AddNew()
            rs.fields("LogName") = LogName
            rs.fields("LogKeyID") = Trim(tostr(LogKeyID, 50))
            rs.fields("LogID") = Trim(tostr(LogID, 50))
            rs.fields("LogKeyIDext1") = Trim(tostr(LogKeyIDExt1, 50))
            rs.fields("LogIDExt1") = Trim(tostr(LogIDExt1, 50))
            rs.fields("LogKeyIDExt2") = Trim(tostr(LogKeyIDExt2, 50))
            rs.fields("LogIDExt2") = Trim(tostr(LogIDExt2, 50))
            rs.fields("LogKeyIDExt3") = Trim(tostr(LogKeyIDExt3, 50))
            rs.fields("LogIDExt3") = Trim(tostr(LogIDExt3, 50))
            rs.fields("Last_Opr") = dbUser
            rs.fields("Last_Update") = Today.Date
            rs.fields("Last_Time") = Time()
            rs.fields("ColName") = ColName
            rs.fields("ColValueOLD") = Trim(tostr(ColValueOld, 255))
            rs.fields("ColValueNew") = Trim(tostr(ColValueNew, 255))
            rs.fields("Action") = Action
            If Not rs.Update Then GoTo ErrMessage
        End If
        Save_Log_Object = True
        Exit Function
ErrMessage:
        Save_Log_Object = False
    End Function

    Private Function GetPickUpList(ByVal Param2 As String, ByVal Param3 As Long, ByVal UserID As String) As String
        Dim FName As String = ""
        Try
            If Param2 <> "" Then
                Dim rs As RemoteSet
                Dim ASource As String
                Dim Status As String
                Dim AType As String
                Dim Role As String
                rs = db.OpenRemoteSet("Select Role from SysUser Where ID = '" & UserID & "'")
                Role = tostr(rs.fields("Role"))
                Select Case Param2
                    Case "UND", "UND-MW", "IT", "MAR"
                        rs = db.OpenRemoteSet("Select Cover.ASource, Acceptance.AStatus, Acceptance.AType from Acceptance, Cover Where Acceptance.CNO=Cover.CNO and Acceptance.ANO=" & Param3)
                        ASource = rs.fields("ASource")
                        Status = rs.fields("AStatus")
                        AType = rs.fields("AType")

                        If Param2 <> "IT" Then
                            Select Case Role
                                Case "ASSESSOR", "SURVEYOR"
                                    FName = GetAppSetting("CareListPolicy")
                                Case "COMMON", "OPERATION", "MARKETING OFFICER"
                                    If ASource = "MKT" And AType = "N" Then
                                        FName = GetAppSetting("CareListQuotation")
                                    Else
                                        FName = GetAppSetting("CareListPolicy")
                                    End If
                                Case "UNDERWRITER"
                                    FName = GetAppSetting("CareListPolicy")
                                Case Else
                                    If ASource = "MKT" And AType = "N" Then
                                        FName = GetAppSetting("CareListQuotation")
                                    Else
                                        FName = GetAppSetting("CareListPolicy")
                                    End If
                            End Select
                        Else
                            FName = GetAppSetting("CareListIT")
                        End If
                    Case "SVC"
                        rs = db.OpenRemoteSet("Select Service.ASource, Service.SStatus from Service Where Service.SNO=" & Param3)
                        ASource = rs.fields("ASource")
                        Status = rs.fields("SStatus")

                        If Status = "R" Then
                            FName = GetAppSetting("CareListServicing")
                        ElseIf UCase(ASource) = "CLM" Then
                            FName = GetAppSetting("CareListClaim2")
                        Else
                            Select Case Role
                                Case "ASSESSOR", "SURVEYOR"
                                    FName = GetAppSetting("CareListClaim2")
                                Case "MARKETING OFFICER"
                                    If ASource = "MKT" Then
                                        FName = GetAppSetting("CareListQuotation2")
                                    Else
                                        FName = GetAppSetting("CareListPolicy2")
                                    End If
                                Case "COMMON", "UNDERWRITER", "OPERATION"
                                    FName = GetAppSetting("CareListPolicy2")
                                Case Else
                                    If ASource = "MKT" Then
                                        FName = GetAppSetting("CareListQuotation2")
                                    Else
                                        FName = GetAppSetting("CareListPolicy2")
                                    End If
                            End Select
                        End If
                    Case "CLM"
                        Select Case Role
                            Case "ASSESSOR", "SURVEYOR", "COMMON", "OPERATION"
                                FName = GetAppSetting("CareListClaim")
                            Case "MARKETING OFFICER", "UNDERWRITER"
                                FName = GetAppSetting("CareListPolicy")
                            Case Else
                                FName = GetAppSetting("CareListClaim")
                        End Select
                    Case "CSCLM"
                        FName = GetAppSetting("CareListCSClaim")
                    Case "FA"
                        FName = GetAppSetting("CareListVoucher")
                    Case "RQ"
                        FName = GetAppSetting("CareListRequest")
                    Case "POLH"
                        FName = GetAppSetting("CHISSListPolicy")
                    Case "CLMH"
                        FName = GetAppSetting("CHISSListClaim")
                    Case "GLH"
                        FName = GetAppSetting("CHISSListGL")
                End Select
                rs.CloseRemoteSet()
            End If
        Catch ex As Exception
            FName = ""
        End Try
        Return FName
    End Function
    Private Sub ParsingObjectID(ByVal ObjectID As String, ByRef str1 As String, ByRef str2 As String, ByRef str3 As String)
        Try
            Dim str As String = ObjectID
            str1 = Left(str, InStr(str, "-") - 1)
            str = Mid(str, InStr(str, "-") + 1, Len(str) - (InStr(str, "-")))
            str2 = Left(str, InStr(str, "-") - 1)
            str3 = Mid(str, InStr(str, "-") + 1, Len(str) - (InStr(str, "-")))
        Catch ex As Exception

        End Try
    End Sub
    Private Function IPCRequest(ByVal ProcName As String, ByVal Remarks As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As Long, ByVal Param4 As Long, ByVal Param5 As String, ByVal Param6 As String, ByVal UserID As String, ByVal UserDate As Date, ByVal UserTime As String, ByVal ActionCode As String, ByVal ApprovedFName As String, ByVal DueDate As Date, ByVal PID As Long, ByVal RequestingUser As String, ByRef Errmsg As String, ByVal EmailF As Boolean, ByVal EmailSubject As String) As Boolean
        IPCRequest = db.Execute("Insert into SysRequest (ProcName, Remarks, Param1, Param2, Param3, Param4, Param5, Param6, UserID, UserDate, UserTime,ActionCode,ApprovedFName, DueDate, PID, RequestingUser, Status) Values ('" & ProcName & "','" & Remarks & "','" & Param1 & "','" & Param2 & "'," & Param3 & "," & Param4 & ",'" & Param5 & "','" & Param6 & "','" & UserID & "','" & Format(UserDate, "MM/DD/YYYY") & "','" & UserTime & "','" & ActionCode & "','" & ApprovedFName & "','" & Format(DueDate, "MM/DD/YYYY") & "'," & PID & ",'" & RequestingUser & "','')")
        Select Case Param2
            Case "CLM", "CLM-MW"
                CLNotes(toval(Param4), "CLAIM", "Message From  " & UserID & " To " & Param5 & vbCrLf & Remarks)
            Case "UND", "UND-MW", "IT", "MAR"
                PNotes(toval(Param3), "POLICY", "Message From  " & UserID & " To " & Param5 & vbCrLf & Remarks)
            Case "UND-OBJ"
                Dim str1 As String = ""
                Dim str2 As String = ""
                Dim str3 As String = ""
                ParsingObjectID(Param6, str1, str2, str3)
                PNotes(toval(str1), "OBJECT REGISTER", "Message From  " & UserID & " To " & Param5 & vbCrLf & Remarks, str2, str3)
            Case "CLM-EXP"
                Dim str1 As String = ""
                Dim str2 As String = ""
                Dim str3 As String = ""
                ParsingObjectID(Param6, str1, str2, str3)
                CLNotes(toval(str1), "CLAIM EXPENSE", "Message From  " & UserID & " To " & Param5 & vbCrLf & Remarks)
        End Select
        Errmsg = db.ErrDesc
        If Errmsg = "" AndAlso Param5 <> "" Then
            If EmailF Then
                Try
                    Dim rs As RemoteSet
                    Dim EmailFrom As String = UserID
                    Dim EmailTo As String = Param5
                    Dim EmailCC() As String = Nothing
                    If ActionCode <> "" Then
                        rs = db.OpenRemoteSet("Select A1.ActionName from ActionList A1 Where A1.ActionCode='" & ActionCode & "'")
                        EmailF = EmailF OrElse toval(rs.fields("EmailF"))
                        EmailSubject = tostr(rs.fields("ActionName")) & "-" & EmailSubject

                        rs = db.OpenRemoteSet("Select * from ActionUser Where ActionCode='" & ActionCode & "' and ActionUser='" & Param5 & "'")
                        Dim i As Integer = 0
                        Dim j As Integer = 0
                        For i = 1 To 10
                            If rs.fields("ActionUserAuth_" & i) <> "" Then
                                ReDim Preserve EmailCC(j)
                                EmailCC(j) = rs.fields("ActionUserAuth_" & i)
                                j = j + 1
                            Else
                                Exit For
                            End If
                        Next
                    End If

                    rs = db.OpenRemoteSet("Select TOP 1 ID from SysRequest Where ProcName='" & ProcName & "' and Param1='" & Param1 & "' and Param2='" & Param2 & "' and Param3=" & Param3 & " and Param4=" & Param4 & " and Param5='" & Param5 & "' and Param6='" & Param6 & "' and UserID='" & UserID & "' and UserDate='" & Format(UserDate, "MM/DD/YYYY") & "' and UserTime='" & UserTime & "' and PID=" & PID & "")

                    EmailSubject = "Pending Task : " & EmailSubject

                    Dim EmailBody As String = EmailSubject & vbCrLf & Remarks

                    Dim FName As String = GetPickUpList(Param2, Param3, Param5)
                    If FName <> "" Then
                        FName = FName & "&Mode=PickUp&Param=" & tostr(rs.fields("ID"))
                        EmailBody = EmailBody & vbCrLf & vbCrLf & "Follow This URL : " & FName
                    End If
                    SendMail(EmailFrom, EmailTo, EmailCC, EmailSubject, HtmlEncode(EmailBody), Errmsg)

                Catch ex As Exception

                End Try
            End If
        End If

    End Function

    Private Function CLNotes(ByVal CNO As Long, ByVal Category As String, ByVal Notes As String) As Boolean
        Dim db As New ClassLib
        Try
            If Not db.OpenConnection("", dbName, dbServer, dbSysUser, dbSysPassword, dbTimeOut, dbTrusted) Then
                Throw New Exception(db.ErrDesc)
            End If
            db.Var_db_user = dbUser
            db.Var_db_Password = dbPassword
            db.Var_db_User_Type = dbUserType
            db.Begintrans()
            CLNotes = db.Execute("Insert into CLNOtes(CNO,Date,Time, UserID, Category,Notes, SystemF) Select Claim.OCNO,'" & Format(Today.Date, "MM/DD/YYYY") & "','" & Time() & "','" & dbUser & "','" & Category & "','" & tostr(Notes) & "'," & 1 & " From Claim Where CNO=" & CNO)

            db.Committrans()
            db.CloseSession()
            Return True
        Catch ex As Exception
            db.RollBack()
            db.CloseSession()
            Return False
        End Try
    End Function
    Private Function PNotes(ByVal ANO As Long, ByVal Category As String, ByVal Notes As String, Optional ByVal FLDID As String = "", Optional ByVal VALUEID As String = "") As Boolean
        Dim db As New ClassLib
        Try
            If Not db.OpenConnection("", dbName, dbServer, dbSysUser, dbSysPassword, dbTimeOut, dbTrusted) Then
                Throw New Exception(db.ErrDesc)
            End If
            db.Var_db_user = dbUser
            db.Var_db_Password = dbPassword
            db.Var_db_User_Type = dbUserType
            db.Begintrans()
            PNotes = db.Execute("Insert into PNotes(ANO,Date,Time, UserID, Category,Notes, SystemF, FLDID, VALUEID) Select Acceptance.OANO,'" & Format(Today.Date, "MM/DD/YYYY") & "','" & Time() & "','" & dbUser & "','" & Category & "','" & tostr(Notes) & "'," & 1 & ",'" & FLDID & "','" & VALUEID & "' from Acceptance Where Acceptance.ANO=" & ANO)

            db.Committrans()
            db.CloseSession()
            Return True
        Catch ex As Exception
            db.RollBack()
            db.CloseSession()
            Return False
        End Try
    End Function

    Private Function StrError(Optional ByVal errmsg As String = "") As String
        Dim str As String
        Dim ErrDesc As String
        ErrDesc = Err.Description
        On Error GoTo errmsg
        str = errmsg
        If str = "" Then str = ErrDesc
        If str = "" Then str = db.ErrDesc
        StrError = str
        Exit Function
errmsg:
        StrError = "Connection is Broken !"
    End Function
    Private Function GetDocumentNumbering(ByVal AutoNum As String, ByVal PolicyType As String, ByVal ISType As String, ByVal PDate As Date, ByVal Branch As String, ByVal TOC As String, ByVal CT As String, Optional ByVal ApplyPrefixInfixF As Boolean = False) As String
        Dim rs As RemoteSet
        Dim Prefix As String
        Dim InFix As String
        Dim Separator As String
        Dim BranchDigit As Integer
        Dim TOCDigit As Integer
        Dim CTDigit As Integer
        Branch = Get_Branch_Alias_By_TOC(Branch, TOC, TOC)
        If PolicyType <> "" Then
            rs = db.OpenRemoteSet("Select Top 1 P.Prefix, P.InFix, P.Separator, P.BranchDigit, P.TOCDigit, P.CTDigit, P.CheckDigitF, I.Prefix as IPrefix, I.InFix as IInFix, S.ApplyAllPrefixInfixF, S.ApplyAllTOCDigitF from ISType I Left Join PolicyType P On (P.PolicyType='" & PolicyType & "') Inner Join SysTab S on (1=1) Where I.ISType='" & ISType & "' Order by P.PolicyType")
        Else
            rs = db.OpenRemoteSet("Select Top 1 P.Prefix, P.InFix, P.Separator, P.BranchDigit, P.TOCDigit, P.CTDigit, P.CheckDigitF, I.Prefix as IPrefix, I.InFix as IInFix, S.ApplyAllPrefixInfixF, S.ApplyAllTOCDigitF from ISType I Left Join PolicyType P On (P.SType='S' and ((P.ISType='A') or (P.ISType='D' and I.ISType in ('I','L','C')) or (P.ISType='I' and I.ISType in ('R','T')))) Inner Join SysTab S on (1=1) Where I.ISType='" & ISType & "' Order by P.PolicyType")
        End If
        If Not rs.Eof Then
            If toval(rs.fields("ApplyAllPrefixInfixF")) Or ApplyPrefixInfixF Then
                Prefix = IIf(tostr(rs.fields("Prefix")) <> "", tostr(rs.fields("Prefix")), tostr(rs.fields("IPrefix")))
                InFix = IIf(tostr(rs.fields("InFix")) <> "", tostr(rs.fields("InFix")), tostr(rs.fields("IInFix")))
            Else
                Prefix = ""
                InFix = ""
            End If
            Separator = tostr(rs.fields("Separator"))
            TOCDigit = toval(rs.fields("TOCDigit"))
            BranchDigit = toval(rs.fields("BranchDigit"))
            CTDigit = toval(rs.fields("CTDigit"))
            If toval(rs.fields("ApplyAllTOCDigitF")) Then
                GetDocumentNumbering = db.GetNextDocNo(AutoNum, PDate, IIf(Prefix <> "", Prefix & Separator, "") & IIf(BranchDigit <> 0, Mid(Branch, 1, BranchDigit) & Separator, "") & IIf(CTDigit <> 0, Mid(CT, 1, CTDigit) & Separator, "") & IIf(InFix <> "", InFix & Separator, "") & IIf(TOCDigit <> 0, Mid(TOC, 1, TOCDigit) & Separator, ""))
            Else
                GetDocumentNumbering = db.GetNextDocNo(AutoNum, PDate, IIf(Prefix <> "", Prefix & Separator, "") & IIf(BranchDigit <> 0, Mid(Branch, 1, BranchDigit) & Separator, "") & IIf(CTDigit <> 0, Mid(CT, 1, CTDigit) & Separator, "") & IIf(InFix <> "", InFix & Separator, "") & TOC & Separator)
            End If
        Else
            GetDocumentNumbering = db.GetNextDocNo(AutoNum, PDate, IIf(Prefix <> "", Prefix & Separator, "") & Branch & Separator & IIf(InFix <> "", InFix & Separator, "") & TOC & Separator)
        End If
    End Function
    Private Function Get_Branch_Alias_By_TOC(ByVal Branch As String, ByVal TOC As String, ByRef TAlias As String, Optional ByRef POL_AUTONUM As String = "") As String
        Dim rs As RemoteSet
        Dim BAlias As String

        rs = db.OpenRemoteSet("Select BAlias, TAlias, POL_AUTONUM from Branch_TOC Where (Branch='' or Branch='" & Branch & "') and (TOC='' or TOC='" & TOC & "')")
        BAlias = tostr(rs.fields("BAlias"))
        TAlias = tostr(rs.fields("TAlias"))
        POL_AUTONUM = tostr(rs.fields("POL_AUTONUM"))

        If BAlias = "" Then BAlias = Branch
        If TAlias = "" Then TAlias = TOC
        If POL_AUTONUM = "" Then POL_AUTONUM = "POL"
        rs.CloseRemoteSet()
        Get_Branch_Alias_By_TOC = BAlias
    End Function
    Private ReadOnly Property ErrDesc() As String
        Get
            ErrDesc = StrMsg
        End Get
    End Property
    Private Function SendMail(ByVal UserFrom As String, ByVal UserTo As String, ByVal UserCC As Object, ByVal EMailSubject As String, ByVal EMailBody As String, ByRef ErrMsg As String) As Boolean

        Try
            Dim rs As RemoteSet
            Dim SMTPServer As String = ""
            Dim SMTPPort As Integer = 0
            Dim SMTPEmail As String = ""
            Dim EmailFrom As String = ""
            Dim EmailTo As String = ""

            rs = db.OpenRemoteSet("Select SMTPServer, SMTPPort, SMTPEmail from SysTab")
            SMTPServer = rs.fields("SMTPServer")
            SMTPPort = toval(rs.fields("SMTPPort"))
            SMTPEmail = rs.fields("SMTPEmail")


            If SMTPEmail <> "" Then
                rs = db.OpenRemoteSet("Select EMail from SysUser Where ID='" & UserFrom & "'")
                EmailFrom = rs.fields("EMail")
            End If
            If EmailFrom = "" Then EmailFrom = SMTPEmail

            Dim msgMail As New System.Net.Mail.MailMessage()
            Dim smtSender As SmtpClient = New SmtpClient()

            If (SMTPServer = "" AndAlso smtSender.Host = "") Or (SMTPPort = 0 AndAlso toval(smtSender.Port) = 0) Or (EmailFrom = "" AndAlso IsNothing(msgMail.From)) Then
                Throw New Exception("Please Check SMTP Configuration in Parameter/Application Configuration Setting !")
            End If

            If EmailFrom <> "" AndAlso IsNothing(msgMail.From) Then
                msgMail.From = New System.Net.Mail.MailAddress(EmailFrom)
            Else
                If Not IsNothing(msgMail.From) Then
                    EmailFrom = msgMail.From.Address
                End If
            End If

            With smtSender
                .DeliveryMethod = SmtpDeliveryMethod.Network
                If .Host = "" Then
                    .Host = SMTPServer
                End If
                If .Port = 0 Then
                    .Port = SMTPPort
                End If
                If .EnableSsl = False AndAlso CBool(toval(GetAppSetting("SMTPEnableSSL"))) Then
                    .EnableSsl = CBool(toval(GetAppSetting("SMTPEnableSSL")))
                End If

                If .EnableSsl And GetAppSetting("EmailSecurityProtocolType") <> "" Then
                    ServicePointManager.Expect100Continue = True
                    ServicePointManager.DefaultConnectionLimit = 9999
                    ServicePointManager.SecurityProtocol = toval(GetAppSetting("EmailSecurityProtocolType"))
                End If
            End With

            If smtSender.Host <> "" AndAlso EmailFrom <> "" Then

                rs = db.OpenRemoteSet("Select S.Email from fn_UserSendMail('" & UserTo & "') S")
                Do While Not rs.Eof
                    If rs.fields("EMail") <> "" AndAlso ValidateEmail(rs.fields("EMail")) Then
                        EmailTo = rs.fields("EMail")
                        msgMail.To.Add(New System.Net.Mail.MailAddress(EmailTo))
                    End If
                    rs.MoveNext()
                Loop

                If TypeName(UserCC).ToUpper = "STRING" AndAlso UserCC <> "" Then
                    rs = db.OpenRemoteSet("Select S.Email from fn_UserSendMail('" & UserCC & "') S")
                    Do While Not rs.Eof
                        If rs.fields("EMail") <> "" AndAlso ValidateEmail(rs.fields("EMail")) Then
                            EmailTo = rs.fields("EMail")
                            msgMail.CC.Add(New System.Net.Mail.MailAddress(EmailTo))
                        End If
                        rs.MoveNext()
                    Loop
                ElseIf TypeName(UserCC).ToUpper = "STRING()" AndAlso UBound(UserCC) >= 0 Then
                    For i = 0 To UBound(UserCC)
                        rs = db.OpenRemoteSet("Select S.Email from fn_UserSendMail('" & UserCC(i) & "') S")
                        Do While Not rs.Eof
                            If rs.fields("EMail") <> "" AndAlso ValidateEmail(rs.fields("EMail")) Then
                                EmailTo = rs.fields("EMail")
                                msgMail.CC.Add(New System.Net.Mail.MailAddress(EmailTo))
                            End If
                            rs.MoveNext()
                        Loop
                    Next
                End If

                If EmailTo <> "" Then
                    msgMail.Subject = EMailSubject
                    msgMail.IsBodyHtml = True
                    msgMail.BodyEncoding = UTF8Encoding.UTF8
                    Try
                        msgMail.Body = EMailBody
                        AddHandler smtSender.SendCompleted, AddressOf SendCompletedCallback
                        smtSender.SendAsync(msgMail, "")
                    Catch ex As Exception
                        ErrMsg = ex.Message
                        Return False
                    End Try
                End If
            End If
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    Private Sub SendCompletedCallback(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs)
        Dim msgMail As MailMessage = TryCast(e.UserState, MailMessage)
        If msgMail IsNot Nothing Then msgMail.Dispose()
    End Sub

    Private Function EncryptByPassPhrase(ByVal PassPhrase As String, ByVal PlainString As Object) As Object
        Try
            If PassPhrase <> "" Then
                Dim rs As RemoteSet
                Dim SQL As String = ""

                If TypeName(PlainString) = "Date" Then
                    SQL = "Select dbo.fn_ENCRYPTDATE('" & PassPhrase & "','" & Format(PlainString, "MM/DD/YYYY") & "')"
                Else
                    SQL = "Select dbo.fn_ENCRYPTTEXT('" & PassPhrase & "','" & tostr(PlainString) & "')"
                End If

                rs = db.OpenRemoteSet(SQL)
                EncryptByPassPhrase = rs.fields(0)
            Else
                EncryptByPassPhrase = PlainString
            End If
        Catch ex As Exception
            EncryptByPassPhrase = ""
        End Try
    End Function
    Private Function DecryptByPassPhrase(ByVal PassPhrase As String, ByVal EncyrptedString As String, ByVal EncryptedType As Integer) As Object
        Try
            If PassPhrase <> "" Then
                Dim rs As RemoteSet
                Dim SQL As String = ""

                If EncryptedType = 1 Then   ' String
                    SQL = "Select dbo.fn_DECRYPTTEXT('" & PassPhrase & "','" & EncyrptedString & "')"
                Else
                    SQL = "Select dbo.fn_DECRYPTDATE('" & PassPhrase & "','" & EncyrptedString & "')"
                End If

                rs = db.OpenRemoteSet(SQL)
                DecryptByPassPhrase = rs.fields(0)
            Else
                DecryptByPassPhrase = EncyrptedString
            End If
        Catch ex As Exception
            DecryptByPassPhrase = ""
        End Try

    End Function
    Function ValidateEmail(ByVal emailAddress As String) As Boolean
        Dim pattern As String = "^[a-zA-Z][\w\.-_\-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-_\-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            Return True
        End If
    End Function

    <WebMethod(EnableSession:=False)>
    Public Function VerifyCLPmt(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            Dim i As Long
            Dim NumVerify As Long = 0
            Dim skipPolicyCheckF As Boolean = toval(getInParamValue(InParam, 1))
            db.Begintrans()
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                If (Dataset.Tables(0).Rows(i)("ProcessFlag") = True) Then
                    Try
                        Dim rs As RemoteSet
                        Dim PolicyNo As String
                        Dim LetterDate As Date
                        Dim ConfirmationDate As Date
                        Dim HoldF As Boolean

                        StrMsg = ""

                        PolicyNo = tostr(GetItemValue("PolicyNo", Dataset, i))
                        LetterDate = todate(GetItemValue("LetterDate", Dataset, i))
                        ConfirmationDate = todate(GetItemValue("ConfirmationDate", Dataset, i))
                        HoldF = toval(GetItemValue("HoldF", Dataset, i))

                        rs = db.OpenRemoteSet("Select OANO from Acceptance Where PolicyNo='" & PolicyNo & "' and ATYPE Not In ('A','C') ")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Please Check Policy No. " & PolicyNo & " !" & vbCrLf
                        End If

                        rs = db.OpenRemoteSet("Select * from ClaimPaymentSchedule Where ANO=" & rs.fields("OANO"))
                        If rs.Eof Then
                            StrMsg = StrMsg & "No Claim Payment Schedule" & vbCrLf
                        End If
                        If StrMsg = "" Then
                            Dataset.Tables(0).Rows(i)("Status") = "Verified"
                        Else
                            Dataset.Tables(0).Rows(i)("Status") = "failed "
                        End If

                    Catch ex As Exception
                        StrMsg = StrMsg & ex.Message
                        Dataset.Tables(0).Rows(i)("Status") = "failed "
                    End Try
                    Dataset.Tables(0).Rows(i)("Message") = StrMsg
                    NumVerify = NumVerify + 1
                End If
            Next


            ReDim OutParam(1)
            OutParam(0) = NumVerify
            OutParam(1) = ""

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()

            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)>
    Public Function UploadCLPmt(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            Dim ProcessOpt As Integer = InParam(0)
            Dim i As Long
            Dim numProcess As Long = 0

            db.Begintrans()

            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                StrMsg = ""
                If (Dataset.Tables(0).Rows(i)("ProcessFlag") = True) And (Dataset.Tables(0).Rows(i)("Status") = "Verified" Or ProcessOpt <> 0) Then
                    If Dataset.Tables(0).Rows(i)("Status") <> "Verified" And ProcessOpt = 2 Then
                        ErrMsg = "Record No. " & i + 1 & " is not Verified Yet !"
                        Throw New Exception(ErrMsg)
                    End If
                    Try
                        Dim rs As RemoteSet
                        Dim PolicyNo As String
                        Dim LetterDate As Date
                        Dim ConfirmationDate As Date
                        Dim HoldF As Boolean

                        StrMsg = ""

                        PolicyNo = tostr(GetItemValue("PolicyNo", Dataset, i))
                        LetterDate = todate(GetItemValue("LetterDate", Dataset, i))
                        ConfirmationDate = todate(GetItemValue("ConfirmationDate", Dataset, i))
                        HoldF = toval(GetItemValue("HoldF", Dataset, i))

                        rs = db.OpenRemoteSet("Select OANO from Acceptance Where PolicyNo='" & PolicyNo & "'")
                        If rs.Eof Then
                            Throw New Exception("Please Check CI No. " & PolicyNo & " !")
                        End If

                        rs = db.OpenRemoteSet("Select * from ClaimPaymentSchedule Where ANO=" & rs.fields("OANO"))
                        If Not rs.Eof Then
                            rs.Edit()
                            rs.fields("LetterDate") = LetterDate
                            rs.fields("ConfirmationDate") = ConfirmationDate
                            rs.fields("HoldF") = HoldF
                            If Not rs.Update Then Throw New Exception(db.ErrDesc)
                        End If

                        If StrMsg = "" Then
                            Dataset.Tables(0).Rows(i)("Status") = "Processed"
                        Else
                            Dataset.Tables(0).Rows(i)("Status") = "failed "
                        End If
                    Catch ex As Exception
                        StrMsg = ex.Message
                        Dataset.Tables(0).Rows(i)("Status") = "failed "
                    End Try
                    Dataset.Tables(0).Rows(i)("Message") = StrMsg
                    numProcess = numProcess + 1
                End If

            Next

            ReDim OutParam(1)
            OutParam(0) = numProcess

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()

            Return False
        End Try
    End Function
   
    <WebMethod(EnableSession:=False)>
    Public Function GetBusinessSourceAccountExecutive(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim BSType As String


            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("SELECT ANO, ID,BSType, PICName, PICPhone, PICMobile, PICEmail FROM ACCBS WHERE ANO=" & ANO & " AND ID='" & ID & "' and BSType='" & BSType & "'")

            findex = rs.fields_index()
            fbuffer = rs.GetRecord()
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = db.ErrDesc
            db.RollBack()
            CloseConnection()
            If ErrMsg = "" Then ErrMsg = ex.Message
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)>
    Public Function GetAlternateBusinessSourceAccountExecutive(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef findex() As String, ByRef fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim BSType As String


            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)
            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("SELECT ANO, ID,BSType, PICName, PICPhone, PICMobile, PICEmail FROM ACCBS WHERE ANO=" & ANO & " AND ID='" & ID & "' and BSType='" & BSType & "'")

            findex = rs.fields_index()
            fbuffer = rs.GetRecord()
            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = db.ErrDesc
            db.RollBack()
            CloseConnection()
            If ErrMsg = "" Then ErrMsg = ex.Message
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)>
    Public Function SaveBusinessSourceAccountExecutive(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim BSType As String


            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("SELECT ANO, ID,BSType, PICName, PICPhone, PICMobile, PICEmail FROM ACCBS WHERE ANO=" & ANO & " AND ID='" & ID & "' and BSType='" & BSType & "'")
            If rs.Eof Then
                Throw New Exception("Business Source Data Not Found !")
            End If

            rs.Edit()
            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = db.ErrDesc
            db.RollBack()
            CloseConnection()
            If ErrMsg = "" Then ErrMsg = ex.Message
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)>
    Public Function SaveAlternateBusinessSourceAccountExecutive(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            ' Get Input Parameter

            Dim ANO As Long
            Dim ID As String
            Dim BSType As String

            ANO = getInParamValue(InParam, 0)
            ID = getInParamValue(InParam, 1)
            BSType = getInParamValue(InParam, 2)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("SELECT ANO, ID,BSType, PICName, PICPhone, PICMobile, PICEmail FROM  WHERE ANO=" & ANO & " AND ID='" & ID & "' and BSType='" & BSType & "'")
            If rs.Eof Then
                Throw New Exception("Business Source Data Not Found !")
            End If

            rs.Edit()
            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)>
    Public Function UpdateUnderwritingInfo(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByVal findex() As String, ByVal fbuffer() As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim ANO As Long
            Dim CNO As Long

            ANO = getInParamValue(InParam, 0)

            Dim LockedKey As String = ""
            If ANO > 0 Then
                LockedKey = ANO
                If Not db.GetLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)
            End If

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select A.AStatus, A.CNO from Acceptance A where A.ANO=" & ANO)
            If rs.Eof Then
                ErrMsg = "Policy Data Not Found !"
                Throw New Exception(ErrMsg)
            End If
            CNO = rs.fields("CNO")

            rs = db.OpenRemoteSet("Select * from UINFO Where CNO=" & CNO)
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If

            Dim i As Integer
            Dim j As Integer
            i = 0
            j = UBound(findex)
            Do While i < j
                rs.fields(findex(i)) = GetFieldValue(fbuffer(i))
                i = i + 1
            Loop
            rs.fields("CNO") = CNO
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ' Set Output Parameter   
            If LockedKey <> "" Then If Not db.ReleaseLockedObject("Policy" & LockedKey) Then Throw New Exception(db.ErrDesc)

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            db.RollBack()
            CloseConnection()
            If ErrMsg = "" Then ErrMsg = ex.Message
            Return False
        End Try
    End Function
    Private Function GetItemValue(ByVal FName As String, ByVal Dataset As DataSet, ByVal i As Integer) As Object
        Try
            GetItemValue = Dataset.Tables(0).Rows(i)(FName)
        Catch ex As Exception
            GetItemValue = ""
        End Try
    End Function
    Private Function Direct_Call_Process_Execute(ByVal ProcLib As String, ByVal ProcName As String, Optional ByVal Param1 As Object = Nothing, Optional ByVal Param2 As Object = Nothing, Optional ByVal Param3 As Object = Nothing, Optional ByVal Param4 As Object = Nothing, Optional ByVal Param5 As Object = Nothing, Optional ByVal Param6 As Object = Nothing, Optional ByVal NewSessionFlag As Boolean = False) As Object
        Dim PObj As Object = Nothing
        Try
            StrMsg = ""
            Select Case UCase(ProcLib)
                Case "PROCESSFA.NETV4.PROCLIB", "PROCESSFA.NETV3.PROCLIB"
                    PObj = New ProcessFA.NetV4.Proclib
                Case "PROCESSFAEXT.NETV4.PROCLIB", "PROCESSFAEXT.NETV3.PROCLIB"
                    PObj = New ProcessFAExt.NetV4.Proclib
                Case "PROCESS.NETV4.PROCLIB", "PROCESS.NETV3.PROCLIB"
                    PObj = New Process.NetV4.Proclib
                Case "PROCESSEXT.NETV4.PROCLIB", "PROCESSEXT.NETV3.PROCLIB"
                    PObj = New ProcessExt.NetV4.Proclib
                Case "PROCESSCLEXT.NETV4.PROCLIB", "PROCESSCLEXT.NETV3.PROCLIB"
                    PObj = New ProcessCLExt.NetV4.Proclib
                Case "PROCESSUL.NETV4.PROCLIB", "PROCESSUL.NETV3.PROCLIB"
                    PObj = New ProcessUL.NetV4.Proclib
                Case "PROCESSHK.NETV4.PROCLIB", "PROCESSHK.NETV3.PROCLIB"
                    PObj = New ProcessHK.NetV4.Proclib
                Case "PROCESSSC.NETV4.PROCLIB", "PROCESSSC.NETV3.PROCLIB"
                    PObj = New ProcessSC.NetV4.Proclib
            End Select

            PObj.Var_db_User = db.Var_db_user
            PObj.Var_db_User_Type = db.Var_db_User_Type
            PObj.Var_db_User_CCode = db.Var_db_User_CCode
            PObj.Var_Db_Name = dbName
            PObj.Var_Db_Server = dbServer
            PObj.Var_db_Password = dbPassword
            PObj.Var_db_SysUser = dbSysUser
            PObj.Var_db_SysPassword = dbSysPassword
            PObj.Var_db_TimeOut = dbTimeOut
            PObj.Var_db_Trusted = dbTrusted
            PObj.Var_ObjSession = db.Var_ObjSession

            If NewSessionFlag Then
                Direct_Call_Process_Execute = PObj.Call_Process("", ProcName, Param1, Param2, Param3, Param4, Param5, Param6)
            Else
                Direct_Call_Process_Execute = PObj.Call_Process_Execute("", ProcName, Param1, Param2, Param3, Param4, Param5, Param6, NewSessionFlag)
            End If
            StrMsg = PObj.ErrDesc
            PObj = Nothing
        Catch ex As Exception
            StrMsg = ex.Message
            PObj = Nothing
            Return False
        End Try
    End Function
    Private Function Direct_Call_Ext_Process(ByVal ProcLib As String, ByVal ProcName As String, Optional ByVal Param1 As Object = Nothing, Optional ByVal Param2 As Object = Nothing, Optional ByVal Param3 As Object = Nothing, Optional ByVal Param4 As Object = Nothing, Optional ByVal Param5 As Object = Nothing, Optional ByVal Param6 As Object = Nothing) As Object
        Direct_Call_Ext_Process = Direct_Call_Process_Execute(ProcLib, ProcName, Param1, Param2, Param3, Param4, Param5, Param6, True)
    End Function
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class