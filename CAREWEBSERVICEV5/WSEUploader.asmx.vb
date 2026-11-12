Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports CareWebLibrary
Imports CareWebLibrary.CursorTypeEnum
Imports CareWebLibrary.EditTypeEnum
Imports CareWebLibrary.DataTypeEnum
Imports System.Threading
Imports System.Globalization
Imports System.Net.Mail
Imports System.IO
Imports System.Math
Imports System.Net
Imports System.Web.Script.Services
Imports System.Web.Script.Serialization
Imports Newtonsoft.Json
Imports System.Xml.Serialization

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class WSEUploader
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
    Shadows Culture As CultureInfo
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
            ErrMsg = db.ErrDesc
            CloseConnection()
            If ErrMsg = "" Then ErrMsg = ex.Message
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
    <WebMethod(EnableSession:=False)> _
    Public Function Check_Excel_column(ByVal dbUser As String, ByVal dbPassword As String, ByVal inparam() As Object, ByVal dataset As DataSet, ByRef outparam As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If


            Return True

        Catch Ex As Exception
            db.RollBack()
            CloseConnection()
            ErrMsg = Ex.Message
            Return False
        End Try
    End Function
    Private Function Data_Check_From(ByVal dbuser As String, ByVal dbpassword As String, ByVal inparam() As Object, ByVal dataset As DataSet, ByRef outparam As Object, ByRef ErrMsg As String) As Boolean
        Dim rs, rst As RemoteSet
        Dim TBName As String = ""
        Dim Sql As String = ""
        Dim i As Integer = 0
        Dim Field As String = ""
        Dim FieldName As String = ""
        Dim ColumnName As String = "'"

        Dim FileTitle As String = ""
        Dim MapCode As String = ""
        Try

            FileTitle = inparam(0)
            MapCode = inparam(1)

            For i = 0 To dataset.Tables(0).Columns.Count - 1
                If dataset.Tables(0).Columns.Item(i).DataType.ToString = "System.DateTime" Then
                    Field = Field & IIf(i = 0, "", ",") & " cast('' as SmallDateTime) as [" & dataset.Tables(0).Columns.Item(i).ToString & "]"
                ElseIf dataset.Tables(0).Columns.Item(i).DataType.ToString = "System.Double" Then
                    Field = Field & IIf(i = 0, "", ",") & " cast('' as int) as [" & dataset.Tables(0).Columns.Item(i).ToString & "]"
                Else
                    Field = Field & IIf(i = 0, "", ",") & " '' as [" & dataset.Tables(0).Columns.Item(i).ToString & "]"
                End If
            Next i

            Sql = "Select " & Field
            rst = db.OpenRemoteSet(Sql)

            Sql = "select * from map_field where fieldfrom not like '%dummy%' and mapcode = '" & FileTitle & "' ORDER BY ORDERNO"
            rs = db.OpenRemoteSet(Sql, adOpenDynamic)

            Data_Check_From = False
            StrMsg = ""
            rs.MoveFirst()
            Do While Not rs.Eof
                FieldName = rs.fields("FieldFrom")
                ColumnName = rst.fields_Name(FieldName)
                If ColumnName <> "" Then
                    If rs.fields("FieldType") = "D" And (rst.fields_Type(ColumnName) <> adDate And rst.fields_Type(ColumnName) <> adDBDate And UCase(tostr(rs.fields("CompulsoryF"))) = "TRUE") Then
                        StrMsg = StrMsg & "Fields " & rs.fields("FieldFrom") & " Does not use Date Format!" & Chr(13)
                    End If
                Else
                    StrMsg = StrMsg & "Fields " & rs.fields("FieldFrom") & " Does not Exist in Excel!" & Chr(13)
                End If
                rs.MoveNext()
            Loop

            If StrMsg = "" Then
                rs.CloseRemoteSet()
                rst.CloseRemoteSet()
                Return True
            Else
                Throw New Exception(StrMsg)
            End If
            rs.CloseRemoteSet()
            rst.CloseRemoteSet()

            Exit Function
        Catch Ex As Exception
            ErrMsg = StrMsg
            If ErrMsg = "" Then
                ErrMsg = Ex.Message
            End If
            Return False
        End Try
    End Function
    Private Function Data_Check_To(ByVal dbuser As String, ByVal dbpassword As String, ByVal inparam() As Object, ByVal dataset As DataSet, ByRef outparam As Object, ByRef ErrMsg As String) As Boolean
        Try
            Dim rs As RemoteSet
            Dim rsup As RemoteSet
            Dim rsupExt As RemoteSet
            Dim rsupExt_1 As RemoteSet
            Dim FileTitle As String = ""
            Dim MapGroup As String = ""
            Dim strMsg As String = ""
            Dim SysBatchExtension As Integer = toval(GetAppSetting("SysBatchExtension"))
            If SysBatchExtension = 0 Then SysBatchExtension = 1
            Dim SQL As String

            FileTitle = inparam(0)
            MapGroup = inparam(1)

            SQL = "select * from SysBatch" & MapGroup & " where 1<>1"
            rsup = db.OpenRemoteSet(SQL)

            If SysBatchExtension >= 1 Then
                SQL = "select * from SysBatch" & MapGroup & "_Ext where 1<>1"
                rsupExt = db.OpenRemoteSet(SQL)
            End If

            If SysBatchExtension >= 2 Then
                SQL = "select * from SysBatch" & MapGroup & "_Ext_1 where 1<>1"
                rsupExt_1 = db.OpenRemoteSet(SQL)
            End If

            SQL = "select * from map_field where mapcode = '" & FileTitle & "' ORDER BY ORDERNO"
            rs = db.OpenRemoteSet(SQL)

            Dim DataCheckF As Boolean = False

            Do While Not rs.Eof
                For i = 0 To rsup.fields_Count - 1
                    If UCase(rs.fields("FieldTo")) = UCase(rsup.fields_Name(i)) Then
                        DataCheckF = True
                        Exit For
                    End If
                    DataCheckF = False
                Next

                If SysBatchExtension >= 1 Then
                    If Not DataCheckF Then
                        For i = 0 To rsupExt.fields_Count - 1
                            If UCase(rs.fields("FieldTo")) = UCase(rsupExt.fields_Name(i)) Then
                                DataCheckF = True
                                Exit For
                            End If
                            DataCheckF = False
                        Next
                    End If
                End If

                If SysBatchExtension >= 2 Then
                    If Not DataCheckF Then
                        For i = 0 To rsupExt_1.fields_Count - 1
                            If UCase(rs.fields("FieldTo")) = UCase(rsupExt_1.fields_Name(i)) Then
                                DataCheckF = True
                                Exit For
                            End If
                            DataCheckF = False
                        Next
                    End If
                End If

                If DataCheckF = False Then
                    StrMsg = StrMsg & "Fields " & rs.fields("FieldFrom") & " Does not Exist in SysBatch" & MapGroup & " !" & vbCrLf
                End If
                rs.MoveNext()
            Loop
            rs.CloseRemoteSet()

            If StrMsg <> "" Then
                Throw New Exception(StrMsg)
            End If
            Return True
        Catch Ex As Exception
            ErrMsg = Ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function Upload_Excel(ByVal dbuser As String, ByVal dbpassword As String, ByRef inparam() As Object, ByVal dataset As DataSet, ByRef outparam() As Object, ByRef ErrMsg As String) As Boolean
        Dim message As String = ""
        Try
            If Not Data_Upload(dbuser, dbpassword, inparam, dataset, outparam, message) Then
                Throw New Exception(message)
            End If
            ErrMsg = message

            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function GetUploadDataSetFormat(ByRef dbuser As String, ByRef dbpassword As String, ByRef inparam() As Object, ByRef dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet = Nothing
        Dim Sql As String = ""
        Dim FileTitle As String = ""
        Dim MapGroup As String
        Dim Batchno As String = ""
        Try
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            FileTitle = inparam(0)
            MapGroup = inparam(1)
            Batchno = inparam(2)

            Sql = "select * from SysBatchOriginal" & MapGroup & " where 1<>1"
            rs = db.OpenRemoteSet(Sql)
            dataset = rs.GetDataSet

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message & vbCrLf & "Get Data Set Format Failed"
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function
    Function getInParamValue(ByVal InParam As Object, ByVal index As Integer) As Object
        If index <= InParam.GetUpperBound(0) And index >= InParam.GetLowerBound(0) Then
            Return InParam(index)
        Else
            Return ""
        End If
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function Data_Query(ByRef dbuser As String, ByRef dbpassword As String, ByRef inparam() As Object, ByRef outparam As Object, ByRef DataSet As DataSet, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet = Nothing
        Dim BatchNo As String = ""
        Dim RefNo As String = ""
        Dim RefNoSQL As String = ""
        Dim MapGroup As String = ""
        Dim SQL As String = ""
        Dim StatusH As String = ""
        Dim StatusUpload As String = ""
        Dim ErrorMsg As String = ""

        Try
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If

            BatchNo = tostr(inparam(0))
            RefNo = tostr(inparam(1))
            RefNoSQL = tostr(inparam(2))

            rs = db.OpenRemoteSet("Select M.MapGroup, S.Status, S.ErrMsg From SysBatchH S, MAP_HEADER M Where S.MapCode = M.MapCode And S.BatchNo = '" & BatchNo & "'")
            If Not rs.Eof Then
                MapGroup = tostr(rs.fields("MapGroup"))
                StatusH = tostr(rs.fields("Status"))
                ErrorMsg = tostr(rs.fields("ErrMsg"))
            End If

            SQL = "Select coalesce(S2.Status, S1.Status) as StatusCode, coalesce(S2.ErrMsg, S1.ErrMsg) as ErrDesc, "
            If MapGroup <> "SC" Then
                SQL = SQL & " S2.RefNo as RefNo,"
            Else
                SQL = SQL & " S1.ContractNo as RefNo,"
            End If
            If MapGroup <> "BC" Then
                SQL = SQL & " S1.Inception,"
                SQL = SQL & " S1.Expiry,"
            End If
            SQL = SQL & " Case When coalesce(A.Astatus,'') not in ('W','T') Then cast(1 as bit) else cast(0 as bit) end as AlreadyInforcedF, A1.VOUCHER, A.POLICYNO, A.CERTIFICATENO, A.ANO, A.TOC, "
            SQL = SQL & " coalesce(A1.Currency, P.Currency) as Currency, "

            SQL = SQL & " coalesce(A1.AMOUNT_1,Pc.Premium) as Amount_1,  "
            SQL = SQL & " coalesce(A2.AMOUNT_2,-Pc.Commission) as Amount_2,  "
            SQL = SQL & " coalesce(A1.AMOUNT_3,Pc.FeeDuty) as Amount_3,  "
            SQL = SQL & " coalesce(A1.AMOUNT_4,0) as Amount_4,  "
            SQL = SQL & " coalesce(A1.AMOUNT_5,-Pc.Discount)+coalesce(A2.AMOUNT_5,0) as Amount_5,  "
            SQL = SQL & " coalesce(A2.Amount_13,-Pc.VAT) as Amount_13,  "
            SQL = SQL & " coalesce(A2.Amount_14,Pc.Tax) as Amount_14,  "
            SQL = SQL & " coalesce(A1.Amount_20,0) as Amount_20 "
            If RefNoSQL <> "" Then
                SQL = SQL & ", " & RefNoSQL
            End If

            SQL = SQL & " from SysBatchOriginal" & MapGroup & " S1 "
            SQL = SQL & " Left Join SysBatch" & MapGroup & " S2 on (S2.BID=S1.BID) "

            If MapGroup <> "SC" Then
                If MapGroup = "BC" Then
                    SQL = SQL & " Left Join Acceptance A on (A.AType='C' and A.BatchNo = S2.BatchNo and ((A.RefNo=S2.RefNo) and (A.BID=S2.BID))) "
                ElseIf MapGroup = "BE" Then
                    SQL = SQL & " Left Join Acceptance A on (A.AType IN('E','C') and A.BatchNo = S2.BatchNo and ((A.BID=S2.BID)))   "
                Else ' MapGroup='UP'
                    SQL = SQL & " Left Join Acceptance A on (A.AType='N' and A.BatchNo = S2.BatchNo and ((A.RefNo=S2.RefNo) and (A.BID=S2.BID))) "
                End If
            Else
                SQL = SQL & " Left Join Scheme_transaction S3 on (S3.BatchNo=S2.BatchNo and S3.ContractNo=S2.ContractNo) "
                SQL = SQL & " Left Join Acceptance A on (A.ANO=S3.ANO) "
            End If
            SQL = SQL & " Left Join Cover C on (C.CNO=A.CNO) "
            SQL = SQL & " Left Join PCalc P on (P.ANO=A.ANO)"
            SQL = SQL & " Outer Apply (Select Top 1 AccAss.AdmNo from AccAss Where AccAss.ANO=A.ANO and AccAss.Type='DI' and AccAss.Code='P' and AccAss.Currency=P.Currency) AccAss"
            SQL = SQL & " Outer Apply (Select A1.Voucher, A1.DocNo, A1.RefNo, A1.Currency, A1.Date, A1.BillDate, A1.Subject, A1.Branch, A1.Segment, A1.A_PolicyNo, A1.Amount_1, A1.Amount_3, A1.Amount_4, A1.Amount_5, A1.Amount_20 from Admlink A1 Where A1.AdmNo=AccAss.AdmNo) A1"
            SQL = SQL & " Outer Apply (Select A2.Amount_2, A2.Amount_5, A2.Amount_13, A2.Amount_14 from Admlink A2 Where A2.DocNo=A1.DocNo and A2.RefNo=A1.RefNo and A2.Currency=A1.Currency and A2.Date=A1.Date and A2.BillDate=A1.BillDate and A2.Subject=A1.Subject and A2.Branch=A1.Branch and A2.Segment=A1.Segment and A2.A_PolicyNo=A1.A_PolicyNo and A2.Type in ('BB','BA','BO','BC','BP','BD') and A2.ID=A.Source) A2"
            SQL = SQL & " Outer Apply (Select Top 1 Premium, Commission, FeeDuty, Discount, VAT, TAX from PREMIUM_CALCULATION(P.ANO) D Where D.Currency=P.Currency) Pc"
            SQL = SQL & " Where S1.BatchNo = '" & BatchNo & "' "
            If RefNo <> "" Then
                If MapGroup <> "SC" Then
                    SQL = SQL & " and S1.RefNo='" & RefNo & "'"
                Else
                    SQL = SQL & " and S1.ContractNo='" & RefNo & "'"
                End If
            End If

            rs = db.OpenRemoteSet(SQL, adOpenKeyset)
            DataSet = rs.GetDataSet

            If StatusH <> "C" Then
                StatusUpload = "Process Upload Fail."
            Else
                StatusUpload = "Process Upload Success."
            End If

            ' Set Output Parameter
            ReDim outparam(2)
            outparam(0) = StatusUpload
            outparam(1) = ErrorMsg
            outparam(2) = StatusH

            CloseConnection()
            Return True
        Catch ex As Exception
            CloseConnection()
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function Data_Upload(ByRef dbuser As String, ByRef dbpassword As String, ByRef inparam() As Object, ByVal dataset As DataSet, ByRef outparam As Object, ByRef ErrMsg As String) As Boolean
        Dim rsto As RemoteSet = Nothing
        Dim rstoExt As RemoteSet = Nothing
        Dim rstoExt_1 As RemoteSet = Nothing
        Dim rs As RemoteSet = Nothing
        Dim FileTitle As String = ""
        Dim Sql As String = ""
        Dim Batchno As String = ""
        Dim StrMsg As String
        Dim FieldTo() As String = Nothing
        Dim FieldFrom() As String = Nothing
        Dim VariableT() As String = Nothing
        Dim DefValue() As String = Nothing
        Dim DefObjInfo(14) As Object
        Dim blankrecord As Long = 0
        Dim TotalUpload As Long = 0
        Dim TotalColumn As Long = 0
        Dim i As Integer = 0
        Dim PartialF As Boolean
        Dim MapGroup As String
        Dim ImmediateF As Boolean
        Dim UploadF As Boolean
        Dim tPrefix As String
        Dim tBatchNo As String = ""
        Dim SysBatchExtension As Integer = toval(GetAppSetting("SysBatchExtension"))
        If SysBatchExtension = 0 Then SysBatchExtension = 1
        Try
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()

            FileTitle = getInParamValue(inparam, 0)
            MapGroup = getInParamValue(inparam, 1)
            Batchno = getInParamValue(inparam, 2)
            ImmediateF = toval(getInParamValue(inparam, 3))
            tPrefix = tostr(getInParamValue(inparam, 4))
            UploadF = toval(getInParamValue(inparam, 5))

            If dataset.Tables(0).Rows.Count = 0 Then
                Throw New Exception("No Data Upload !")
            End If

            Sql = "Select PartialF from MAP_HEADER where MapCode = '" & FileTitle & "'"
            rs = db.OpenRemoteSet(Sql)
            If Not rs.Eof Then PartialF = toval(rs.fields("PartialF"))

            Sql = "Select * from map_field where mapcode = '" & FileTitle & "' order by orderno "
            rs = db.OpenRemoteSet(Sql)
            rs.MoveFirst()
            ReDim FieldTo(rs.RecordCount)
            ReDim FieldFrom(rs.RecordCount)
            ReDim VariableT(rs.RecordCount)
            ReDim DefValue(rs.RecordCount)
            TotalColumn = 0
            Do While Not rs.Eof
                FieldTo(TotalColumn) = tostr(rs.fields("FieldTo"))
                FieldFrom(TotalColumn) = Trim(tostr(rs.fields("FieldFrom")))
                VariableT(TotalColumn) = tostr(rs.fields("FieldType"))
                DefValue(TotalColumn) = tostr(rs.fields("DefaultValue"))
                TotalColumn = TotalColumn + 1
                rs.MoveNext()
            Loop

            Sql = "select * into #SysBatch from " & tPrefix & "SysBatchOriginal" & MapGroup & " where 1<>1"
            If Not db.Execute(Sql) Then Throw New Exception(db.ErrDesc)
            Sql = "select * from #SysBatch where 1<>1"
            rsto = db.OpenRemoteSet(Sql)

            If SysBatchExtension >= 1 Then
                Sql = "select * into #SysBatch_Ext from " & tPrefix & "SysBatchOriginal" & MapGroup & "_Ext where 1<>1"
                If Not db.Execute(Sql) Then Throw New Exception(db.ErrDesc)
                Sql = "select * from #SysBatch_Ext where 1<>1"
                rstoExt = db.OpenRemoteSet(Sql)
            End If

            If SysBatchExtension >= 2 Then
                Sql = "select * into #SysBatch_Ext_1 from " & tPrefix & "SysBatchOriginal" & MapGroup & "_Ext_1 where 1<>1"
                If Not db.Execute(Sql) Then Throw New Exception(db.ErrDesc)
                Sql = "select * from #SysBatch_Ext_1 where 1<>1"
                rstoExt_1 = db.OpenRemoteSet(Sql)
            End If

            Dim ListBatchNo As New ListDictionary
            For i = 0 To dataset.Tables(0).Rows.Count - 1
                rsto.AddNew()
                If SysBatchExtension >= 1 Then rstoExt.AddNew()
                If SysBatchExtension >= 2 Then rstoExt_1.AddNew()

                StrMsg = ""
                blankrecord = 0
                For Y = 0 To TotalColumn - 1
                    Dim Value As String

                    If InStr(UCase(FieldFrom(Y)), "DUMMY") < 1 Then
                        Select Case TypeName(dataset.Tables(0).Rows(i)(FieldFrom(Y)))
                            Case "Date"
                                Value = Format(dataset.Tables(0).Rows(i)(FieldFrom(Y)), "MM/DD/YYYY")
                            Case "Double", "Single"
                                Value = tovalSF(dataset.Tables(0).Rows(i)(FieldFrom(Y)))
                            Case Else
                                Value = tostr(dataset.Tables(0).Rows(i)(FieldFrom(Y)))
                        End Select
                        If Value = "" Then
                            blankrecord = blankrecord + 1
                        End If
                    Else
                        Value = DefValue(Y)
                    End If
                    If rsto.fields_Name(FieldTo(Y)) <> "" Then
                        rsto.fields(FieldTo(Y)) = Value
                    ElseIf SysBatchExtension >= 1 AndAlso rstoExt.fields_Name(FieldTo(Y)) <> "" Then
                        rstoExt.fields(FieldTo(Y)) = Value
                    ElseIf SysBatchExtension >= 2 AndAlso rstoExt_1.fields_Name(FieldTo(Y)) <> "" Then
                        rstoExt_1.fields(FieldTo(Y)) = Value
                    End If

                Next

                tBatchNo = Batchno
                If tBatchNo = "" Then
                    tBatchNo = tostr(GetItemValue("BatchNo", dataset, i))
                End If
                If tBatchNo = "" Then
                    tBatchNo = rsto.fields("Batchno")
                End If
                If tBatchNo = "" Then
                    If Batchno = "" Then Batchno = Format(db.GetIdentity(tPrefix & "BatchNo") + 1, "0000000000")
                    tBatchNo = Batchno
                End If
                If Not ListBatchNo.Contains(tBatchNo) Then
                    ListBatchNo.Add(tBatchNo, 0)
                End If

                rsto.fields("Batchno") = tBatchNo
                rsto.fields("TRDate") = Today.Date
                rsto.fields("Status") = "W"
                rsto.fields("errmsg") = ""
                If blankrecord >= TotalColumn Then
                    If Not rsto.DeleteRow() Then Throw New Exception(rsto.ErrDesc)
                    If SysBatchExtension >= 1 Then If Not rstoExt.DeleteRow() Then Throw New Exception(rstoExt.ErrDesc)
                    If SysBatchExtension >= 2 Then If Not rstoExt_1.DeleteRow() Then Throw New Exception(rstoExt_1.ErrDesc)
                Else
                    Dim BID As Long
                    BID = db.GetIdentity("SysBatch" & MapGroup) + 1
                    rsto.fields("BID") = BID
                    If Not rsto.UpdateRow Then Throw New Exception(rsto.ErrDesc)

                    If SysBatchExtension >= 1 Then
                        rstoExt.fields("BID") = BID
                        If Not rstoExt.UpdateRow Then Throw New Exception(rstoExt.ErrDesc)
                    End If

                    If SysBatchExtension >= 2 Then
                        rstoExt_1.fields("BID") = BID
                        If Not rstoExt_1.UpdateRow Then Throw New Exception(rstoExt_1.ErrDesc)
                    End If

                    TotalUpload = TotalUpload + 1
                    ListBatchNo(tBatchNo) = toval(ListBatchNo(tBatchNo)) + 1
                End If
            Next i

            If Not rsto.Eof Then
                If Not rsto.Update Then Throw New Exception(rsto.ErrDesc)
            End If

            If SysBatchExtension >= 1 Then
                If Not rstoExt.Eof Then
                    If Not rstoExt.Update Then Throw New Exception(rstoExt.ErrDesc)
                End If
            End If

            If SysBatchExtension >= 2 Then
                If Not rstoExt_1.Eof Then
                    If Not rstoExt_1.Update Then Throw New Exception(rstoExt_1.ErrDesc)
                End If
            End If

            Dim dBatchNo As DictionaryEntry
            For Each dBatchNo In ListBatchNo
                tBatchNo = dBatchNo.Key
                If Not db.GetLockedObject(tPrefix & "SysBatch" & tBatchNo) Then Throw New Exception(db.ErrDesc)

                rs = db.OpenRemoteSet("Select * from " & tPrefix & "SysBatchH where batchno='" & tBatchNo & "'")
                If rs.Eof Then
                    rs.AddNew()
                    rs.fields("Completed") = 0
                    rs.fields("NotCompleted") = 0
                    rs.fields("TotalRecord") = 0
                Else
                    rs.Edit()
                End If
                rs.fields("TotalUpload") = toval(rs.fields("TotalUpload")) + toval(ListBatchNo(tBatchNo))
                rs.fields("BatchNo") = tBatchNo
                rs.fields("MapCode") = FileTitle
                rs.fields("UploadF") = UploadF
                rs.fields("Proclib") = "ProcessBatch.NetV3.ProcLib"
                rs.fields("ProcName") = "Verify_" & MapGroup
                rs.fields("Param1") = MapGroup
                rs.fields("Param2") = dbuser
                rs.fields("Param3") = FileTitle
                rs.fields("Param4") = tBatchNo
                rs.fields("Param5") = PartialF
                rs.fields("db_User") = UCase(IIf(dbuser = "", dbuser, tostr(dbuser)))
                rs.fields("db_password") = Encrypt(tostr(dbpassword), CCode)
                rs.fields("PDate") = Today.Date
                rs.fields("PTime") = Time()
                rs.fields("ErrMsg") = ""

                Sql = ""
                For i = 0 To rsto.fields_Count - 1
                    If Sql = "" Then
                        Sql = "[" & rsto.fields_Name(i) & "]"
                    Else
                        Sql = Sql & "," & "[" & rsto.fields_Name(i) & "]"
                    End If
                Next
                Sql = "Insert into " & tPrefix & "SysBatchOriginal" & MapGroup & "(" & Sql & ") Select " & Sql & " from #SysBatch Where BatchNo='" & tBatchNo & "' order by BID"
                If Not db.Execute(Sql) Then Throw New Exception(db.ErrDesc)

                If SysBatchExtension >= 1 Then
                    Sql = ""
                    For i = 0 To rstoExt.fields_Count - 1
                        If Sql = "" Then
                            Sql = "S1.[" & rstoExt.fields_Name(i) & "]"
                        Else
                            Sql = Sql & "," & "S1.[" & rstoExt.fields_Name(i) & "]"
                        End If
                    Next
                    Sql = "Insert into " & tPrefix & "SysBatchOriginal" & MapGroup & "_Ext (" & Sql & ") Select " & Sql & " from #SysBatch_Ext S1, #SysBatch S2 Where S1.BID=S2.BID and S2.BatchNo='" & tBatchNo & "' order by S2.BID"
                    If Not db.Execute(Sql) Then Throw New Exception(db.ErrDesc)
                End If

                If SysBatchExtension >= 2 Then
                    Sql = ""
                    For i = 0 To rstoExt_1.fields_Count - 1
                        If Sql = "" Then
                            Sql = "S1.[" & rstoExt_1.fields_Name(i) & "]"
                        Else
                            Sql = Sql & "," & "S1.[" & rstoExt_1.fields_Name(i) & "]"
                        End If
                    Next
                    Sql = "Insert into " & tPrefix & "SysBatchOriginal" & MapGroup & "_Ext_1 (" & Sql & ") Select " & Sql & " from #SysBatch_Ext_1 S1, #SysBatch S2 Where S1.BID=S2.BID and S2.BatchNo='" & tBatchNo & "' order by S2.BID"
                    If Not db.Execute(Sql) Then Throw New Exception(db.ErrDesc)
                End If

                If ImmediateF Or tPrefix <> "" Then
                    rs.fields("Status") = "N"
                    rs.fields("Priority") = 10
                    If Not rs.Update Then Throw New Exception(rs.ErrDesc)
                Else
                    rs.fields("Status") = "R"
                    rs.fields("Priority") = 10
                    If Not rs.Update Then Throw New Exception(rs.ErrDesc)
                End If
                If Not db.ReleaseLockedObject(tPrefix & "SysBatch" & tBatchNo) Then Throw New Exception(db.ErrDesc)
            Next

            If Not db.Execute("Drop Table #SysBatch") Then Throw New Exception(db.ErrDesc)
            db.Committrans()

            ' Set Output Parameter
            ReDim outparam(4)
            outparam(0) = tBatchNo
            outparam(1) = TotalUpload
            outparam(2) = ""
            outparam(3) = ""
            outparam(4) = ""

            If ImmediateF Or tPrefix <> "" Then
                Try
                    Dim strBatchNo As String = ""
                    For Each dBatchNo In ListBatchNo
                        tBatchNo = dBatchNo.Key
                        If strBatchNo = "" Then
                            strBatchNo = "'" & tBatchNo & "'"
                        Else
                            strBatchNo = strBatchNo & ",'" & tBatchNo & "'"
                        End If

                        Try
                            If Not VerifyExecute(tBatchNo, ErrMsg, tPrefix) Then Throw New Exception(ErrMsg)
                            If Not UploadF AndAlso Not BatchExecute(tBatchNo, ErrMsg, tPrefix) Then Throw New Exception(ErrMsg)
                        Catch ex As Exception

                        End Try
                    Next

                    Dim PolicyNo As String
                    Dim CertificateNo As String
                    Dim dsResult As DataSet

                    Sql = "Select S1.BatchNo, coalesce(Admlink.Voucher,'') as Voucher, coalesce(A.PolicyNo, coalesce(S3.PolicyNo, S2.PolicyNo)) as PolicyNo, coalesce(A.CertificateNo, coalesce(S3.CertificateNo, S2.CertificateNo)) as CertificateNo, coalesce(S3.Status, S2.Status) as Status,  coalesce(S3.ErrMsg, S2.ErrMsg) as ErrMsg"
                    Sql = Sql & " from sysbatchH S1"
                    Sql = Sql & " Inner Join " & tPrefix & "sysbatchOriginal" & MapGroup & " S2 on (S2.BatchNo=S1.BatchNo)"
                    Sql = Sql & " Left Join " & tPrefix & "sysbatch" & MapGroup & " S3 on (S3.BID=S2.BID)"
                    If MapGroup <> "SC" Then
                        If MapGroup = "BC" Then
                            Sql = Sql & " Left Join " & tPrefix & "Acceptance A on (A.BID=S3.BID) "
                        ElseIf MapGroup = "BE" Then
                            Sql = Sql & " Left Join " & tPrefix & "Acceptance A on (A.BID=S3.BID) "
                        Else ' MapGroup='UP'
                            Sql = Sql & " Left Join " & tPrefix & "Acceptance A on (A.BID=S3.BID) "
                        End If
                    Else
                        Sql = Sql & " Left  Join " & tPrefix & "Scheme_transaction S4 on (S4.BID=S3.BID) "
                        Sql = Sql & " Left  Join " & tPrefix & "Acceptance A on (A.ANO=S4.ANO) "
                    End If
                    Sql = Sql & " Outer Apply (Select Top 1 Admlink.Voucher from AccAss, Admlink Where Admlink.AdmNo=AccAss.AdmNo and AccAss.Type in ('DI','IC','IR','IT','IX') and AccAss.Code='P' and AccAss.ANO=A.ANO) Admlink"
                    Sql = Sql & " Where S1.BatchNo in (" & strBatchNo & ")"

                    rs = db.OpenRemoteSet(Sql)
                    dsResult = rs.GetDataSet
                    PolicyNo = tostr(rs.fields("PolicyNo"))
                    CertificateNo = tostr(rs.fields("CertificateNo"))

                    outparam(2) = PolicyNo
                    outparam(3) = CertificateNo

                    ' to write
                    Dim serializer As New XmlSerializer(GetType(DataSet))
                    Dim Writer As New StringWriter
                    serializer.Serialize(Writer, dsResult)
                    outparam(4) = Writer.GetStringBuilder.ToString

                    ' to read
                    'Dim serializer As New XmlSerializer(GetType(DataSet))
                    'Dim fs As New StringReader(outparam(4))
                    'dsResult = CType(serializer.Deserialize(fs), DataSet)




                Catch ex As Exception

                End Try
            End If

            CloseConnection()
            Data_Upload = True
        Catch ex As Exception
            ErrMsg = ex.Message & vbCrLf & "Process Upload Excel Failed"
            db.RollBack()

            CloseConnection()
            Data_Upload = False
        Finally
        End Try
    End Function
    Private Function VerifyExecute(ByVal BatchNo As String, ByRef strmsg As String, Optional ByVal tPrefix As String = "") As Boolean

        Dim db_user As String
        Dim db_password As String
        Dim db_user_type As String
        Dim db_user_ccode As String
        StrMsg = ""
        Try
            Dim ProcLib As String
            Dim ProcName As String
            Dim MainProcName As String
            Dim MapGroup As String
            Dim Param(5) As Object
            Dim fIndex() As String
            Dim fBuffer() As String
            Dim FCount As Integer
            Dim fIndex_1() As String
            Dim FCount_1 As Integer
            Dim rsBatchH As RemoteSet
            Dim rsBatchD As RemoteSet
            Dim rAffected As Integer = 0
            Dim PartialF As Boolean
            Dim SysBatchExtension As Integer = toval(GetAppSetting("SysBatchExtension"))
            If SysBatchExtension = 0 Then SysBatchExtension = 1

            Dim SQL As String

            db.Begintrans()
            SQL = "Select top 1 S.* from " & tPrefix & "sysBatchH S with (ReadCommitted)"
            SQL = SQL & " Where S.status = 'N' "
            SQL = SQL & " and S.BatchNo='" & BatchNo & "'"
            rsBatchH = db.OpenRemoteSet(SQL)
            If rsBatchH.Eof Then
                db.Committrans()
                rsBatchH.CloseRemoteSet()
                VerifyExecute = True
                Exit Function
            End If

            If Not db.GetLockedObject(tPrefix & "SysBatch" & BatchNo) Then Throw New Exception(db.ErrDesc)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select MapGroup from MAP_Header Where MapCode='" & rsBatchH.fields("MapCode") & "'")
            MapGroup = tostr(rs.fields("MapGroup"))
            If MapGroup = "" Then MapGroup = "D"

            Dim UploadF As Boolean
            ProcLib = tostr(rsBatchH.fields("ProcLib"))
            ProcName = tostr(rsBatchH.fields("ProcName"))
            UploadF = toval(rsBatchH.fields("UploadF"))
            MainProcName = tostr(rsBatchH.fields("MainProcName"))
            If MainProcName = "" Then MainProcName = ProcName
            MainProcName = "Verify_" & MapGroup
            Param(0) = GetParamValue(tostr(rsBatchH.fields("Param1")), toval(rsBatchH.fields("ISDate1")))
            Param(1) = GetParamValue(tostr(rsBatchH.fields("Param2")), toval(rsBatchH.fields("ISDate2")))
            Param(2) = GetParamValue(tostr(rsBatchH.fields("Param3")), toval(rsBatchH.fields("ISDate3")))
            Param(3) = GetParamValue(tostr(rsBatchH.fields("Param4")), toval(rsBatchH.fields("ISDate4")))
            Param(4) = GetParamValue(tostr(rsBatchH.fields("Param5")), toval(rsBatchH.fields("ISDate5")))
            Param(5) = GetParamValue(tostr(rsBatchH.fields("Param6")), toval(rsBatchH.fields("ISDate6")))

            PartialF = toval(Param(4))
            db_user = tostr(rsBatchH.fields("db_user"))
            db.Var_db_user = db_user
            db_user_type = db.GetUserType
            db_user_ccode = db.GetUserCCode
            If db_user_ccode = "" Then db_user_ccode = CCode
            db_password = Decrypt(tostr(rsBatchH.fields("db_Password")), db_user_ccode)

            SQL = "Select S1.* "
            rs = db.OpenRemoteSet("Select * from " & tPrefix & "sysBatchOriginal" & MapGroup & " Where 1<>1")
            fIndex = rs.fields_index
            FCount = UBound(fIndex)

            If SysBatchExtension >= 1 Then
                rs = db.OpenRemoteSet("Select * from " & tPrefix & "sysBatchOriginal" & MapGroup & "_Ext Where 1<>1")
                fIndex_1 = rs.fields_index
                FCount_1 = UBound(fIndex_1)

                For i = 0 To rs.fields_Count - 1
                    If GetFieldIndex(rs.fields_Name(i), fIndex) >= FCount Then
                        SQL = SQL & ", S2.[" & rs.fields_Name(i) & "]"
                    End If
                Next
            End If

            If SysBatchExtension >= 2 Then
                rs = db.OpenRemoteSet("Select * from " & tPrefix & "sysBatchOriginal" & MapGroup & "_Ext_1 Where 1<>1")
                For i = 0 To rs.fields_Count - 1
                    If GetFieldIndex(rs.fields_Name(i), fIndex) >= FCount AndAlso GetFieldIndex(rs.fields_Name(i), fIndex_1) >= FCount_1 Then
                        SQL = SQL & ", S3.[" & rs.fields_Name(i) & "]"
                    End If
                Next
            End If

            SQL = SQL & " from " & tPrefix & "sysBatchOriginal" & MapGroup & " S1 "
            If SysBatchExtension >= 1 Then
                SQL = SQL & " Left Join " & tPrefix & "sysBatchOriginal" & MapGroup & "_Ext S2 on (S2.BID=S1.BID)"
            End If
            If SysBatchExtension >= 2 Then
                SQL = SQL & " Left Join " & tPrefix & "sysBatchOriginal" & MapGroup & "_Ext_1 S3 on (S3.BID=S1.BID)"
            End If
            SQL = SQL & " Where S1.Status='W' and S1.BatchNo='" & BatchNo & "' "
            SQL = SQL & " Order by S1.BID"
            rsBatchD = db.OpenRemoteSet(SQL)

            If Not rsBatchD.Eof Then
                'Start Process..
                Dim TotalRecord As Long = 0
                Dim TotalNotVerify As Long = 0
                Dim Completed As Integer = 0
                Dim NotCompleted As Integer = 0

                Dim dbLib As New ClassLib

                Try
                    If dbLib.OpenConnection("", "", "", "", "") Then
                        dbLib.Var_db_user = db_user
                        dbLib.Var_db_Password = db_password

                        Dim PObj As New ProcessBatch.NetV4.Proclib
                        PObj.Var_db_User = db.Var_db_user
                        PObj.Var_db_User_Type = db.Var_db_User_Type
                        PObj.Var_db_User_CCode = db.Var_db_User_CCode
                        PObj.Var_Db_Name = db.Var_Db_Name
                        PObj.Var_Db_Server = db.Var_Db_Server
                        PObj.Var_db_Password = db_password
                        PObj.Var_ObjSession = dbLib.Var_ObjSession

                        fIndex = rsBatchD.fields_index
                        Do While Not rsBatchD.Eof
                            dbLib.Begintrans()
                            fBuffer = rsBatchD.GetRecord

                            ' Non-Completed
                            If Not dbLib.Execute("Update " & tPrefix & "SysBatchOriginal" & MapGroup & " Set Status='N', ErrMsg='On-Progress...' Where Status='W' and BID=" & rsBatchD.fields("BID")) Then
                                strmsg = "Update SysBatchOriginal" & MapGroup & " Status Error : " & dbLib.ErrDesc
                                Throw New Exception(strmsg)
                            End If
                            If Not dbLib.Execute("Update " & tPrefix & "SysBatchH Set TotalNotVerify=TotalNotVerify+1 Where BatchNo='" & BatchNo & "'") Then
                                strmsg = "Update SysBatch Status Error : " & dbLib.ErrDesc
                                Throw New Exception(strmsg)
                            End If

                            Dim PMainProcName As String = ""

                            If tostr(rsBatchD.fields("TType")) = "" Then
                                PMainProcName = MainProcName
                            Else
                                Select Case tostr(rsBatchD.fields("TType"))
                                    Case "R"
                                        PMainProcName = "Verify_UP"
                                    Case "E"
                                        PMainProcName = "Verify_BE"
                                    Case "C"
                                        PMainProcName = "Verify_BC"
                                    Case Else
                                        PMainProcName = MainProcName
                                End Select
                            End If

                            If PObj.Call_Process_Execute("", PMainProcName, Param, fIndex, fBuffer, tPrefix, Nothing, Nothing) Then
                                strmsg = PObj.ErrDesc
                                ' Completed
                                If Not dbLib.Execute("Update " & tPrefix & "SysBatchOriginal" & MapGroup & " Set Status='C', ErrMsg='' Where BID=" & rsBatchD.fields("BID")) Then
                                    strmsg = "Update SysBatchOriginal" & MapGroup & " Status Error : " & dbLib.ErrDesc
                                    Throw New Exception(strmsg)
                                End If
                                If Not dbLib.Execute("Update " & tPrefix & "SysBatchH Set TotalRecord=TotalRecord+1, TotalNotVerify=TotalNotVerify-1 Where BatchNo='" & BatchNo & "'") Then
                                    strmsg = "Update SysBatch Status Error : " & dbLib.ErrDesc
                                    Throw New Exception(strmsg)
                                End If

                                ' Continue with Upload
                                If UploadF Then
                                    Dim rsBatchDUP As RemoteSet
                                    Dim fIndexUP() As String
                                    Dim fBufferUP() As String
                                    Dim FCountUP As Integer
                                    Dim fIndexUP_1() As String
                                    Dim FCountUP_1 As Integer

                                    SQL = "Select S1.* "
                                    rs = db.OpenRemoteSet("Select * from " & tPrefix & "sysBatch" & MapGroup & " Where 1<>1")
                                    fIndexUP = rs.fields_index
                                    FCountUP = UBound(fIndexUP)

                                    If SysBatchExtension >= 1 Then
                                        rs = db.OpenRemoteSet("Select * from " & tPrefix & "sysBatch" & MapGroup & "_Ext Where 1<>1")
                                        fIndexUP_1 = rs.fields_index
                                        FCountUP_1 = UBound(fIndexUP)

                                        For i = 0 To rs.fields_Count - 1
                                            If GetFieldIndex(rs.fields_Name(i), fIndexUP) >= FCountUP Then
                                                SQL = SQL & ", S2.[" & rs.fields_Name(i) & "]"
                                            End If
                                        Next
                                    End If

                                    If SysBatchExtension >= 2 Then
                                        rs = db.OpenRemoteSet("Select * from " & tPrefix & "sysBatch" & MapGroup & "_Ext_1 Where 1<>1")
                                        For i = 0 To rs.fields_Count - 1
                                            If GetFieldIndex(rs.fields_Name(i), fIndexUP) >= FCountUP AndAlso GetFieldIndex(rs.fields_Name(i), fIndexUP_1) >= FCountUP_1 Then
                                                SQL = SQL & ", S3.[" & rs.fields_Name(i) & "]"
                                            End If
                                        Next
                                    End If

                                    SQL = SQL & " from " & tPrefix & "sysBatch" & MapGroup & " S1 "
                                    If SysBatchExtension >= 1 Then
                                        SQL = SQL & " Left Join " & tPrefix & "sysBatch" & MapGroup & "_Ext S2 on (S2.BID=S1.BID)"
                                    End If

                                    If SysBatchExtension >= 2 Then
                                        SQL = SQL & " Left Join " & tPrefix & "sysBatch" & MapGroup & "_Ext_1 S3 on (S3.BID=S1.BID)"
                                    End If
                                    SQL = SQL & " Where S1.Status='W' and S1.BID=" & rsBatchD.fields("BID")

                                    rsBatchDUP = db.OpenRemoteSet(SQL)
                                    fBufferUP = rsBatchDUP.GetRecord
                                    fIndexUP = rsBatchDUP.fields_index

                                    ' Non-Completed
                                    If Not dbLib.Execute("Update " & tPrefix & "sysBatch" & MapGroup & " Set Status='N', ErrMsg='On-Progress...' Where Status='W' and BID=" & rsBatchDUP.fields("BID")) Then
                                        strmsg = "Update SysBatchOriginal" & MapGroup & " Status Error : " & dbLib.ErrDesc
                                        Throw New Exception(strmsg)
                                    End If
                                    If Not dbLib.Execute("Update " & tPrefix & "SysBatchH Set NotCompleted=NotCompleted+1 Where BatchNo='" & BatchNo & "'") Then
                                        strmsg = "Update SysBatch Status Error : " & dbLib.ErrDesc
                                        Throw New Exception(strmsg)
                                    End If

                                    PMainProcName = Replace(PMainProcName, "Verify_", "Upload_")
                                    If PObj.Call_Process_Execute("", PMainProcName, Param, fIndexUP, fBufferUP, Nothing, Nothing, Nothing) Then
                                        strmsg = PObj.ErrDesc
                                        ' Completed
                                        If Not dbLib.Execute("Update " & tPrefix & "sysBatch" & MapGroup & " Set Status='C', ErrMsg='' Where BID=" & rsBatchD.fields("BID")) Then
                                            strmsg = "Update SysBatch" & MapGroup & " Status Error : " & dbLib.ErrDesc
                                            Throw New Exception(strmsg)
                                        End If
                                        If Not dbLib.Execute("Update " & tPrefix & "SysBatchH Set Completed=Completed+1, NotCompleted=NotCompleted-1 Where BatchNo='" & BatchNo & "'") Then
                                            strmsg = "Update SysBatch Status Error : " & dbLib.ErrDesc
                                            Throw New Exception(strmsg)
                                        End If

                                        TotalRecord = TotalRecord + 1
                                        Completed = Completed + 1
                                        dbLib.Committrans()
                                    Else
                                        strmsg = PObj.ErrDesc
                                        dbLib.RollBack()
                                        ' Non-Completed
                                        If Not dbLib.Execute("Update " & tPrefix & "SysBatchOriginal" & MapGroup & " Set Status='N', ErrMsg='" & tostr(Mid(strmsg, 1, 255)) & "' Where BID=" & rsBatchD.fields("BID")) Then
                                            strmsg = "Update SysBatchOriginal" & MapGroup & " Status Error : " & dbLib.ErrDesc
                                            Throw New Exception(strmsg)
                                        End If
                                        If Not dbLib.Execute("Update " & tPrefix & "SysBatchH Set TotalNotVerify=TotalNotVerify+1 Where BatchNo='" & BatchNo & "'") Then
                                            strmsg = "Update SysBatch Status Error : " & dbLib.ErrDesc
                                            Throw New Exception(strmsg)
                                        End If
                                        TotalNotVerify = TotalNotVerify + 1
                                    End If
                                Else
                                    TotalRecord = TotalRecord + 1
                                    dbLib.Committrans()
                                End If
                            Else
                                strmsg = PObj.ErrDesc
                                dbLib.RollBack()
                                ' Non-Completed
                                If Not dbLib.Execute("Update " & tPrefix & "SysBatchOriginal" & MapGroup & " Set Status='N', ErrMsg='" & tostr(Mid(strmsg, 1, 255)) & "' Where BID=" & rsBatchD.fields("BID")) Then
                                    strmsg = "Update SysBatchOriginal" & MapGroup & " Status Error : " & dbLib.ErrDesc
                                    Throw New Exception(strmsg)
                                End If
                                If Not dbLib.Execute("Update " & tPrefix & "SysBatchH Set TotalNotVerify=TotalNotVerify+1 Where BatchNo='" & BatchNo & "'") Then
                                    strmsg = "Update SysBatch Status Error : " & dbLib.ErrDesc
                                    Throw New Exception(strmsg)
                                End If
                                TotalNotVerify = TotalNotVerify + 1
                            End If

                            rAffected = rAffected + 1
                            rsBatchD.MoveNext()
                        Loop
                        dbLib.CloseSession()
                        dbLib = Nothing
                    Else
                        strmsg = "Connection Error !" & dbLib.ErrDesc
                        dbLib = Nothing
                        Return False
                    End If
                Catch ex As Exception
                    dbLib.CloseSession()
                    dbLib = Nothing
                    Throw New Exception(ex.Message)
                End Try
            End If

            SQL = "Select * from " & tPrefix & "sysBatchH S "
            SQL = SQL & " Where S.BatchNo='" & BatchNo & "'"
            rsBatchH = db.OpenRemoteSet(SQL)
            rsBatchH.Edit()
            If rsBatchH.fields("TotalRecord") = rsBatchH.fields("TotalUpload") Then
                rsBatchH.fields("Status") = "T"
                rsBatchH.fields("ProcName") = "Upload_" & MapGroup
            ElseIf rsBatchH.fields("TotalUpload") = rsBatchH.fields("TotalNotVerify") Then
                rsBatchH.fields("Status") = "N"
            ElseIf rsBatchH.fields("TotalNotVerify") <> 0 AndAlso rsBatchH.fields("TotalRecord") + rsBatchH.fields("TotalNotVerify") = rsBatchH.fields("TotalUpload") Then
                If PartialF Then
                    rsBatchH.fields("Status") = "T"
                    rsBatchH.fields("ProcName") = "Upload_" & MapGroup
                Else
                    rsBatchH.fields("Status") = "N"
                End If
            Else
                rsBatchH.fields("Status") = "N"
            End If

            If UploadF Then
                If rsBatchH.fields("Completed") <> 0 Then
                    rsBatchH.fields("Status") = "C"
                Else
                    rsBatchH.fields("Status") = "N"
                End If
            End If

            If Not rsBatchH.Update Then
                strmsg = "Update Batch Status Error : " & db.ErrDesc
                Throw New Exception(strmsg)
            End If

            If tPrefix = "" Then
                If Not UploadF Then
                    If rsBatchH.fields("Status") = "T" Or rsBatchH.fields("Status") = "N" Then
                        If rsBatchH.fields("Status") = "T" Then
                            IPCRequest(db, ProcLib, ProcName + " Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalUpload").ToString + " Total Verify : " + rsBatchH.fields("TotalRecord").ToString + " Total Not Verify : " + rsBatchH.fields("TotalNotVerify").ToString, "SEND", "", 0, 0, rsBatchH.fields("db_User"), "", db_user, Today.Date, Time())
                            If CBool(toval(GetAppSetting("EmailTaskF"))) Then
                                SendMail("", db_user, "CARESVC - " & ProcName, ProcName + " Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, strmsg)
                            End If
                            IPCRequest(db, ProcLib, ProcName + " Not Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalUpload").ToString + " Total Verify : " + rsBatchH.fields("TotalRecord").ToString + " Total Not Verify : " + rsBatchH.fields("TotalNotVerify").ToString, "SEND", "", 0, 0, rsBatchH.fields("db_User"), "", db_user, Today.Date, Time())
                            If CBool(toval(GetAppSetting("EmailTaskF"))) Then
                                SendMail("", db_user, "CARESVC - " & ProcName, ProcName + " Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, strmsg)
                            End If
                        End If
                    End If
                Else
                    If rsBatchH.fields("Status") = "C" Or rsBatchH.fields("Status") = "N" Then
                        If rsBatchD.fields("Status") = "C" Then
                            IPCRequest(db, ProcLib, ProcName + " Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, "SEND", "", 0, 0, rsBatchH.fields("db_User"), "", db_user, Today.Date, Time())
                            SendMail("", db_user, "CARESVC - " & ProcName, ProcName + " Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, strmsg)
                        Else
                            IPCRequest(db, ProcLib, ProcName + " Not Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, "SEND", "", 0, 0, rsBatchH.fields("db_User"), "", db_user, Today.Date, Time())
                            SendMail("", db_user, "CARESVC - " & ProcName, ProcName + " Not Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, strmsg)
                        End If
                    End If
                End If
            End If

            If Not db.ReleaseLockedObject(tPrefix & "SysBatch" & BatchNo) Then Throw New Exception(db.ErrDesc)
            db.Committrans()

            If rsBatchH.fields("Status") = "N" Then
                VerifyExecute = False
            Else
                VerifyExecute = True
            End If
        Catch ex As Exception
            db.RollBack()
            strmsg = ex.Message
            VerifyExecute = False
        End Try
    End Function
    Private Function BatchExecute(ByVal BatchNo As String, ByRef strmsg As String, Optional ByVal tPrefix As String = "") As Boolean
        Dim db_user As String
        Dim db_password As String
        Dim db_user_type As String
        Dim db_user_ccode As String

        strmsg = ""
        Try
            Dim ProcLib As String
            Dim ProcName As String
            Dim MainProcName As String
            Dim MapGroup As String
            Dim Param(5) As Object
            Dim fIndex() As String
            Dim fBuffer() As String
            Dim FCount As Integer
            Dim fIndex_1() As String
            Dim FCount_1 As Integer
            Dim rsBatchH As RemoteSet
            Dim rsBatchD As RemoteSet
            Dim rAffected As Integer = 0
            Dim PartialF As Boolean
            Dim SysBatchExtension As Integer = toval(GetAppSetting("SysBatchExtension"))
            If SysBatchExtension = 0 Then SysBatchExtension = 1
            Dim SQL As String

            db.Begintrans()
            SQL = "Select top 1 S.* from " & tPrefix & "sysBatchH S with (ReadCommitted)"
            SQL = SQL & " Where S.status = 'T' "
            SQL = SQL & " and S.BatchNo='" & BatchNo & "'"
            rsBatchH = db.OpenRemoteSet(SQL)
            If rsBatchH.Eof Then
                db.Committrans()
                rsBatchH.CloseRemoteSet()
                BatchExecute = True
                Exit Function
            End If

            If Not db.GetLockedObject(tPrefix & "SysBatch" & BatchNo) Then Throw New Exception(db.ErrDesc)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select MapGroup from MAP_Header Where MapCode='" & rsBatchH.fields("MapCode") & "'")
            MapGroup = tostr(rs.fields("MapGroup"))
            If MapGroup = "" Then MapGroup = "D"

            ProcLib = tostr(rsBatchH.fields("ProcLib"))
            ProcName = tostr(rsBatchH.fields("ProcName"))
            MainProcName = tostr(rsBatchH.fields("MainProcName"))
            If MainProcName = "" Then MainProcName = ProcName
            MainProcName = "Upload_" & MapGroup
            Param(0) = GetParamValue(tostr(rsBatchH.fields("Param1")), toval(rsBatchH.fields("ISDate1")))
            Param(1) = GetParamValue(tostr(rsBatchH.fields("Param2")), toval(rsBatchH.fields("ISDate2")))
            Param(2) = GetParamValue(tostr(rsBatchH.fields("Param3")), toval(rsBatchH.fields("ISDate3")))
            Param(3) = GetParamValue(tostr(rsBatchH.fields("Param4")), toval(rsBatchH.fields("ISDate4")))
            Param(4) = GetParamValue(tostr(rsBatchH.fields("Param5")), toval(rsBatchH.fields("ISDate5")))
            Param(5) = GetParamValue(tostr(rsBatchH.fields("Param6")), toval(rsBatchH.fields("ISDate6")))
            PartialF = toval(Param(4))

            db_user = tostr(rsBatchH.fields("db_user"))
            db_user_type = db.GetUserType
            db_user_ccode = db.GetUserCCode
            If db_user_ccode = "" Then db_user_ccode = CCode
            db_password = Decrypt(tostr(rsBatchH.fields("db_Password")), db_user_ccode)

            SQL = "Select S1.* "
            rs = db.OpenRemoteSet("Select * from " & tPrefix & "sysBatch" & MapGroup & " Where 1<>1")
            fIndex = rs.fields_index
            FCount = UBound(fIndex)

            If SysBatchExtension >= 1 Then
                rs = db.OpenRemoteSet("Select * from " & tPrefix & "sysBatch" & MapGroup & "_Ext Where 1<>1")
                fIndex_1 = rs.fields_index
                FCount_1 = UBound(fIndex_1)

                For i = 0 To rs.fields_Count - 1
                    If GetFieldIndex(rs.fields_Name(i), fIndex) >= FCount Then
                        SQL = SQL & ", S2.[" & rs.fields_Name(i) & "]"
                    End If
                Next
            End If

            If SysBatchExtension >= 2 Then
                rs = db.OpenRemoteSet("Select * from " & tPrefix & "sysBatch" & MapGroup & "_Ext_1 Where 1<>1")
                For i = 0 To rs.fields_Count - 1
                    If GetFieldIndex(rs.fields_Name(i), fIndex) >= FCount AndAlso GetFieldIndex(rs.fields_Name(i), fIndex_1) >= FCount_1 Then
                        SQL = SQL & ", S3.[" & rs.fields_Name(i) & "]"
                    End If
                Next
            End If

            SQL = SQL & " from " & tPrefix & "sysBatch" & MapGroup & " S1 "
            If SysBatchExtension >= 1 Then
                SQL = SQL & " Left Join " & tPrefix & "sysBatch" & MapGroup & "_Ext S2 on (S2.BID=S1.BID)"
            End If

            If SysBatchExtension >= 2 Then
                SQL = SQL & " Left Join " & tPrefix & "sysBatch" & MapGroup & "_Ext_1 S3 on (S3.BID=S1.BID)"
            End If
            SQL = SQL & " Where S1.Status='W' and S1.BatchNo='" & BatchNo & "' "
            SQL = SQL & " Order by S1.BID"
            rsBatchD = db.OpenRemoteSet(SQL)
            If Not rsBatchD.Eof Then
                'Start Process..
                Dim Completed As Long = 0
                Dim NotCompleted As Long = 0
                Dim dbLib As New ClassLib

                Try
                    If dbLib.OpenConnection("", "", "", "", "") Then
                        dbLib.Var_db_user = db_user
                        dbLib.Var_db_Password = db_password

                        Dim PObj As New ProcessBatch.NetV4.Proclib
                        PObj.Var_db_User = db.Var_db_user
                        PObj.Var_db_User_Type = db.Var_db_User_Type
                        PObj.Var_db_User_CCode = db.Var_db_User_CCode
                        PObj.Var_Db_Name = db.Var_Db_Name
                        PObj.Var_Db_Server = db.Var_Db_Server
                        PObj.Var_db_Password = db_password
                        PObj.Var_ObjSession = dbLib.Var_ObjSession

                        fIndex = rsBatchD.fields_index
                        Do While Not rsBatchD.Eof
                            dbLib.Begintrans()
                            fBuffer = rsBatchD.GetRecord

                            ' Non-Completed
                            If Not dbLib.Execute("Update " & tPrefix & "sysBatch" & MapGroup & " Set Status='N', ErrMsg='On-Progress...' Where Status='W' and BID=" & rsBatchD.fields("BID")) Then
                                strmsg = "Update SysBatchOriginal" & MapGroup & " Status Error : " & dbLib.ErrDesc
                                Throw New Exception(strmsg)
                            End If
                            If Not dbLib.Execute("Update " & tPrefix & "SysBatchH Set NotCompleted=NotCompleted+1 Where BatchNo='" & BatchNo & "'") Then
                                strmsg = "Update SysBatch Status Error : " & dbLib.ErrDesc
                                Throw New Exception(strmsg)
                            End If

                            Dim PMainProcName As String = ""

                            If tostr(rsBatchD.fields("TType")) = "" Then
                                PMainProcName = MainProcName
                            Else
                                Select Case tostr(rsBatchD.fields("TType"))
                                    Case "R"
                                        PMainProcName = "Upload_UP"
                                    Case "E"
                                        PMainProcName = "Upload_BE"
                                    Case "C"
                                        PMainProcName = "Upload_BC"
                                    Case Else
                                        PMainProcName = MainProcName
                                End Select
                            End If

                            If PObj.Call_Process_Execute("", PMainProcName, Param, fIndex, fBuffer, tPrefix, Nothing, Nothing) Then
                                strmsg = PObj.ErrDesc
                                ' Completed
                                If Not dbLib.Execute("Update " & tPrefix & "sysBatch" & MapGroup & " Set Status='C', ErrMsg='' Where BID=" & rsBatchD.fields("BID")) Then
                                    strmsg = "Update SysBatch" & MapGroup & " Status Error : " & dbLib.ErrDesc
                                    Throw New Exception(strmsg)
                                End If
                                If Not dbLib.Execute("Update " & tPrefix & "SysBatchH Set Completed=Completed+1, NotCompleted=NotCompleted-1 Where BatchNo='" & BatchNo & "'") Then
                                    strmsg = "Update SysBatch Status Error : " & dbLib.ErrDesc
                                    Throw New Exception(strmsg)
                                End If
                                Completed = Completed + 1
                                dbLib.Committrans()
                            Else
                                strmsg = PObj.ErrDesc
                                dbLib.RollBack()
                                ' Non-Completed
                                If Not dbLib.Execute("Update " & tPrefix & "sysBatch" & MapGroup & " Set Status='N', ErrMsg='" & tostr(Mid(strmsg, 1, 255)) & "' Where BID=" & rsBatchD.fields("BID")) Then
                                    strmsg = "Update SysBatch" & MapGroup & " Status Error : " & dbLib.ErrDesc
                                    Throw New Exception(strmsg)
                                End If
                                If Not dbLib.Execute("Update " & tPrefix & "SysBatchH Set NotCompleted=NotCompleted+1 Where BatchNo='" & BatchNo & "'") Then
                                    strmsg = "Update SysBatch Status Error : " & dbLib.ErrDesc
                                    Throw New Exception(strmsg)
                                End If
                                NotCompleted = NotCompleted + 1
                            End If

                            rAffected = rAffected + 1
                            rsBatchD.MoveNext()
                        Loop
                        dbLib.CloseSession()
                        dbLib = Nothing
                    Else
                        strmsg = "Connection Error !" & dbLib.ErrDesc
                        dbLib = Nothing
                        Return False
                    End If
                Catch ex As Exception
                    dbLib.CloseSession()
                    dbLib = Nothing
                    Throw New Exception(ex.Message)
                End Try
            End If

            SQL = "Select * from " & tPrefix & "sysBatchH S "
            SQL = SQL & " Where S.BatchNo='" & BatchNo & "'"
            rsBatchH = db.OpenRemoteSet(SQL)
            rsBatchH.Edit()

            If rsBatchH.fields("Completed") <> 0 Then
                rsBatchH.fields("Status") = "C"
            Else
                rsBatchH.fields("Status") = "N"
            End If

            If Not rsBatchH.Update Then
                strmsg = "Update Batch Status Error : " & db.ErrDesc
                Throw New Exception(strmsg)
            End If

            If tPrefix = "" Then
                If rsBatchH.fields("Status") = "C" Or rsBatchH.fields("Status") = "N" Then
                    If rsBatchD.fields("Status") = "C" Then
                        IPCRequest(db, ProcLib, ProcName + " Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, "SEND", "", 0, 0, rsBatchH.fields("db_User"), "", db_user, Today.Date, Time())
                        If CBool(toval(GetAppSetting("EmailTaskF"))) Then
                            SendMail("", db_user, "CARESVC - " & ProcName, ProcName + " Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, strmsg)
                        End If
                    Else
                        IPCRequest(db, ProcLib, ProcName + " Not Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, "SEND", "", 0, 0, rsBatchH.fields("db_User"), "", db_user, Today.Date, Time())
                        If CBool(toval(GetAppSetting("EmailTaskF"))) Then
                            SendMail("", db_user, "CARESVC - " & ProcName, ProcName + " Completed : " + tostr(rsBatchH.fields("Param1")) + " " + tostr(rsBatchH.fields("Param2")) + " " + tostr(rsBatchH.fields("Param3")) + " " + tostr(rsBatchH.fields("Param4")) + " " + tostr(rsBatchH.fields("Param5")) + " " + tostr(rsBatchH.fields("Param6")) + vbCrLf + " Total Record : " & rsBatchH.fields("TotalRecord").ToString + " Completed : " + rsBatchH.fields("Completed").ToString + " Not Completed : " + rsBatchH.fields("NotCompleted").ToString, strmsg)
                        End If
                    End If
                End If
            End If
            If Not db.ReleaseLockedObject(tPrefix & "SysBatch" & BatchNo) Then Throw New Exception(db.ErrDesc)
            db.Committrans()

            If rsBatchH.fields("Status") = "N" Then
                BatchExecute = False
            Else
                BatchExecute = True
            End If
        Catch ex As Exception
            db.RollBack()
            strmsg = ex.Message
            BatchExecute = False
        End Try
    End Function
    Private Function GetParamValue(ByVal param As String, ByVal IsDate As Boolean) As Object
        If IsDate Then
            GetParamValue = DateSerial(Mid(param, 1, 4), Mid(param, 6, 2), Mid(param, 9, 2))
        Else
            GetParamValue = param
        End If
    End Function
    Private Function IPCRequest(ByVal db As ClassLib, ByVal ProcName As String, ByVal Remarks As String, ByVal Param1 As String, ByVal Param2 As String, ByVal Param3 As Long, ByVal Param4 As Long, ByVal Param5 As String, ByVal Param6 As String, ByVal UserID As String, ByVal UserDate As Date, ByVal UserTime As String) As Boolean
        IPCRequest = db.Execute("Insert into SysRequest (ProcName, Remarks, Param1, Param2, Param3, Param4, Param5, Param6, UserID, UserDate, UserTime,Status) Values ('" & ProcName & "','" & Remarks & "','" & Param1 & "','" & Param2 & "'," & Param3 & "," & Param4 & ",'" & Param5 & "','" & Param6 & "','" & UserID & "','" & Format(UserDate, "MM/DD/YYYY") & "','" & UserTime & "','')")
    End Function
    Private Function SendMail(ByVal UserFrom As String, ByVal UserTo As String, ByVal EMailSubject As String, ByVal EMailBody As String, ByRef ErrMsg As String) As Boolean
        Try
            Dim rs As RemoteSet
            Dim SMTPServer As String = ""
            Dim SMTPPort As String = ""
            Dim SMTPEmail As String = ""
            Dim EmailFrom As String = ""
            Dim EmailTo As String = ""

            rs = db.OpenRemoteSet("Select SMTPServer, SMTPPort, SMTPEmail from SysTab")
            SMTPServer = rs.fields("SMTPServer")
            SMTPPort = rs.fields("SMTPPort")
            SMTPEmail = rs.fields("SMTPEmail")

            If SMTPEmail <> "" Then
                rs = db.OpenRemoteSet("Select EMail from SysUser Where ID='" & UserFrom & "'")
                EmailFrom = rs.fields("EMail")
            End If
            If EmailFrom = "" Then EmailFrom = SMTPEmail

            Dim msgMail As New System.Net.Mail.MailMessage()
            Dim smtSender As SmtpClient = New SmtpClient()

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

    <WebMethod(EnableSession:=False)> _
    Public Function Delete_Upload_Data(ByVal dbuser As String, ByVal dbpassword As String, ByVal inparam() As Object, ByRef outparam As Object, ByRef ErrMsg As String) As Boolean
        Try
            Dim flag As Integer = 0
            Dim batchno As String = ""
            Dim mapcode As String = ""
            Dim Sql As String = ""
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            batchno = inparam(0)
            flag = inparam(1)
            mapcode = inparam(2)
            '1=delete all
            '2=delete error data
            '3=delete error with upload

            Dim LockedKey As String = batchno
            If Not db.GetLockedObject("SysBatch" & LockedKey) Then Throw New Exception(db.ErrDesc)

            If flag = 1 Then
                Sql = "delete from SysBatchOriginal" & mapcode & " where batchno='" & batchno & "'"
                If Not db.Execute(Sql) Then
                    Throw New Exception(db.ErrDesc)
                End If
                Sql = "delete from SysBatch" & mapcode & " where batchno='" & batchno & "'"
                If Not db.Execute(Sql) Then
                    Throw New Exception(db.ErrDesc)
                End If
                Sql = "Update SysBatchH set totalUpload=0, totalRecord=0, status='P' where batchno='" & batchno & "'"
                If Not db.Execute(Sql) Then
                    Throw New Exception(db.ErrDesc)
                End If
            ElseIf flag = 2 Then
                Sql = "delete from SysBatchOriginal" & mapcode & " where Status ='N' and  batchno='" & batchno & "'"
                If Not db.Execute(Sql) Then
                    Throw New Exception(db.ErrDesc)
                End If
                Sql = "Update SysBatchH set totalUpload=totalrecord,status='P' where batchno='" & batchno & "'"
                If Not db.Execute(Sql) Then
                    Throw New Exception(db.ErrDesc)
                End If
            End If

            If Not db.ReleaseLockedObject("SysBatch" & LockedKey) Then Throw New Exception(db.ErrDesc)
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
    Public Function Export_Upload_ListView_To_Excel(ByVal dbuser As String, ByVal dbpassword As String, ByVal MapGroup As String, ByVal Batchno As String, ByVal MapCode As String, ByVal Status As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Dim message As String = ""
        Dim rsd As RemoteSet
        Dim Field As String
        Dim SQL As String = ""
        Try
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If
            Field = ""

            If Left(MapCode, 2).ToUpper = "BC" And Status.ToUpper = "P" Then
                rsd = db.OpenRemoteSet("Select FieldFrom,FieldTo,FieldType from map_field where mapcode = '" & MapCode & "' and FieldFrom not like '(Dummy%' and FieldTo not in ('RefNo') order by OrderNo,FieldFrom")
                Do While Not rsd.Eof
                    Field = Field & IIf(rsd.Bookmark = 0, "", ",") & "[" & rsd.fields("FieldTo") & "] as [" & rsd.fields("FieldFrom") & "]"
                    rsd.MoveNext()
                Loop

                SQL = "select D.BID, D.REFNO, " & Field & " ,dbo.getrate(a.ano,1,1) as rate1,dbo.getrate(a.ano,2,1) as rate2,dbo.getrate(a.ano,3,1) as rate3,dbo.getrate(a.ano,4,1) as rate4,dbo.getrate(a.ano,5,1) as rate5,dbo.getrate(a.ano,1,4) as GrossPremi1,dbo.getrate(a.ano,2,4) as GrossPremi2,dbo.getrate(a.ano,3,4) as GrossPremi3,dbo.getrate(a.ano,4,4) as GrossPremi4,dbo.getrate(a.ano,5,4) as GrossPremi5,dbo.getrate(a.ano,1,5) as Discount1,dbo.getrate(a.ano,2,5) as Discount2,dbo.getrate(a.ano,3,5) as Discount3,dbo.getrate(a.ano,4,5) as Discount4,dbo.getrate(a.ano,5,5) as Discount5,-(dbo.getrate(a.ano,1,4)-dbo.getrate(a.ano,1,5)) as NettPremi1,-(dbo.getrate(a.ano,2,4)-dbo.getrate(a.ano,2,5)) as NettPremi2,-(dbo.getrate(a.ano,3,4)-dbo.getrate(a.ano,3,5)) as NettPremi3,-(dbo.getrate(a.ano,4,4)-dbo.getrate(a.ano,4,5)) as NettPremi4,-(dbo.getrate(a.ano,5,4)-dbo.getrate(a.ano,5,5)) as NettPremi5,isnull(f.amount,0) as adm "
                SQL = SQL & " from SysBatchOriginal" & MapGroup & " D"
                SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                SQL = SQL & " Cross Apply (Select Top 1 D1.BID from SysBatch" & MapGroup & " D1 Where D1.BID=D.BID and D1.Status='C') D1"
                SQL = SQL & " Cross Apply (Select Top 1 A.ANO from Acceptance A Where A.BID=D.BID) A"
                SQL = SQL & " Outer Apply (Select Top 1 f.Amount from ACCFD f Where f.ANO=a.ANO) f"
                SQL = SQL & " Where D.BatchNo ='" & Batchno & "'"
                SQL = SQL & " order by D.BID"
            Else
                Select Case Status.ToUpper
                    Case "E" 'Error
                        rsd = db.OpenRemoteSet("Select FieldFrom,FieldTo,FieldType from map_field where mapcode = '" & MapCode & "' and FieldFrom not like '(Dummy%' order by OrderNo,FieldFrom")
                        Do While Not rsd.Eof
                            Field = Field & IIf(rsd.Bookmark = 0, "", ",") & "[" & rsd.fields("FieldTo") & "] as [" & rsd.fields("FieldFrom") & "]"
                            rsd.MoveNext()
                        Loop

                        SQL = "Select D.BID, D.[ErrMsg]," & Field & " "
                        SQL = SQL & " from SysBatchOriginal" & MapGroup & " D "
                        SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID) "
                        SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                        SQL = SQL & " Where D.BatchNo ='" & Batchno & "' and D.Status = 'N' "
                        SQL = SQL & " order by D.BID"
                    Case "C" 'Completed without premi
                        rsd = db.OpenRemoteSet("Select FieldFrom,FieldTo,FieldType from map_field where mapcode = '" & MapCode & "' and FieldFrom not like '(Dummy%' and FieldTo not in ('PolicyNo','CertificateNo') order by OrderNo,FieldFrom")
                        Do While Not rsd.Eof
                            Field = Field & IIf(rsd.Bookmark = 0, "", ",") & "[" & rsd.fields("FieldTo") & "] as [" & rsd.fields("FieldFrom") & "]"
                            rsd.MoveNext()
                        Loop

                        If MapGroup = "SC" Then
                            SQL = "Select D.BID, A.Policyno, A.Certificateno," & Field & " , D.TrDate "
                            SQL = SQL & " from SysBatchOriginal" & MapGroup & " D"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                            SQL = SQL & " Cross Apply (Select Top 1 D1.BID from SysBatch" & MapGroup & " D1 Where D1.BID=D.BID and D1.Status='C') D1"
                            SQL = SQL & " Cross Apply (Select Top 1 A.PolicyNo, A.CertificateNo from Acceptance A Where A.BID=D.BID) A"
                            SQL = SQL & " Where D.BatchNo ='" & Batchno & "'"
                            SQL = SQL & " order by D.BID"
                        ElseIf MapGroup = "UP" Then
                            SQL = "Select D.BID, A.Policyno, A.Certificateno," & Field & "  "
                            SQL = SQL & " from SysBatchOriginal" & MapGroup & " D"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                            SQL = SQL & " Cross Apply (Select Top 1 D1.BID from SysBatch" & MapGroup & " D1 Where D1.BID=D.BID and D1.Status='C') D1"
                            SQL = SQL & " Cross Apply (Select Top 1 A.PolicyNo, A.CertificateNo from Acceptance A Where A.BID=D.BID) A"
                            SQL = SQL & " Where D.BatchNo ='" & Batchno & "'"
                            SQL = SQL & " order by D.BID"
                        Else
                            SQL = "Select D.BID, A.Policyno, A.Certificateno," & Field & "  "
                            SQL = SQL & " from SysBatchOriginal" & MapGroup & " D"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                            SQL = SQL & " Cross Apply (Select Top 1 D1.BID from SysBatch" & MapGroup & " D1 Where D1.BID=D.BID and D1.Status='C') D1"
                            SQL = SQL & " Cross Apply (Select Top 1 A.PolicyNo, A.CertificateNo from Acceptance A Where A.BID=D.BID) A"
                            SQL = SQL & " Where D.BatchNo ='" & Batchno & "'"
                            SQL = SQL & " order by D.BID"
                        End If
                    Case "N" 'Not Completed
                        rsd = db.OpenRemoteSet("Select FieldFrom,FieldTo,FieldType from map_field where mapcode = '" & MapCode & "' and FieldFrom not like '(Dummy%' order by OrderNo,FieldFrom")
                        Do While Not rsd.Eof
                            Field = Field & IIf(rsd.Bookmark = 0, "", ",") & "[" & rsd.fields("FieldTo") & "] as [" & rsd.fields("FieldFrom") & "]"
                            rsd.MoveNext()
                        Loop

                        SQL = "Select D.BID, D1.[ErrMsg]," & Field & " "
                        SQL = SQL & " from SysBatchOriginal" & MapGroup & " D "
                        SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                        SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                        SQL = SQL & " Cross Apply (Select Top 1 ErrMsg from SysBatch" & MapGroup & " D1 Where D1.BID=D.BID and D1.Status='N') D1"
                        SQL = SQL & " Where D.BatchNo ='" & Batchno & "' "
                        SQL = SQL & " order by D.BID"
                    Case "W" 'Waiting
                        rsd = db.OpenRemoteSet("Select FieldFrom,FieldTo,FieldType from map_field where mapcode = '" & MapCode & "' and FieldFrom not like '(Dummy%' order by OrderNo,FieldFrom")
                        Do While Not rsd.Eof
                            Field = Field & IIf(rsd.Bookmark = 0, "", ",") & "[" & rsd.fields("FieldTo") & "] as [" & rsd.fields("FieldFrom") & "]"
                            rsd.MoveNext()
                        Loop

                        SQL = "Select D.BID, D1.[ErrMsg]," & Field & " "
                        SQL = SQL & " from SysBatchOriginal" & MapGroup & " D "
                        SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                        SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                        SQL = SQL & " Cross Apply (Select Top 1 ErrMsg from SysBatch" & MapGroup & " D1 Where D1.BID=D.BID and D1.Status='W') D1"
                        SQL = SQL & " Where D.BatchNo ='" & Batchno & "' "
                        SQL = SQL & " order by D.BID"
                    Case "F" 'Full
                        rsd = db.OpenRemoteSet("Select FieldFrom,FieldTo,FieldType from map_field where mapcode = '" & MapCode & "' and FieldFrom not like '(Dummy%' order by OrderNo,FieldFrom")
                        Do While Not rsd.Eof
                            Field = Field & IIf(rsd.Bookmark = 0, "", ",") & "[" & rsd.fields("FieldTo") & "] as [" & rsd.fields("FieldFrom") & "]"
                            rsd.MoveNext()
                        Loop

                        SQL = "Select D.BID, D.[ErrMsg]," & Field & " "
                        SQL = SQL & " from SysBatchOriginal" & MapGroup & " D "
                        SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                        SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                        SQL = SQL & " Where D.BatchNo ='" & Batchno & "' "
                        SQL = SQL & " order by D.BID"
                    Case "P" 'Completed with premi
                        rsd = db.OpenRemoteSet("Select FieldFrom,FieldTo,FieldType from map_field where mapcode = '" & MapCode & "' and FieldFrom not like '(Dummy%' and FieldTo not in ('PolicyNo','CertificateNo') order by OrderNo,FieldFrom")
                        Do While Not rsd.Eof
                            Field = Field & IIf(rsd.Bookmark = 0, "", ",") & "[" & rsd.fields("FieldTo") & "] as [" & rsd.fields("FieldFrom") & "]"
                            rsd.MoveNext()
                        Loop

                        If MapGroup = "SC" Then
                            SQL = "select D.BID, A.Policyno, A.Certificateno, " & Field & " ,dbo.getrate(a.ano,1,1) as rate1,dbo.getrate(a.ano,2,1) as rate2,dbo.getrate(a.ano,3,1) as rate3,dbo.getrate(a.ano,4,1) as rate4,dbo.getrate(a.ano,5,1) as rate5,dbo.getrate(a.ano,6,1) as rate6,dbo.getrate(a.ano,1,2) as GrossPremi1,dbo.getrate(a.ano,2,2) as GrossPremi2,dbo.getrate(a.ano,3,2) as GrossPremi3,dbo.getrate(a.ano,4,2) as GrossPremi4,dbo.getrate(a.ano,5,2) as GrossPremi5,dbo.getrate(a.ano,6,2) as GrossPremi6,dbo.getrate(a.ano,1,3) as Discount1,dbo.getrate(a.ano,2,3) as Discount2,dbo.getrate(a.ano,3,3) as Discount3,dbo.getrate(a.ano,4,3) as Discount4,dbo.getrate(a.ano,5,3) as Discount5,dbo.getrate(a.ano,6,3) as Discount6,dbo.getrate(a.ano,1,2)-dbo.getrate(a.ano,1,3) as NettPremi1,dbo.getrate(a.ano,2,2)-dbo.getrate(a.ano,2,3) as NettPremi2,dbo.getrate(a.ano,3,2)-dbo.getrate(a.ano,3,3) as NettPremi3,dbo.getrate(a.ano,4,2)-dbo.getrate(a.ano,4,3) as NettPremi4,dbo.getrate(a.ano,5,2)-dbo.getrate(a.ano,5,3) as NettPremi5,dbo.getrate(a.ano,6,2)-dbo.getrate(a.ano,6,3) as NettPremi6, isnull(f.amount,0) as adm "
                            SQL = SQL & " from SysBatchOriginal" & MapGroup & " D"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                            SQL = SQL & " Cross Apply (Select Top 1 D1.BID from SysBatch" & MapGroup & " D1 Where D1.BID=D.BID and D1.Status='C') D1"
                            SQL = SQL & " Cross Apply (Select Top 1 A.ANO, A.PolicyNo, A.CertificateNo from Acceptance A Where A.BID=D.BID) A"
                            SQL = SQL & " Outer Apply (Select Top 1 f.Amount from ACCFD f Where f.ANO=a.ANO) f"
                            SQL = SQL & " Where D.BatchNo ='" & Batchno & "'"
                            SQL = SQL & " order by D.BID"
                        ElseIf MapGroup = "UP" Then
                            SQL = "select D.BID, A.Policyno, A.Certificateno, " & Field & " ,dbo.getrate(a.ano,1,1) as rate1,dbo.getrate(a.ano,2,1) as rate2,dbo.getrate(a.ano,3,1) as rate3,dbo.getrate(a.ano,4,1) as rate4,dbo.getrate(a.ano,5,1) as rate5,dbo.getrate(a.ano,6,1) as rate6,dbo.getrate(a.ano,1,2) as GrossPremi1,dbo.getrate(a.ano,2,2) as GrossPremi2,dbo.getrate(a.ano,3,2) as GrossPremi3,dbo.getrate(a.ano,4,2) as GrossPremi4,dbo.getrate(a.ano,5,2) as GrossPremi5,dbo.getrate(a.ano,6,2) as GrossPremi6,dbo.getrate(a.ano,1,3) as Discount1,dbo.getrate(a.ano,2,3) as Discount2,dbo.getrate(a.ano,3,3) as Discount3,dbo.getrate(a.ano,4,3) as Discount4,dbo.getrate(a.ano,5,3) as Discount5,dbo.getrate(a.ano,6,3) as Discount6,dbo.getrate(a.ano,1,2)-dbo.getrate(a.ano,1,3) as NettPremi1,dbo.getrate(a.ano,2,2)-dbo.getrate(a.ano,2,3) as NettPremi2,dbo.getrate(a.ano,3,2)-dbo.getrate(a.ano,3,3) as NettPremi3,dbo.getrate(a.ano,4,2)-dbo.getrate(a.ano,4,3) as NettPremi4,dbo.getrate(a.ano,5,2)-dbo.getrate(a.ano,5,3) as NettPremi5,dbo.getrate(a.ano,6,2)-dbo.getrate(a.ano,6,3) as NettPremi6 "
                            SQL = SQL & " from SysBatchOriginal" & MapGroup & " D"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                            SQL = SQL & " Cross Apply (Select Top 1 D1.BID from SysBatch" & MapGroup & " D1 Where D1.BID=D.BID and D1.Status='C') D1"
                            SQL = SQL & " Cross Apply (Select Top 1 A.ANO, A.PolicyNo, A.CertificateNo from Acceptance A Where A.BID=D.BID) A"
                            SQL = SQL & " Outer Apply (Select Top 1 f.Amount from ACCFD f Where f.ANO=a.ANO) f"
                            SQL = SQL & " Where D.BatchNo ='" & Batchno & "'"
                            SQL = SQL & " order by D.BID"
                        Else
                            SQL = "select D.BID, A.Policyno, A.Certificateno, " & Field & " ,dbo.getrate(a.ano,1,1) as rate1,dbo.getrate(a.ano,2,1) as rate2,dbo.getrate(a.ano,3,1) as rate3,dbo.getrate(a.ano,4,1) as rate4,dbo.getrate(a.ano,5,1) as rate5,dbo.getrate(a.ano,6,1) as rate6,dbo.getrate(a.ano,1,2) as GrossPremi1,dbo.getrate(a.ano,2,2) as GrossPremi2,dbo.getrate(a.ano,3,2) as GrossPremi3,dbo.getrate(a.ano,4,2) as GrossPremi4,dbo.getrate(a.ano,5,2) as GrossPremi5,dbo.getrate(a.ano,6,2) as GrossPremi6,dbo.getrate(a.ano,1,3) as Discount1,dbo.getrate(a.ano,2,3) as Discount2,dbo.getrate(a.ano,3,3) as Discount3,dbo.getrate(a.ano,4,3) as Discount4,dbo.getrate(a.ano,5,3) as Discount5,dbo.getrate(a.ano,6,3) as Discount6,dbo.getrate(a.ano,1,2)-dbo.getrate(a.ano,1,3) as NettPremi1,dbo.getrate(a.ano,2,2)-dbo.getrate(a.ano,2,3) as NettPremi2,dbo.getrate(a.ano,3,2)-dbo.getrate(a.ano,3,3) as NettPremi3,dbo.getrate(a.ano,4,2)-dbo.getrate(a.ano,4,3) as NettPremi4,dbo.getrate(a.ano,5,2)-dbo.getrate(a.ano,5,3) as NettPremi5,dbo.getrate(a.ano,6,2)-dbo.getrate(a.ano,6,3) as NettPremi6 "
                            SQL = SQL & " from SysBatchOriginal" & MapGroup & " D"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext D_Ext on (D_Ext.BID=D.BID)"
                            SQL = SQL & " Left Join SysBatchOriginal" & MapGroup & "_Ext_1 D_Ext_1 on (D_Ext_1.BID=D.BID)"
                            SQL = SQL & " Cross Apply (Select Top 1 D1.BID from SysBatch" & MapGroup & " D1 Where D1.BID=D.BID and D1.Status='C') D1"
                            SQL = SQL & " Cross Apply (Select Top 1 A.ANO, A.PolicyNo, A.CertificateNo from Acceptance A Where A.BID=D.BID) A"
                            SQL = SQL & " Outer Apply (Select Top 1 f.Amount from ACCFD f Where f.ANO=a.ANO) f"
                            SQL = SQL & " Where D.BatchNo ='" & Batchno & "'"
                            SQL = SQL & " order by D.BID"
                        End If
                    Case "ALL"
                        SQL = "select SysBatchOri.BID, SysBatchOri.errmsg, case when(SysBatch.status = 'N') then 'Failed' when(SysBatch.status = 'C') then 'Success' when(SysBatch.status = 'W') then 'Not yet process' else 'Failed' end as Status,"
                        SQL = SQL & " SysBatchOri.batchno,SysBatchOri.[LOCATION] as [BRANCH],SysBatchOri.[CONTRACTNO] as [CONTRACTNO], isnull(a.PolicyNo,'') as PolicyNo, isnull(a.Certificateno,'') as CertificateNo, SysBatchOri.[INCEPTION] as [INCEPTION],SysBatchOri.[EXPIRY] as [EXPIRY], "
                        SQL = SQL & " SUM(CASE WHEN A2.code='P' And A2.TYPE In ('DI') THEN A2.Amount ELSE 0 END) As Premium,"
                        SQL = SQL & " SUM(CASE WHEN A2.code='C' And A2.TYPE In ('BA','BB','BD','BP','BO','BC') THEN A2.Amount ELSE 0 END) As Commision,"
                        SQL = SQL & " SUM(CASE WHEN A2.code='O' And A2.TYPE In ('DI') THEN A2.Amount ELSE 0 END) As Discount, "
                        SQL = SQL & " SUM("
                        SQL = SQL & " CASE "
                        SQL = SQL & " WHEN coalesce((A3.Amount-A3.Amount_8-A3.Amount_11-A3.Amount_13-A3.Amount_14-A3.Amount_17-A3.Amount_18),0)<>0 THEN "
                        SQL = SQL & " (A3.Amount_8+A3.Amount_13)*A2.amount/(A3.Amount-A3.Amount_8-A3.Amount_11-A3.Amount_13-A3.Amount_14-A3.Amount_17-A3.Amount_18)"
                        SQL = SQL & " ELSE 0 "
                        SQL = SQL & " End"
                        SQL = SQL & " ) As VAT, "
                        SQL = SQL & " SUM("
                        SQL = SQL & " CASE "
                        SQL = SQL & " WHEN coalesce((A3.Amount-A3.Amount_8-A3.Amount_11-A3.Amount_13-A3.Amount_14-A3.Amount_17-A3.Amount_18),0)<>0 THEN "
                        SQL = SQL & " (A3.Amount_14)*A2.amount/(A3.Amount-A3.Amount_8-A3.Amount_11-A3.Amount_13-A3.Amount_14-A3.Amount_17-A3.Amount_18)"
                        SQL = SQL & " ELSE 0 "
                        SQL = SQL & " End"
                        SQL = SQL & " ) As TAX, "
                        SQL = SQL & " SUM(CASE WHEN A2.code='F' And A2.TYPE In ('DI') THEN A2.Amount ELSE 0 END) As BiayaAdmin, "
                        SQL = SQL & " "
                        SQL = SQL & " SUM(CASE WHEN A2.code='P' And A2.TYPE In ('DI') THEN A2.Amount ELSE 0 END) +"
                        SQL = SQL & " SUM(CASE WHEN A2.code='C' And A2.TYPE In ('BA','BB','BD','BP','BO','BC') THEN A2.Amount ELSE 0 END) +"
                        SQL = SQL & " SUM(CASE WHEN A2.code='O' And A2.TYPE In ('DI') THEN A2.Amount ELSE 0 END) +"
                        SQL = SQL & " SUM("
                        SQL = SQL & " CASE "
                        SQL = SQL & " WHEN coalesce((A3.Amount-A3.Amount_8-A3.Amount_11-A3.Amount_13-A3.Amount_14-A3.Amount_17-A3.Amount_18),0)<>0 THEN "
                        SQL = SQL & " (A3.Amount_8+A3.Amount_13)*A2.amount/(A3.Amount-A3.Amount_8-A3.Amount_11-A3.Amount_13-A3.Amount_14-A3.Amount_17-A3.Amount_18)"
                        SQL = SQL & " ELSE 0 "
                        SQL = SQL & " End"
                        SQL = SQL & " ) +"
                        SQL = SQL & " SUM("
                        SQL = SQL & " CASE "
                        SQL = SQL & " WHEN coalesce((A3.Amount-A3.Amount_8-A3.Amount_11-A3.Amount_13-A3.Amount_14-A3.Amount_17-A3.Amount_18),0)<>0 THEN "
                        SQL = SQL & " (A3.Amount_14)*A2.amount/(A3.Amount-A3.Amount_8-A3.Amount_11-A3.Amount_13-A3.Amount_14-A3.Amount_17-A3.Amount_18)"
                        SQL = SQL & " ELSE 0 "
                        SQL = SQL & " End"
                        SQL = SQL & " ) +"
                        SQL = SQL & " SUM(CASE WHEN A2.code='F' And A2.TYPE In ('DI') THEN A2.Amount ELSE 0 END) As [Nett Premi], "
                        SQL = SQL & " A.CDate as EntryDate,A.ADate as ApprovalDate "
                        SQL = SQL & " from SysBatchOriginalSC SysBatchOri "
                        SQL = SQL & " left join SysBatchSC SysBatch on SysBatch.BID = SysBatchOri.BID  "
                        SQL = SQL & " left join Scheme_Transaction S on s.BatchNo=SysBatch.BatchNo and s.PolicyNo=SysBatch.PolicyNo  and s.ContractNo=SysBatch.ContractNo "
                        SQL = SQL & " left join ACCEPTANCE a on a.ANO=s.ANO "
                        SQL = SQL & " Left Join ACCASS A2 On A2.ANO = a.ANO "
                        SQL = SQL & " Left Join AdmLink A3 on A3.ADMNO = A2.ADMNO "
                        SQL = SQL & " where SysBatchOri.batchno = '" & Batchno & "'"
                        SQL = SQL & " Group By SysBatchOri.BID, SysBatchOri.errmsg, case when(SysBatch.status = 'N') then 'Failed' when(SysBatch.status = 'C') then 'Success' when(SysBatch.status = 'W') then 'Not yet process' else 'Failed' end, SysBatchOri.batchno,SysBatchOri.[LOCATION],SysBatchOri.[CONTRACTNO] , isnull(a.PolicyNo,'') , isnull(a.Certificateno,'') , SysBatchOri.[INCEPTION] ,SysBatchOri.[EXPIRY] ,A.CDate ,A.ADate "
                        SQL = SQL & " Order By SysBatchOri.BID"

                    Case "PC"
                        SQL = "select " & Field & " ,dbo.getrate(a.ano,1,1) as rate1,dbo.getrate(a.ano,2,1) as rate2,dbo.getrate(a.ano,3,1) as rate3,dbo.getrate(a.ano,4,1) as rate4,dbo.getrate(a.ano,5,1) as rate5,dbo.getrate(a.ano,6,1) as rate6,dbo.getrate(a.ano,1,2) as GrossPremi1,dbo.getrate(a.ano,2,2) as GrossPremi2,dbo.getrate(a.ano,3,2) as GrossPremi3,dbo.getrate(a.ano,4,2) as GrossPremi4,dbo.getrate(a.ano,5,2) as GrossPremi5,dbo.getrate(a.ano,6,2) as GrossPremi6,dbo.getrate(a.ano,1,3) as Discount1,dbo.getrate(a.ano,2,3) as Discount2,dbo.getrate(a.ano,3,3) as Discount3,dbo.getrate(a.ano,4,3) as Discount4,dbo.getrate(a.ano,5,3) as Discount5,dbo.getrate(a.ano,6,3) as Discount6,dbo.getrate(a.ano,1,2)-dbo.getrate(a.ano,1,3) as NettPremi1,dbo.getrate(a.ano,2,2)-dbo.getrate(a.ano,2,3) as NettPremi2,dbo.getrate(a.ano,3,2)-dbo.getrate(a.ano,3,3) as NettPremi3,dbo.getrate(a.ano,4,2)-dbo.getrate(a.ano,4,3) as NettPremi4,dbo.getrate(a.ano,5,2)-dbo.getrate(a.ano,5,3) as NettPremi5,dbo.getrate(a.ano,6,2)-dbo.getrate(a.ano,6,3) as NettPremi6,isnull(f.amount,0) as adm,a.certificateno from SysBatchSC SysBatch inner join SysBatchOriginalSC d on SysBatch.BID = d.BID left join Scheme_Transaction s on s.BatchNo=SysBatch.BatchNo and s.PolicyNo=SysBatch.PolicyNo  and s.ContractNo=SysBatch.ContractNo Inner join ACCEPTANCE a on a.ANO=s.ANO left join PAYOR p on a.ANO=p.ano left join ACCFD f on f.ANO=a.ANO where SysBatch.batchno='" & Batchno & "' and SysBatch.status = 'C' order by SysBatch.ContractNo"
                End Select
            End If

            rsd = db.OpenRemoteSet(SQL)
            Dataset = rsd.GetDataSet

            CloseConnection()
            Return True
        Catch ex As Exception
            CloseConnection()
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function Export_Upload_ListView_To_Excel_FeedBack(ByVal dbuser As String, ByVal dbpassword As String, ByVal MapGroup As String, ByVal Batchno As String, ByVal MapCode As String, ByVal Status As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Dim rsd As RemoteSet
        Dim SQL As String = ""
        Try
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If
            
            rsd = db.OpenQueryDef("qDownload_Feedback", adOpenKeyset, Batchno)
            Dataset = rsd.GetDataSet

            CloseConnection()
            Return True
        Catch ex As Exception
            CloseConnection()
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
   Public Function Export_Upload_ListView_To_Excel_DN(ByVal dbuser As String, ByVal dbpassword As String, ByVal MapGroup As String, ByVal BatchNo As String, ByVal MapCode As String, ByVal Status As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim message As String = ""
        Dim SQL As String = ""
        Try
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If
            If Status = "C" Then
                SQL = "select ANO into #Temp from ACCEPTANCE A2 where a2.ATYPE <> 'C' and a2.ANO in (select MAX(ano) from ACCEPTANCE where batchno = '" & BatchNo & "' group by acceptance.oano)"
                If Not db.Execute(SQL) Then Throw New Exception(db.ErrDesc)

                SQL = "select distinct Min(v.Voucher) Voucher ,Min(v.Date) Date,sysbatch.Extra_3 ,'D' as paymenttype,v.refno ,max(v.Nominal_CC -v.Diff_CC) as totalamount, m.id,Min(year(v.date)) as no,a.policyno,a.certificateno from SysBatch" & MapGroup & " sysbatch left join Scheme_Transaction s on SysBatch.ANO=s.LANO and SysBatch.SchemeNO =s.SchemeNo and SysBatch.ContractNo=s.ContractNo and SysBatch.PolicyNo =s.PolicyNo  inner join ACCEPTANCE a on a.batchno = '" & BatchNo & "' and s.ANO=a.ANO and a.LANO=s.LANO and a.POLICYNO=s.PolicyNo and SysBatch.PolicyNo =a.POLICYNO and SysBatch.ANO=a.LANO and A.ANO in (select ANO from #Temp) left join ACCASS c on a.ANO=c.ANO left join AdmLink m on c.ADMNO=m.admno   left join Voucher v on v.Voucher =m.VOUCHER  where   'C' in (select Status from SysBatchH where BatchNo='" & BatchNo & "') and c.TYPE ='DI' and SysBatch.batchno='" & BatchNo & "' and m.WRITEF=0 and c.CODE ='P' And YEAR(V.Date) >= YEAR(GETDATE()) Group by sysbatch.Extra_3, v.refno, m.id, a.policyno, a.certificateno  order by Min(year(v.date))"
            ElseIf Status = "P" Then
                SQL = "select distinct  v.Voucher ,v.Date,sysbatch.Extra_3 ,'D' as paymenttype,v.refno ,v.Nominal_CC -v.Diff_CC as totalamount, m.id,year(v.date) as no,a.policyno,a.certificateno from SysBatch" & MapGroup & " sysbatch left join Scheme_Transaction s on SysBatch.ANO=s.LANO and SysBatch.SchemeNO =s.SchemeNo and SysBatch.ContractNo=s.ContractNo and SysBatch.PolicyNo =s.PolicyNo  inner join ACCEPTANCE a on a.batchno = '" & BatchNo & "' and s.ANO=a.ANO and a.LANO=s.LANO and a.POLICYNO=s.PolicyNo and SysBatch.PolicyNo =a.POLICYNO and SysBatch.ANO=a.LANO  left join ACCASS c on a.ANO=c.ANO left join AdmLink m on c.ADMNO=m.admno   left join Voucher v on v.Voucher =m.VOUCHER  where   'C' in (select Status from SysBatchH where BatchNo='" & BatchNo & "') and c.TYPE ='DI' and SysBatch.batchno='" & BatchNo & "' and m.WRITEF=0 and c.CODE ='P' And YEAR(V.Date) >= YEAR(GETDATE()) order by v.date"
            Else
                SQL = "select distinct  v.Voucher ,v.Date,sysbatch.Extra_3 ,'D' as paymenttype,v.refno ,v.Nominal_CC -v.Diff_CC as totalamount, m.id,year(v.date) as no,a.policyno,a.certificateno from SysBatch" & MapGroup & " sysbatch left join Scheme_Transaction s on SysBatch.ANO=s.LANO and SysBatch.SchemeNO =s.SchemeNo and SysBatch.ContractNo=s.ContractNo and SysBatch.PolicyNo =s.PolicyNo  inner join ACCEPTANCE a on a.batchno = '" & BatchNo & "' and s.ANO=a.ANO and a.LANO=s.LANO and a.POLICYNO=s.PolicyNo and SysBatch.PolicyNo =a.POLICYNO and SysBatch.ANO=a.LANO  left join ACCASS c on a.ANO=c.ANO left join AdmLink m on c.ADMNO=m.admno   left join Voucher v on v.Voucher =m.VOUCHER  where   'C' in (select Status from SysBatchH where BatchNo='" & BatchNo & "') and c.TYPE ='DI' and SysBatch.batchno='" & BatchNo & "' and m.WRITEF=0 and c.CODE ='P' order by v.date"
            End If

            rs = db.OpenRemoteSet(SQL)
            Dataset = rs.GetDataSet

            CloseConnection()
            Return True
        Catch ex As Exception
            CloseConnection()
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function Export_Upload_ListView_To_Excel_CN(ByVal dbuser As String, ByVal dbpassword As String, ByVal MapGroup As String, ByVal BatchNo As String, ByVal MapCode As String, ByVal Status As String, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet
        Dim message As String = ""
        Dim SQL As String = ""
        Try
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If
            If Left(MapCode, 2) = "BC" Then
                SQL = "SELECT distinct adm1.docno,Max(SysBatch.Batchno) BatchNo,Max(adm1.Refno) Refno, count(adm1.voucher) vou into #B from SysBatchBC SysBatch inner join ACCEPTANCE a on SysBatch.batchno =a.batchno and sysbatch.policyno = a.policyno and sysbatch.CERTIFICATENO = a.CERTIFICATENO inner join AdmLink Adm1 on adm1.Refno=A.RefNo and adm1.amount_1 <> 0 and adm1.amount_3 <> 0  and adm1.TYPE ='DI' and adm1.extra_4 = 'cancellation' and 'C' in (select Status from SysBatchH where BatchNo='" & BatchNo & "') and SysBatch.batchno='" & BatchNo & "' and adm1.WRITEF=0 group by adm1.docno having count(adm1.voucher) = 1"
                db.Execute(SQL)
                db.Execute("GO")
                SQL = " Select * into #c from ( Select distinct min(SysBatch.BID) as BID, min(adm1.Voucher) as Voucher ,min(adm1.Date) as Date, min(sysbatch.Extra_3) as Batch_No,'C' as paymenttype,min(adm1.refno) as refno ,abs(sum(adm1.amount)) as totalamount, min(adm1.id) ID, min(year(adm1.date)) as no, min(a.policyno) as policyno,min(a.certificateno) as certificateno,min(SysBatch.EXTRA_2) as EXTRA_2, min(a.pdate) as ENDORSE_Date from SysBatchBC SysBatch inner join ACCEPTANCE a on SysBatch.batchno =a.batchno and sysbatch.policyno = a.policyno and sysbatch.CERTIFICATENO = a.CERTIFICATENO inner join #B B on SysBatch.batchno = B.batchno inner join AdmLink Adm1 on Adm1.Refno = A.Refno and adm1.Docno=B.Docno and adm1.AMOUNT_1 <> 0 and adm1.amount_3 <> 0 and SysBatch.batchno='" & BatchNo & "' and adm1.WRITEF=0 group by  adm1.refno  Union ALL SELECT distinct min(SysBatch.BID) as BID, min(adm1.Voucher) as Voucher ,min(adm1.Date) as Date, min(sysbatch.Extra_3) as Batch_No,'C' as paymenttype,min(adm1.refno) as refno ,abs(sum(adm1.amount)) as totalamount, min(adm1.id) ID, min(year(adm1.date)) as no, min(a.policyno) as policyno,min(a.certificateno) as certificateno,min(SysBatch.EXTRA_2) as EXTRA_2, min(a.pdate) as ENDORSE_Date from SysBatchBC SysBatch inner join ACCEPTANCE a on SysBatch.batchno =a.batchno and sysbatch.policyno = a.policyno and sysbatch.CERTIFICATENO = a.CERTIFICATENO inner join AdmLink Adm1 on Adm1.Refno = A.Refno and adm1.extra_4 = 'cancellation' Inner join admlink Adm2 on Adm1.docno = Adm2.Docno and Adm1.Billdate < Adm2.billdate  and adm2.amount_1 <> 0 and adm2.amount_3 = 0 and Adm1.refno not in (select B.refno from #B B) and SysBatch.batchno='" & BatchNo & "' and adm1.WRITEF=0 group by  adm1.refno) A "
                db.Execute(SQL)
                db.Execute("GO")
                SQL = " Select Voucher , Date, Batch_No, paymenttype,refno ,totalamount, ID, no, policyno,certificateno,EXTRA_2, ENDORSE_Date from (select * from #c union all SELECT distinct  min(sysbatch.bid), min(adm1.Voucher) as Voucher ,min(adm1.Date) as Date, min(sysbatch.Extra_3) as Batch_No,'C' as paymenttype,min(adm1.refno) as refno ,abs(sum(adm1.amount)) as totalamount, min(adm1.id) ID, min(year(adm1.date)) as no, min(a.policyno) as policyno,min(a.certificateno) as certificateno,min(SysBatch.EXTRA_2) as EXTRA_2, min(a.pdate) as ENDORSE_Date from SysBatchBC SysBatch inner join ACCEPTANCE a on SysBatch.batchno =a.batchno and sysbatch.policyno = a.policyno and sysbatch.CERTIFICATENO = a.CERTIFICATENO inner join AdmLink Adm1 on Adm1.Refno = A.Refno and adm1.extra_4 = 'cancellation' and Adm1.refno not in (select refno from #C) and SysBatch.batchno='" & BatchNo & "' and adm1.WRITEF=0 group by  adm1.refno ) A order by BID "
            Else
                SQL = "select distinct v.Voucher ,v.Date,SysBatch.batchno,'D' as paymenttype,a.REMARKS ,v.Nominal_CC -v.Diff_CC as totalamount, m.id  from SysBatch" & MapGroup & " SysBatch left join Scheme_Transaction s on SysBatch.ANO=s.LANO and SysBatch.SchemeNO =s.SchemeNo and SysBatch.ContractNo=s.ContractNo and SysBatch.PolicyNo =s.PolicyNo  inner join ACCEPTANCE a on a.Batchno = '" & BatchNo & "' and s.ANO=a.ANO and a.LANO=s.LANO and a.POLICYNO=s.PolicyNo and SysBatch.PolicyNo =a.POLICYNO and SysBatch.ANO=a.LANO  left join ACCASS c on a.ANO=c.ANO left join RCOVER r on r.ANO=a.ano  left join AdmLink m on c.ADMNO=m.admno   left join Voucher v on v.Voucher =m.VOUCHER  where 'C' in (select Status from SysBatchH where BatchNo='" & BatchNo & "') and c.TYPE ='DI' and SysBatch.batchno='" & BatchNo & "' and m.WRITEF=0 and c.CODE ='P'"
            End If

            rs = db.OpenRemoteSet(SQL)
            Dataset = rs.GetDataSet

            CloseConnection()
            Return True

        Catch ex As Exception
            CloseConnection()
            ErrMsg = ex.Message
            Return False
        End Try
    End Function
    
    <WebMethod(EnableSession:=False)> _
    Public Function ConvertFile(ByVal dbuser As String, ByVal dbPassword As String, ByVal inparam() As String, ByRef OutParam As String, ByRef errMsg As String) As Boolean
        Try
            Dim input As String = inparam(0)
            Dim output As String = inparam(1)
            Dim myProcess As Diagnostics.Process = New Diagnostics.Process()
            'OutParam = output
            myProcess.StartInfo.FileName = GetAppSetting("CareConvert") & inparam(2) & ".exe"
            myProcess.StartInfo.Arguments = "/I:" & inparam(0) & " /O:" & output
            myProcess.StartInfo.UseShellExecute = False
            'myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            output = myProcess.StandardOutput.ReadToEnd()
            Return True
        Catch ex As Exception
            errMsg = ex.Message

            If errMsg = "" Then errMsg = ex.Message
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)> _
    Public Function ProcessSuspend(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()
            Dim batchno As String = InParam(0)

            If Not db.Execute("update SysBatchh set status='S' where batchno='" & batchno & "' and status in ('P','T')") Then
                Throw New Exception(db.ErrDesc)
            End If
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
    <WebMethod(EnableSession:=False)> _
    Public Function ProcessResume(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()
            Dim batchno As String = InParam(0)
            Dim rs As RemoteSet = Nothing
            rs = db.OpenRemoteSet("select * from SysBatchH where batchno='" & batchno & "' and status='S'")
            If Not rs Is Nothing Then
                If rs.RecordCount > 0 Then
                    rs.Edit()
                    rs.fields("Status") = "P"
                    If Not rs.Update Then
                        Throw New Exception(rs.ErrDesc)
                    End If
                End If
            Else
                Throw New Exception(rs.ErrDesc)
            End If
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
    <WebMethod(EnableSession:=False)> _
    Public Function SavePriority(ByVal dbuser As String, ByVal dbpassword As String, ByVal BatchNo As String, ByVal Priority As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            Dim Sql As String
            Dim rs As RemoteSet
            Sql = "Select * from SysBatchh where batchno='" & BatchNo & "'"
            rs = db.OpenRemoteSet(Sql)
            If Not rs.Eof Then
                rs.Edit()
                rs.fields("Priority") = tostr(Priority)
                If Not rs.Update Then
                    Throw New Exception(rs.ErrDesc)
                End If
            Else
                Throw New Exception("Please check batchno.")
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
    Public Function ForceComplete(ByVal dbuser As String, ByVal dbpassword As String, ByVal BatchNo As String, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbuser, dbpassword, ErrMsg) Then
                Return False
            End If
            db.Begintrans()
            Dim Sql As String
            Dim rs As RemoteSet
            Sql = "Select * from SysBatchh where batchno='" & BatchNo & "' and status='P'"
            rs = db.OpenRemoteSet(Sql)
            If Not rs.Eof Then
                rs.Edit()
                rs.fields("Status") = "C"
                If Not rs.Update Then
                    Throw New Exception(rs.ErrDesc)
                End If
            Else
                Throw New Exception("Please check batchno.")
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
    Public Function Verify_Upload_Profile(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            Dim i As Integer
            Dim StrMsg As String
            Dim rs As RemoteSet
            Dim LineNo As String
            Dim CheckF As Boolean
            Dim rsMaxLength As RemoteSet
            Dim LengthRefID As Integer
            Dim LengthName As Integer
            Dim LengthTaxID As Integer
            Dim Status As Integer

            Dim PType As String
            Dim LOB As String
            Dim IntermediaryF As Integer
            Dim CorporateF As Integer
            Dim ID As String
            Dim RefID As String
            Dim PName As String
            Dim FName As String
            Dim MName As String
            Dim LName As String
            Dim Country As String
            Dim City As String
            Dim Province As String
            Dim Area As String
            Dim Branch As String
            Dim BIType As String
            Dim HSCode As String
            Dim BankCode As String
            Dim BankBranch As String
            Dim MOID As String
            Dim Address As String
            Dim Currency As String
            Dim RefType As String
            Dim BCTransferBy As Integer
            Dim PRDays As Integer

            Dim TaxID As String
            Dim TaxType As String

            Dim TaxF As Boolean
            Dim UploadAgencyF As Boolean

            ' Get Max Length From SysDefinitions
            rsMaxLength = db.OpenRemoteSet("Select Length from SysDefinitions Where TBName='Profile' and ColName='RefID'")
            LengthRefID = rsMaxLength.fields("Length")

            rsMaxLength = db.OpenRemoteSet("Select Length from SysDefinitions Where TBName='Profile' and ColName='Name'")
            LengthName = rsMaxLength.fields("Length")

            LengthTaxID = toval(GetAppSetting("CareTaxIDMaxLength"))
            ' End Get Max Length

            CheckF = False
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                ReDim Preserve OutParam(i)
                CheckF = True
                StrMsg = ""
                Status = 1

                PType = tostr(GetItemValue("PType", Dataset, i))
                LOB = tostr(GetItemValue("LOB", Dataset, i))
                IntermediaryF = toval(GetItemValue("IntermediaryF", Dataset, i))
                CorporateF = toval(GetItemValue("CorporateF", Dataset, i))
                ID = tostr(GetItemValue("ID", Dataset, i))
                RefID = tostr(GetItemValue("RefID", Dataset, i))
                PName = tostr(GetItemValue("Name", Dataset, i))
                FName = tostr(GetItemValue("FirstName", Dataset, i))
                MName = tostr(GetItemValue("MidName", Dataset, i))
                LName = tostr(GetItemValue("LastName", Dataset, i))
                If PName = "" Then
                    PName = LTrim(FName & " " & MName & " " & LName)
                End If

                Country = tostr(GetItemValue("Country", Dataset, i))
                City = tostr(GetItemValue("City", Dataset, i))
                Province = tostr(GetItemValue("Province", Dataset, i))
                Area = tostr(GetItemValue("Area", Dataset, i))
                Branch = tostr(GetItemValue("Branch", Dataset, i))
                BIType = tostr(GetItemValue("BIType", Dataset, i))
                HSCode = tostr(GetItemValue("HSCode", Dataset, i))
                BankCode = tostr(GetItemValue("BankCode", Dataset, i))
                BankBranch = tostr(GetItemValue("BankBranch", Dataset, i))
				Currency = tostr(GetItemValue("Currency", Dataset, i))
                MOID = tostr(GetItemValue("MOID", Dataset, i))
                RefType = tostr(GetItemValue("RefType", Dataset, i))
                BCTransferBy = toval(GetItemValue("BCTransferBy", Dataset, i))
                PRDays = toval(GetItemValue("PRDays", Dataset, i))

                'Taxation
                TaxID = tostr(GetItemValue("TaxID", Dataset, i))
                TaxType = tostr(GetItemValue("TaxType", Dataset, i))
                TaxF = toval(GetItemValue("TaxF", Dataset, i))
                Address = tostr(GetItemValue("Address_1", Dataset, i))

                UploadAgencyF = toval(GetItemValue("UploadAgencyF", Dataset, i))
                'Cek Duplicate Viet
                If CBool(toval(GetAppSetting("CareProfileVietF"))) Then
                    Dim ID_Type, ID_No As String
                    ID_Type = tostr(GetItemValue("ID_Type", Dataset, i))
                    ID_No = tostr(GetItemValue("ID_No", Dataset, i))
                    'Check Duplicate Profile + Address
                    If Not ID <> "" And CorporateF = 1 Then
                        rs = db.OpenRemoteSet("Select * from Profile Where Name='" & PName & "' and Address_1='" & Address & "'")
                        If Not rs.Eof Then
                            StrMsg = StrMsg & "Duplicate Profile Entry (see ID : " & rs.fields("ID") & " or Address: " & rs.fields("Address_1") & ") !" & vbCrLf
                        End If
                    Else
                        'Check Duplicate ID Non corporate
                        If ID_Type <> "" And ID_No <> "" And CorporateF = 0 And Not ID <> "" Then
                            rs = db.OpenRemoteSet("Select Top 1 ID from Profile_Ext Where ID_Type='" & ID_Type & "' and ID_No='" & ID_No & "'")
                            If Not rs.Eof Then
                                StrMsg = StrMsg & "Duplicate ID Number (see ID : " & rs.fields("ID") & " !" & vbCrLf
                            End If
                        End If
                    End If
                End If

                ' Check Mandatory Fields
                If CBool(toval(GetAppSetting("ApplyVerifyProfile"))) Then

                    ' Profile Name
                    If PName = "" Then
                        StrMsg = StrMsg & "Incorrect Profile Name !" & vbCrLf
                    End If

                    ' Address
                    If tostr(GetItemValue("Address_1", Dataset, i)) = "" Then
                        StrMsg = StrMsg & "Incorrect Address !" & vbCrLf
                    End If

                    ' Country
                    If CBool(toval(GetAppSetting("CountryMandatoryF"))) And Country = "" Then
                        StrMsg = StrMsg & "Incorrect Country !" & vbCrLf
                    End If

                    If Country <> "" Then
                        rs = db.OpenRemoteSet("Select Country, Description from Country Where country='" & Country & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Invalid Country Reference!" & Country & vbCrLf
                        End If
                    End If

                    ' Province
                    If Province = "" Then
                        StrMsg = StrMsg & "Incorrect Province !" & vbCrLf
                    Else
                        rs = db.OpenRemoteSet("Select Province from Province Where Province='" & Province & "' and Country ='" & Country & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Invalid Province Reference !" & Province & vbCrLf
                        End If
                    End If

                    ' City
                    If City = "" Then
                        StrMsg = StrMsg & "Incorrect City !" & vbCrLf
                    Else
                        rs = db.OpenRemoteSet("Select City from City Where City='" & City & "' and Province ='" & Province & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Invalid City Reference !" & City & vbCrLf
                        End If
                    End If

                    ' Profile Type
                    rs = db.OpenRemoteSet("Select Method, Name from Method where MType = 'PT' and Method='" & PType & "'")
                    If rs.Eof Then
                        StrMsg = StrMsg & "Incorrect Profile Type ('C','D','I','M','O','T','F') !" & vbCrLf
                    End If

                    If CorporateF = 0 Then
                        ' BirthDate
                        If tostr(GetItemValue("BirthDate", Dataset, i)) = "" Or Not IsDate(GetItemValue("BirthDate", Dataset, i)) Then
                            StrMsg = StrMsg & "Invalid Birth Date !" & vbCrLf
                        End If

                    End If

                Else
                    ' Profile Type
                    rs = db.OpenRemoteSet("Select Method, Name from Method where MType = 'PT' and Method='" & PType & "'")
                    If rs.Eof Then
                        StrMsg = StrMsg & "Incorrect Profile Type ('C','D','I','M','O','T','F') !" & vbCrLf
                    End If

                    ' Branch
                    If CBool(toval(GetAppSetting("ThaiProfile"))) Then
                        If Branch <> "" And Not isNumberString(Branch) Then
                            StrMsg = StrMsg & "Incorrect Branch ID, Must Contain Number Only (0..9)!" & vbCrLf
                        End If
                        ' Length Branch
                        If Len(Branch) > 5 Then
                            StrMsg = StrMsg & "Incorrect Branch Length (Max. 5 character(s)) !" & vbCrLf
                        End If
                    Else
                        If Branch <> "" Then
                            rs = db.OpenRemoteSet("Select Branch, Name from Branch where Op_Flag = 1 and Branch='" & Branch & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Branch ID !" & vbCrLf
                            End If
                        End If
                    End If

                    ' LOB
                    rs = db.OpenRemoteSet("Select LOB from LOB Where LOB='" & LOB & "'")
                    If rs.Eof Then
                        StrMsg = StrMsg & "Incorrect Line Of Business !" & vbCrLf
                    End If

                    ' Currency
                    rs = db.OpenRemoteSet("Select Currency from Currency Where Currency='" & Currency & "'")
                    If rs.Eof Then
                        StrMsg = StrMsg & "Incorrect Currency !" & vbCrLf
                    End If

                    If RefType <> "" And RefType.ToUpper <> "BANK" And RefType.ToUpper <> "VIRTUAL" Then
                        StrMsg = StrMsg & "Incorrect RefType !" & vbCrLf
                    End If

                    ' Bank Charges
                    If BCTransferBy <> 0 And BCTransferBy <> 1 And BCTransferBy <> 2 Then
                        StrMsg = StrMsg & "Incorrect Bank Charges By !" & vbCrLf
                    End If

                    If PRDays < 0 Then
                        StrMsg = StrMsg & "Incorrect Payment Reminder !" & vbCrLf
                    End If

                    ' IntermediaryF
                    If IntermediaryF <> 0 And IntermediaryF <> 1 Then
                        StrMsg = StrMsg & "Incorrect IntermediaryF (0 / 1) !" & vbCrLf
                    End If

                    ' CorporateF
                    If CorporateF <> 0 And CorporateF <> 1 Then
                        StrMsg = StrMsg & "Incorrect CorporateF (0 / 1) !" & vbCrLf
                    End If

                    ' Duplicate Ref ID or Profile
                    If ID = "" Then
                        If RefID <> "" Then
                            rs = db.OpenRemoteSet("Select * from Profile Where RefID='" & RefID & "'")
                        Else
                            rs = db.OpenRemoteSet("Select * from Profile Where Name='" & PName & "' and TaxID='" & TaxID & "'")
                        End If
                        If Not rs.Eof Then
                            StrMsg = StrMsg & "Duplicate Profile Entry (see ID : " & rs.fields("ID") & " or Ref. ID : " & rs.fields("RefID") & ") !" & vbCrLf
                        End If
                    End If

                    ' Profile Name
                    If PName = "" Then
                        StrMsg = StrMsg & "Incorrect Profile Name !" & vbCrLf
                    End If

                    ' Address
                    If tostr(GetItemValue("Address_1", Dataset, i)) = "" Then
                        StrMsg = StrMsg & "Incorrect Address !" & vbCrLf
                    End If

                    ' City
                    'If tostr(Dataset.Tables(0).Rows(i)("City")) = "" Then
                    '    StrMsg = StrMsg & "Incorrect City !" & vbCrLf
                    'End If

                    ' Telephone / Mobile Number
                    If tostr(GetItemValue("Phone_1", Dataset, i)) = "" AndAlso tostr(GetItemValue("Phone_2", Dataset, i)) = "" AndAlso tostr(GetItemValue("Mobile_1", Dataset, i)) = "" AndAlso tostr(GetItemValue("Mobile_2", Dataset, i)) = "" Then
                        StrMsg = StrMsg & "Incorrect Telephone or Mobile Phone !" & vbCrLf
                    End If

                    ' Area
                    If Area <> "" Then
                        rs = db.OpenRemoteSet("Select Area from Area Where Area='" & Area & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Incorrect Area !" & vbCrLf
                        End If
                    End If

                    'BI Type
                    If BIType <> "" Then
                        rs = db.OpenRemoteSet("Select BIType from BIType Where BIType='" & BIType & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Incorrect BI Type !" & vbCrLf
                        End If
                    End If

                    'HS Code
                    If HSCode <> "" Then
                        rs = db.OpenRemoteSet("Select HSCode from HSCode Where HSCode='" & HSCode & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Incorrect HS Code !" & vbCrLf
                        End If
                    End If

                    'Bank Code
                    If BankCode <> "" Then
                        rs = db.OpenRemoteSet("Select Bank from Bank Where Bank='" & BankCode & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Incorrect Bank Code !" & vbCrLf
                        End If
                    End If

                    'Bank Branch
                    If BankBranch <> "" Then
                        rs = db.OpenRemoteSet("Select BankBranch from BankBranch Where Bank='" & BankCode & "' and BankBranch = '" & BankBranch & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Incorrect Bank Branch !" & vbCrLf
                        End If
                    End If

                    If MOID <> "" Then
                        rs = db.OpenRemoteSet("Select ID from sysuser Where Role='MARKETING OFFICER' and ID='" & MOID & "'")
                        If rs.Eof Then
                            StrMsg = StrMsg & "Incorrect Marketing Officer ID !" & vbCrLf
                        End If
                    End If

                    'Validate Agency Data
                    If UploadAgencyF Then
                        If tostr(GetItemValue("AgentName", Dataset, i)) = "" Then
                            StrMsg = StrMsg & "Incorrect Agent Name !" & vbCrLf
                        End If

                        If tostr(GetItemValue("AgentType", Dataset, i)) = "" Then
                            StrMsg = StrMsg & "Incorrect Agent Type !" & vbCrLf
                        Else
                            rs = db.OpenRemoteSet("select TOP 1 * From AgentType Where AgentType = '" & tostr(GetItemValue("AgentType", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Agent Type !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("AgentLevel", Dataset, i)) = "" Then
                            StrMsg = StrMsg & "Incorrect Agent Level !" & vbCrLf
                        Else
                            rs = db.OpenRemoteSet("select TOP 1 * From AgentLevel Where AgentLevel = '" & tostr(GetItemValue("AgentLevel", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Agent Level !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("RecruiterID", Dataset, i)) = "" Then
                            StrMsg = StrMsg & "Incorrect Recruiter !" & vbCrLf
                        Else
                            rs = db.OpenRemoteSet("select TOP 1 * From Profile Where ID = '" & tostr(GetItemValue("RecruiterID", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Recruiter !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("ULAgentID", Dataset, i)) <> "" Then
                            rs = db.OpenRemoteSet("select TOP 1 * From AIF Where ID = '" & tostr(GetItemValue("ULAgentID", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Up Line Agent !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("AgentStatus", Dataset, i)) <> "" Then
                            If UCase(tostr(GetItemValue("AgentStatus", Dataset, i))) <> "A" And UCase(tostr(GetItemValue("AgentStatus", Dataset, i))) <> "N" Then
                                StrMsg = StrMsg & "Incorrect Agent Status !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("RegisterStartDate", Dataset, i)) <> "" Then
                            If Not IsDate(GetItemValue("RegisterStartDate", Dataset, i)) Then
                                StrMsg = StrMsg & "Invalid Register Start Date Agent !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("RegisterExpiryDate", Dataset, i)) <> "" Then
                            If Not IsDate(GetItemValue("RegisterExpiryDate", Dataset, i)) Then
                                StrMsg = StrMsg & "Invalid Register Expiry Date Agent !" & vbCrLf
                            End If
                        End If

                    End If

                    ' No need to proceed data verify if any mandatory fields not fulfilled
                    If StrMsg <> "" Then
                        Status = 0
                    End If

                    If Status = 1 Then
                        ' Proceed data verify for fields validation
                        ' Length RefID
                        If RefID <> "" And Len(RefID) > LengthRefID Then
                            StrMsg = StrMsg & "Incorrect Ref. ID Length (Max. " & LengthRefID & " character(s)) !" & vbCrLf
                        End If

                        ' Length Name
                        If Len(PName) > LengthName Then
                            StrMsg = StrMsg & "Incorrect Name Length (Max. " & LengthName & " character(s)) !" & vbCrLf
                        End If

                        ' Phone_1                                        
                        If tostr(GetItemValue("Phone_1", Dataset, i)) <> "" And Not isNumberString(tostr(GetItemValue("Phone_1", Dataset, i)), "-") Then
                            StrMsg = StrMsg & "Incorrect Phone_1, Must Contain Number Only (0..9) or '-' !" & vbCrLf
                        End If
                        ' Phone_2
                        If tostr(GetItemValue("Phone_2", Dataset, i)) <> "" And Not isNumberString(tostr(GetItemValue("Phone_2", Dataset, i)), "-") Then
                            StrMsg = StrMsg & "Incorrect Phone_2, Must Contain Number Only (0..9) or '-' !" & vbCrLf
                        End If
                        ' Mobile_1
                        If tostr(GetItemValue("Mobile_1", Dataset, i)) <> "" And Not isNumberString(tostr(GetItemValue("Mobile_1", Dataset, i)), "-") Then
                            StrMsg = StrMsg & "Incorrect Mobile_1, Must Contain Number Only (0..9) or '-' !" & vbCrLf
                        End If
                        ' Mobile_2
                        If tostr(GetItemValue("Mobile_2", Dataset, i)) <> "" And Not isNumberString(tostr(GetItemValue("Mobile_2", Dataset, i)), "-") Then
                            StrMsg = StrMsg & "Incorrect Mobile_2, Must Contain Number Only (0..9) or '-' !" & vbCrLf
                        End If
                        ' Fax_1
                        If tostr(GetItemValue("Fax_1", Dataset, i)) <> "" And Not isNumberString(tostr(GetItemValue("Fax_1", Dataset, i)), "-") Then
                            StrMsg = StrMsg & "Incorrect Fax_1, Must Contain Number Only (0..9) or '-' !" & vbCrLf
                        End If
                        ' Fax_2
                        If tostr(GetItemValue("Fax_2", Dataset, i)) <> "" And Not isNumberString(tostr(GetItemValue("Fax_2", Dataset, i)), "-") Then
                            StrMsg = StrMsg & "Incorrect Fax_2, Must Contain Number Only (0..9) or '-' !" & vbCrLf
                        End If

                        ' Validation for Corporate
                        If CorporateF = 1 Then
                            Dim CGroup As String
                            Dim SCGroup As String

                            CGroup = tostr(GetItemValue("CGroup", Dataset, i))
                            SCGroup = tostr(GetItemValue("SCGroup", Dataset, i))

                            ' PIC
                            If tostr(GetItemValue("PIC", Dataset, i)) = "" Then
                                StrMsg = StrMsg & "Incorrect PIC !" & vbCrLf
                            End If

                            ' CType
                            If tostr(GetItemValue("CType", Dataset, i)) = "" Then
                                StrMsg = StrMsg & "Incorrect Company Type !" & vbCrLf
                            End If

                            ' CGroup
                            rs = db.OpenRemoteSet("Select CGroup from CGroup Where CGroup='" & CGroup & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Company Group !" & vbCrLf
                            End If

                            ' SCGroup
                            If SCGroup <> "" Then
                                rs = db.OpenRemoteSet("Select SCGroup from SCGroup Where SCGroup='" & SCGroup & "' and CGroup='" & CGroup & "'")
                                If rs.Eof Then
                                    StrMsg = StrMsg & "Incorrect Sub Company Group !" & vbCrLf
                                End If
                            End If

                        
                        End If

                        ' Validation for Non-Corporate and Profile Type Direct
                        If CorporateF = 0 And PType = "D" Then
                            Dim Gender As String
                            Dim ID_Type As String
                            Dim ID_No As String

                            Gender = UCase(tostr(GetItemValue("Gender", Dataset, i)))
                            ID_Type = tostr(GetItemValue("ID_Type", Dataset, i))
                            ID_No = tostr(GetItemValue("ID_No", Dataset, i))

                            ' Gender
                            If Gender <> "MALE" And Gender <> "FEMALE" Then
                                StrMsg = StrMsg & "Incorrect Gender !" & vbCrLf
                            End If

                            ' Birth Date
                            If tostr(GetItemValue("BirthDate", Dataset, i)) = "" Or Not IsDate(GetItemValue("BirthDate", Dataset, i)) Then
                                StrMsg = StrMsg & "Invalid Birth Date !" & vbCrLf
                            End If
                            ' Birth Place
                            If tostr(GetItemValue("BirthPlace", Dataset, i)) = "" Then
                                StrMsg = StrMsg & "Invalid Birth Place !" & vbCrLf
                            End If

                            ' ID No
                            If ID = "" Then
                                If ID_Type <> "" And ID_No <> "" Then
                                    rs = db.OpenRemoteSet("Select Top 1 ID from Profile_Ext Where ID_Type='" & ID_Type & "' and ID_No='" & ID_No & "'")
                                    If Not rs.Eof Then
                                        StrMsg = StrMsg & "Duplicate ID Number (see ID : " & rs.fields("ID") & " !" & vbCrLf
                                    End If
                                End If
                            End If
                        End If

                        ' Taxation
                        If TaxF And TaxType = "" Then
                            StrMsg = StrMsg & "Incorrect Tax Type !" & vbCrLf
                        End If

                        ' Tax Type
                        If TaxType <> "" Then
                            rs = db.OpenRemoteSet("Select Top 1 TaxType from TaxType Where TaxType='" & TaxType & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Tax Type !" & vbCrLf
                            End If
                        End If

                        ' TaxID
                        'validation TAX ID Length
                        If TaxID <> "" And (Not isNumberString(TaxID, "-") Or Len(TaxID) <> LengthTaxID) Then
                           StrMsg = StrMsg & "Incorrect Tax ID, Must Contain Number Only and must " & LengthTaxID & " Chars" & vbCrLf
                        End If

                        ' VAT
                        If tostr(GetItemValue("VAT", Dataset, i)) <> "" And Not IsNumeric(tostr(GetItemValue("VAT", Dataset, i))) Then
                            StrMsg = StrMsg & "Incorrect VAT !" & vbCrLf
                        End If
                        ' TAX
                        If tostr(GetItemValue("TAX", Dataset, i)) <> "" And Not IsNumeric(tostr(GetItemValue("TAX", Dataset, i))) Then
                            StrMsg = StrMsg & "Incorrect Tax !" & vbCrLf
                        End If
                        ' TAX Loading
                        If tostr(GetItemValue("TaxLoading", Dataset, i)) <> "" And Not IsNumeric(tostr(GetItemValue("TaxLoading", Dataset, i))) Then
                            StrMsg = StrMsg & "Incorrect Tax Loading !" & vbCrLf
                        End If
                    End If
                End If

                'CEK Master Customer Consent
                Dim IDConsent As String = ""

                For j = 0 To Dataset.Tables(0).Columns.Count() - 1
                    If Dataset.Tables(0).Columns(j).ColumnName.StartsWith("IDConsent") Then
                        If tostr(GetItemValue(Dataset.Tables(0).Columns(j).ColumnName, Dataset, i)) <> "" Then
                            IDConsent = tostr(GetItemValue(Dataset.Tables(0).Columns(j).ColumnName, Dataset, i))
                            rs = db.OpenRemoteSet("Select ConsentID from Consent Where ConsentID='" & IDConsent & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect ID Customer Consent !" & vbCrLf
                            End If
                        End If
                    End If
                Next


                OutParam(i) = StrMsg
            Next

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


    <WebMethod(EnableSession:=False)> _
    Public Function Process_Upload_Profile(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            Dim i As Integer
            Dim rs As RemoteSet
            Dim rsDst As RemoteSet
            Dim rsDst2 As RemoteSet
            Dim rsDst3 As RemoteSet
            Dim LineNo As String

            Dim PName As String
            Dim FName As String
            Dim mName As String
            Dim lName As String
            Dim ID As String
            Dim RefID As String
            Dim XProfileID As String
            Dim Branch As String
            Dim PType As String
            Dim LOB As String
            Dim IntermediaryF As Integer
            Dim CorporateF As Integer
			Dim VatSubsidiesF As Integer
            Dim TaxID As String
            Dim TaxIDOLD As String
            Dim TaxType As String
            Dim TaxLoading As Double
            Dim BankCode As String
            Dim BankBranch As String
            Dim BankAccountNo As String
            Dim BankAccountName As String
            Dim BankName As String
            Dim RefType As String
            Dim BCTransferBy As Integer
            Dim PRDays As Integer
            Dim Remarks As String
            Dim br_branch As String = ""
            Dim br_Currency As String = ""
            Dim Currency As String = ""
            Dim SYNCDATE As String = ""
            Dim SYNCTIME As String = ""
            Dim message As String = ""
            Dim TAXF As Boolean
            Dim VATF As Boolean
            Dim BVATF As Boolean
            Dim TVATF As Boolean
            Dim VATCOMMEXEMPTIF As Boolean
            Dim RedFlag As Boolean
            Dim ConsentOnDataKeepByCompanyF As Boolean
            Dim ConsentOnDataPrivacyToThirdPartyF As Boolean
            Dim ConsentOnProductAndServiceOfferingF As Boolean
            Dim Remark As String


            LineNo = ""



            rs = db.OpenRemoteSet("Select Currency, Branch from SysTab")
            br_Currency = tostr(rs.fields("Currency"))
            br_branch = tostr(rs.fields("Branch"))

            If Dataset.Tables(0).Columns("ID") Is Nothing Then
                ' Add Column Profile ID
                Dim dc As System.Data.DataColumn
                dc = New DataColumn
                dc.ColumnName = "ID"
                dc.DataType = Type.GetType("System.String")
                dc.DefaultValue = ""
                Dataset.Tables(0).Columns.Add(dc)
            End If

            Dim rsP As RemoteSet
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Try
                    db.Begintrans()
                    ReDim Preserve OutParam(i)
                    If Dataset.Tables(0).Rows(i)("Status") = "Verified" Then

                        If tostr(GetItemValue("SYNCDATE", Dataset, i)) = "" Then
                            SYNCDATE = "1900-01-01"
                        Else
                            SYNCDATE = todate(GetItemValue("SYNCDATE", Dataset, i))
                        End If

                        If tostr(GetItemValue("SYNCTIME", Dataset, i)) = "" Then
                            SYNCTIME = "00:00:00"
                        Else
                            SYNCTIME = tostr(GetItemValue("SYNCTIME", Dataset, i))
                        End If

                        PName = tostr(GetItemValue("Name", Dataset, i))
                        FName = tostr(GetItemValue("FirstName", Dataset, i))
                        mName = tostr(GetItemValue("MidName", Dataset, i))
                        lName = tostr(GetItemValue("LastName", Dataset, i))

                        RefID = tostr(GetItemValue("RefID", Dataset, i))
                        XProfileID = tostr(GetItemValue("XProfileID", Dataset, i))
                        If CBool(toval(GetAppSetting("ThaiProfile"))) Then
                            Dim tmp_br As String = tostr(GetItemValue("Branch", Dataset, i))
                            Dim tmp_branch As String = Right("00000" & tmp_br, 5)
                            Branch = IIf(tostr(GetItemValue("Branch", Dataset, i)) <> "", tmp_branch, "")
                        Else
                            Branch = IIf(tostr(GetItemValue("Branch", Dataset, i)) <> "", tostr(GetItemValue("Branch", Dataset, i)), br_branch)
                        End If
                        PType = tostr(GetItemValue("PType", Dataset, i))
                        LOB = tostr(GetItemValue("LOB", Dataset, i))
                        IntermediaryF = toval(GetItemValue("IntermediaryF", Dataset, i))
                        CorporateF = toval(GetItemValue("CorporateF", Dataset, i))
                        VATSubsidiesF = toval(GetItemValue("VATSubsidiesF", Dataset, i))
                        TaxID = tostr(GetItemValue("TaxID", Dataset, i))
                        TaxIDOLD = tostr(GetItemValue("TaxIDOLD", Dataset, i))
                        TaxType = tostr(GetItemValue("TaxType", Dataset, i))
                        TaxLoading = IIf(toval(GetItemValue("TAX", Dataset, i)) = 0, 0, IIf(toval(GetItemValue("TaxLoading", Dataset, i)) = 0, 100, toval(GetItemValue("TaxLoading", Dataset, i))))

                        Currency = tostr(GetItemValue("Currency", Dataset, i))
                        If Currency = "" Then Currency = br_Currency

                        BankCode = tostr(GetItemValue("BankCode", Dataset, i))
                        BankBranch = tostr(GetItemValue("BankBranch", Dataset, i))
                        BankAccountNo = tostr(GetItemValue("BankAccountNo", Dataset, i))
                        BankAccountName = tostr(GetItemValue("BankAccountName", Dataset, i))
                        BankName = tostr(GetItemValue("BankName", Dataset, i))
                        RefType = tostr(GetItemValue("RefType", Dataset, i))
                        BCTransferBy = toval(GetItemValue("BCTransferBy", Dataset, i))
                        PRDays = toval(GetItemValue("PRDays", Dataset, i))
                        Remarks = tostr(GetItemValue("Remarks", Dataset, i))
                        TAXF = toval(GetItemValue("TAXF", Dataset, i))
                        VATF = toval(GetItemValue("VATF", Dataset, i))
                        BVATF = toval(GetItemValue("BVATF", Dataset, i))
                        TVATF = toval(GetItemValue("TVATF", Dataset, i))
                        VATCOMMEXEMPTIF = toval(GetItemValue("VATCOMMEXEMPTIF", Dataset, i))
                        RedFlag = toval(GetItemValue("RedFlag", Dataset, i))
                        ConsentOnDataKeepByCompanyF = toval(GetItemValue("ConsentOnDataKeepByCompanyF", Dataset, i))
                        ConsentOnDataPrivacyToThirdPartyF = toval(GetItemValue("ConsentOnDataPrivacyToThirdPartyF", Dataset, i))
                        ConsentOnProductAndServiceOfferingF = toval(GetItemValue("ConsentOnProductAndServiceOfferingF", Dataset, i))
                        Remark = tostr(GetItemValue("Remark", Dataset, i))


                        If PName = "" Then
                            PName = LTrim(FName & " " & mName & " " & lName)
                        Else
                            Profile_Name(PName, FName, mName, lName)
                        End If

                        ID = tostr(GetItemValue("ID", Dataset, i))
                        If ID = "" Then
                            If CBool(toval(GetAppSetting("ThaiProfile"))) Then
                                ID = Trim(db.GetNextDocNo("Pro", Today.Date, Mid(Get_AlphaNum(PType), 1, 1)))
                            Else
                                If Len(Get_Branch_Alias(Branch)) > 2 Then
                                    ID = Trim(db.GetNextDocNo("Pro", Today.Date, Mid(Get_AlphaNum(PType), 1, 1) + Mid(Get_AlphaNum(FName), 1, 1) + IIf(Mid(Get_AlphaNum(mName), 1, 1) = "", Mid(Get_AlphaNum(FName), 2, 1), Mid(Get_AlphaNum(mName), 1, 1))))
                                Else
                                    ID = Trim(db.GetNextDocNo("Pro", Today.Date, Mid(Get_AlphaNum(PType), 1, 1) + Get_Branch_Alias(Branch) + Mid(Get_AlphaNum(FName), 1, 1) + IIf(Mid(Get_AlphaNum(mName), 1, 1) = "", Mid(Get_AlphaNum(FName), 2, 1), Mid(Get_AlphaNum(mName), 1, 1))))
                                End If
                            End If
                        End If
                        If ID = "" Then Throw New Exception("Generate Profile ID Failed !")

                        Dataset.Tables(0).Rows(i)("ID") = ID

                        If Not db.Execute("Select * into #Profile from Profile Where ID='" & ID & "'") Then Throw New Exception(db.ErrDesc)
                        rsDst = db.OpenRemoteSet("Select * from #Profile")
                        If rsDst.Eof Then
                            rsDst.AddNew()
                        Else
                            rsDst.Edit()
                        End If
                        rsDst.fields("PType") = PType
                        rsDst.fields("LOB") = LOB
                        rsDst.fields("IntermediaryF") = IIf(IntermediaryF = 0, False, True)
                        rsDst.fields("CorporateF") = IIf(CorporateF = 0, False, True)
						rsDst.fields("VatSubsidiesF") = IIf(VatSubsidiesF = 0, False, True)
                        rsDst.fields("Name") = PName
                        rsDst.fields("FirstName") = FName
                        rsDst.fields("MidName") = mName
                        rsDst.fields("LastName") = lName
                        rsDst.fields("Salutation") = tostr(GetItemValue("Salutation", Dataset, i))
                        rsDst.fields("Initial") = tostr(GetItemValue("Initial", Dataset, i))
                        rsDst.fields("Address_1") = tostr(GetItemValue("Address_1", Dataset, i))
                        rsDst.fields("Address_2") = tostr(GetItemValue("Address_2", Dataset, i))
                        rsDst.fields("Address_3") = tostr(GetItemValue("Address_3", Dataset, i))
                        rsDst.fields("City") = tostr(GetItemValue("City", Dataset, i))
                        rsDst.fields("ZipCode") = tostr(GetItemValue("ZipCode", Dataset, i))
                        rsDst.fields("Phone_1") = tostr(GetItemValue("Phone_1", Dataset, i))
                        rsDst.fields("Phone_2") = tostr(GetItemValue("Phone_2", Dataset, i))
                        rsDst.fields("Mobile_1") = tostr(GetItemValue("Mobile_1", Dataset, i))
                        rsDst.fields("Mobile_2") = tostr(GetItemValue("Mobile_2", Dataset, i))
                        rsDst.fields("Fax_1") = tostr(GetItemValue("Fax_1", Dataset, i))
                        rsDst.fields("Fax_2") = tostr(GetItemValue("Fax_2", Dataset, i))
                        rsDst.fields("EMail") = tostr(GetItemValue("Email_1", Dataset, i))
                        rsDst.fields("EMail_2") = tostr(GetItemValue("Email_2", Dataset, i))
                        rsDst.fields("Country") = tostr(GetItemValue("Country", Dataset, i))
                        rsDst.fields("Area") = tostr(GetItemValue("Area", Dataset, i))
                        rsDst.fields("Contact") = tostr(GetItemValue("PIC", Dataset, i))
                        rsDst.fields("Title") = tostr(GetItemValue("PICTitle", Dataset, i))
                        rsDst.fields("ContactAddress") = tostr(GetItemValue("PICAddress", Dataset, i))
                        rsDst.fields("ContactPhone") = tostr(GetItemValue("PICPhone", Dataset, i))
                        rsDst.fields("CType") = tostr(GetItemValue("CType", Dataset, i))
                        rsDst.fields("CGroup") = tostr(GetItemValue("CGroup", Dataset, i))
                        rsDst.fields("SCGroup") = tostr(GetItemValue("SCGroup", Dataset, i))
                        rsDst.fields("TaxType") = TaxType
                        rsDst.fields("TaxID") = TaxID
                        rsDst.fields("TaxIDOLD") = TaxIDOLD
                        rsDst.fields("TaxName") = tostr(GetItemValue("TaxName", Dataset, i))
                        rsDst.fields("TaxAddress") = tostr(GetItemValue("TaxAddress", Dataset, i))

                        rsDst.fields("VAT") = toval(GetItemValue("VAT", Dataset, i))
                        rsDst.fields("TAX") = toval(GetItemValue("TAX", Dataset, i))
                        rsDst.fields("PID") = tostr(GetItemValue("ReferalID", Dataset, i))
                        rsDst.fields("RType") = tostr(GetItemValue("ReferalType", Dataset, i))

                        If toval(GetItemValue("AccruedTax", Dataset, i)) <> 0 Then
                            rsDst.fields("AccruedTax") = toval(GetItemValue("AccruedTax", Dataset, i))
                        End If

                        If toval(rsDst.fields("VAT")) <> 0 Then VATF = True
                        If toval(rsDst.fields("TAX")) <> 0 Then TAXF = True

                        rsDst.fields("VatF") = VATF
                        rsDst.fields("TaxF") = TAXF
                        rsDst.fields("BVATF") = BVATF
                        rsDst.fields("TVATF") = TVATF
                        rsDst.fields("VATCOMMEXEMPTIF") = VATCOMMEXEMPTIF
                        rsDst.fields("RedFlag") = RedFlag
                        rsDst.fields("ConsentOnDataKeepByCompanyF") = ConsentOnDataKeepByCompanyF
                        rsDst.fields("ConsentOnDataPrivacyToThirdPartyF") = ConsentOnDataPrivacyToThirdPartyF
                        rsDst.fields("ConsentOnProductAndServiceOfferingF") = ConsentOnProductAndServiceOfferingF
                        rsDst.fields("Remark") = Remark

                        rsDst.fields("TaxPct") = TaxLoading
                        rsDst.fields("LicenseNo") = tostr(GetItemValue("LicenseNo", Dataset, i))
                        rsDst.fields("ID") = ID
                        rsDst.fields("RefID") = IIf(RefID <> "", RefID, ID & "/" & LOB)
                        rsDst.fields("XProfileID") = XProfileID
                        rsDst.fields("Owner") = GetUserOwner(dbUser)
                        rsDst.fields("CDate") = Today.Date
                        rsDst.fields("UserID") = dbUser
                        rsDst.fields("Last_Update") = Today.Date
                        rsDst.fields("Last_Time") = Time()
                        rsDst.fields("SYNCDATE") = SYNCDATE
                        rsDst.fields("SYNCTIME") = SYNCTIME
                        rsDst.fields("Last_Opr") = dbUser
                        rsDst.fields("Branch") = Branch
                        rsDst.fields("BIType") = tostr(GetItemValue("BIType", Dataset, i))
                        rsDst.fields("HSCode") = tostr(GetItemValue("HSCode", Dataset, i))
                        rsDst.fields("RestrictedF") = toval(GetItemValue("RestrictedF", Dataset, i))
                        rsDst.fields("Province") = tostr(GetItemValue("Province", Dataset, i))
                        rsDst.fields("Building") = tostr(GetItemValue("Building", Dataset, i))
                        rsDst.fields("Village") = tostr(GetItemValue("Village", Dataset, i))
                        rsDst.fields("RoomNo") = tostr(GetItemValue("RoomNo", Dataset, i))
                        rsDst.fields("FloorNo") = tostr(GetItemValue("FloorNo", Dataset, i))
                        rsDst.fields("HomeNo") = tostr(GetItemValue("HomeNo", Dataset, i))
                        rsDst.fields("MOO") = tostr(GetItemValue("MOO", Dataset, i))
                        rsDst.fields("SOI") = tostr(GetItemValue("SOI", Dataset, i))
                        rsDst.fields("Road") = tostr(GetItemValue("Road", Dataset, i))
                        rsDst.fields("SubDistrict") = tostr(GetItemValue("SubDistrict", Dataset, i))
                        rsDst.fields("District") = tostr(GetItemValue("District", Dataset, i))
                        rsDst.fields("MOID") = tostr(GetItemValue("MOID", Dataset, i))


                        If tostr(GetItemValue("Address_1", Dataset, i)) = "" Then
                            If tostr(GetItemValue("Building", Dataset, i)) <> "" Then
                                rsDst.fields("Address_1") = tostr(GetItemValue("Building", Dataset, i)) & " ,"
                            End If
                            If tostr(GetItemValue("Village", Dataset, i)) <> "" Then
                                rsDst.fields("Address_1") = tostr(GetItemValue("Village", Dataset, i)) & " ,"
                            End If
                        End If

                        If tostr(GetItemValue("Address_2", Dataset, i)) = "" Then
                            If tostr(GetItemValue("HomeNo", Dataset, i)) <> "" Then
                                rsDst.fields("Address_2") = tostr(GetItemValue("HomeNo", Dataset, i)) + " ,"
                            End If

                            If tostr(GetItemValue("RoomNo", Dataset, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("RoomNo", Dataset, i)) & " ,"
                            End If

                            If tostr(GetItemValue("FloorNo", Dataset, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("FloorNo", Dataset, i)) & " ,"
                            End If

                            If tostr(GetItemValue("MOO", Dataset, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("Moo", Dataset, i))
                            End If

                            If tostr(GetItemValue("Road", Dataset, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("Road", Dataset, i))
                            End If

                            If tostr(GetItemValue("SOI", Dataset, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("Soi", Dataset, i))
                            End If
                        End If


                        If tostr(GetItemValue("Address_3", Dataset, i)) = "" Then
                            If CBool(toval(GetAppSetting("ProfileAddressExtension"))) Then
                                'NO overwrite for address_2 and address_3
                                If tostr(GetItemValue("SubDistrict", Dataset, i)) <> "" Then
                                    rsDst.fields("SubDistrict") = rsDst.fields("SubDistrict") & tostr(GetItemValue("SubDistrict", Dataset, i)) & " ,"
                                End If
                                If tostr(GetItemValue("District", Dataset, i)) <> "" Then
                                    rsDst.fields("District") = rsDst.fields("District") & tostr(GetItemValue("District", Dataset, i))
                                End If
                            Else
                                'overwrite for address_2 and address_3
                                If tostr(GetItemValue("SubDistrict", Dataset, i)) <> "" Then
                                    rsDst.fields("Address_3") = rsDst.fields("Address_3") & tostr(GetItemValue("SubDistrict", Dataset, i)) & " ,"
                                End If
                                If tostr(GetItemValue("District", Dataset, i)) <> "" Then
                                    rsDst.fields("Address_3") = rsDst.fields("Address_3") & tostr(GetItemValue("District", Dataset, i))
                                End If
                            End If
                        End If

                        rsDst.fields("LicenseDate") = todate(GetItemValue("LicenseDate", Dataset, i))
                        rsDst.fields("LicenseExpiry") = todate(GetItemValue("LicenseExpiry", Dataset, i))
                        rsDst.fields("WPC") = toval(GetItemValue("WPC", Dataset, i))
                        rsDst.fields("Grace") = toval(GetItemValue("Grace", Dataset, i))
                        rsDst.fields("Disc") = toval(GetItemValue("Disc", Dataset, i))
                        rsDst.fields("Comm") = toval(GetItemValue("Comm", Dataset, i))
                        rsDst.fields("NETCOMMF") = toval(GetItemValue("NETCOMMF", Dataset, i))
                        rsDst.fields("MaxOutGo") = toval(GetItemValue("MaxOutGo", Dataset, i))
                        rsDst.fields("Rating") = toval(GetItemValue("Rating", Dataset, i))

                        If CBool(toval(GetAppSetting("ApplyVerifyProfile"))) Then
                            If PName <> "" AndAlso tostr(GetItemValue("Address_1", Dataset, i)) <> "" AndAlso tostr(GetItemValue("City", Dataset, i)) <> "" AndAlso tostr(GetItemValue("Province", Dataset, i)) <> "" AndAlso PType <> "" Then
                                If CorporateF = 0 Then
                                    If tostr(GetItemValue("BirthDate", Dataset, i)) <> "" AndAlso LOB <> "" AndAlso tostr(GetItemValue("Country", Dataset, i)) <> "" AndAlso tostr(GetItemValue("ID_Type", Dataset, i)) <> "" AndAlso tostr(GetItemValue("ID_No", Dataset, i)) <> "" Then
                                        rsDst.fields("VerifiedF") = True
                                        rsDst.fields("VerifiedBy") = dbUser
                                        rsDst.fields("VerifiedDate") = Today.Date
                                        rsDst.fields("VerifiedTime") = Time()
                                    Else
                                        rsDst.fields("VerifiedF") = False
                                    End If
                                Else
                                    If LOB <> "" AndAlso tostr(GetItemValue("CGroup", Dataset, i)) <> "" AndAlso tostr(GetItemValue("Country", Dataset, i)) <> "" Then
                                        rsDst.fields("VerifiedF") = True
                                        rsDst.fields("VerifiedBy") = dbUser
                                        rsDst.fields("VerifiedDate") = Today.Date
                                        rsDst.fields("VerifiedTime") = Time()
                                    Else
                                        rsDst.fields("VerifiedF") = False
                                    End If

                                End If
                            Else
                                rsDst.fields("VerifiedF") = False
                            End If
                        Else
                            rsDst.fields("VerifiedF") = False
                        End If

                        If CBool(toval(GetAppSetting("ProfileConsentMandatoryF"))) Then
                            rsDst.fields("ConsentOnDataKeepByCompanyF") = 1
                        Else
                            rsDst.fields("ConsentOnDataKeepByCompanyF") = 0
                        End If

                        Dim EditMode As Integer
                        rsP = db.OpenRemoteSet("Select * from Profile Where ID='" & ID & "'")
                        If rsP.Eof Then
                            rsP.AddNew()
                        Else
                            rsP.Edit()
                        End If
                        EditMode = rsP.editmode

                        Dim j As Integer
                        For j = 0 To rsP.fields_Count - 1
                            If EditMode = AdEditInProgress Then
                                If rsP.fields(j) <> rsDst.fields(j) Then
                                    If Not Save_Log_Object("PROFILE", "ID", rsP.fields("ID"), "", "", "", "", "", "", rsP.fields_Name(j), rsP.fields(j), rsDst.fields(j), "Update") Then Throw New Exception("Error Save Log Object")
                                End If
                            End If
                            rsP.fields(j) = rsDst.fields(j)
                        Next
                        If Not rsP.Update Then Throw New Exception(rsP.ErrDesc)

                        ' Create Profile_Ext
                        If Not db.Execute("Select * into #Profile_Ext from Profile_Ext Where ID='" & ID & "'") Then Throw New Exception(db.ErrDesc)
                        rsDst2 = db.OpenRemoteSet("Select * from #Profile_Ext")
                        If rsDst2.Eof Then
                            rsDst2.AddNew()
                        Else
                            rsDst2.Edit()
                        End If
                        rsDst2.fields("ID") = ID
                        rsDst2.fields("ID_Type") = tostr(GetItemValue("ID_Type", Dataset, i))
                        rsDst2.fields("ID_No") = tostr(GetItemValue("ID_No", Dataset, i))
                        rsDst2.fields("ID_Name") = tostr(GetItemValue("ID_Name", Dataset, i))
                        rsDst2.fields("Gender") = tostr(GetItemValue("Gender", Dataset, i))
                        rsDst2.fields("BirthDate") = todate(GetItemValue("BirthDate", Dataset, i))
                        rsDst2.fields("BirthPlace") = tostr(GetItemValue("BirthPlace", Dataset, i))
                        rsDst2.fields("WNIF") = toval(GetItemValue("WNIF", Dataset, i))
                        rsDst2.fields("CITIZEN") = tostr(GetItemValue("CITIZEN", Dataset, i))
                        rsDst2.fields("BLNO") = tostr(GetItemValue("BLNo", Dataset, i))
                        rsDst2.fields("Income") = toval(GetItemValue("Income", Dataset, i))
                        rsDst2.fields("CAddress") = tostr(GetItemValue("CAddress", Dataset, i))
                        rsDst2.fields("CPhone_1") = tostr(GetItemValue("CPhone_1", Dataset, i))
                        rsDst2.fields("Marital") = tostr(GetItemValue("Marital", Dataset, i))
                        rsDst2.fields("MothersName") = tostr(GetItemValue("MothersName", Dataset, i))

                        rsDst2.fields("DomicileAddress") = tostr(GetItemValue("DomicileAddress", Dataset, i))
                        rsDst2.fields("AddressSameIDF") = toval(GetItemValue("AddressSameIDF", Dataset, i))
                        rsDst2.fields("PIDExpDate") = todate(GetItemValue("PIDExpDate", Dataset, i))
                        rsDst2.fields("NoPIDExpDateF") = toval(GetItemValue("NoPIDExpDateF", Dataset, i))
                        rsDst2.fields("LastEducation") = tostr(GetItemValue("LastEducation", Dataset, i))
                        rsDst2.fields("MonthlyIncome") = tostr(GetItemValue("MonthlyIncome", Dataset, i))
                        rsDst2.fields("MonthlyExpenses") = tostr(GetItemValue("MonthlyExpenses", Dataset, i))


                        'SLIK
                        rsDst2.fields("Num_Dependent") = toval(GetItemValue("Num_Dependent", Dataset, i))
                        rsDst2.fields("Relation") = tostr(GetItemValue("Relation", Dataset, i))
                        rsDst2.fields("Debtor_Class") = tostr(GetItemValue("Debtor_Class", Dataset, i))
                        rsDst2.fields("BirthDate_1") = todate(GetItemValue("BirthDate_1", Dataset, i))
                        rsDst2.fields("Rating_Agency") = tostr(GetItemValue("Rating_Agency", Dataset, i))
                        rsDst2.fields("Rating_Date") = todate(GetItemValue("Rating_Date", Dataset, i))
                        rsDst2.fields("ID_No_Ext_1") = tostr(GetItemValue("ID_No_Ext_1", Dataset, i))
                        rsDst2.fields("ID_No_Ext_2") = tostr(GetItemValue("ID_No_Ext_2", Dataset, i))
                        rsDst2.fields("ID_No_Ext_3") = tostr(GetItemValue("ID_No_Ext_3", Dataset, i))
                        rsDst2.fields("ID_Type_Ext_1") = tostr(GetItemValue("ID_Type_Ext_1", Dataset, i))
                        rsDst2.fields("ID_Type_Ext_2") = tostr(GetItemValue("ID_Type_Ext_2", Dataset, i))
                        rsDst2.fields("ID_Type_Ext_3") = tostr(GetItemValue("ID_Type_Ext_3", Dataset, i))
                        rsDst2.fields("BOD_Gender_1") = tostr(GetItemValue("BOD_Gender_1", Dataset, i))
                        rsDst2.fields("BOD_Gender_2") = tostr(GetItemValue("BOD_Gender_2", Dataset, i))
                        rsDst2.fields("BOD_Gender_3") = tostr(GetItemValue("BOD_Gender_3", Dataset, i))
                        rsDst2.fields("BOD_Address_1") = tostr(GetItemValue("BOD_Address_1", Dataset, i))
                        rsDst2.fields("BOD_Address_2") = tostr(GetItemValue("BOD_Address_2", Dataset, i))
                        rsDst2.fields("BOD_Address_3") = tostr(GetItemValue("BOD_Address_3", Dataset, i))
                        rsDst2.fields("BOD_SubDistrict_1") = tostr(GetItemValue("BOD_SubDistrict_1", Dataset, i))
                        rsDst2.fields("BOD_SubDistrict_2") = tostr(GetItemValue("BOD_SubDistrict_2", Dataset, i))
                        rsDst2.fields("BOD_SubDistrict_3") = tostr(GetItemValue("BOD_SubDistrict_3", Dataset, i))
                        rsDst2.fields("BOD_District_1") = tostr(GetItemValue("BOD_District_1", Dataset, i))
                        rsDst2.fields("BOD_District_2") = tostr(GetItemValue("BOD_District_2", Dataset, i))
                        rsDst2.fields("BOD_District_3") = tostr(GetItemValue("BOD_District_3", Dataset, i))
						rsDst2.fields("BOD_Village_1") = tostr(GetItemValue("BOD_Village_1", Dataset, i))
						rsDst2.fields("BOD_Village_2") = tostr(GetItemValue("BOD_Village_2", Dataset, i))
						rsDst2.fields("BOD_Village_3") = tostr(GetItemValue("BOD_Village_3", Dataset, i))
						   

                        rsDst2.fields("BOD_City_1") = tostr(GetItemValue("BOD_City_1", Dataset, i))
                        rsDst2.fields("BOD_City_2") = tostr(GetItemValue("BOD_City_2", Dataset, i))
                        rsDst2.fields("BOD_City_3") = tostr(GetItemValue("BOD_City_3", Dataset, i))
                        rsDst2.fields("Own_Share") = toval(GetItemValue("Own_Share", Dataset, i))
                        rsDst2.fields("Own_Status") = tostr(GetItemValue("Own_Status", Dataset, i))
						
						rsDst2.fields("GoPublicF") = toval(GetItemValue("GoPublicF", Dataset, i))
						rsDst2.fields("PropertySeparationF") = toval(GetItemValue("PropertySeparationF", Dataset, i))
                        
						rsP = db.OpenRemoteSet("Select * from Profile_Ext Where ID='" & ID & "'")
                        If rsP.Eof Then
                            rsP.AddNew()
                        Else
                            rsP.Edit()
                        End If
                        EditMode = rsP.editmode

                        For j = 0 To rsP.fields_Count - 1
                            If EditMode = AdEditInProgress Then
                                If rsP.fields(j) <> rsDst2.fields(j) Then
                                    If Not Save_Log_Object("PROFILE_EXT", "ID", rsP.fields("ID"), "", "", "", "", "", "", rsP.fields_Name(j), rsP.fields(j), rsDst2.fields(j), "Update") Then Throw New Exception("Error Save Log Object")
                                End If
                            End If
                            rsP.fields(j) = rsDst2.fields(j)
                        Next
                        If Not rsP.Update Then Throw New Exception(rsP.ErrDesc)

                        'Create Profile Account
                        If Not Create_ProAcc(ID, Currency, BankAccountNo, BankAccountName, BankName, BankCode, BankBranch, RefType, BCTransferBy, PRDays, Remarks) Then Throw New Exception("Create Profile Account Failed !")

                        'Create Agency
                        If toval(GetItemValue("UploadAgencyF", Dataset, i)) Then
                            If Not db.Execute("Select * into #AIF from AIF Where ID='" & ID & "'") Then Throw New Exception(db.ErrDesc)
                            rsDst3 = db.OpenRemoteSet("Select * from #AIF")
                            If rsDst3.Eof Then
                                rsDst3.AddNew()
                            Else
                                rsDst3.Edit()
                            End If

                            rsDst3.fields("ID") = ID
                            rsDst3.fields("AgentID") = ID
                            rsDst3.fields("AgentName") = tostr(GetItemValue("AgentName", Dataset, i))
                            rsDst3.fields("ContractNo") = tostr(GetItemValue("ContractNo", Dataset, i))
                            rsDst3.fields("AgentLevel") = tostr(GetItemValue("AgentLevel", Dataset, i))
                            rsDst3.fields("AgentType") = tostr(GetItemValue("AgentType", Dataset, i))
                            rsDst3.fields("SDate") = IIf(tostr(GetItemValue("RegisterStartDate", Dataset, i)) <> "", todate(GetItemValue("RegisterStartDate", Dataset, i)), Today.Date)
                            rsDst3.fields("EDate") = IIf(tostr(GetItemValue("RegisterExpiryDate", Dataset, i)) <> "", todate(GetItemValue("RegisterExpiryDate", Dataset, i)), Today.Date)
                            rsDst3.fields("RecruiterID") = tostr(GetItemValue("RecruiterID", Dataset, i))
                            rsDst3.fields("RecruiterType") = "P"
                            rsDst3.fields("AgentStatus") = IIf(tostr(GetItemValue("AgentStatus", Dataset, i)) <> "", tostr(GetItemValue("AgentStatus", Dataset, i)), "N")
                            rsDst3.fields("ULAgentID") = tostr(GetItemValue("ULAgentID", Dataset, i))
                            rsDst3.fields("Status") = "R"
                            rsDst3.fields("CDate") = Today.Date
                            rsDst3.fields("UserID") = dbUser
                            rsDst3.fields("Last_Opr") = dbUser
                            rsDst3.fields("Last_Update") = Today.Date

                            rsP = db.OpenRemoteSet("Select * from AIF Where ID='" & ID & "'")
                            If rsP.Eof Then
                                rsP.AddNew()
                            Else
                                rsP.Edit()
                            End If
                            EditMode = rsP.editmode

                            For j = 0 To rsP.fields_Count - 1
                                If EditMode = AdEditInProgress Then
                                    If rsP.fields(j) <> rsDst3.fields(j) Then
                                        If Not Save_Log_Object("AIF", "AgentID", rsP.fields("ID"), "", "", "", "", "", "", rsP.fields_Name(j), rsP.fields(j), rsDst3.fields(j), "Update") Then Throw New Exception("Error Save Log Object")
                                    End If
                                End If
                                rsP.fields(j) = rsDst3.fields(j)
                            Next
                            If Not rsP.Update Then Throw New Exception(rsP.ErrDesc)

                            If Not db.Execute("Drop Table #AIF") Then Throw New Exception(db.ErrDesc)
                        End If

                        If Not db.Execute("Drop Table #Profile") Then Throw New Exception(db.ErrDesc)
                        If Not db.Execute("Drop Table #Profile_Ext") Then Throw New Exception(db.ErrDesc)
                    Else
                        OutParam(i) = "Failed"
                    End If

                    db.Committrans()
                Catch ex As Exception
                    OutParam(i) = ex.Message
                    db.RollBack()
                End Try
            Next

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
    Public Function Verify_Upload_Profile_Mandatory(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef DatasetResult As DataSet, ByRef ErrMsg As String) As Boolean
        Dim Status As String
        Dim SStatus As String
        Dim BatchNo As String
        Dim BID As Integer
        Dim TOTALF As Integer = 0
        Dim TOTALC As Integer = 0

        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            Dim i As Integer
            Dim StrMsg As String
            Dim rs As RemoteSet
            Dim LineNo As String
            Dim CheckF As Boolean
            Dim rsMaxLength As RemoteSet
            Dim Rs_Mandatory As RemoteSet
            Dim Rs_Mandatory_1 As RemoteSet
            Dim Rs_Mandatory_2 As RemoteSet
            Dim Rs_Mandatory_3 As RemoteSet
            Dim Rs_Mandatory_4 As RemoteSet
            Dim Rs_Mandatory_5 As RemoteSet
            Dim Rs_Profile_Batch As RemoteSet
            Dim Rs_Profile_BatchNew As RemoteSet

            Dim DataSetNew As DataSet
            Dim DataSetNew_1 As DataSet
            Dim DataSetNew_2 As DataSet
            Dim DataSetNew_3 As DataSet
            Dim DataSetNew_4 As DataSet
            Dim DataSetNew_5 As DataSet
            Dim DS_Mandatory As DataSet
            Dim HeaderColumn As String
            Dim HeaderColumn_Ext As String
            Dim HeaderColumn_Ext_1 As String

            Dim Message As String
            Dim FieldName As String

            Dim CGroup As String
            Dim SCGroup As String
            Dim Gender As String
            Dim ID_Type As String
            Dim ID_No As String

            Dim LengthRefID As Integer
            Dim LengthName As Integer
            Dim LengthTaxID As Integer

            Dim PType As String
            Dim LOB As String
            Dim IntermediaryF As Integer
            Dim CorporateF As Integer
            Dim ID As String
            Dim RefID As String
            Dim PName As String
            Dim FName As String
            Dim MName As String
            Dim LName As String
            Dim Country As String
            Dim City As String
            Dim Province As String
            Dim Area As String
            Dim Branch As String
            Dim BIType As String
            Dim HSCode As String
            Dim BankCode As String
            Dim BankBranch As String
            Dim MOID As String
            Dim Address As String
            Dim Currency As String
            Dim RefType As String
            Dim BCTransferBy As Integer
            Dim PRDays As Integer

            Dim TaxID As String
            Dim TaxType As String
            Dim TaxF As Boolean
            Dim Dataset As DataSet

            ' Get Max Length From SysDefinitions
            rsMaxLength = db.OpenRemoteSet("Select Length from SysDefinitions Where TBName='Profile' and ColName='RefID'")
            LengthRefID = rsMaxLength.fields("Length")

            rsMaxLength = db.OpenRemoteSet("Select Length from SysDefinitions Where TBName='Profile' and ColName='Name'")
            LengthName = rsMaxLength.fields("Length")

            LengthTaxID = toval(GetAppSetting("CareTaxIDMaxLength"))
            ' End Get Max Length

            Rs_Mandatory = db.OpenRemoteSet("Select * from profile_mandatory Where CorporateF= 1 and IndividualF=1 and IntermediaryF=0")
            DataSetNew = Rs_Mandatory.GetDataSet

            Rs_Mandatory_1 = db.OpenRemoteSet("Select * from profile_mandatory Where CorporateF= 1 and IntermediaryF=0")
            DataSetNew_1 = Rs_Mandatory_1.GetDataSet

            Rs_Mandatory_2 = db.OpenRemoteSet("Select * from profile_mandatory Where IndividualF=1 and IntermediaryF=0")
            DataSetNew_2 = Rs_Mandatory_2.GetDataSet

            Rs_Mandatory_3 = db.OpenRemoteSet("Select * from profile_mandatory Where CorporateF= 1 and IntermediaryF=1")
            DataSetNew_3 = Rs_Mandatory_3.GetDataSet

            Rs_Mandatory_4 = db.OpenRemoteSet("Select * from profile_mandatory Where IndividualF=1 and IntermediaryF=1")
            DataSetNew_4 = Rs_Mandatory_4.GetDataSet

            Rs_Mandatory_5 = db.OpenRemoteSet("Select * from profile_mandatory Where CorporateF= 1 and IndividualF=1 and IntermediaryF=1")
            DataSetNew_5 = Rs_Mandatory_5.GetDataSet

            BatchNo = InParam(0)
            Rs_Profile_Batch = db.OpenRemoteSet("Select * from Profile_Batch Where BatchNo='" & BatchNo & "' AND STATUS='R'")
            Dataset = Rs_Profile_Batch.GetDataSet
            TOTALC = (Dataset.Tables(0).Rows.Count)

            CheckF = False
            StrMsg = ""
            ErrMsg = ""
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Try
                    ReDim Preserve OutParam(3)
                    CheckF = True

                    SStatus = 1

                    BID = toval(GetItemValue("BID", Dataset, i))
                    PType = tostr(GetItemValue("PType", Dataset, i))
                    LOB = tostr(GetItemValue("LOB", Dataset, i))
                    IntermediaryF = toval(GetItemValue("IntermediaryF", Dataset, i))
                    CorporateF = toval(GetItemValue("CorporateF", Dataset, i))
                    ID = tostr(GetItemValue("ID", Dataset, i))
                    RefID = tostr(GetItemValue("RefID", Dataset, i))
                    PName = tostr(GetItemValue("Name", Dataset, i))
                    FName = tostr(GetItemValue("FirstName", Dataset, i))
                    MName = tostr(GetItemValue("MidName", Dataset, i))
                    LName = tostr(GetItemValue("LastName", Dataset, i))
                    If PName = "" Then
                        PName = LTrim(FName & " " & MName & " " & LName)
                    End If

                    Country = tostr(GetItemValue("Country", Dataset, i))
                    City = tostr(GetItemValue("City", Dataset, i))
                    Province = tostr(GetItemValue("Province", Dataset, i))
                    Area = tostr(GetItemValue("Area", Dataset, i))
                    Branch = tostr(GetItemValue("Branch", Dataset, i))
                    BIType = tostr(GetItemValue("BIType", Dataset, i))
                    HSCode = tostr(GetItemValue("HSCode", Dataset, i))
                    BankCode = tostr(GetItemValue("BankCode", Dataset, i))
                    BankBranch = tostr(GetItemValue("BankBranch", Dataset, i))
                    Currency = tostr(GetItemValue("Currency", Dataset, i))
                    MOID = tostr(GetItemValue("MOID", Dataset, i))
                    RefType = tostr(GetItemValue("RefType", Dataset, i))
                    BCTransferBy = toval(GetItemValue("BCTransferBy", Dataset, i))
                    PRDays = toval(GetItemValue("PRDays", Dataset, i))

                    'Taxation
                    TaxID = tostr(GetItemValue("TaxID", Dataset, i))
                    TaxType = tostr(GetItemValue("TaxType", Dataset, i))
                    TaxF = toval(GetItemValue("TaxF", Dataset, i))
                    Address = tostr(GetItemValue("Address_1", Dataset, i))

                    CGroup = tostr(GetItemValue("CGroup", Dataset, i))
                    SCGroup = tostr(GetItemValue("SCGroup", Dataset, i))

                    Gender = UCase(tostr(GetItemValue("Gender", Dataset, i)))
                    ID_Type = tostr(GetItemValue("ID_Type", Dataset, i))
                    ID_No = tostr(GetItemValue("ID_No", Dataset, i))

                    StrMsg = ""

                    If CBool(toval(GetAppSetting("ProfileMandatoryTableF"))) Then

                        If IntermediaryF Then
                            DS_Mandatory = DataSetNew_5
                            If DataSetNew_5.Tables(0).Rows.Count = 0 OrElse DataSetNew_3.Tables(0).Rows.Count <> 0 OrElse DataSetNew_4.Tables(0).Rows.Count <> 0 Then
                                If CorporateF Then
                                    DS_Mandatory = DataSetNew_3
                                Else
                                    DS_Mandatory = DataSetNew_4
                                End If
                            End If
                        Else
                            DS_Mandatory = DataSetNew
                            If DataSetNew.Tables(0).Rows.Count = 0 OrElse DataSetNew_1.Tables(0).Rows.Count <> 0 OrElse DataSetNew_2.Tables(0).Rows.Count <> 0 Then
                                If CorporateF Then
                                    DS_Mandatory = DataSetNew_1
                                Else
                                    DS_Mandatory = DataSetNew_2
                                End If
                            End If
                        End If

                        For j = 0 To DS_Mandatory.Tables(0).Rows.Count - 1

                            HeaderColumn = tostr(DS_Mandatory.Tables(0).Rows(j)("HeaderColumn"))
                            HeaderColumn_Ext = tostr(DS_Mandatory.Tables(0).Rows(j)("HeaderColumn_Ext"))
                            HeaderColumn_Ext_1 = tostr(DS_Mandatory.Tables(0).Rows(j)("HeaderColumn_Ext_1"))
                            FieldName = tostr(DS_Mandatory.Tables(0).Rows(j)("FieldName"))
                            Message = tostr(DS_Mandatory.Tables(0).Rows(j)("Message"))

                            If HeaderColumn <> "" AndAlso tostr(GetItemValue(HeaderColumn, Dataset, i)) = "" Then

                                If HeaderColumn_Ext = "" AndAlso HeaderColumn_Ext_1 = "" Then
                                    StrMsg = StrMsg & Message & vbCrLf
                                ElseIf HeaderColumn_Ext <> "" AndAlso tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)) = "" Then
                                    StrMsg = StrMsg & Message & vbCrLf
                                ElseIf HeaderColumn_Ext_1 <> "" AndAlso tostr(GetItemValue(HeaderColumn_Ext_1, Dataset, i)) = "" Then
                                    StrMsg = StrMsg & Message & vbCrLf
                                End If
                            End If

                            If toval(DS_Mandatory.Tables(0).Rows(j)("MaxLen")) <> 0 Then
                                Dim Email_1Maxlen As Integer = 0
                                Dim PICPhoneMaxlen As Integer = 0
                                Dim Phone_1Maxlen As Integer = 0
                                Dim Mobile_1Maxlen As Integer = 0
                                Dim TAXIDMaxlen As Integer = 0
                                Dim Phone_1Maxlen_Ext As Integer = 0
                                Dim Mobile_1Maxlen_Ext As Integer = 0
                                Dim TAXIDMaxlen_Ext As Integer = 0
                                Dim TAXIDMaxlen_Ext_1 As Integer = 0
                                Dim TaxNameMaxlen As Integer = 0

                                If FieldName = "Email_1" Then
                                    Email_1Maxlen = toval(DS_Mandatory.Tables(0).Rows(j)("MaxLen"))

                                    If tostr(GetItemValue(HeaderColumn, Dataset, i)) <> "" AndAlso Email_1Maxlen <> 0 AndAlso Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) < Email_1Maxlen Then
                                        StrMsg = StrMsg & "Incorrect Email Address (min " & tostr(Email_1Maxlen) & " chars) !" & vbCrLf
                                    End If

                                    If tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)) <> "" AndAlso Email_1Maxlen <> 0 AndAlso Len(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) < Email_1Maxlen Then
                                        StrMsg = StrMsg & "Incorrect Email Address (min " & tostr(Email_1Maxlen) & " chars) !" & vbCrLf
                                    End If

                                    If tostr(GetItemValue(HeaderColumn_Ext_1, Dataset, i)) <> "" AndAlso Email_1Maxlen <> 0 AndAlso Len(tostr(GetItemValue(HeaderColumn_Ext_1, Dataset, i))) < Email_1Maxlen Then
                                        StrMsg = StrMsg & "Incorrect Email Address (min " & tostr(Email_1Maxlen) & " chars) !" & vbCrLf
                                    End If
                                End If

                                If FieldName = "PICPhone" Then
                                    PICPhoneMaxlen = toval(DS_Mandatory.Tables(0).Rows(j)("MaxLen"))

                                    If tostr(GetItemValue(HeaderColumn, Dataset, i)) <> "" AndAlso PICPhoneMaxlen <> 0 AndAlso Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) < PICPhoneMaxlen Then
                                        StrMsg = StrMsg & "Incorrect PIC Phone Number, Must Contain Number Only (0..9) ! " & vbCrLf & " For Example : 0211234567 and not Less Then " & tostr(PICPhoneMaxlen) & " Chars" & vbCrLf
                                    End If
                                End If

                                If FieldName = "Phone_1" Then
                                    Phone_1Maxlen = toval(DS_Mandatory.Tables(0).Rows(j)("MaxLen"))
                                    Phone_1Maxlen_Ext = toval(DS_Mandatory.Tables(0).Rows(j)("Maxlen_Ext"))

                                    If Phone_1Maxlen <> 0 AndAlso Phone_1Maxlen = Phone_1Maxlen_Ext Then
                                        If tostr(GetItemValue(HeaderColumn, Dataset, i)) <> "" AndAlso (Not isNumberString(tostr(GetItemValue(HeaderColumn, Dataset, i)), "-") OrElse ((Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) < Phone_1Maxlen OrElse Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) > Phone_1Maxlen))) Then
                                            StrMsg = StrMsg & "Incorrect Telephone Number, Must Contain Number Only (0..9) ! " & vbCrLf & " For Example  0211234567 And Not Less Then Or More " & tostr(Phone_1Maxlen) & " Chars" & vbCrLf

                                        End If
                                        If tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)) <> "" AndAlso (Not isNumberString(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)), "-") OrElse ((Len(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) < Phone_1Maxlen OrElse Len(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) > Phone_1Maxlen))) Then
                                            StrMsg = StrMsg & "Incorrect Telephone Number, Must Contain Number Only (0..9) ! " & vbCrLf & " for Example  0211234567 And Not Less Then Or More " & tostr(Phone_1Maxlen) & " Chars" & vbCrLf
                                        End If
                                    ElseIf Phone_1Maxlen <> 0 AndAlso Phone_1Maxlen < Phone_1Maxlen_Ext Then
                                        If tostr(GetItemValue(HeaderColumn, Dataset, i)) <> "" AndAlso (Not isNumberString(tostr(GetItemValue(HeaderColumn, Dataset, i)), "-") OrElse (Not (Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) <= Phone_1Maxlen_Ext AndAlso Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) >= Phone_1Maxlen))) Then
                                            StrMsg = StrMsg & "Incorrect Telephone Number, Must Contain Number Only (0..9) ! " & vbCrLf & " for Example  0211234567 And Must Between " & tostr(Phone_1Maxlen) & " Chars And  " & tostr(Phone_1Maxlen_Ext) & "Chars" & vbCrLf
                                        End If
                                        If tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)) <> "" AndAlso (Not isNumberString(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)), "-") OrElse (Not (Len(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) <= Phone_1Maxlen_Ext AndAlso Len(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) >= Phone_1Maxlen))) Then
                                            StrMsg = StrMsg & "Incorrect Telephone Number, Must Contain Number Only (0..9) ! " & vbCrLf & " for Example  0211234567 And Must Between " & tostr(Phone_1Maxlen) & " Chars And  " & tostr(Phone_1Maxlen_Ext) & "Chars" & vbCrLf
                                        End If
                                    End If
                                End If

                                If FieldName = "Mobile_1" Then
                                    Mobile_1Maxlen = toval(DS_Mandatory.Tables(0).Rows(j)("MaxLen"))
                                    Mobile_1Maxlen_Ext = toval(DS_Mandatory.Tables(0).Rows(j)("Maxlen_Ext"))

                                    If Mobile_1Maxlen <> 0 AndAlso Mobile_1Maxlen = Mobile_1Maxlen_Ext Then
                                        If tostr(GetItemValue(HeaderColumn, Dataset, i)) <> "" AndAlso (Not isNumberString(tostr(GetItemValue(HeaderColumn, Dataset, i)), "-") OrElse ((Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) < Mobile_1Maxlen OrElse Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) > Mobile_1Maxlen))) Then
                                            StrMsg = StrMsg & "Incorrect Mobile Phone Number, Must Contain Number Only (0..9) ! " & vbCrLf & " For Example  0211234567 And Not Less Then Or More " & tostr(Mobile_1Maxlen) & " Chars" & vbCrLf

                                        End If
                                        If tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)) <> "" AndAlso (Not isNumberString(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)), "-") OrElse ((Len(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) < Mobile_1Maxlen OrElse Len(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) > Mobile_1Maxlen))) Then
                                            StrMsg = StrMsg & "Incorrect Mobile Phone Number, Must Contain Number Only (0..9) ! " & vbCrLf & " for Example  0211234567 And Not Less Then Or More " & tostr(Mobile_1Maxlen) & " Chars" & vbCrLf
                                        End If
                                    ElseIf Mobile_1Maxlen <> 0 AndAlso Mobile_1Maxlen < Mobile_1Maxlen_Ext Then
                                        If tostr(GetItemValue(HeaderColumn, Dataset, i)) <> "" AndAlso (Not isNumberString(tostr(GetItemValue(HeaderColumn, Dataset, i)), "-") OrElse (Not (Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) <= Mobile_1Maxlen_Ext AndAlso Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) >= Mobile_1Maxlen))) Then
                                            StrMsg = StrMsg & "Incorrect Mobile Phone Number, Must Contain Number Only (0..9) ! " & vbCrLf & " for Example  0211234567 And Must Between " & tostr(Mobile_1Maxlen) & " Chars And  " & tostr(Mobile_1Maxlen_Ext) & "Chars" & vbCrLf
                                        End If
                                        If tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)) <> "" AndAlso (Not isNumberString(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i)), "-") OrElse (Not (Len(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) <= Mobile_1Maxlen_Ext AndAlso Len(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) >= Mobile_1Maxlen))) Then
                                            StrMsg = StrMsg & "Incorrect Mobile Phone Number, Must Contain Number Only (0..9) ! " & vbCrLf & " for Example  0211234567 And Must Between " & tostr(Mobile_1Maxlen) & " Chars And  " & tostr(Mobile_1Maxlen_Ext) & "Chars" & vbCrLf
                                        End If
                                    End If
                                End If

                                If FieldName = "TAXID" Then
                                    TAXIDMaxlen = toval(DS_Mandatory.Tables(0).Rows(j)("MaxLen"))
                                    TAXIDMaxlen_Ext = toval(DS_Mandatory.Tables(0).Rows(j)("Maxlen_Ext"))
                                    TAXIDMaxlen_Ext_1 = toval(DS_Mandatory.Tables(0).Rows(j)("Maxlen_Ext_1"))

                                    If TAXIDMaxlen <> 0 AndAlso TAXIDMaxlen = TAXIDMaxlen_Ext Then
                                        If tostr(GetItemValue(HeaderColumn, Dataset, i)) <> "" AndAlso Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) <> TAXIDMaxlen Then
                                            StrMsg = StrMsg & "Incorrect Tax ID (" & tostr(TAXIDMaxlen) & " chars) ! - " & Replace(Replace(Replace(Replace(tostr(GetItemValue(HeaderColumn, Dataset, i)), ".", ""), "-", ""), "_", ""), ",", "") & " For TaxType " & tostr(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) & vbCrLf
                                        End If
                                    ElseIf TAXIDMaxlen <> 0 AndAlso TAXIDMaxlen < TAXIDMaxlen_Ext Then
                                        If tostr(GetItemValue(HeaderColumn, Dataset, i)) <> "" AndAlso Not (Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) <= TAXIDMaxlen_Ext AndAlso Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) >= toval(TAXIDMaxlen)) Then
                                            StrMsg = StrMsg & "Incorrect Tax ID, Tax ID Must Between " & tostr(TAXIDMaxlen) & " chars " & tostr(TAXIDMaxlen_Ext) & " chars ! - " & Replace(Replace(Replace(Replace(tostr(GetItemValue(HeaderColumn, Dataset, i)), ".", ""), "-", ""), "_", ""), ",", "") & " For TaxType " & tostr(tostr(GetItemValue(HeaderColumn_Ext, Dataset, i))) & vbCrLf
                                        End If
                                    End If

                                End If

                                If FieldName = "TaxName" Then
                                    TaxNameMaxlen = toval(DS_Mandatory.Tables(0).Rows(j)("MaxLen"))

                                    If tostr(GetItemValue(HeaderColumn, Dataset, i)) <> "" AndAlso TaxNameMaxlen <> 0 AndAlso Len(tostr(GetItemValue(HeaderColumn, Dataset, i))) < TaxNameMaxlen Then
                                        StrMsg = StrMsg & "Incorrect Tax Name (min " & tostr(TaxNameMaxlen) & " chars) !" & vbCrLf
                                    End If
                                End If
                            End If

                        Next j

                        'Privilege
                        Select Case PType
                            Case "D"
                                If Not db.GetPrivileges("PROFILE_DIRECT") Then
                                    StrMsg = StrMsg & "Not Allowed For Direct Type Profile !" & vbCrLf
                                End If
                            Case "C"
                                If Not db.GetPrivileges("PROFILE_CAPTIVE") Then
                                    StrMsg = StrMsg & "Not Allowed For Captive Type Profile !" & vbCrLf
                                End If
                            Case "I"
                                If Not db.GetPrivileges("PROFILE_INWARD") Then
                                    StrMsg = StrMsg & "Not Allowed For Inward Type Profile !" & vbCrLf
                                End If
                            Case "M"
                                If Not db.GetPrivileges("PROFILE_INTER") Then
                                    StrMsg = StrMsg & "Not Allowed For Intermediaries Type Profile !" & vbCrLf
                                End If
                            Case "O"
                                If Not db.GetPrivileges("PROFILE_OUTWARD") Then
                                    StrMsg = StrMsg & "Not Allowed For Outward Type Profile !" & vbCrLf
                                End If
                            Case "T"
                                If Not db.GetPrivileges("PROFILE_OTHERS") Then
                                    StrMsg = StrMsg & "Not Allowed For Others Type Profile !" & vbCrLf
                                End If
                            Case "F"
                                If Not db.GetPrivileges("PROFILE_FINANCE") Then
                                    StrMsg = StrMsg & "Not Allowed For Finance Type Profile !" & vbCrLf
                                End If
                        End Select

                        If Country <> "" Then
                            rs = db.OpenRemoteSet("Select Country, Description from Country Where country='" & Country & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Invalid Country Reference!" & Country & vbCrLf
                            End If
                        End If

                        If PType <> "" Then
                            ' Profile Type
                            rs = db.OpenRemoteSet("Select Method, Name from Method where MType = 'PT' and Method='" & PType & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Profile Type ('C','D','I','M','O','T','F') !" & vbCrLf
                            End If
                        End If

                        ' Branch
                        If CBool(toval(GetAppSetting("ThaiProfile"))) Then
                            If Branch <> "" And Not isNumberString(Branch) Then
                                StrMsg = StrMsg & "Incorrect Branch ID, Must Contain Number Only (0..9)!" & vbCrLf
                            End If
                            ' Length Branch
                            If Len(Branch) > 5 Then
                                StrMsg = StrMsg & "Incorrect Branch Length (Max. 5 character(s)) !" & vbCrLf
                            End If
                        Else
                            If Branch <> "" Then
                                rs = db.OpenRemoteSet("Select Branch, Name from Branch where Op_Flag = 1 and Branch='" & Branch & "'")
                                If rs.Eof Then
                                    StrMsg = StrMsg & "Incorrect Branch ID !" & vbCrLf
                                End If
                            End If
                        End If

                        ' LOB
                        If LOB <> "" Then
                            rs = db.OpenRemoteSet("Select LOB from LOB Where LOB='" & LOB & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Line Of Business !" & vbCrLf
                            End If
                        End If

                        ' Currency
                        If Currency <> "" Then
                            rs = db.OpenRemoteSet("Select Currency from Currency Where Currency='" & Currency & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Currency !" & vbCrLf
                            End If
                        End If

                        If RefType <> "" And RefType.ToUpper <> "BANK" And RefType.ToUpper <> "VIRTUAL" Then
                            StrMsg = StrMsg & "Incorrect RefType !" & vbCrLf
                        End If

                        ' Bank Charges
                        If BCTransferBy <> 0 And BCTransferBy <> 1 And BCTransferBy <> 2 Then
                            StrMsg = StrMsg & "Incorrect Bank Charges By !" & vbCrLf
                        End If

                        If PRDays < 0 Then
                            StrMsg = StrMsg & "Incorrect Payment Reminder !" & vbCrLf
                        End If


                        ' Duplicate Ref ID or Profile
                        If ID = "" Then
                            If RefID <> "" Then
                                rs = db.OpenRemoteSet("Select * from Profile Where RefID='" & RefID & "'")
                            Else
                                rs = db.OpenRemoteSet("Select * from Profile Where Name='" & PName & "' and TaxID='" & TaxID & "'")
                            End If
                            If Not rs.Eof Then
                                StrMsg = StrMsg & "Duplicate Profile Entry (see ID : " & rs.fields("ID") & " or Ref. ID : " & rs.fields("RefID") & ") !" & vbCrLf
                            End If
                        End If

                        ' Area
                        If Area <> "" Then
                            rs = db.OpenRemoteSet("Select Area from Area Where Area='" & Area & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Area !" & vbCrLf
                            End If
                        End If

                        'BI Type
                        If BIType <> "" Then
                            rs = db.OpenRemoteSet("Select BIType from BIType Where BIType='" & BIType & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect BI Type !" & vbCrLf
                            End If
                        End If

                        'HS Code
                        If HSCode <> "" Then
                            rs = db.OpenRemoteSet("Select HSCode from HSCode Where HSCode='" & HSCode & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect HS Code !" & vbCrLf
                            End If
                        End If

                        'Bank Code
                        If BankCode <> "" Then
                            rs = db.OpenRemoteSet("Select Bank from Bank Where Bank='" & BankCode & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Bank Code !" & vbCrLf
                            End If
                        End If

                        'Bank Branch
                        If BankBranch <> "" Then
                            rs = db.OpenRemoteSet("Select BankBranch from BankBranch Where Bank='" & BankCode & "' and BankBranch = '" & BankBranch & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Bank Branch !" & vbCrLf
                            End If
                        End If

                        If MOID <> "" Then
                            rs = db.OpenRemoteSet("Select ID from sysuser Where Role='MARKETING OFFICER' and ID='" & MOID & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Marketing Officer ID !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("AgentType", Dataset, i)) <> "" Then
                            rs = db.OpenRemoteSet("select TOP 1 * From AgentType Where AgentType = '" & tostr(GetItemValue("AgentType", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Agent Type !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("AgentLevel", Dataset, i)) <> "" Then
                            rs = db.OpenRemoteSet("select TOP 1 * From AgentLevel Where AgentLevel = '" & tostr(GetItemValue("AgentLevel", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Agent Level !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("RecruiterID", Dataset, i)) <> "" Then
                            rs = db.OpenRemoteSet("select TOP 1 * From Profile Where ID = '" & tostr(GetItemValue("RecruiterID", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Recruiter !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("ULAgentID", Dataset, i)) <> "" Then
                            rs = db.OpenRemoteSet("select TOP 1 * From AIF Where ID = '" & tostr(GetItemValue("ULAgentID", Dataset, i)) & "'")
                            If rs.Eof Then
                                StrMsg = StrMsg & "Incorrect Up Line Agent !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("AgentStatus", Dataset, i)) <> "" Then
                            If UCase(tostr(GetItemValue("AgentStatus", Dataset, i))) <> "A" And UCase(tostr(GetItemValue("AgentStatus", Dataset, i))) <> "N" Then
                                StrMsg = StrMsg & "Incorrect Agent Status !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("RegisterStartDate", Dataset, i)) <> "" Then
                            If Not IsDate(GetItemValue("RegisterStartDate", Dataset, i)) Then
                                StrMsg = StrMsg & "Invalid Register Start Date Agent !" & vbCrLf
                            End If
                        End If

                        If tostr(GetItemValue("RegisterExpiryDate", Dataset, i)) <> "" Then
                            If Not IsDate(GetItemValue("RegisterExpiryDate", Dataset, i)) Then
                                StrMsg = StrMsg & "Invalid Register Expiry Date Agent !" & vbCrLf
                            End If
                        End If

                        Status = "P"
                        ' No need to proceed data verify if any mandatory fields not fulfilled
                        If StrMsg.Replace(vbCrLf, "") <> "" Then
                            StrMsg = StrMsg & "For BID " & tostr(GetItemValue("BID", Dataset, i)) & vbCrLf
                            If StrMsg = "" Then StrMsg = db.ErrDesc
                            Status = "N"
                            SStatus = 0
                            TOTALF = TOTALF + 1
                            db.Execute("Update Profile_Batch set status = 'N',DumpF = 1,  Errmsg ='" & StrMsg & "',  Last_Opr ='" & dbUser & "',  Last_Update = '" & Today.Date & "',  Last_Time = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)
                        Else
                            StrMsg = ""
                        End If

                        If SStatus = 1 Then
                            ' Proceed data verify for fields validation
                            ' Length RefID
                            If RefID <> "" And Len(RefID) > LengthRefID Then
                                StrMsg = StrMsg & "Incorrect Ref. ID Length (Max. " & LengthRefID & " character(s)) !" & vbCrLf
                            End If

                            ' Length Name
                            If Len(PName) > LengthName Then
                                StrMsg = StrMsg & "Incorrect Name Length (Max. " & LengthName & " character(s)) !" & vbCrLf
                            End If


                            If CGroup <> "" Then
                                ' CGroup
                                rs = db.OpenRemoteSet("Select CGroup from CGroup Where CGroup='" & CGroup & "'")
                                If rs.Eof Then
                                    StrMsg = StrMsg & "Incorrect Company Group !" & vbCrLf
                                End If
                            End If

                            ' SCGroup
                            If SCGroup <> "" Then
                                rs = db.OpenRemoteSet("Select SCGroup from SCGroup Where SCGroup='" & SCGroup & "' and CGroup='" & CGroup & "'")
                                If rs.Eof Then
                                    StrMsg = StrMsg & "Incorrect Sub Company Group !" & vbCrLf
                                End If
                            End If

                            ' Gender
                            If Gender <> "" Then
                                rs = db.OpenRemoteSet("Select Gender from Gender Where Gender='" & Gender & "'")
                                If rs.Eof Then
                                    StrMsg = StrMsg & "Incorrect Gender !" & vbCrLf
                                End If
                            End If

                            ' Birth Date
                            If tostr(GetItemValue("BirthDate", Dataset, i)) <> "" AndAlso Not IsDate(GetItemValue("BirthDate", Dataset, i)) Then
                                StrMsg = StrMsg & "Invalid Birth Date !" & vbCrLf
                            End If

                            If ID_Type <> "" AndAlso ID_No <> "" Then
                                rs = db.OpenRemoteSet("Select Top 1 ID from Profile_Ext Where ID_Type='" & ID_Type & "' and ID_No='" & ID_No & "'")
                                If Not rs.Eof Then
                                    StrMsg = StrMsg & "Duplicate ID Number (see ID : " & rs.fields("ID") & " !" & vbCrLf
                                End If
                            End If

                            ' Taxation
                            If TaxF And TaxType = "" Then
                                StrMsg = StrMsg & "Incorrect Tax Type !" & vbCrLf
                            End If

                            ' Tax Type
                            If TaxType <> "" Then
                                rs = db.OpenRemoteSet("Select Top 1 TaxType from TaxType Where TaxType='" & TaxType & "'")
                                If rs.Eof Then
                                    StrMsg = StrMsg & "Incorrect Tax Type !" & vbCrLf
                                End If
                            End If

                            ' VAT
                            If tostr(GetItemValue("VAT", Dataset, i)) <> "" And Not IsNumeric(tostr(GetItemValue("VAT", Dataset, i))) Then
                                StrMsg = StrMsg & "Incorrect VAT !" & vbCrLf
                            End If
                            ' TAX
                            If tostr(GetItemValue("TAX", Dataset, i)) <> "" And Not IsNumeric(tostr(GetItemValue("TAX", Dataset, i))) Then
                                StrMsg = StrMsg & "Incorrect Tax !" & vbCrLf
                            End If
                            ' TAX Loading
                            If tostr(GetItemValue("TaxLoading", Dataset, i)) <> "" And Not IsNumeric(tostr(GetItemValue("TaxLoading", Dataset, i))) Then
                                StrMsg = StrMsg & "Incorrect Tax Loading !" & vbCrLf
                            End If
                            If StrMsg.Replace(vbCrLf, "") <> "" Then
                                StrMsg = StrMsg & "For BID " & tostr(GetItemValue("BID", Dataset, i)) & vbCrLf
                                If StrMsg = "" Then StrMsg = db.ErrDesc
                                Status = "N"
                                TOTALF = TOTALF + 1
                                db.Execute("Update Profile_Batch set status = 'N',  Errmsg ='" & StrMsg & "',  Last_Opr ='" & dbUser & "',  Last_Update = '" & Today.Date & "',  Last_Time = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)
                            Else
                                Status = "P"
                                If db.GetPrivileges("AllowChangeDumpF") Then
                                    db.Execute("Update Profile_Batch set status = 'P',  Errmsg = '',DumpF = 0,  Last_Opr ='" & dbUser & "',  Last_Update = '" & Today.Date & "',  Last_Time = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)
                                Else
                                    If db.GetPrivileges("Forceprofileverification") Then
                                        db.Execute("Update Profile_Batch set status = 'P',  Errmsg = '',DumpF = 1,  Last_Opr ='" & dbUser & "',  Last_Update = '" & Today.Date & "',  Last_Time = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)
                                    Else
                                        db.Execute("Update Profile_Batch set status = 'P',  Errmsg = '',  Last_Opr ='" & dbUser & "',  Last_Update = '" & Today.Date & "',  Last_Time = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        StrMsg = ""
                    End If

                    ErrMsg = ErrMsg & StrMsg & vbCrLf

                Catch ex As Exception
                    StrMsg = StrMsg & "For BID " & tostr(GetItemValue("BID", Dataset, i)) & vbCrLf
                    If StrMsg = "" Then StrMsg = db.ErrDesc
                    db.RollBack()

                    Status = "N"
                    db.Execute("Update Profile_Batch set status = '" & Status & "',DumpF = 1,  Errmsg = '" & StrMsg & "',  Last_Opr ='" & dbUser & "',  Last_Update = '" & Today.Date & "',  Last_Time = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)
                End Try
            Next


            Rs_Profile_BatchNew = db.OpenRemoteSet("Select * from Profile_Batch Where BatchNo='" & BatchNo & "'")
            DatasetResult = Rs_Profile_BatchNew.GetDataSet


            OutParam(0) = BatchNo
            OutParam(1) = TOTALF
            OutParam(2) = TOTALC
            OutParam(3) = ErrMsg

            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = ErrMsg & StrMsg & vbCrLf
            db.RollBack()
            If ErrMsg = "" Then ErrMsg = ex.Message
            Status = "N"
            db.Execute("Update Profile_Batch set status = '" & Status & "',DumpF = 1,  Errmsg = '" & ErrMsg & "',  Last_Opr ='" & dbUser & "',  Last_Update = '" & Today.Date & "',  Last_Time = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)

            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)>
    Public Function Process_Upload_Profile_Mandatory(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef DatasetResult As DataSet, ByRef ErrMsg As String) As Boolean
        Dim BatchNo As String
        Dim BID As Integer
        Dim Status As String
        Dim TOTALF As Integer = 0
        Dim TOTALC As Integer = 0
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If
            Dim i As Integer
            Dim rs As RemoteSet
            Dim rsDst As RemoteSet
            Dim rsDst2 As RemoteSet
            Dim rsDst3 As RemoteSet
            Dim LineNo As String

            Dim PName As String
            Dim FName As String
            Dim mName As String
            Dim lName As String
            Dim ID As String
            Dim RefID As String
            Dim XProfileID As String
            Dim Branch As String
            Dim PType As String
            Dim LOB As String
            Dim IntermediaryF As Integer
            Dim CorporateF As Integer
            Dim VatSubsidiesF As Integer
            Dim TaxID As String
            Dim TaxIDOLD As String
            Dim TaxType As String
            Dim TaxLoading As Double
            Dim BankCode As String
            Dim BankBranch As String
            Dim BankAccountNo As String
            Dim BankAccountName As String
            Dim BankName As String
            Dim RefType As String
            Dim BCTransferBy As Integer
            Dim PRDays As Integer
            Dim Remarks As String
            Dim br_branch As String = ""
            Dim br_Currency As String = ""
            Dim Currency As String = ""
            Dim SYNCDATE As String = ""
            Dim SYNCTIME As String = ""
            Dim message As String = ""
            Dim TAXF As Boolean
            Dim VATF As Boolean
            Dim BVATF As Boolean
            Dim TVATF As Boolean
            Dim VATCOMMEXEMPTIF As Boolean
            Dim RedFlag As Boolean
            Dim ConsentOnDataKeepByCompanyF As Boolean
            Dim ConsentOnDataPrivacyToThirdPartyF As Boolean
            Dim ConsentOnProductAndServiceOfferingF As Boolean
            Dim Remark As String
            Dim PIC As String

            Dim rsP As RemoteSet
            Dim Rs_Profile_Batch As RemoteSet
            Dim Rs_Profile_BatchNEW As RemoteSet
            Dim DATASET As DataSet

            LineNo = ""

            rs = db.OpenRemoteSet("Select Currency, Branch from SysTab")
            br_Currency = tostr(rs.fields("Currency"))
            br_branch = tostr(rs.fields("Branch"))

            BatchNo = InParam(0)

            Rs_Profile_Batch = db.OpenRemoteSet("Select * from Profile_Batch Where BatchNo='" & BatchNo & "' AND STATUS='P'")
            DATASET = Rs_Profile_Batch.GetDataSet
            TOTALC = (DataSet.Tables(0).Rows.Count)

            If DataSet.Tables(0).Columns("ID") Is Nothing Then
                ' Add Column Profile ID
                Dim dc As System.Data.DataColumn
                dc = New DataColumn
                dc.ColumnName = "ID"
                dc.DataType = Type.GetType("System.String")
                dc.DefaultValue = ""
                DataSet.Tables(0).Columns.Add(dc)
            End If

            ErrMsg = ""
            StrMsg = ""
            For i = 0 To DataSet.Tables(0).Rows.Count - 1
                Try
                    db.Begintrans()
                    ReDim Preserve OutParam(3)
                    If DataSet.Tables(0).Rows(i)("Status") = "P" Then

                        If tostr(GetItemValue("SYNCDATE", DataSet, i)) = "" Then
                            SYNCDATE = "1900-01-01"
                        Else
                            SYNCDATE = todate(GetItemValue("SYNCDATE", DataSet, i))
                        End If

                        If tostr(GetItemValue("SYNCTIME", DataSet, i)) = "" Then
                            SYNCTIME = "00:00:00"
                        Else
                            SYNCTIME = tostr(GetItemValue("SYNCTIME", DataSet, i))
                        End If

                        BID = toval(GetItemValue("BID", DataSet, i))
                        PName = tostr(GetItemValue("Name", DataSet, i))
                        FName = tostr(GetItemValue("FirstName", DataSet, i))
                        mName = tostr(GetItemValue("MidName", DataSet, i))
                        lName = tostr(GetItemValue("LastName", DataSet, i))

                        RefID = tostr(GetItemValue("RefID", DataSet, i))
                        XProfileID = tostr(GetItemValue("XProfileID", DataSet, i))
                        If CBool(toval(GetAppSetting("ThaiProfile"))) Then
                            Dim tmp_br As String = tostr(GetItemValue("Branch", DataSet, i))
                            Dim tmp_branch As String = Right("00000" & tmp_br, 5)
                            Branch = IIf(tostr(GetItemValue("Branch", DataSet, i)) <> "", tmp_branch, "")
                        Else
                            Branch = IIf(tostr(GetItemValue("Branch", DataSet, i)) <> "", tostr(GetItemValue("Branch", DataSet, i)), br_branch)
                        End If
                        PType = tostr(GetItemValue("PType", DataSet, i))
                        LOB = tostr(GetItemValue("LOB", DataSet, i))
                        IntermediaryF = toval(GetItemValue("IntermediaryF", DataSet, i))
                        CorporateF = toval(GetItemValue("CorporateF", DataSet, i))
                        VatSubsidiesF = toval(GetItemValue("VATSubsidiesF", DataSet, i))
                        TaxID = tostr(GetItemValue("TaxID", DataSet, i))
                        TaxIDOLD = tostr(GetItemValue("TaxIDOLD", DataSet, i))
                        TaxType = tostr(GetItemValue("TaxType", DataSet, i))
                        TaxLoading = IIf(toval(GetItemValue("TAX", DataSet, i)) = 0, 0, IIf(toval(GetItemValue("TaxLoading", DataSet, i)) = 0, 100, toval(GetItemValue("TaxLoading", DataSet, i))))

                        Currency = tostr(GetItemValue("Currency", DataSet, i))
                        If Currency = "" Then Currency = br_Currency

                        BankCode = tostr(GetItemValue("BankCode", DataSet, i))
                        BankBranch = tostr(GetItemValue("BankBranch", DataSet, i))
                        BankAccountNo = tostr(GetItemValue("BankAccountNo", DataSet, i))
                        BankAccountName = tostr(GetItemValue("BankAccountName", DataSet, i))
                        BankName = tostr(GetItemValue("BankName", DataSet, i))
                        RefType = tostr(GetItemValue("RefType", DataSet, i))
                        BCTransferBy = toval(GetItemValue("BCTransferBy", DataSet, i))
                        PRDays = toval(GetItemValue("PRDays", DataSet, i))
                        Remarks = tostr(GetItemValue("Remarks", DataSet, i))
                        TAXF = toval(GetItemValue("TAXF", DataSet, i))
                        VATF = toval(GetItemValue("VATF", DataSet, i))
                        BVATF = toval(GetItemValue("BVATF", DataSet, i))
                        TVATF = toval(GetItemValue("TVATF", DataSet, i))
                        VATCOMMEXEMPTIF = toval(GetItemValue("VATCOMMEXEMPTIF", DataSet, i))
                        RedFlag = toval(GetItemValue("RedFlag", DataSet, i))
                        ConsentOnDataKeepByCompanyF = toval(GetItemValue("ConsentOnDataKeepByCompanyF", DataSet, i))
                        ConsentOnDataPrivacyToThirdPartyF = toval(GetItemValue("ConsentOnDataPrivacyToThirdPartyF", DataSet, i))
                        ConsentOnProductAndServiceOfferingF = toval(GetItemValue("ConsentOnProductAndServiceOfferingF", DataSet, i))
                        Remark = tostr(GetItemValue("Remark", DATASET, i))
                        PIC = tostr(GetItemValue("PIC", DATASET, i))

                        If PName = "" Then
                            PName = LTrim(FName & " " & mName & " " & lName)
                        Else
                            Profile_Name(PName, FName, mName, lName)
                        End If

                        ID = tostr(GetItemValue("ID", DataSet, i))
                        If ID = "" Then
                            If CBool(toval(GetAppSetting("ThaiProfile"))) Then
                                ID = Trim(db.GetNextDocNo("Pro", Today.Date, Mid(Get_AlphaNum(PType), 1, 1)))
                            Else
                                If Len(Get_Branch_Alias(Branch)) > 2 Then
                                    ID = Trim(db.GetNextDocNo("Pro", Today.Date, Mid(Get_AlphaNum(PType), 1, 1) + Mid(Get_AlphaNum(FName), 1, 1) + IIf(Mid(Get_AlphaNum(mName), 1, 1) = "", Mid(Get_AlphaNum(FName), 2, 1), Mid(Get_AlphaNum(mName), 1, 1))))
                                Else
                                    ID = Trim(db.GetNextDocNo("Pro", Today.Date, Mid(Get_AlphaNum(PType), 1, 1) + Get_Branch_Alias(Branch) + Mid(Get_AlphaNum(FName), 1, 1) + IIf(Mid(Get_AlphaNum(mName), 1, 1) = "", Mid(Get_AlphaNum(FName), 2, 1), Mid(Get_AlphaNum(mName), 1, 1))))
                                End If
                            End If
                        End If
                        If ID = "" Then Throw New Exception("Generate Profile ID Failed !")

                        DataSet.Tables(0).Rows(i)("ID") = ID

                        If Not db.Execute("Select * into #Profile from Profile Where ID='" & ID & "'") Then Throw New Exception(db.ErrDesc)
                        rsDst = db.OpenRemoteSet("Select * from #Profile")
                        If rsDst.Eof Then
                            rsDst.AddNew()
                        Else
                            rsDst.Edit()
                        End If
                        rsDst.fields("BatchNo") = BatchNo
                        rsDst.fields("PType") = PType
                        rsDst.fields("LOB") = LOB
                        rsDst.fields("IntermediaryF") = IIf(IntermediaryF = 0, False, True)
                        rsDst.fields("CorporateF") = IIf(CorporateF = 0, False, True)
                        rsDst.fields("VatSubsidiesF") = IIf(VatSubsidiesF = 0, False, True)
                        rsDst.fields("Name") = PName
                        rsDst.fields("FirstName") = FName
                        rsDst.fields("MidName") = mName
                        rsDst.fields("LastName") = lName
                        rsDst.fields("Salutation") = tostr(GetItemValue("Salutation", DataSet, i))
                        rsDst.fields("Initial") = tostr(GetItemValue("Initial", DataSet, i))
                        rsDst.fields("Address_1") = tostr(GetItemValue("Address_1", DataSet, i))
                        rsDst.fields("Address_2") = tostr(GetItemValue("Address_2", DataSet, i))
                        rsDst.fields("Address_3") = tostr(GetItemValue("Address_3", DataSet, i))
                        rsDst.fields("City") = tostr(GetItemValue("City", DataSet, i))
                        rsDst.fields("ZipCode") = tostr(GetItemValue("ZipCode", DataSet, i))
                        rsDst.fields("Phone_1") = tostr(GetItemValue("Phone_1", DataSet, i))
                        rsDst.fields("Phone_2") = tostr(GetItemValue("Phone_2", DataSet, i))
                        rsDst.fields("Mobile_1") = tostr(GetItemValue("Mobile_1", DataSet, i))
                        rsDst.fields("Mobile_2") = tostr(GetItemValue("Mobile_2", DataSet, i))
                        rsDst.fields("Fax_1") = tostr(GetItemValue("Fax_1", DataSet, i))
                        rsDst.fields("Fax_2") = tostr(GetItemValue("Fax_2", DataSet, i))
                        rsDst.fields("EMail") = tostr(GetItemValue("Email_1", DataSet, i))
                        rsDst.fields("EMail_2") = tostr(GetItemValue("Email_2", DataSet, i))
                        rsDst.fields("Country") = tostr(GetItemValue("Country", DataSet, i))
                        rsDst.fields("Area") = tostr(GetItemValue("Area", DataSet, i))
                        rsDst.fields("PICName") = tostr(GetItemValue("PICName", DATASET, i))
                        rsDst.fields("PICTitle") = tostr(GetItemValue("PICTitle", DATASET, i))
                        rsDst.fields("PICAddress") = tostr(GetItemValue("PICAddress", DATASET, i))
                        rsDst.fields("PICPhone") = tostr(GetItemValue("PICPhone", DATASET, i))
                        rsDst.fields("CType") = tostr(GetItemValue("CType", DataSet, i))
                        rsDst.fields("CGroup") = tostr(GetItemValue("CGroup", DataSet, i))
                        rsDst.fields("SCGroup") = tostr(GetItemValue("SCGroup", DataSet, i))
                        rsDst.fields("TaxType") = TaxType
                        rsDst.fields("TaxID") = TaxID
                        rsDst.fields("TaxIDOLD") = TaxIDOLD
                        rsDst.fields("TaxName") = tostr(GetItemValue("TaxName", DataSet, i))
                        rsDst.fields("TaxAddress") = tostr(GetItemValue("TaxAddress", DataSet, i))

                        rsDst.fields("VAT") = toval(GetItemValue("VAT", DataSet, i))
                        rsDst.fields("TAX") = toval(GetItemValue("TAX", DataSet, i))
                        rsDst.fields("PID") = tostr(GetItemValue("ReferalID", DataSet, i))
                        rsDst.fields("RType") = tostr(GetItemValue("ReferalType", DataSet, i))

                        If toval(GetItemValue("AccruedTax", DataSet, i)) <> 0 Then
                            rsDst.fields("AccruedTax") = toval(GetItemValue("AccruedTax", DataSet, i))
                        End If

                        If toval(rsDst.fields("VAT")) <> 0 Then VATF = True
                        If toval(rsDst.fields("TAX")) <> 0 Then TAXF = True

                        rsDst.fields("VatF") = VATF
                        rsDst.fields("TaxF") = TAXF
                        rsDst.fields("BVATF") = BVATF
                        rsDst.fields("TVATF") = TVATF
                        rsDst.fields("VATCOMMEXEMPTIF") = VATCOMMEXEMPTIF
                        rsDst.fields("RedFlag") = RedFlag
                        rsDst.fields("ConsentOnDataKeepByCompanyF") = ConsentOnDataKeepByCompanyF
                        rsDst.fields("ConsentOnDataPrivacyToThirdPartyF") = ConsentOnDataPrivacyToThirdPartyF
                        rsDst.fields("ConsentOnProductAndServiceOfferingF") = ConsentOnProductAndServiceOfferingF
                        rsDst.fields("Remark") = Remark

                        rsDst.fields("TaxPct") = TaxLoading
                        rsDst.fields("LicenseNo") = tostr(GetItemValue("LicenseNo", DataSet, i))
                        rsDst.fields("ID") = ID
                        rsDst.fields("RefID") = IIf(RefID <> "", RefID, ID & "/" & LOB)
                        rsDst.fields("XProfileID") = XProfileID
                        rsDst.fields("Owner") = GetUserOwner(dbUser)
                        rsDst.fields("CDate") = Today.Date
                        rsDst.fields("UserID") = dbUser
                        rsDst.fields("Last_Update") = Today.Date
                        rsDst.fields("Last_Time") = Time()
                        rsDst.fields("SYNCDATE") = SYNCDATE
                        rsDst.fields("SYNCTIME") = SYNCTIME
                        rsDst.fields("Last_Opr") = dbUser
                        rsDst.fields("Branch") = Branch
                        rsDst.fields("BIType") = tostr(GetItemValue("BIType", DataSet, i))
                        rsDst.fields("HSCode") = tostr(GetItemValue("HSCode", DataSet, i))
                        rsDst.fields("RestrictedF") = toval(GetItemValue("RestrictedF", DataSet, i))
                        rsDst.fields("Province") = tostr(GetItemValue("Province", DataSet, i))
                        rsDst.fields("Building") = tostr(GetItemValue("Building", DataSet, i))
                        rsDst.fields("Village") = tostr(GetItemValue("Village", DataSet, i))
                        rsDst.fields("RoomNo") = tostr(GetItemValue("RoomNo", DataSet, i))
                        rsDst.fields("FloorNo") = tostr(GetItemValue("FloorNo", DataSet, i))
                        rsDst.fields("HomeNo") = tostr(GetItemValue("HomeNo", DataSet, i))
                        rsDst.fields("MOO") = tostr(GetItemValue("MOO", DataSet, i))
                        rsDst.fields("SOI") = tostr(GetItemValue("SOI", DataSet, i))
                        rsDst.fields("Road") = tostr(GetItemValue("Road", DataSet, i))
                        rsDst.fields("SubDistrict") = tostr(GetItemValue("SubDistrict", DataSet, i))
                        rsDst.fields("District") = tostr(GetItemValue("District", DataSet, i))
                        rsDst.fields("MOID") = tostr(GetItemValue("MOID", DATASET, i))
                        rsDst.fields("Contact") = tostr(GetItemValue("PIC", DATASET, i))
                        rsDst.fields("Title") = tostr(GetItemValue("PICTitle", DATASET, i))
                        rsDst.fields("ContactAddress") = tostr(GetItemValue("PICAddress", DATASET, i))
                        rsDst.fields("ContactPhone") = tostr(GetItemValue("PICPhone", DATASET, i))


                        If tostr(GetItemValue("Address_1", DataSet, i)) = "" Then
                            If tostr(GetItemValue("Building", DataSet, i)) <> "" Then
                                rsDst.fields("Address_1") = tostr(GetItemValue("Building", DataSet, i)) & " ,"
                            End If
                            If tostr(GetItemValue("Village", DataSet, i)) <> "" Then
                                rsDst.fields("Address_1") = tostr(GetItemValue("Village", DataSet, i)) & " ,"
                            End If
                        End If

                        If tostr(GetItemValue("Address_2", DataSet, i)) = "" Then
                            If tostr(GetItemValue("HomeNo", DataSet, i)) <> "" Then
                                rsDst.fields("Address_2") = tostr(GetItemValue("HomeNo", DataSet, i)) + " ,"
                            End If

                            If tostr(GetItemValue("RoomNo", DataSet, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("RoomNo", DataSet, i)) & " ,"
                            End If

                            If tostr(GetItemValue("FloorNo", DataSet, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("FloorNo", DataSet, i)) & " ,"
                            End If

                            If tostr(GetItemValue("MOO", DataSet, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("Moo", DataSet, i))
                            End If

                            If tostr(GetItemValue("Road", DataSet, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("Road", DataSet, i))
                            End If

                            If tostr(GetItemValue("SOI", DataSet, i)) <> "" Then
                                rsDst.fields("Address_2") = rsDst.fields("Address_2") & tostr(GetItemValue("Soi", DataSet, i))
                            End If
                        End If

                        If tostr(GetItemValue("Address_3", DataSet, i)) = "" Then
                            If CBool(toval(GetAppSetting("ProfileAddressExtension"))) Then
                                'NO overwrite for address_2 and address_3
                                If tostr(GetItemValue("SubDistrict", DataSet, i)) <> "" Then
                                    rsDst.fields("SubDistrict") = rsDst.fields("SubDistrict") & tostr(GetItemValue("SubDistrict", DataSet, i)) & " ,"
                                End If
                                If tostr(GetItemValue("District", DataSet, i)) <> "" Then
                                    rsDst.fields("District") = rsDst.fields("District") & tostr(GetItemValue("District", DataSet, i))
                                End If
                            Else
                                'overwrite for address_2 and address_3
                                If tostr(GetItemValue("SubDistrict", DataSet, i)) <> "" Then
                                    rsDst.fields("Address_3") = rsDst.fields("Address_3") & tostr(GetItemValue("SubDistrict", DataSet, i)) & " ,"
                                End If
                                If tostr(GetItemValue("District", DataSet, i)) <> "" Then
                                    rsDst.fields("Address_3") = rsDst.fields("Address_3") & tostr(GetItemValue("District", DataSet, i))
                                End If
                            End If
                        End If

                        rsDst.fields("LicenseDate") = todate(GetItemValue("LicenseDate", DataSet, i))
                        rsDst.fields("LicenseExpiry") = todate(GetItemValue("LicenseExpiry", DataSet, i))
                        rsDst.fields("WPC") = toval(GetItemValue("WPC", DataSet, i))
                        rsDst.fields("Grace") = toval(GetItemValue("Grace", DataSet, i))
                        rsDst.fields("Disc") = toval(GetItemValue("Disc", DataSet, i))
                        rsDst.fields("Comm") = toval(GetItemValue("Comm", DataSet, i))
                        rsDst.fields("NETCOMMF") = toval(GetItemValue("NETCOMMF", DataSet, i))
                        rsDst.fields("MaxOutGo") = toval(GetItemValue("MaxOutGo", DataSet, i))
                        rsDst.fields("Rating") = toval(GetItemValue("Rating", DataSet, i))

                        If CBool(toval(GetAppSetting("ProfileConsentMandatoryF"))) Then
                            rsDst.fields("ConsentOnDataKeepByCompanyF") = 1
                        Else
                            rsDst.fields("ConsentOnDataKeepByCompanyF") = 0
                        End If

                        If db.GetPrivileges("Forceprofileverification") Then
                            rsDst.fields("DumpF") = True
                        End If

                        If db.GetPrivileges("AllowChangeDumpF") Then
                            rsDst.fields("DumpF") = False
                            rsDst.fields("VerifiedF") = True
                            rsDst.fields("VerifiedBy") = dbUser
                            rsDst.fields("VerifiedDate") = Today.Date
                            rsDst.fields("VerifiedTime") = Time()
                        Else
                            rsDst.fields("VerifiedF") = False
                            rsDst.fields("VerifiedBy") = ""
                            rsDst.fields("VerifiedDate") = "1/1/1900"
                            rsDst.fields("VerifiedTime") = ""
                        End If

                        Dim EditMode As Integer
                        rsP = db.OpenRemoteSet("Select * from Profile Where ID='" & ID & "'")
                        If rsP.Eof Then
                            rsP.AddNew()
                        Else
                            rsP.Edit()
                        End If
                        EditMode = rsP.editmode

                        Dim j As Integer
                        For j = 0 To rsP.fields_Count - 1
                            If EditMode = AdEditInProgress Then
                                If rsP.fields(j) <> rsDst.fields(j) Then
                                    If Not Save_Log_Object("PROFILE", "ID", rsP.fields("ID"), "", "", "", "", "", "", rsP.fields_Name(j), rsP.fields(j), rsDst.fields(j), "Update") Then Throw New Exception("Error Save Log Object")
                                End If
                            End If
                            rsP.fields(j) = rsDst.fields(j)
                        Next
                        If Not rsP.Update Then Throw New Exception(rsP.ErrDesc)

                        ' Create Profile_Ext
                        If Not db.Execute("Select * into #Profile_Ext from Profile_Ext Where ID='" & ID & "'") Then Throw New Exception(db.ErrDesc)
                        rsDst2 = db.OpenRemoteSet("Select * from #Profile_Ext")
                        If rsDst2.Eof Then
                            rsDst2.AddNew()
                        Else
                            rsDst2.Edit()
                        End If
                        rsDst2.fields("ID") = ID
                        rsDst2.fields("ID_Type") = tostr(GetItemValue("ID_Type", DataSet, i))
                        rsDst2.fields("ID_No") = tostr(GetItemValue("ID_No", DataSet, i))
                        rsDst2.fields("ID_Name") = tostr(GetItemValue("ID_Name", DataSet, i))
                        rsDst2.fields("Gender") = tostr(GetItemValue("Gender", DataSet, i))
                        rsDst2.fields("BirthDate") = todate(GetItemValue("BirthDate", DataSet, i))
                        rsDst2.fields("BirthPlace") = tostr(GetItemValue("BirthPlace", DataSet, i))
                        rsDst2.fields("WNIF") = toval(GetItemValue("WNIF", DataSet, i))
                        rsDst2.fields("CITIZEN") = tostr(GetItemValue("CITIZEN", DataSet, i))
                        rsDst2.fields("BLNO") = tostr(GetItemValue("BLNo", DataSet, i))
                        rsDst2.fields("Income") = toval(GetItemValue("Income", DataSet, i))
                        rsDst2.fields("CAddress") = tostr(GetItemValue("CAddress", DataSet, i))
                        rsDst2.fields("CPhone_1") = tostr(GetItemValue("CPhone_1", DataSet, i))
                        rsDst2.fields("Marital") = tostr(GetItemValue("Marital", DataSet, i))
                        rsDst2.fields("MothersName") = tostr(GetItemValue("MothersName", DataSet, i))

                        rsDst2.fields("DomicileAddress") = tostr(GetItemValue("DomicileAddress", DataSet, i))
                        rsDst2.fields("AddressSameIDF") = toval(GetItemValue("AddressSameIDF", DataSet, i))
                        rsDst2.fields("PIDExpDate") = todate(GetItemValue("PIDExpDate", DataSet, i))
                        rsDst2.fields("NoPIDExpDateF") = toval(GetItemValue("NoPIDExpDateF", DataSet, i))
                        rsDst2.fields("LastEducation") = tostr(GetItemValue("LastEducation", DataSet, i))
                        rsDst2.fields("MonthlyIncome") = tostr(GetItemValue("MonthlyIncome", DataSet, i))
                        rsDst2.fields("MonthlyExpenses") = tostr(GetItemValue("MonthlyExpenses", DATASET, i))

                        rsDst2.fields("BLDate") = todate(GetItemValue("BLDate", DATASET, i))
                        rsDst2.fields("BLPlace") = tostr(GetItemValue("BLPlace", DATASET, i))


                        'SLIK
                        rsDst2.fields("Num_Dependent") = toval(GetItemValue("Num_Dependent", DataSet, i))
                        rsDst2.fields("Relation") = tostr(GetItemValue("Relation", DataSet, i))
                        rsDst2.fields("Debtor_Class") = tostr(GetItemValue("Debtor_Class", DataSet, i))
                        rsDst2.fields("BirthDate_1") = todate(GetItemValue("BirthDate_1", DataSet, i))
                        rsDst2.fields("Rating_Agency") = tostr(GetItemValue("Rating_Agency", DataSet, i))
                        rsDst2.fields("Rating_Date") = todate(GetItemValue("Rating_Date", DataSet, i))
                        rsDst2.fields("ID_No_Ext_1") = tostr(GetItemValue("ID_No_Ext_1", DataSet, i))
                        rsDst2.fields("ID_No_Ext_2") = tostr(GetItemValue("ID_No_Ext_2", DataSet, i))
                        rsDst2.fields("ID_No_Ext_3") = tostr(GetItemValue("ID_No_Ext_3", DataSet, i))
                        rsDst2.fields("ID_Type_Ext_1") = tostr(GetItemValue("ID_Type_Ext_1", DataSet, i))
                        rsDst2.fields("ID_Type_Ext_2") = tostr(GetItemValue("ID_Type_Ext_2", DataSet, i))
                        rsDst2.fields("ID_Type_Ext_3") = tostr(GetItemValue("ID_Type_Ext_3", DataSet, i))
                        rsDst2.fields("BOD_Gender_1") = tostr(GetItemValue("BOD_Gender_1", DataSet, i))
                        rsDst2.fields("BOD_Gender_2") = tostr(GetItemValue("BOD_Gender_2", DataSet, i))
                        rsDst2.fields("BOD_Gender_3") = tostr(GetItemValue("BOD_Gender_3", DataSet, i))
                        rsDst2.fields("BOD_Address_1") = tostr(GetItemValue("BOD_Address_1", DataSet, i))
                        rsDst2.fields("BOD_Address_2") = tostr(GetItemValue("BOD_Address_2", DataSet, i))
                        rsDst2.fields("BOD_Address_3") = tostr(GetItemValue("BOD_Address_3", DataSet, i))
                        rsDst2.fields("BOD_SubDistrict_1") = tostr(GetItemValue("BOD_SubDistrict_1", DataSet, i))
                        rsDst2.fields("BOD_SubDistrict_2") = tostr(GetItemValue("BOD_SubDistrict_2", DataSet, i))
                        rsDst2.fields("BOD_SubDistrict_3") = tostr(GetItemValue("BOD_SubDistrict_3", DataSet, i))
                        rsDst2.fields("BOD_District_1") = tostr(GetItemValue("BOD_District_1", DataSet, i))
                        rsDst2.fields("BOD_District_2") = tostr(GetItemValue("BOD_District_2", DataSet, i))
                        rsDst2.fields("BOD_District_3") = tostr(GetItemValue("BOD_District_3", DataSet, i))

                        rsDst2.fields("BOD_City_1") = tostr(GetItemValue("BOD_City_1", DataSet, i))
                        rsDst2.fields("BOD_City_2") = tostr(GetItemValue("BOD_City_2", DataSet, i))
                        rsDst2.fields("BOD_City_3") = tostr(GetItemValue("BOD_City_3", DataSet, i))
                        rsDst2.fields("Own_Share") = toval(GetItemValue("Own_Share", DataSet, i))
                        rsDst2.fields("Own_Status") = tostr(GetItemValue("Own_Status", DataSet, i))

                        rsP = db.OpenRemoteSet("Select * from Profile_Ext Where ID='" & ID & "'")
                        If rsP.Eof Then
                            rsP.AddNew()
                        Else
                            rsP.Edit()
                        End If
                        EditMode = rsP.editmode

                        For j = 0 To rsP.fields_Count - 1
                            If EditMode = AdEditInProgress Then
                                If rsP.fields(j) <> rsDst2.fields(j) Then
                                    If Not Save_Log_Object("PROFILE_EXT", "ID", rsP.fields("ID"), "", "", "", "", "", "", rsP.fields_Name(j), rsP.fields(j), rsDst2.fields(j), "Update") Then Throw New Exception("Error Save Log Object")
                                End If
                            End If
                            rsP.fields(j) = rsDst2.fields(j)
                        Next
                        If Not rsP.Update Then Throw New Exception(rsP.ErrDesc)

                        'Create Profile Account
                        If Not Create_ProAcc(ID, Currency, BankAccountNo, BankAccountName, BankName, BankCode, BankBranch, RefType, BCTransferBy, PRDays, Remarks) Then Throw New Exception("Create Profile Account Failed !")

                        'Create Agency
                        If toval(GetItemValue("UploadAgencyF", DataSet, i)) Then
                            If Not db.Execute("Select * into #AIF from AIF Where ID='" & ID & "'") Then Throw New Exception(db.ErrDesc)
                            rsDst3 = db.OpenRemoteSet("Select * from #AIF")
                            If rsDst3.Eof Then
                                rsDst3.AddNew()
                            Else
                                rsDst3.Edit()
                            End If

                            rsDst3.fields("ID") = ID
                            rsDst3.fields("AgentID") = ID
                            rsDst3.fields("AgentName") = tostr(GetItemValue("AgentName", DataSet, i))
                            rsDst3.fields("ContractNo") = tostr(GetItemValue("ContractNo", DataSet, i))
                            rsDst3.fields("AgentLevel") = tostr(GetItemValue("AgentLevel", DataSet, i))
                            rsDst3.fields("AgentType") = tostr(GetItemValue("AgentType", DataSet, i))
                            rsDst3.fields("SDate") = IIf(tostr(GetItemValue("RegisterStartDate", DataSet, i)) <> "", todate(GetItemValue("RegisterStartDate", DataSet, i)), Today.Date)
                            rsDst3.fields("EDate") = IIf(tostr(GetItemValue("RegisterExpiryDate", DataSet, i)) <> "", todate(GetItemValue("RegisterExpiryDate", DataSet, i)), Today.Date)
                            rsDst3.fields("RecruiterID") = tostr(GetItemValue("RecruiterID", DataSet, i))
                            rsDst3.fields("RecruiterType") = "P"
                            rsDst3.fields("AgentStatus") = IIf(tostr(GetItemValue("AgentStatus", DataSet, i)) <> "", tostr(GetItemValue("AgentStatus", DataSet, i)), "N")
                            rsDst3.fields("ULAgentID") = tostr(GetItemValue("ULAgentID", DataSet, i))
                            rsDst3.fields("Status") = "R"
                            rsDst3.fields("CDate") = Today.Date
                            rsDst3.fields("UserID") = dbUser
                            rsDst3.fields("Last_Opr") = dbUser
                            rsDst3.fields("Last_Update") = Today.Date

                            rsP = db.OpenRemoteSet("Select * from AIF Where ID='" & ID & "'")
                            If rsP.Eof Then
                                rsP.AddNew()
                            Else
                                rsP.Edit()
                            End If
                            EditMode = rsP.editmode

                            For j = 0 To rsP.fields_Count - 1
                                If EditMode = AdEditInProgress Then
                                    If rsP.fields(j) <> rsDst3.fields(j) Then
                                        If Not Save_Log_Object("AIF", "AgentID", rsP.fields("ID"), "", "", "", "", "", "", rsP.fields_Name(j), rsP.fields(j), rsDst3.fields(j), "Update") Then Throw New Exception("Error Save Log Object")
                                    End If
                                End If
                                rsP.fields(j) = rsDst3.fields(j)
                            Next
                            If Not rsP.Update Then Throw New Exception(rsP.ErrDesc)

                            If Not db.Execute("Drop Table #AIF") Then Throw New Exception(db.ErrDesc)
                        End If

                        If Not db.Execute("Drop Table #Profile") Then Throw New Exception(db.ErrDesc)
                        If Not db.Execute("Drop Table #Profile_Ext") Then Throw New Exception(db.ErrDesc)
                    Else
                        ErrMsg = "Failed"
                    End If
                    Status = "C"
                    db.Committrans()
                    db.Execute("Update Profile_Batch set ID = '" & ID & "', RefID= '" & IIf(RefID <> "", RefID, ID & "/" & LOB) & "', status = '" & Status & "',  Errmsg = '',  PUSERID  ='" & dbUser & "',  PDATE = '" & Today.Date & "',  PTIME = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)

                Catch ex As Exception
                    ErrMsg = db.ErrDesc
                    If ErrMsg = "" Then ErrMsg = ex.Message

                    db.RollBack()
                    Status = "N"
                    TOTALF = TOTALF + 1
                    db.Execute("Update Profile_Batch set status = '" & Status & "',  Errmsg = '" & ErrMsg & "',  Last_Opr ='" & dbUser & "',  Last_Update = '" & Today.Date & "',  Last_Time = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)

                End Try
            Next

            Rs_Profile_BatchNEW = db.OpenRemoteSet("Select * from Profile_Batch Where BatchNo='" & BatchNo & "'")
            DatasetResult = Rs_Profile_BatchNEW.GetDataSet

            OutParam(0) = BatchNo
            OutParam(1) = TOTALF
            OutParam(2) = TOTALC
            OutParam(3) = ErrMsg

            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = db.ErrDesc
            db.RollBack()
            If ErrMsg = "" Then ErrMsg = ex.Message
            Status = "N"
            db.Execute("Update Profile_Batch set status = '" & Status & "',  Errmsg = '" & ErrMsg & "',  Last_Opr ='" & dbUser & "',  Last_Update = '" & Today.Date & "',  Last_Time = '" & Time() & "'  where batchno = '" & BatchNo & "' AND BID = " & BID)
            CloseConnection()
            Return False
        End Try
    End Function
    <WebMethod(EnableSession:=False)>
    Public Function InsertToProfileBatch(ByVal dbUser As String, ByVal dbPassword As String, ByRef InParam() As Object, ByRef OutParam() As Object, ByRef BatchNo As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet

        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            Dim PName As String
            Dim FName As String
            Dim mName As String
            Dim lName As String
            Dim ID As String
            Dim RefID As String
            Dim XProfileID As String
            Dim Branch As String
            Dim PType As String
            Dim LOB As String
            Dim IntermediaryF As Integer
            Dim CorporateF As Integer
            Dim VatSubsidiesF As Integer
            Dim TaxID As String
            Dim TaxIDOLD As String
            Dim TaxType As String
            Dim TaxLoading As Double
            Dim BankCode As String
            Dim BankBranch As String
            Dim BankAccountNo As String
            Dim BankAccountName As String
            Dim BankName As String
            Dim RefType As String
            Dim BCTransferBy As Integer
            Dim PRDays As Integer
            Dim Remarks As String
            Dim br_branch As String = ""
            Dim br_Currency As String = ""
            Dim Currency As String = ""
            Dim SYNCDATE As String = ""
            Dim SYNCTIME As String = ""

            Dim TAXF As Boolean
            Dim VATF As Boolean
            Dim BVATF As Boolean
            Dim TVATF As Boolean
            Dim VATCOMMEXEMPTIF As Boolean
            Dim RedFlag As Boolean
            Dim ConsentOnDataKeepByCompanyF As Boolean
            Dim ConsentOnDataPrivacyToThirdPartyF As Boolean
            Dim ConsentOnProductAndServiceOfferingF As Boolean
            Dim Remark As String
            Dim PIC As String


            Dim i As Integer
            Dim User As String
            Dim Message As String = ""
            Dim BIRTHDATE As String = ""

            User = InParam(1)
            BatchNo = Format(db.GetIdentity("BatchNoProfile") + 1, "000000")

            rs = db.OpenRemoteSet("Select * from Profile_Batch Where 1<>1")
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Try
                    db.Begintrans()

                    rs.AddNew()

                    If tostr(GetItemValue("SYNCDATE", Dataset, i)) = "" Then
                        SYNCDATE = "1900-01-01"
                    Else
                        SYNCDATE = todate(GetItemValue("SYNCDATE", Dataset, i))
                    End If

                    If tostr(GetItemValue("SYNCTIME", Dataset, i)) = "" Then
                        SYNCTIME = "00:00:00"
                    Else
                        SYNCTIME = tostr(GetItemValue("SYNCTIME", Dataset, i))
                    End If

                    PName = tostr(GetItemValue("Name", Dataset, i))
                    FName = tostr(GetItemValue("FirstName", Dataset, i))
                    mName = tostr(GetItemValue("MidName", Dataset, i))
                    lName = tostr(GetItemValue("LastName", Dataset, i))

                    RefID = tostr(GetItemValue("RefID", Dataset, i))
                    XProfileID = tostr(GetItemValue("XProfileID", Dataset, i))
                    If CBool(toval(GetAppSetting("ThaiProfile"))) Then
                        Dim tmp_br As String = tostr(GetItemValue("Branch", Dataset, i))
                        Dim tmp_branch As String = Right("00000" & tmp_br, 5)
                        Branch = IIf(tostr(GetItemValue("Branch", Dataset, i)) <> "", tmp_branch, "")
                    Else
                        Branch = IIf(tostr(GetItemValue("Branch", Dataset, i)) <> "", tostr(GetItemValue("Branch", Dataset, i)), br_branch)
                    End If
                    PType = tostr(GetItemValue("PType", Dataset, i))
                    LOB = tostr(GetItemValue("LOB", Dataset, i))
                    IntermediaryF = toval(GetItemValue("IntermediaryF", Dataset, i))
                    CorporateF = toval(GetItemValue("CorporateF", Dataset, i))
                    VatSubsidiesF = toval(GetItemValue("VATSubsidiesF", Dataset, i))
                    TaxID = tostr(GetItemValue("TaxID", Dataset, i))
                    TaxIDOLD = tostr(GetItemValue("TaxIDOLD", Dataset, i))
                    TaxType = tostr(GetItemValue("TaxType", Dataset, i))
                    TaxLoading = IIf(toval(GetItemValue("TAX", Dataset, i)) = 0, 0, IIf(toval(GetItemValue("TaxLoading", Dataset, i)) = 0, 100, toval(GetItemValue("TaxLoading", Dataset, i))))

                    Currency = tostr(GetItemValue("Currency", Dataset, i))

                    BankCode = tostr(GetItemValue("BankCode", Dataset, i))
                    BankBranch = tostr(GetItemValue("BankBranch", Dataset, i))
                    BankAccountNo = tostr(GetItemValue("BankAccountNo", Dataset, i))
                    BankAccountName = tostr(GetItemValue("BankAccountName", Dataset, i))
                    BankName = tostr(GetItemValue("BankName", Dataset, i))
                    RefType = tostr(GetItemValue("RefType", Dataset, i))
                    BCTransferBy = toval(GetItemValue("BCTransferBy", Dataset, i))
                    PRDays = toval(GetItemValue("PRDays", Dataset, i))
                    Remarks = tostr(GetItemValue("Remarks", Dataset, i))
                    TAXF = toval(GetItemValue("TAXF", Dataset, i))
                    VATF = toval(GetItemValue("VATF", Dataset, i))
                    BVATF = toval(GetItemValue("BVATF", Dataset, i))
                    TVATF = toval(GetItemValue("TVATF", Dataset, i))
                    VATCOMMEXEMPTIF = toval(GetItemValue("VATCOMMEXEMPTIF", Dataset, i))
                    RedFlag = toval(GetItemValue("RedFlag", Dataset, i))
                    ConsentOnDataKeepByCompanyF = toval(GetItemValue("ConsentOnDataKeepByCompanyF", Dataset, i))
                    ConsentOnDataPrivacyToThirdPartyF = toval(GetItemValue("ConsentOnDataPrivacyToThirdPartyF", Dataset, i))
                    ConsentOnProductAndServiceOfferingF = toval(GetItemValue("ConsentOnProductAndServiceOfferingF", Dataset, i))
                    Remark = tostr(GetItemValue("Remark", Dataset, i))
                    ID = tostr(GetItemValue("ID", Dataset, i))

                    rs.fields("PType") = PType
                    rs.fields("LOB") = LOB
                    rs.fields("IntermediaryF") = IIf(IntermediaryF = 0, False, True)
                    rs.fields("CorporateF") = IIf(CorporateF = 0, False, True)
                    rs.fields("VatSubsidiesF") = IIf(VatSubsidiesF = 0, False, True)
                    rs.fields("Name") = PName
                    rs.fields("FirstName") = FName
                    rs.fields("MidName") = mName
                    rs.fields("LastName") = lName
                    rs.fields("Salutation") = tostr(GetItemValue("Salutation", Dataset, i))
                    rs.fields("Initial") = tostr(GetItemValue("Initial", Dataset, i))
                    rs.fields("Address_1") = tostr(GetItemValue("Address_1", Dataset, i))
                    rs.fields("Address_2") = tostr(GetItemValue("Address_2", Dataset, i))
                    rs.fields("Address_3") = tostr(GetItemValue("Address_3", Dataset, i))
                    rs.fields("City") = tostr(GetItemValue("City", Dataset, i))
                    rs.fields("ZipCode") = tostr(GetItemValue("ZipCode", Dataset, i))
                    rs.fields("Phone_1") = tostr(GetItemValue("Phone_1", Dataset, i))
                    rs.fields("Phone_2") = tostr(GetItemValue("Phone_2", Dataset, i))
                    rs.fields("Mobile_1") = tostr(GetItemValue("Mobile_1", Dataset, i))
                    rs.fields("Mobile_2") = tostr(GetItemValue("Mobile_2", Dataset, i))
                    rs.fields("Fax_1") = tostr(GetItemValue("Fax_1", Dataset, i))
                    rs.fields("Fax_2") = tostr(GetItemValue("Fax_2", Dataset, i))
                    rs.fields("EMail_1") = tostr(GetItemValue("Email_1", Dataset, i))
                    rs.fields("EMail_2") = tostr(GetItemValue("Email_2", Dataset, i))
                    rs.fields("Country") = tostr(GetItemValue("Country", Dataset, i))
                    rs.fields("Area") = tostr(GetItemValue("Area", Dataset, i))
                    rs.fields("PICName") = tostr(GetItemValue("PICName", Dataset, i))
                    rs.fields("PICTitle") = tostr(GetItemValue("PICTitle", Dataset, i))
                    rs.fields("PICAddress") = tostr(GetItemValue("PICAddress", Dataset, i))
                    rs.fields("PICPhone") = tostr(GetItemValue("PICPhone", Dataset, i))
                    rs.fields("CType") = tostr(GetItemValue("CType", Dataset, i))
                    rs.fields("CGroup") = tostr(GetItemValue("CGroup", Dataset, i))
                    rs.fields("SCGroup") = tostr(GetItemValue("SCGroup", Dataset, i))
                    rs.fields("TaxType") = TaxType
                    rs.fields("TaxID") = TaxID
                    rs.fields("TaxIDOLD") = TaxIDOLD
                    rs.fields("TaxName") = tostr(GetItemValue("TaxName", Dataset, i))
                    rs.fields("TaxAddress") = tostr(GetItemValue("TaxAddress", Dataset, i))

                    rs.fields("VAT") = toval(GetItemValue("VAT", Dataset, i))
                    rs.fields("TAX") = toval(GetItemValue("TAX", Dataset, i))
                    rs.fields("PID") = tostr(GetItemValue("ReferalID", Dataset, i))
                    rs.fields("RType") = tostr(GetItemValue("ReferalType", Dataset, i))

                    If toval(GetItemValue("AccruedTax", Dataset, i)) <> 0 Then
                        rs.fields("AccruedTax") = toval(GetItemValue("AccruedTax", Dataset, i))
                    End If

                    If toval(rs.fields("VAT")) <> 0 Then VATF = True
                    If toval(rs.fields("TAX")) <> 0 Then TAXF = True

                    rs.fields("VatF") = VATF
                    rs.fields("TaxF") = TAXF
                    rs.fields("BVATF") = BVATF
                    rs.fields("TVATF") = TVATF
                    rs.fields("VATCOMMEXEMPTIF") = VATCOMMEXEMPTIF
                    rs.fields("RedFlag") = RedFlag
                    rs.fields("ConsentOnDataKeepByCompanyF") = ConsentOnDataKeepByCompanyF
                    rs.fields("ConsentOnDataPrivacyToThirdPartyF") = ConsentOnDataPrivacyToThirdPartyF
                    rs.fields("ConsentOnProductAndServiceOfferingF") = ConsentOnProductAndServiceOfferingF
                    rs.fields("Remark") = Remark
                    rs.fields("PIC") = PIC

                    rs.fields("TaxPct") = TaxLoading
                    rs.fields("LicenseNo") = tostr(GetItemValue("LicenseNo", Dataset, i))
                    rs.fields("ID") = ID
                    rs.fields("RefID") = RefID
                    rs.fields("XProfileID") = XProfileID
                    rs.fields("Owner") = GetUserOwner(dbUser)
                    rs.fields("CDate") = Today.Date
                    rs.fields("UserID") = dbUser
                    rs.fields("Last_Update") = Today.Date
                    rs.fields("Last_Time") = Time()
                    rs.fields("SYNCDATE") = SYNCDATE
                    rs.fields("SYNCTIME") = SYNCTIME
                    rs.fields("Last_Opr") = dbUser
                    rs.fields("Branch") = Branch
                    rs.fields("BIType") = tostr(GetItemValue("BIType", Dataset, i))
                    rs.fields("HSCode") = tostr(GetItemValue("HSCode", Dataset, i))
                    rs.fields("RestrictedF") = toval(GetItemValue("RestrictedF", Dataset, i))
                    rs.fields("Province") = tostr(GetItemValue("Province", Dataset, i))
                    rs.fields("Building") = tostr(GetItemValue("Building", Dataset, i))
                    rs.fields("Village") = tostr(GetItemValue("Village", Dataset, i))
                    rs.fields("RoomNo") = tostr(GetItemValue("RoomNo", Dataset, i))
                    rs.fields("FloorNo") = tostr(GetItemValue("FloorNo", Dataset, i))
                    rs.fields("HomeNo") = tostr(GetItemValue("HomeNo", Dataset, i))
                    rs.fields("MOO") = tostr(GetItemValue("MOO", Dataset, i))
                    rs.fields("SOI") = tostr(GetItemValue("SOI", Dataset, i))
                    rs.fields("Road") = tostr(GetItemValue("Road", Dataset, i))
                    rs.fields("SubDistrict") = tostr(GetItemValue("SubDistrict", Dataset, i))
                    rs.fields("District") = tostr(GetItemValue("District", Dataset, i))
                    rs.fields("MOID") = tostr(GetItemValue("MOID", Dataset, i))


                    rs.fields("LicenseDate") = todate(GetItemValue("LicenseDate", Dataset, i))
                    rs.fields("LicenseExpiry") = todate(GetItemValue("LicenseExpiry", Dataset, i))
                    rs.fields("WPC") = toval(GetItemValue("WPC", Dataset, i))
                    rs.fields("Grace") = toval(GetItemValue("Grace", Dataset, i))
                    rs.fields("Disc") = toval(GetItemValue("Disc", Dataset, i))
                    rs.fields("Comm") = toval(GetItemValue("Comm", Dataset, i))
                    rs.fields("NETCOMMF") = toval(GetItemValue("NETCOMMF", Dataset, i))
                    rs.fields("MaxOutGo") = toval(GetItemValue("MaxOutGo", Dataset, i))
                    rs.fields("Rating") = toval(GetItemValue("Rating", Dataset, i))

                    rs.fields("ID_Type") = tostr(GetItemValue("ID_Type", Dataset, i))
                    rs.fields("ID_No") = tostr(GetItemValue("ID_No", Dataset, i))
                    rs.fields("ID_Name") = tostr(GetItemValue("ID_Name", Dataset, i))
                    rs.fields("Gender") = tostr(GetItemValue("Gender", Dataset, i))
                    rs.fields("BirthDate") = todate(GetItemValue("BirthDate", Dataset, i))
                    rs.fields("BirthPlace") = tostr(GetItemValue("BirthPlace", Dataset, i))
                    rs.fields("WNIF") = toval(GetItemValue("WNIF", Dataset, i))
                    rs.fields("CITIZEN") = tostr(GetItemValue("CITIZEN", Dataset, i))
                    rs.fields("BLNO") = tostr(GetItemValue("BLNo", Dataset, i))
                    rs.fields("Income") = toval(GetItemValue("Income", Dataset, i))
                    rs.fields("CAddress") = tostr(GetItemValue("CAddress", Dataset, i))
                    rs.fields("CPhone_1") = tostr(GetItemValue("CPhone_1", Dataset, i))
                    rs.fields("Marital") = tostr(GetItemValue("Marital", Dataset, i))
                    rs.fields("MothersName") = tostr(GetItemValue("MothersName", Dataset, i))

                    rs.fields("DomicileAddress") = tostr(GetItemValue("DomicileAddress", Dataset, i))
                    rs.fields("AddressSameIDF") = toval(GetItemValue("AddressSameIDF", Dataset, i))
                    rs.fields("PIDExpDate") = todate(GetItemValue("PIDExpDate", Dataset, i))
                    rs.fields("NoPIDExpDateF") = toval(GetItemValue("NoPIDExpDateF", Dataset, i))
                    rs.fields("LastEducation") = tostr(GetItemValue("LastEducation", Dataset, i))
                    rs.fields("MonthlyIncome") = tostr(GetItemValue("MonthlyIncome", Dataset, i))
                    rs.fields("MonthlyExpenses") = tostr(GetItemValue("MonthlyExpenses", Dataset, i))

                    rs.fields("BLDate") = todate(GetItemValue("BLDate", Dataset, i))
                    rs.fields("BLPlace") = tostr(GetItemValue("BLPlace", Dataset, i))

                    'SLIK
                    rs.fields("Num_Dependent") = toval(GetItemValue("Num_Dependent", Dataset, i))
                    rs.fields("Relation") = tostr(GetItemValue("Relation", Dataset, i))
                    rs.fields("Debtor_Class") = tostr(GetItemValue("Debtor_Class", Dataset, i))
                    rs.fields("BirthDate_1") = todate(GetItemValue("BirthDate_1", Dataset, i))
                    rs.fields("Rating_Agency") = tostr(GetItemValue("Rating_Agency", Dataset, i))
                    rs.fields("Rating_Date") = todate(GetItemValue("Rating_Date", Dataset, i))
                    rs.fields("ID_No_Ext_1") = tostr(GetItemValue("ID_No_Ext_1", Dataset, i))
                    rs.fields("ID_No_Ext_2") = tostr(GetItemValue("ID_No_Ext_2", Dataset, i))
                    rs.fields("ID_No_Ext_3") = tostr(GetItemValue("ID_No_Ext_3", Dataset, i))
                    rs.fields("ID_Type_Ext_1") = tostr(GetItemValue("ID_Type_Ext_1", Dataset, i))
                    rs.fields("ID_Type_Ext_2") = tostr(GetItemValue("ID_Type_Ext_2", Dataset, i))
                    rs.fields("ID_Type_Ext_3") = tostr(GetItemValue("ID_Type_Ext_3", Dataset, i))
                    rs.fields("BOD_Gender_1") = tostr(GetItemValue("BOD_Gender_1", Dataset, i))
                    rs.fields("BOD_Gender_2") = tostr(GetItemValue("BOD_Gender_2", Dataset, i))
                    rs.fields("BOD_Gender_3") = tostr(GetItemValue("BOD_Gender_3", Dataset, i))
                    rs.fields("BOD_Address_1") = tostr(GetItemValue("BOD_Address_1", Dataset, i))
                    rs.fields("BOD_Address_2") = tostr(GetItemValue("BOD_Address_2", Dataset, i))
                    rs.fields("BOD_Address_3") = tostr(GetItemValue("BOD_Address_3", Dataset, i))
                    rs.fields("BOD_SubDistrict_1") = tostr(GetItemValue("BOD_SubDistrict_1", Dataset, i))
                    rs.fields("BOD_SubDistrict_2") = tostr(GetItemValue("BOD_SubDistrict_2", Dataset, i))
                    rs.fields("BOD_SubDistrict_3") = tostr(GetItemValue("BOD_SubDistrict_3", Dataset, i))
                    rs.fields("BOD_District_1") = tostr(GetItemValue("BOD_District_1", Dataset, i))
                    rs.fields("BOD_District_2") = tostr(GetItemValue("BOD_District_2", Dataset, i))
                    rs.fields("BOD_District_3") = tostr(GetItemValue("BOD_District_3", Dataset, i))

                    rs.fields("BOD_City_1") = tostr(GetItemValue("BOD_City_1", Dataset, i))
                    rs.fields("BOD_City_2") = tostr(GetItemValue("BOD_City_2", Dataset, i))
                    rs.fields("BOD_City_3") = tostr(GetItemValue("BOD_City_3", Dataset, i))
                    rs.fields("Own_Share") = toval(GetItemValue("Own_Share", Dataset, i))
                    rs.fields("Own_Status") = tostr(GetItemValue("Own_Status", Dataset, i))

                    rs.fields("Address_1") = tostr(GetItemValue("Address_1", Dataset, i))
                    rs.fields("Address_2") = tostr(GetItemValue("Address_2", Dataset, i))
                    rs.fields("Address_3") = tostr(GetItemValue("Address_3", Dataset, i))

                    rs.fields("Currency") = tostr(GetItemValue("Currency", Dataset, i))
                    rs.fields("BankCode") = tostr(GetItemValue("BankCode", Dataset, i))
                    rs.fields("BankBranch") = tostr(GetItemValue("BankBranch", Dataset, i))
                    rs.fields("BankAccountName") = tostr(GetItemValue("BankAccountName", Dataset, i))
                    rs.fields("BankAccountNo") = tostr(GetItemValue("BankAccountNo", Dataset, i))
                    rs.fields("BankName") = tostr(GetItemValue("BankName", Dataset, i))
                    rs.fields("RefType") = tostr(GetItemValue("RefType", Dataset, i))
                    rs.fields("BCTransferBy") = toval(GetItemValue("BCTransferBy", Dataset, i))
                    rs.fields("PRDays") = toval(GetItemValue("PRDays", Dataset, i))
                    rs.fields("Remarks") = tostr(GetItemValue("Remarks", Dataset, i))
                    rs.fields("PIC") = tostr(GetItemValue("PIC", Dataset, i))

                    rs.fields("BatchNo") = BatchNo
                    rs.fields("Status") = "R"
                    rs.fields("CUserID") = User
                    rs.fields("CDate") = Today.Date
                    rs.fields("CTime") = Time()

                    If Not rs.Update Then Throw New Exception(rs.ErrDesc)
                    db.Committrans()

                Catch ex As Exception
                    Message = Message & ex.Message & vbCrLf
                    db.RollBack()
                End Try
            Next i

            OutParam(0) = BatchNo
            OutParam(1) = Message

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

    <WebMethod(EnableSession:=False)> _
    Public Function LoadArBucketFromExcel(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef dsExcel As DataSet, ByRef Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Dim i As Integer
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            Dim rs As RemoteSet
            Dim rscomacc As RemoteSet

            For i = 0 To dsExcel.Tables(0).Rows.Count - 1
                ReDim Preserve OutParam(i)
                'validasi
                If tostr(dsExcel.Tables(0).Rows(i)("A_PolicyNo")) = "" And tostr(dsExcel.Tables(0).Rows(i)("Voucher")) = "" And tostr(dsExcel.Tables(0).Rows(i)("RefNo")) = "" And (tostr(dsExcel.Tables(0).Rows(i)("AccountNo")) = "" And tostr(dsExcel.Tables(0).Rows(i)("Bank")) <> "") Then
                    ErrMsg = "No PolicyNo/Voucher/Refno Data/AccountNo"
                    Throw New Exception(ErrMsg)
                End If
                If tostr(dsExcel.Tables(0).Rows(i)("Bank")) <> "" Then
                    rs = db.OpenRemoteSet("Select * from Bank where Bank = '" & tostr(dsExcel.Tables(0).Rows(i)("Bank")) & "'")
                    If rs.Eof Then
                        ErrMsg = "Invalid Bank Code"
                        Throw New Exception(ErrMsg)
                    End If
                End If
                If tostr(dsExcel.Tables(0).Rows(i)("comacc")) <> "" Then
                    rs = db.OpenRemoteSet("Select * from comacc where comacc = '" & tostr(dsExcel.Tables(0).Rows(i)("comacc")) & "'")
                    If rs.Eof Then
                        ErrMsg = "Invalid ComAcc"
                        Throw New Exception(ErrMsg)
                    End If
                End If
                If tostr(dsExcel.Tables(0).Rows(i)("Branch")) <> "" Then
                    rs = db.OpenRemoteSet("Select * from Branch where Branch = '" & tostr(dsExcel.Tables(0).Rows(i)("Branch")) & "'")
                    If rs.Eof Then
                        ErrMsg = "Invalid Branch"
                        Throw New Exception(ErrMsg)
                    End If
                End If
                'end validation

                rs = db.OpenRemoteSet("Select * from ARBucket where 1 <> 1")
                If rs.Eof Then
                    rs.AddNew()
                Else
                    rs.Edit()
                End If
                rs.fields("A_PolicyNo") = tostr(dsExcel.Tables(0).Rows(i)("A_PolicyNo"))
                rs.fields("Bank") = tostr(dsExcel.Tables(0).Rows(i)("Bank"))
                rs.fields("AccountNo") = tostr(dsExcel.Tables(0).Rows(i)("AccountNo"))
                rs.fields("Voucher") = tostr(dsExcel.Tables(0).Rows(i)("Voucher"))
                rscomacc = db.OpenRemoteSet("Select * from comacc where bankf = 1 and transferf = 1")
                If tostr(dsExcel.Tables(0).Rows(i)("comacc")) = "" Then
                    rs.fields("ComAcc") = rscomacc.fields("ComAcc")
                Else
                    rs.fields("ComAcc") = tostr(dsExcel.Tables(0).Rows(i)("comacc"))
                End If
                If tostr(dsExcel.Tables(0).Rows(i)("Branch")) = "" Then
                    rs.fields("Branch") = rscomacc.fields("Branch")
                Else
                    rs.fields("Branch") = tostr(dsExcel.Tables(0).Rows(i)("Branch"))
                End If
                rs.fields("RefType") = tostr(dsExcel.Tables(0).Rows(i)("RefType"))
                rs.fields("RefNo") = tostr(dsExcel.Tables(0).Rows(i)("RefNo"))
                rs.fields("RefDate") = todate(dsExcel.Tables(0).Rows(i)("RefDate"))
                rs.fields("AmountPaid") = toval(dsExcel.Tables(0).Rows(i)("AmountPaid"))
                rs.fields("Rate") = toval(dsExcel.Tables(0).Rows(i)("Rate"))
                rs.fields("Status") = tostr("R")
                rs.fields("Last_Opr") = dbUser
                rs.fields("PDate") = todate(Date.Today)
                If Not rs.Update Then Throw New Exception(ErrMsg)
                OutParam(i) = ""
            Next

            db.Committrans()
            CloseConnection()
            Return True
        Catch ex As Exception
            'ErrMsg = ex.Message
            OutParam(i) = ErrMsg
            db.RollBack()
            CloseConnection()
            Return False
        End Try
    End Function

    <WebMethod(EnableSession:=False)> _
    Public Function Process_Receipt_Payment(ByVal dbUser As String, ByVal dbPassword As String, ByRef InParam() As Object, ByRef OutParam() As Object, ByRef BatchNo As String, ByVal Dataset As DataSet, ByRef ErrMsg As String) As Boolean
        Dim rs As RemoteSet

        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            Dim i As Integer
            Dim BackDatedF As Boolean
            Dim User As String

            User = InParam(1)
            BatchNo = Format(db.GetIdentity("BatchNo") + 1, "000000")

            rs = db.OpenRemoteSet("Select * from sysuser Where id = '" & User & "'")
            BackDatedF = rs.fields("BackDatedF")

            rs = db.OpenRemoteSet("Select * from RPBucket Where 1<>1")
            For i = 0 To Dataset.Tables(0).Rows.Count - 1
                Try
                    db.Begintrans()

                    rs.AddNew()
                    rs.fields("Branch") = tostr(Dataset.Tables(0).Rows(i)("Branch"))
                    rs.fields("CT") = tostr(Dataset.Tables(0).Rows(i)("CT"))
                    rs.fields("ID") = tostr(Dataset.Tables(0).Rows(i)("ID"))
                    rs.fields("Date") = todate(Dataset.Tables(0).Rows(i)("Date"))
                    rs.fields("VType") = tostr(Dataset.Tables(0).Rows(i)("VType"))
                    rs.fields("ComAcc") = tostr(Dataset.Tables(0).Rows(i)("ComAcc"))
                    rs.fields("remarks") = tostr(Dataset.Tables(0).Rows(i)("Remarks"))
                    rs.fields("refType") = tostr(Dataset.Tables(0).Rows(i)("RefType"))
                    rs.fields("RefDate") = todate(Dataset.Tables(0).Rows(i)("RefDate"))
                    rs.fields("RefNo") = tostr(Dataset.Tables(0).Rows(i)("RefNo"))
                    rs.fields("Voucher") = tostr(Dataset.Tables(0).Rows(i)("Voucher"))
                    rs.fields("Currency") = tostr(Dataset.Tables(0).Rows(i)("Currency"))
                    rs.fields("Nominal") = toval(Dataset.Tables(0).Rows(i)("Nominal"))
                    rs.fields("BatchNo") = BatchNo
                    rs.fields("Last_opr") = User
                    rs.fields("last_date") = Today.Date
                    rs.fields("account") = tostr(Dataset.Tables(0).Rows(i)("Account"))
                    rs.fields("account_1") = tostr(Dataset.Tables(0).Rows(i)("Account_1"))
                    rs.fields("Vaccount") = tostr(Dataset.Tables(0).Rows(i)("VAccount"))
                    rs.fields("DocNo") = tostr(Dataset.Tables(0).Rows(i)("refno_polis"))
                    rs.fields("Rate") = toval(Dataset.Tables(0).Rows(i)("Rate"))
                    rs.fields("Payment") = toval(Dataset.Tables(0).Rows(i)("Payment"))
                    rs.fields("Nvoucher") = tostr(Dataset.Tables(0).Rows(i)("Nvoucher"))
                    rs.fields("Payment_date") = todate(Dataset.Tables(0).Rows(i)("Payment_date"))
                    rs.fields("Status") = "T"

                    If Not rs.Update Then Throw New Exception(rs.ErrDesc)
                    db.Committrans()

                Catch ex As Exception
                    OutParam(i) = ex.Message
                    db.RollBack()
                End Try
            Next

            db.Begintrans()
            If db.GetUserType() <> "SUPERVISOR" Then
                If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Branch ', 512)  where batchno = '" & BatchNo & "' and branch not in (select branch from sysuserbranch where id = '" & User & "')") Then Throw New Exception(db.ErrDesc)
                If db.RecordsAffected > 0 Then GoTo ErrorMsg
            Else
                If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Branch ', 512)  where batchno = '" & BatchNo & "' and branch not in (select branch from branch where Op_flag=1 and Active_Flag = 1)") Then Throw New Exception(db.ErrDesc)
                If db.RecordsAffected > 0 Then GoTo ErrorMsg
            End If

            If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong CT ', 512)  where batchno = '" & BatchNo & "' and CT<>'' and not exists (select * from CT where RPBucket.CT=CT.CT and CT.AllowedF=1 and CT.Post_Flag=1)") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Profile ', 512)  where batchno = '" & BatchNo & "' and ID not in (Select ID from Profile)") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket Set Status = 'C',  errmsg = left(errmsg  + ' RefDate cannot higher than Date ', 512)  where batchno = '" & BatchNo & "' and DATEDIFF(day,Date,RefDate)>0") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Voucher Type ', 512)  where batchno = '" & BatchNo & "' and Nominal < 0 and VType not in (Select VType from Vtype Where sl_cat = 'P' and creditorF=1)") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Voucher Type ', 512)  where batchno = '" & BatchNo & "' and Nominal >= 0 and VType not in (Select VType from Vtype Where sl_cat = 'P' and DebtorF=1)") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket Set Status = 'C',  errmsg = left(errmsg  + ' Wrong ComAcc for User : ',  255)  from RPBucket where batchno = '" & BatchNo & "' and (ComAcc='' or (ComAcc<>'' and not exists (Select * from FN_UserComAcc(Last_Opr,Branch) C where RpBucket.ComAcc=C.ComAcc)))") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Ref. Type ', 512)  where batchno = '" & BatchNo & "' and RefType not in (Select PaymentType from PaymentType)") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Reference No ', 512)  where batchno = '" & BatchNo & "' and RefType <> 'Cash' and RefNo = ''") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Currency ', 512)  where batchno = '" & BatchNo & "' and Currency not in (Select Currency from Currency Where AllowedF=1)") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Nominal ', 512)  where batchno = '" & BatchNo & "' and Nominal = 0") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If Not db.Execute("Update RPBucket set status = 'C',  errmsg = left(errmsg  + ' Wrong Account Suspend ', 512)  where batchno = '" & BatchNo & "' and not exists (select * from Account A where RpBucket.Account=A.Account and suspend_flag=1 and Post_Flag=1)") Then Throw New Exception(db.ErrDesc)
            If db.RecordsAffected > 0 Then GoTo ErrorMsg

            If BackDatedF Then
                If Not db.Execute("Update R Set Status='C',ErrMsg=left(errmsg  + 'BackDated Date Constraint : ', 512)  from RPBucket R, Branch B where batchno = '" & BatchNo & "' and R.Branch=B.Branch and B.Active_Flag=1 and B.Op_Flag=1 and DATEDIFF(day,B.End_Date,R.Date)<=0 and DATEDIFF(day,B.BackDatedDate,R.Date)>0") Then Throw New Exception(db.ErrDesc)
                If db.RecordsAffected > 0 Then GoTo ErrorMsg

                If Not db.Execute("Update R Set BackDatedF=1 from RPBucket R, Branch B where batchno = '" & BatchNo & "' and R.Branch=B.Branch and B.Active_Flag=1 and B.Op_Flag=1 and Status='T' and DATEDIFF(day,B.End_Date,R.Date)<=0 and not(DATEDIFF(day,B.BackDatedDate,R.Date)>0)") Then Throw New Exception(db.ErrDesc)
            Else
                If Not db.Execute("Update R Set Status='C',ErrMsg=left(errmsg  + 'BackDated Transaction : ', 512)  from RPBucket R, Branch B where batchno = '" & BatchNo & "' and R.Branch=B.Branch and B.Active_Flag=1 and B.Op_Flag=1 and DATEDIFF(day,B.End_Date,R.Date)<=0") Then Throw New Exception(db.ErrDesc)
                If db.RecordsAffected > 0 Then GoTo ErrorMsg
            End If

            If Not db.Execute("Update RPBucket set status = 'R' where batchno = '" & BatchNo & "' and status = 'T' ") Then Throw New Exception(db.ErrDesc)

            db.Committrans()

            rs = db.OpenRemoteSet("select count(PID) as 'Total' from rpbucket where batchno = '" & BatchNo & "' and status = 'R' ")
            OutParam(0) = rs.fields("Total")

            rs = db.OpenRemoteSet("select count(PID) as 'Total' from rpbucket where batchno = '" & BatchNo & "' and status = 'C' ")
            OutParam(1) = rs.fields("Total")

            If InParam(0) = 1 Then  '1 : Immediate
                Me.Direct_Call_Process_Execute("ProcessSC.NetV3.Proclib", "AUTORECEIPTPAYMENT_PROCESS", Today.Date, BatchNo)

                rs = db.OpenRemoteSet("select count(PID) as 'Total' from rpbucket where batchno = '" & BatchNo & "' and status = 'P' ")
                OutParam(2) = rs.fields("Total")
                OutParam(3) = OutParam(0) - OutParam(2)
            End If

            CloseConnection()
            Return True
        Catch ex As Exception
            ErrMsg = db.ErrDesc
            db.RollBack()
            CloseConnection()
            If ErrMsg = "" Then ErrMsg = ex.Message
            Return False
        End Try

ErrorMsg:
        rs = db.OpenRemoteSet("select count(PID) as 'Total' from rpbucket where batchno = '" & BatchNo & "' and status = 'R' ")
        OutParam(0) = rs.fields("Total")

        rs = db.OpenRemoteSet("select count(PID) as 'Total' from rpbucket where batchno = '" & BatchNo & "' and status = 'C' ")
        OutParam(1) = rs.fields("Total")

        If Not db.Execute("Update RPBucket set status = '' where batchno = '" & BatchNo & "' and status in ('R','T') ") Then Throw New Exception(db.ErrDesc)

        db.Committrans()
        CloseConnection()
        Return True
    End Function

    <WebMethod(EnableSession:=False)>
    Public Function SaveUploadExcel(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim BatchID As Long
            Dim ImageID As Long
            Dim SaveLocation As Object
            Dim FileName As String
            Dim FileType As String
            Dim rs As RemoteSet

            BatchID = getInParamValue(InParam, 0)
            FileName = getInParamValue(InParam, 1)
            SaveLocation = getInParamValue(InParam, 2)
            FileType = getInParamValue(InParam, 3)

            rs = db.OpenRemoteSet("Select * from SysbatchUpload Where BatchID=" & BatchID & "")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            rs.fields("BatchID") = BatchID
            rs.fields("FileName") = FileName
            rs.fields("SaveLocation") = SaveLocation
            rs.fields("FileType") = FileType
            rs.fields("Last_Update") = Today.Date
            rs.fields("Last_opr") = dbUser
            rs.fields("Last_Time") = Time()
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ReDim OutParam(0)
            OutParam(0) = ImageID
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
    Public Function SplitFileExcel(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef OutParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim BatchID As Long
            Dim BID As Long
            Dim Part As Long
            Dim SaveLocation As Object
            Dim FileName As String
            Dim FileType As String
            Dim rs As RemoteSet

            BatchID = getInParamValue(InParam, 0)
            FileName = getInParamValue(InParam, 1)
            SaveLocation = getInParamValue(InParam, 2)
            FileType = getInParamValue(InParam, 3)
            Part = getInParamValue(InParam, 4)

            BID = db.GetIdentity("SysbatchDetUpload") + 1

            rs = db.OpenRemoteSet("Select * from SysbatchDetUpload Where BatchID=" & BatchID & " and BID=" & BID & "")
            If rs.Eof Then
                rs.AddNew()
            Else
                rs.Edit()
            End If
            rs.fields("BID") = BID
            rs.fields("BatchID") = BatchID
            rs.fields("FileName") = FileName
            rs.fields("SaveLocation") = SaveLocation
            rs.fields("FileType") = FileType
            rs.fields("Part") = Part
            rs.fields("CreateDate") = Today.Date
            rs.fields("CreateUser") = dbUser
            rs.fields("CreateTime") = Time()
            rs.fields("VerifyF") = False
            If Not rs.Update Then Throw New Exception(db.ErrDesc)

            ReDim OutParam(0)
            OutParam(0) = BID
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
    Public Function DropFileExcel(ByVal dbUser As String, ByVal dbPassword As String, ByVal InParam() As Object, ByRef ErrMsg As String) As Boolean
        Try
            If Not OpenConnection(dbUser, dbPassword, ErrMsg) Then
                Return False
            End If

            db.Begintrans()

            ' Get Input Parameter
            Dim BID As Long
            Dim SaveLocation As String

            BID = getInParamValue(InParam, 0)
            SaveLocation = getInParamValue(InParam, 1)

            Dim rs As RemoteSet
            rs = db.OpenRemoteSet("Select * from SysbatchDetUpload Where BID=" & BID & "")
            If rs.Eof Then
                ErrMsg = "No File Excel Data !"
                Throw New Exception(ErrMsg)
            End If

            If Dir(SaveLocation) <> "" Then
                Kill(SaveLocation)
            End If

            If Not rs.Delete Then Throw New Exception(db.ErrDesc)

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
    Public Function isNumberString(ByVal str As String, Optional ByVal ch As String = "") As Boolean
        Dim c As String
        Dim i As Integer
        Dim j As Integer
        For i = 1 To Len(str)
            c = Mid(str, i, 1)
            If InStr("0123456789" & ch, c) = 0 Then
                isNumberString = False
                Exit Function
            End If
        Next
        isNumberString = True
    End Function
    Public Sub Profile_Name(ByVal PName As String, ByRef FName As String, ByRef mName As String, ByRef lName As String)
        Dim TStr As String
        Dim tpos As Integer
        FName = ""
        mName = ""
        lName = ""
        TStr = PName
        tpos = InStr(TStr, " ")
        If toval(tpos) <> 0 Then
            FName = Left(TStr, tpos - 1)
            TStr = Trim(Mid(TStr, tpos + 1, Len(TStr) - tpos))
            tpos = InStr(TStr, " ")
            If toval(tpos) <> 0 Then
                mName = Left(TStr, tpos - 1)
                TStr = Trim(Mid(TStr, tpos + 1, Len(TStr) - tpos))
                lName = TStr
            Else
                mName = TStr
            End If
        Else
            FName = TStr
        End If
    End Sub
    Private Function Get_AlphaNum(ByVal str As String) As String
        Dim i As Integer
        Dim j As Integer
        Dim C As String
        Dim NewStr As String
        NewStr = ""
        j = Len(str)
        For i = 1 To j
            C = Mid(str, i, 1)
            If (C >= "A" And C <= "Z") Or (C >= "a" And C <= "z") Or (C >= "0" And C <= "9") Then
                NewStr = NewStr + C
            End If
        Next i
        Get_AlphaNum = NewStr
    End Function
    Private Function Get_Branch_Alias(ByVal Branch As String) As String
        Dim rs As RemoteSet
        Dim BAlias As String

        rs = db.OpenRemoteSet("Select * from Branch Where Branch='" & Branch & "'")
        BAlias = tostr(rs.fields("BAlias"))
        If BAlias = "" Then BAlias = Branch
        Get_Branch_Alias = BAlias
        rs.CloseRemoteSet()
    End Function
    Private Function GetUserOwner(ByVal ID As String) As String
        Dim rs As RemoteSet
        rs = db.OpenRemoteSet("Select Owner from SysUser Where ID='" & ID & "'")
        If Not rs.Eof Then
            GetUserOwner = tostr(rs.fields("Owner"))
            Exit Function
        End If
        GetUserOwner = ""
    End Function
    Private Function Create_ProAcc(ByVal ID As String, ByVal CCY As String, ByVal AccountNo As String, ByVal BankAccount As String, ByVal BankName As String, ByVal Bank As String, ByVal BankBranch As String, ByVal RefType As String, ByVal BCTransferby As Integer, ByVal PRDays As Integer, ByVal Remarks As String) As Boolean
        On Error GoTo errmsg

        Dim rs As RemoteSet
        Dim Description As String

        If Bank <> "" Then
            rs = db.OpenRemoteSet("Select Bank, Description From Bank Where Bank = '" & Bank & "'")
            Bank = tostr(rs.fields("Bank"))
            If Not rs.Eof Then
                BankName = rs.fields("Description")
                If BankBranch <> "" Then
                    BankName = BankName & " " & BankBranch
                End If
            End If
        End If

        rs = db.OpenRemoteSet("Select Name from Profile Where ID='" & ID & "'")
        Description = rs.fields("Name")

        Dim rsP As RemoteSet
        Dim RestrictedF As Boolean = False
        rsP = db.OpenRemoteSet("Select * from ProAcc Where ID='" & ID & "' and Currency='" & CCY & "'")
        If rsP.Eof Then
            rsP.AddNew()
            rsP.fields("ID") = ID
            rsP.fields("Currency") = CCY
            rsP.fields("ProAcc") = ID + "-" + Left(CCY, 3)
            rsP.fields("Description") = Description
            rsP.fields("UserID") = dbUser
            rsP.fields("CDate") = Today.Date
        Else
            rsP.Edit()
        End If
        Dim EditMode As Integer
        EditMode = rsP.editmode

        If DecryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), rsP.fields("AccountNo"), 1) <> AccountNo Then
            If EditMode = AdEditInProgress AndAlso Not Save_Log_Object("PROACC", "PROACC", rsP.fields("ProAcc"), "", "", "", "", "", "", "AccountNo", DecryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), rsP.fields("AccountNo"), 1), AccountNo, "Update") Then Throw New Exception("Error Save Log Object")
            rsP.fields("AccountNo") = EncryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), AccountNo)
            RestrictedF = True
        End If

        If DecryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), rsP.fields("BankAccount"), 1) <> BankAccount Then
            If EditMode = AdEditInProgress AndAlso Not Save_Log_Object("PROACC", "PROACC", rsP.fields("ProAcc"), "", "", "", "", "", "", "BankAccount", DecryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), rsP.fields("BankAccount"), 1), BankAccount, "Update") Then Throw New Exception("Error Save Log Object")
            rsP.fields("BankAccount") = EncryptByPassPhrase(GetAppSetting("EncryptionPassPhrase"), BankAccount)
            RestrictedF = True
        End If

        If rsP.fields("BankName") <> BankName Then
            If EditMode = AdEditInProgress AndAlso Not Save_Log_Object("PROACC", "PROACC", rsP.fields("PROACC"), "", "", "", "", "", "", "BankName", rs.fields("BankName"), BankName, "Update") Then Throw New Exception("Error Save Log Object")
            rsP.fields("BankName") = BankName
            RestrictedF = True
        End If
        If rsP.fields("Bank") <> Bank Then
            If EditMode = AdEditInProgress AndAlso Not Save_Log_Object("PROACC", "PROACC", rsP.fields("PROACC"), "", "", "", "", "", "", "Bank", rs.fields("Bank"), Bank, "Update") Then Throw New Exception("Error Save Log Object")
            rsP.fields("Bank") = Bank
            RestrictedF = True
        End If
        If rsP.fields("BankBranch") <> BankBranch Then
            If EditMode = AdEditInProgress AndAlso Not Save_Log_Object("PROACC", "PROACC", rsP.fields("PROACC"), "", "", "", "", "", "", "BankBranch", rs.fields("BankBranch"), BankBranch, "Update") Then Throw New Exception("Error Save Log Object")
            rsP.fields("BankBranch") = BankBranch
            RestrictedF = True
        End If
        If rsP.fields("RefType") <> RefType Then
            If EditMode = AdEditInProgress AndAlso Not Save_Log_Object("PROACC", "PROACC", rsP.fields("PROACC"), "", "", "", "", "", "", "RefType", rs.fields("RefType"), RefType, "Update") Then Throw New Exception("Error Save Log Object")
            rsP.fields("RefType") = RefType
            RestrictedF = True
        End If
        If toval(rsP.fields("BCTransferby")) <> BCTransferby Then
            If EditMode = AdEditInProgress AndAlso Not Save_Log_Object("PROACC", "PROACC", rsP.fields("PROACC"), "", "", "", "", "", "", "BCTransferby", rs.fields("BCTransferby"), BCTransferby, "Update") Then Throw New Exception("Error Save Log Object")
            rsP.fields("BCTransferby") = BCTransferby
            RestrictedF = True
        End If
        If toval(rsP.fields("PRDays")) <> PRDays Then
            If EditMode = AdEditInProgress AndAlso Not Save_Log_Object("PROACC", "PROACC", rsP.fields("PROACC"), "", "", "", "", "", "", "PRDays", rs.fields("PRDays"), PRDays, "Update") Then Throw New Exception("Error Save Log Object")
            rsP.fields("PRDays") = PRDays
            RestrictedF = True
        End If
        If rsP.fields("Remarks") <> Remarks Then
            If EditMode = AdEditInProgress AndAlso Not Save_Log_Object("PROACC", "PROACC", rsP.fields("PROACC"), "", "", "", "", "", "", "Remarks", rs.fields("Remarks"), Remarks, "Update") Then Throw New Exception("Error Save Log Object")
            rsP.fields("Remarks") = Remarks
            RestrictedF = True
        End If


        If Not db.GetPrivileges("UNRESTRICTPROACC") Then
            rsP.fields("RestrictedF") = RestrictedF
            rsP.fields("PUserDate") = Today.Date
            rsP.fields("PUserID") = dbUser
        Else
            rsP.fields("RestrictedF") = False
            rsP.fields("PUserDate") = Today.Date
            rsP.fields("PUserID") = dbUser
        End If
        rsP.fields("Last_Opr") = dbUser
        rsP.fields("Last_Update") = Today.Date
        '' rsP.fields("Last_Time") = Time()
        If Not rsP.Update Then GoTo errmsg

        Create_ProAcc = True
        Exit Function
errmsg:
        Create_ProAcc = False
    End Function
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
    Public Function DecryptByPassPhrase(ByVal PassPhrase As String, ByVal EncyrptedString As String, ByVal EncryptedType As Integer) As Object
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
    Private Function Save_Log_Object(ByVal LogName As String, ByVal LogKeyID As String, ByVal LogID As String, ByVal LogKeyIDExt1 As String, ByVal LogIDExt1 As String, ByVal LogKeyIDExt2 As String, ByVal LogIDExt2 As String, ByVal LogKeyIDExt3 As String, ByVal LogIDExt3 As String, ByVal ColName As String, ByVal ColValueOld As String, ByVal ColValueNew As String, ByVal Action As String) As Boolean
        Dim rs As Remoteset
        On Error GoTo errmessage
        If LogName <> "" And (ColValueOld <> ColValueNew Or Action <> "Update") Then
            rs = db.OpenRemoteSet("Select * from SysLogObject Where 1<>1")
            rs.addnew()
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
            If Not rs.Update Then GoTo errmessage
        End If
        Save_Log_Object = True
        Exit Function
errmessage:
        Save_Log_Object = False
    End Function
    Private Function GetItemValue(ByVal FName As String, ByVal Dataset As DataSet, ByVal i As Integer) As Object
        Try
            GetItemValue = Dataset.Tables(0).Rows(i)(FName)
        Catch ex As Exception
            GetItemValue = ""
        End Try
    End Function
    Private ReadOnly Property ErrDesc() As String
        Get
            ErrDesc = StrMsg
        End Get
    End Property
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