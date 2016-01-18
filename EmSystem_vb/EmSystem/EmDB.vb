Imports System.Web.UI.WebControls
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Data
Imports System.Data.Linq
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class EmDB

    Dim emsys As New StrTool()
    ' 是否是使用在公司(True=抓iniKernal(預設) , False=必須自己設定連線字串參數connStrSelf)
    Dim IsCompany As Boolean = True
    ' 是否是 windows form
    Dim IsForm As Boolean = False
    ' 是否存 SQLLog (只存執行成功的)
    Dim IsSaveSql As Boolean = False
    ' 是否啟用阻擋 SQL injection
    Dim IsSqlAtk As Boolean = True
    ' 自定義連線字串，當isCompany為False才會使用
    Dim ConnStrSelf As String = ""


    Dim sql As String = ""

    Public Sub New()
        Dim section As EmDBSection = TryCast(ConfigurationManager.GetSection("EmDB"), EmDBSection)
        If section IsNot Nothing Then
            IsCompany = section.GetIsCompany
            IsForm = section.GetIsForm
            IsSaveSql = section.GetIsSaveSql
            IsSqlAtk = section.GetIsSqlAtk
            If Not IsCompany Then
                ConnStrSelf = section.GetConnStr
            End If
        End If
    End Sub

    'Power By JKY
    ''' <summary>
    ''' 使用的連線字串
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ConnStr As String
        Get
            If IsCompany Then
                Dim sb As New StringBuilder(255)
                Dim strTool As New StrTool()
                strTool.isForm = IsForm
                'Dim iniLocation As String = "D:\Web\JKY WEB\IWKernel.ini"
                Dim str3 As String = strTool.DecodeStr(strTool.GetIniInfo("IntraWare", "Server_Name").ToString())
                Dim str2 As String = strTool.DecodeStr(strTool.GetIniInfo("IntraWare", "Database_Name").ToString())
                Dim str4 As String = strTool.DecodeStr(strTool.GetIniInfo("IntraWare", "SQL_ID").ToString())
                Dim str5 As String = strTool.DecodeStr(strTool.GetIniInfo("IntraWare", "SQL_Password").ToString())
                Return "data source=" & str3 & ";initial catalog=" & str2 & ";user id=" & str4 & ";password=" & str5
            Else
                Return ConnStrSelf
            End If
        End Get
    End Property


    Private mySQLDB As New SqlCommand()


#Region "SELECT"

    ''' <summary>
    ''' 查詢全部資料，查不到或錯誤則回傳空List(count=0)
    ''' </summary>
    ''' <typeparam name="T">型態</typeparam>
    ''' <returns></returns>
    Public Function SelectAll(Of T As New)() As List(Of T)
        Try
            Dim myDB As New System.Data.Linq.DataContext(connStr)
            Dim list As List(Of T) = myDB.ExecuteQuery(Of T)("SELECT * FROM " & GetType(T).Name).ToList()
            Return list
        Catch e As Exception
            emsys.Debuger(e.ToString())
            Return New List(Of T)
            'myDB.Connection.Dispose();
        Finally
        End Try
    End Function

    ''' <summary>
    ''' 查詢(使用查詢物件)，查不到或錯誤則回傳空List(count=0)
    ''' </summary>
    ''' <typeparam name="T">型態</typeparam>
    ''' <param name="spArray">查詢參數陣列</param>
    ''' <returns></returns>
    Public Function [Select](Of T As New)(ParamArray spArray As SelectParams()) As List(Of T)
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            Dim operation As String
            sql = "SELECT * FROM " & GetType(T).Name & " WHERE "
            For Each sp As SelectParams In spArray
                If sp.key IsNot Nothing AndAlso sp.value IsNot Nothing Then
                    If sp.type Then
                        operation = "="
                    Else
                        operation = "!="
                    End If
                    sql = sql & sp.key & operation & "'" & sp.value & "'"
                    If Not sp.Equals(spArray(UBound(spArray))) Then
                        sql = sql & " AND "
                    End If
                End If
            Next
            Dim list As List(Of T) = myDB.ExecuteQuery(Of T)(sql).ToList()
            Return list
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return New List(Of T)
        Finally
        End Try
    End Function

    ''' <summary>
    ''' 查詢(SQL語法)，查不到或錯誤則回傳空List(count=0)
    ''' </summary>
    ''' <typeparam name="T">型態</typeparam>
    ''' <param name="sqlCommon">SQL語法</param>
    ''' <returns></returns>
    Public Function [Select](Of T As New)(sqlCommon As String) As List(Of T)
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            Dim list As List(Of T) = myDB.ExecuteQuery(Of T)(sqlCommon).ToList()
            Return list
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return New List(Of T)
        Finally
        End Try
    End Function

    ''' <summary>
    ''' 查詢(單一查詢條件)，查不到或錯誤則回傳空List(count=0)
    ''' </summary>
    ''' <typeparam name="T">型態</typeparam>
    ''' <param name="key">欄位名稱</param>
    ''' <param name="value">值</param>
    ''' <param name="type">true:等於、false:不等於</param>
    ''' <returns></returns>
    Public Function [Select](Of T As New)(key As String, value As String, type As Boolean) As List(Of T)
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try

            Dim operation As String
            If key IsNot Nothing AndAlso value IsNot Nothing Then

                If type Then
                    operation = "="
                Else
                    operation = "!="
                End If
                sql = "SELECT * FROM " & GetType(T).Name & " WHERE " & key & operation & "'" & value & "'"
            Else
                sql = "SELECT * FROM " & GetType(T).Name
            End If

            Dim list As List(Of T) = myDB.ExecuteQuery(Of T)(sql).ToList()
            Return list
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return New List(Of T)
        Finally
        End Try
    End Function

    ''' <summary>
    ''' 查詢(回傳第一筆first)，查不到或錯誤則回傳NULL
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="key"></param>
    ''' <param name="value"></param>
    ''' <param name="type"></param>
    ''' <returns></returns>
    Public Function SelectOne(Of T As New)(key As String, value As String, type As Boolean) As T
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            Dim operation As String
            If key IsNot Nothing AndAlso value IsNot Nothing Then

                If type Then
                    operation = "="
                Else
                    operation = "!="
                End If
                sql = "SELECT * FROM " & GetType(T).Name & " WHERE " & key & operation & "'" & value & "'"
            Else
                sql = "SELECT * FROM " & GetType(T).Name
            End If

            Dim list As List(Of T) = myDB.ExecuteQuery(Of T)(sql).ToList()
            If list.Count > 0 Then
                Return list.First()
            Else
                Return Nothing
            End If
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return Nothing
        Finally
        End Try
    End Function

    ''' <summary>
    ''' 查詢(回傳第一筆first，使用查詢物件)，查不到或錯誤則回傳NULL
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="spArray"></param>
    ''' <returns></returns>
    Public Function SelectOne(Of T As New)(ParamArray spArray As SelectParams()) As T
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            Dim operation As String
            sql = "SELECT * FROM " & GetType(T).Name & " WHERE "
            For Each sp As SelectParams In spArray
                If sp.key IsNot Nothing AndAlso sp.value IsNot Nothing Then
                    If sp.type Then
                        operation = "="
                    Else
                        operation = "!="
                    End If
                    sql = sql & sp.key & operation & "'" & sp.value & "'"
                    If Not sp.Equals(spArray(UBound(spArray))) Then
                        sql = sql & " AND "
                    End If
                End If
            Next
            Dim list As List(Of T) = myDB.ExecuteQuery(Of T)(sql).ToList()
            If list.Count > 0 Then
                Return list.First()
            Else
                Return Nothing
            End If
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return Nothing
        Finally
        End Try
    End Function

    ''' <summary>
    ''' 查詢(回傳第一筆 TOP 1)，查不到或錯誤則回傳NULL
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="spArray"></param>
    ''' <returns></returns>
    Public Function SelectTopOne(Of T As New)(ParamArray spArray As SelectParams()) As T
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            Dim operation As String
            sql = "SELECT TOP 1 * FROM " & GetType(T).Name & " WHERE "
            For Each sp As SelectParams In spArray
                If sp.key IsNot Nothing AndAlso sp.value IsNot Nothing Then
                    If sp.type Then
                        operation = "="
                    Else
                        operation = "!="
                    End If
                    sql = sql & sp.key & operation & "'" & sp.value & "'"
                    If Not sp.Equals(spArray(UBound(spArray))) Then
                        sql = sql & " AND "
                    End If
                End If
            Next
            Dim list As List(Of T) = myDB.ExecuteQuery(Of T)(sql).ToList()
            If list.Count() <> 0 Then
                Return list.First()
            Else
                Return Nothing
            End If
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return Nothing
            'myDB.Dispose();
        Finally
        End Try
    End Function


    ''' <summary>
    ''' 查詢 TOP 1 回傳一筆，查不到或錯誤則回傳NULL
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="key"></param>
    ''' <param name="value"></param>
    ''' <param name="type"></param>
    ''' <returns></returns>
    Public Function SelectTopOne(Of T As New)(key As String, value As String, type As Boolean) As T
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            Dim operation As String
            'myDB.Open();
            If key IsNot Nothing AndAlso value IsNot Nothing Then
                If type Then
                    operation = "="
                Else
                    operation = "!="
                End If
                sql = "SELECT TOP 1 * FROM " & GetType(T).Name & " WHERE " & key & operation & "'" & value & "'"
            Else
                sql = "SELECT TOP 1 * FROM " & GetType(T).Name
            End If
            Dim list As List(Of T) = myDB.ExecuteQuery(Of T)(sql).ToList()
            If list.Count() <> 0 Then
                Return list.First()
            Else
                Return New T
            End If
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return New T
            'myDB.Dispose();
        Finally
        End Try
    End Function

    Private Class dataCount
        Public Property count As Integer
    End Class

    ''' <summary>
    ''' 查詢筆數，如果錯誤則回傳-1
    ''' </summary>
    ''' <typeparam name="T">型態</typeparam>
    ''' <param name="key">查詢的欄位</param>
    ''' <param name="value">查詢的值</param>
    ''' <param name="type">相等(true)或不相等(false)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SelectCount(Of T As New)(key As String, value As String, type As Boolean) As Integer
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            Dim operation As String
            'myDB.Open();

            If key IsNot Nothing AndAlso value IsNot Nothing Then
                If type Then
                    operation = "="
                Else
                    operation = "!="
                End If
                sql = "SELECT COUNT(*) AS count FROM " & GetType(T).Name & " WHERE " & key & operation & "'" & value & "'"
            Else
                sql = "SELECT COUNT(*) AS count FROM " & GetType(T).Name
            End If
            Dim dc As dataCount = myDB.ExecuteQuery(Of dataCount)(sql).First()
            Return dc.count
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return -1
            'myDB.Dispose();
        Finally
        End Try
    End Function

    ''' <summary>
    ''' 查詢筆數(使用查詢物件)，如果錯誤則回傳-1
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="spArray">查詢物件</param>
    ''' <returns></returns>
    Public Function SelectCount(Of T As New)(ParamArray spArray As SelectParams()) As Integer
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            Dim operation As String
            sql = "SELECT COUNT(*) AS count FROM " & GetType(T).Name & " WHERE "
            For Each sp As SelectParams In spArray
                If sp.key IsNot Nothing AndAlso sp.value IsNot Nothing Then
                    If sp.type Then
                        operation = "="
                    Else
                        operation = "!="
                    End If
                    sql = sql & sp.key & operation & "'" & sp.value & "'"
                    If Not sp.Equals(spArray(UBound(spArray))) Then
                        sql = sql & " AND "
                    End If
                End If
            Next
            Dim dc As dataCount = myDB.ExecuteQuery(Of dataCount)(sql).First()
            Return dc.count
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return -1
            'myDB.Dispose();
        Finally
        End Try
    End Function

    ''' <summary>
    ''' 查詢Table的識別項號
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Public Function SelectIdentity(ByVal tableName As String) As String
        Try
            'identity_seed
            Dim myDBs As New SqlDataSource()
            myDBs.SelectCommand = "SELECT IDENT_SEED('" & tableName & "') AS identity_seed;"
            Dim dv As DataView = CType(myDBs.Select(DataSourceSelectArguments.Empty), DataView)
            Return dv(0).Item(0).ToString()
        Catch e As Exception
            Return e.ToString()
        Finally
        End Try
    End Function
#End Region

#Region "UPDATE"
    ''' <summary>
    ''' 更新某欄位值
    ''' </summary>
    ''' <typeparam name="T">物件</typeparam>
    ''' <param name="setId">欲更新的欄位名稱</param>
    ''' <param name="setVal">欲更新的職</param>
    ''' <param name="key">條件的欄位名稱</param>
    ''' <param name="value">條件的值</param>
    ''' <param name="type">條件為相等(true)或不相等(false)</param>
    ''' <returns></returns>
    Public Function Update(Of T)(setId As String, setVal As String, key As String, value As String, type As Boolean) As Boolean
        Dim myDB As New DataContext(ConnStr)
        Try
            Dim operation As String
            'myDB.Open();
            If key IsNot Nothing AndAlso value IsNot Nothing Then

                If type Then
                    operation = "="
                Else
                    operation = "!="
                End If
                sql = "UPDATE " & GetType(T).Name & " SET " & setId & "=" & "'" & setVal.Replace("'", "''") & "'" & " WHERE " & key & operation & "'" & value.Replace("'", "''") & "'"
            Else
                sql = "UPDATE " & GetType(T).Name & " SET " & setId & "=" & "'" & setVal.Replace("'", "''") & "'"
            End If
            myDB.ExecuteCommand(sql)
            Return True
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return False
            'myDB.Dispose();
        Finally
        End Try
    End Function

    Public Function update(Of T)(obj As T, ParamArray spArray As SelectParams()) As Boolean
        Dim myDB As New DataContext(connStr)
        Dim isSuccess As Boolean = False
        Dim trueValue As String = ""
        Dim tp As Type = obj.[GetType]()
        Dim properties As PropertyInfo() = tp.GetProperties()
        Dim keyIds As ArrayList = New ArrayList()
        For Each sp In spArray
            keyIds.Add(sp.key)
        Next
        Try
            sql = "UPDATE " & tp.Name & " SET "
            For i As Integer = 0 To properties.Length - 1
                If Not keyIds.Contains(properties(i).Name.ToString()) Then
                    trueValue = Convert.ToString(properties(i).GetValue(obj, Nothing)).ToString()
                    If (New Date).GetType.Equals(properties(i).PropertyType) Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    ElseIf GetType(Date?).Equals(properties(i).PropertyType) And trueValue <> "" Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    End If
                    sql += properties(i).Name & "=" & "'" & trueValue.Replace("'", "''") & "',"
                End If
            Next
            sql = sql.Substring(0, sql.Length - 1)
            If spArray.Length > 0 Then
                sql += " WHERE "
            End If
            Dim operation As String
            For Each sp As SelectParams In spArray
                If sp.key IsNot Nothing AndAlso sp.value IsNot Nothing Then
                    If sp.type Then
                        operation = "="
                    Else
                        operation = "!="
                    End If
                    sql += sp.key & operation & "'" & sp.value.Replace("'", "''") & "'"
                    If Not sp.Equals(spArray(UBound(spArray))) Then
                        sql += sql & " AND "
                    End If
                End If
            Next
            myDB.ExecuteCommand(sql)
            isSuccess = True
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            isSuccess = False
        Finally

        End Try
        Return isSuccess
    End Function

    Public Function UpdateStr(Of T As New)(setId As String, setVal As String, key As String, value As String, type As Boolean) As String
        Dim operation As String
        If key IsNot Nothing AndAlso value IsNot Nothing Then

            If type Then
                operation = "="
            Else
                operation = "!="
            End If
            sql = "UPDATE " & GetType(T).Name & " SET " & setId & "=" & "'" & setVal.Replace("'", "''") & "'" & " WHERE " & key & operation & "'" & value.Replace("'", "''") & "';"
        Else
            sql = "UPDATE " & GetType(T).Name & " SET " & setId & "=" & "'" & setVal.Replace("'", "''") & "';"
        End If
        Return sql
    End Function

    Public Function UpdateStr(Of T)(obj As T, key As String, val As String, same As Boolean) As String
        Dim type As Type = obj.[GetType]()
        Dim operation As String
        If same Then
            operation = "="
        Else
            operation = "!="
        End If
        Dim trueValue As String = ""
        Dim properties As PropertyInfo() = type.GetProperties()
        Try
            sql = "UPDATE " & type.Name & " SET "
            For i As Integer = 0 To properties.Length - 1
                If key = properties(i).Name Or properties(i).Name = "level_value1" Then
                    Continue For
                End If
                If Not key.Equals(properties(i).Name.ToString()) Then
                    sql += properties(i).Name
                    trueValue = Convert.ToString(properties(i).GetValue(obj, Nothing)).ToString()
                    If (New Date).GetType.Equals(properties(i).PropertyType) Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    ElseIf GetType(Date?).Equals(properties(i).PropertyType) And trueValue <> "" Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    End If
                    sql += "='" & trueValue.Replace("'", "''") & "'"
                    If i <> properties.Length - 1 Then
                        sql += ","
                    End If
                End If
            Next
            sql += " WHERE " & key & operation & "'" & val & "';"
            Return sql
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return e.ToString()
        Finally
        End Try
    End Function
#End Region

#Region "INSERT"
    ''' <summary>
    ''' 新增，回傳true 或 flase
    ''' </summary>
    ''' <typeparam name="T">型態</typeparam>
    ''' <param name="obj">欲新增之物件</param>
    ''' <param name="keyId">Id名稱，或需要忽略的欄位</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert(Of T)(obj As T, ParamArray keyId As String()) As Boolean
        Dim isSuccess As Boolean = False
        Dim sqlconn As New SqlConnection(connStr)
        Dim trueValue As String = ""
        mySQLDB.Connection = sqlconn
        Dim tp As Type = obj.[GetType]()
        Dim properties As PropertyInfo() = tp.GetProperties()
        Try
            sql = "INSERT INTO " & tp.Name & " ("
            Dim param As String = "", value As String = ""
            For i As Integer = 0 To properties.Length - 1
                If Not keyId.Contains(properties(i).Name.ToString()) Then
                    trueValue = Convert.ToString(properties(i).GetValue(obj, Nothing)).ToString()
                    If (New Date).GetType.Equals(properties(i).PropertyType) Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    ElseIf GetType(Date?).Equals(properties(i).PropertyType) And trueValue <> "" Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    End If
                    param += properties(i).Name

                    If properties(i).PropertyType.Equals("".GetType()) Then
                        value += "N"
                    End If

                    value += "'" & trueValue.Replace("'", "''") & "'"
                    param += ","
                    value += ","
                End If
            Next
            param = param.Substring(0, param.Length - 1)
            value = value.Substring(0, value.Length - 1)
            sql += param & ") VALUES (" & value & ") SELECT @@IDENTITY"
            mySQLDB.CommandText() = sql
            mySQLDB.Connection.Open()
            mySQLDB.ExecuteScalar()
            isSuccess = True
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            isSuccess = False
        Finally
            mySQLDB.Connection.Close()
        End Try
        Return isSuccess
    End Function

    ''' <summary>
    ''' 會回傳insert的Id，如果發生錯誤則回傳-1
    ''' </summary>
    ''' <typeparam name="T">型態</typeparam>
    ''' <param name="obj">欲新增之物件</param>
    ''' <param name="keyId">Id名稱，或需要忽略的欄位</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertReturnId(Of T)(obj As T, ParamArray keyId As String()) As String
        Dim returnId As String = "-1"
        Dim sqlconn As New SqlConnection(connStr)
        Dim trueValue As String = ""
        mySQLDB.Connection = sqlconn
        Dim tp As Type = obj.[GetType]()
        Dim properties As PropertyInfo() = tp.GetProperties()
        Try
            sql = "INSERT INTO " & tp.Name & " ("
            Dim param As String = "", value As String = ""
            For i As Integer = 0 To properties.Length - 1
                If Not keyId.Contains(properties(i).Name.ToString()) Then
                    trueValue = Convert.ToString(properties(i).GetValue(obj, Nothing)).ToString()
                    If (New Date).GetType.Equals(properties(i).PropertyType) Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    ElseIf GetType(Date?).Equals(properties(i).PropertyType) And trueValue <> "" Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    End If
                    param += properties(i).Name
                    If properties(i).PropertyType.Equals("".GetType()) Then
                        value += "N"
                    End If
                    value += "'" & trueValue.Replace("'", "''") & "'"
                    param += ","
                    value += ","
                End If
            Next
            param = param.Substring(0, param.Length - 1)
            value = value.Substring(0, value.Length - 1)
            sql += param & ") VALUES (" & value & ") SELECT @@IDENTITY"
            mySQLDB.CommandText() = sql
            mySQLDB.Connection.Open()
            returnId = mySQLDB.ExecuteScalar().ToString()
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            returnId = "-1"
        Finally
            mySQLDB.Connection.Close()
        End Try
        Return returnId
    End Function

    Public Function InsertStr(Of T)(obj As T, ParamArray keyId As String()) As String
        Dim type As Type = obj.[GetType]()
        Dim trueValue As String = ""
        Dim properties As PropertyInfo() = type.GetProperties()
        Try
            sql = "INSERT INTO " & type.Name & " ("
            Dim param As String = "", value As String = ""
            For i As Integer = 0 To properties.Length - 1
                If Not keyId.Contains(properties(i).Name.ToString()) Then
                    trueValue = Convert.ToString(properties(i).GetValue(obj, Nothing)).ToString()
                    If (New Date).GetType.Equals(properties(i).PropertyType) Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    ElseIf GetType(Date?).Equals(properties(i).PropertyType) And trueValue <> "" Then
                        trueValue = CDate(trueValue).ToString("yyyy-MM-dd HH:mm:ss")
                    End If
                    param += properties(i).Name
                    If properties(i).PropertyType.Equals("".GetType()) Then
                        value += "N"
                    End If
                    value += "'" & trueValue.Replace("'", "''") & "'"
                    param += ","
                    value += ","
                End If
            Next
            param = param.Substring(0, param.Length - 1)
            value = value.Substring(0, value.Length - 1)
            sql += param & ") VALUES (" & value & ");"
            Return sql
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return Nothing
        Finally
        End Try
    End Function

    ''' <summary>
    ''' 新增或修改
    ''' </summary>
    ''' <typeparam name="T">型態</typeparam>
    ''' <param name="obj">物件</param>
    ''' <param name="spArray">KEY值(識別項)，最好只有一組，若新增時會忽略此值，更新則由此值當條件</param>
    ''' <returns></returns>
    Public Function insertOrUpdate(Of T As New)(obj As T, ParamArray spArray As SelectParams()) As Boolean
        Dim keys(spArray.Length) As String
        Dim vals(spArray.Length) As String
        For i As Integer = 0 To spArray.Length - 1
            keys(i) = spArray(i).key
            vals(i) = spArray(i).value
        Next
        If spArray.Length = 0 Or vals.Contains("") Or keys.Contains("") Then
            'INSERT
            Return Insert(obj, keys)
        Else
            'UPDATE
            If IsNothing(SelectTopOne(Of T)(spArray)) Then
                Return Insert(obj, keys)
            Else
                Return update(obj, spArray)
            End If
        End If
    End Function
#End Region

#Region "DELETE"
    Public Function DeleteAll(Of T)() As Boolean
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            'myDB.Open();
            Dim deleteCount As Integer = myDB.ExecuteCommand("DELETE FROM " & GetType(T).Name)
            System.Diagnostics.Debug.WriteLine("Delete count : " & deleteCount)
            Return True
        Catch e As Exception
            System.Diagnostics.Debug.WriteLine("Delete Exception: " & Convert.ToString(e))
            Return False
            'myDB.Dispose();
        Finally
        End Try
    End Function

    Public Function DeleteAllStr(Of T)() As String
        Return "DELETE FROM " & GetType(T).Name & ";"
    End Function

    Public Function Delete(Of T)(key As String, value As String) As Boolean
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            sql = "DELETE FROM " & GetType(T).Name & " WHERE " & key & "='" & value & "'"
            myDB.ExecuteCommand(sql)
            Return True
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return False
        Finally
        End Try
    End Function

    Public Function DeleteMulti(Of T)(key As String, ParamArray value As String()) As Boolean
        Dim myDB As New System.Data.Linq.DataContext(connStr)
        Try
            sql = "DELETE FROM " & GetType(T).Name & " WHERE "
            For Each v As String In value
                sql = sql & key & "='" & v & "' OR "
            Next
            sql = sql.Substring(0, sql.Length - 4)
            myDB.ExecuteCommand(sql)
            Return True
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sql)
            Return False
        Finally
        End Try
    End Function

    Public Function DeleteStr(Of T)(key As String, value As String) As String
        Return "DELETE FROM " & GetType(T).Name & " WHERE " & key & "='" & value & "'"
    End Function

    Public Function DeleteMultiStr(Of T)(key As String, ParamArray value As String()) As String
        sql = "DELETE FROM " & GetType(T).Name & " WHERE "
        For Each v As String In value
            sql = sql & key & "='" & v & "' OR "
        Next
        sql = sql.Substring(0, sql.Length - 4)
        Return sql
    End Function
#End Region

    ''' <summary>
    ''' 通用，直接下指令，回傳true or false
    ''' </summary>
    ''' <param name="sqlcomm">SQL</param>
    ''' <returns>boolean</returns>
    Public Function sqlCommand(sqlcomm As String) As Boolean
        Dim isSuccess As Boolean = False
        Dim sqlconn As New SqlConnection(connStr)
        mySQLDB.Connection = sqlconn
        Try
            mySQLDB.CommandText() = sqlcomm
            mySQLDB.Connection.Open()
            mySQLDB.ExecuteScalar()
            isSuccess = True
        Catch e As Exception
            emsys.Debuger(e.ToString())
            emsys.Debuger("Error SQL = " & sqlcomm)
            isSuccess = False
        Finally
            mySQLDB.Connection.Close()
        End Try
        Return isSuccess
    End Function

    ''' <summary>
    ''' 通用，直接下指令，回傳true or false
    ''' </summary>
    ''' <param name="sqlcomm">SQL</param>
    ''' <returns>boolean</returns>
    Public Function sqlCommand(sqlcomm As String, ParamArray sqlParams As SelectParams()) As Boolean
        Dim isSuccess As Boolean = False
        Dim sqlconn As New SqlConnection(connStr)
        mySQLDB.Connection = sqlconn
        Try
            mySQLDB.CommandText() = sqlcomm
            mySQLDB.Connection.Open()
            mySQLDB.Parameters.Clear()
            For Each sqlParam As SelectParams In sqlParams
                mySQLDB.Parameters.AddWithValue(sqlParam.key, sqlParam.value)
            Next
            mySQLDB.ExecuteScalar()
            isSuccess = True
        Catch e As Exception
            emsys.Debuger(e.ToString())
            For Each sqlParam As SelectParams In sqlParams
                sqlcomm = sqlcomm.Replace(sqlParam.key, "'" & sqlParam.value & "'")
            Next
            emsys.Debuger("Error SQL = " & sqlcomm)
            isSuccess = False
        Finally
            mySQLDB.Connection.Close()
        End Try
        Return isSuccess
    End Function

    ''' <summary>
    ''' 查詢用SQL to DataTable
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <returns>回傳datatable</returns>
    ''' <remarks></remarks>
    Public Function SqlToDt(sql As String) As DataTable
        Dim sqlconn As New SqlConnection(connStr)
        Dim sqlcmd As New SqlCommand()
        Dim dt As DataTable = New DataTable()
        Try
            sqlcmd.Connection = sqlconn
            sqlcmd.Connection.Open()
            sqlcmd.CommandText = sql
            Dim sr As SqlDataReader = sqlcmd.ExecuteReader()
            dt.Load(sr)
            Return dt
        Catch ex As Exception
            emsys.Debuger(ex.ToString())
            emsys.Debuger("Error SQL = " & sql)
            isSuccess = False
        Finally
            sqlcmd.Connection.Close()
        End Try
        Return dt
    End Function

    ''' <summary>
    ''' 查詢用SQL to DataTable
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <returns>回傳datatable</returns>
    ''' <remarks></remarks>
    Public Function SqlToDt(sql As String, ParamArray sqlParams As SelectParams()) As DataTable
        Dim sqlconn As New SqlConnection(connStr)
        Dim sqlcmd As New SqlCommand()
        Dim dt As DataTable = New DataTable()
        Try
            sqlcmd.Connection = sqlconn
            sqlcmd.Connection.Open()
            sqlcmd.CommandText = sql
            For Each sqlParam As SelectParams In sqlParams
                sqlcmd.Parameters.AddWithValue(sqlParam.key, sqlParam.value)
            Next
            Dim sr As SqlDataReader = sqlcmd.ExecuteReader()
            dt.Load(sr)
            Return dt
        Catch ex As Exception
            emsys.Debuger(ex.ToString())
            For Each sqlParam As SelectParams In sqlParams
                sql = sql.Replace(sqlParam.key, "'" & sqlParam.value & "'")
            Next
            emsys.Debuger("Error SQL = " & sql)
        Finally
            sqlcmd.Connection.Close()
        End Try
        Return dt
    End Function

    ''' <summary>
    ''' 查詢物件(key=欄位值、value=數值、type=true:等於||false:不等於)
    ''' </summary>
    Public Class SelectParams
        Public key As String
        Public value As String
        Public type As Boolean
        Public Sub New()

        End Sub
        Public Sub New(vKey As String, vValue As String, Optional vType As Boolean = True)
            Me.key = vKey
            Me.value = vValue
            Me.type = vType
        End Sub
    End Class

End Class
