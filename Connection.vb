Public Class Connection
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents VenomRegistry As New Douglas.Venom.Registry.VenomRegistry()
    Public Sub New()
    End Sub
    Private Sub msgProcessLog(strConnectionString As String, msg As String)

        Dim strStoreProcedure As String = "msg_proccesLog"
        Dim strParameter As String = "@log"
        Dim strType As String = SqlDbType.VarChar
        Dim queryString As String = msg


        ExecuteProcessMsgLog(strConnectionString, strStoreProcedure, strParameter, strType, queryString)

        Exit Sub
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Private Sub msgProcessLog(strConnectionString As String, msg As String, logData As String)

        Dim strStoreProcedure As String = "msg_proccesLog"
        Dim strParameter As String = "@log"
        Dim strType As String = SqlDbType.VarChar
        Dim queryString As String = msg & ":" & logData


        ExecuteProcessMsgLog(strConnectionString, strStoreProcedure, strParameter, strType, queryString)

        Exit Sub
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Private Sub msgProcessLog(strConnectionString As String, msg As String, logData As String, logParameter As String)
        Dim para As String = logParameter.Replace("@", String.Empty)
        Dim strStoreProcedure As String = "msg_proccesLog"
        Dim strParameter As String = "@log"
        Dim strType As String = SqlDbType.VarChar
        Dim queryString As String = para & ":" & logData
        queryString = msg & ":" & queryString

        ExecuteProcessMsgLog(strConnectionString, strStoreProcedure, strParameter, strType, queryString)

        Exit Sub
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Private Sub msgProcessLogs(strConnectionString As String, msg As String, logData As ArrayList, logParameter As ArrayList)
        On Error GoTo Err
        Dim strStoreProcedure As String = "msg_proccesLog"
        Dim strParameter As String = "@log"
        Dim strType As String = SqlDbType.Text
        Dim para As String = String.Empty
        Dim queryString As String = String.Empty

        For r As Integer = 0 To logParameter.Count - 1
            para = logParameter(r).Replace("@", String.Empty)
            If r = 0 Then
                queryString = para & "=" & logData(0).ToString
            Else
                queryString = queryString & ":" & para & "=" & logData(r).ToString
            End If
        Next
        queryString = msg & ":" & queryString

        ExecuteProcessMsgLog(strConnectionString, strStoreProcedure, strParameter, strType, queryString)

        Exit Sub
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub ExecuteProcessMsgLog(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal strParameter As String, ByVal strType As String, ByRef queryString As String)
        On Error GoTo Err
        Dim UserLogin As String
        UserLogin = VenomRegistry.GetSetting("LoginInfo", "UserName", "")

        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        If String.IsNullOrEmpty(UserLogin) Then
            cmd.Parameters.AddWithValue(strParameter, strType).Value = queryString
        Else
            cmd.Parameters.AddWithValue(strParameter, strType).Value = queryString + "::" + UserLogin
        End If
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = con
        cmd.ExecuteNonQuery()
        con.Close()



        Exit Sub
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."

        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

    Public Function getSqlDataAdapterWithParameter(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal strParameter As String, ByVal strType As String, ByRef queryString As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue(strParameter, strType).Value = queryString
        cmd.Connection = con

        Dim apt As New SqlClient.SqlDataAdapter(cmd)

        con.Close()

        msgProcessLog(strConnectionString, strStoreProcedure, queryString, strParameter)


        Return apt

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        msgProcessLog(strConnectionString, rtn, queryString, queryString)
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getSqlDataAdapterWithParameters(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal arrParameter As ArrayList, ByVal arrType As ArrayList, ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim arrCount As Integer = arrParameter.Count

        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        cmd.CommandType = CommandType.StoredProcedure
        For r As Integer = 0 To arrCount - 1
            cmd.Parameters.AddWithValue(arrParameter(r), arrType(r)).Value = arrQueryString(r)
        Next
        cmd.Connection = con

        Dim apt As New SqlClient.SqlDataAdapter(cmd)


        con.Close()
        msgProcessLogs(strConnectionString, strStoreProcedure, arrQueryString, arrParameter)

        Return apt
        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        msgProcessLogs(strConnectionString, rtn, arrQueryString, arrQueryString)
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getSqlDataAdapter(ByVal strConnectionString As String, ByVal strStoreProcedure As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = con

        Dim apt As New SqlClient.SqlDataAdapter(cmd)

        con.Close()
        msgProcessLog(strConnectionString, strStoreProcedure)


        Return apt


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        msgProcessLog(strConnectionString, rtn)
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub ExecuteProcess(ByVal strConnectionString As String, ByVal strStoreProcedure As String)

        On Error GoTo Err

        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = con
        cmd.ExecuteNonQuery()
        con.Close()

        msgProcessLog(strConnectionString, strStoreProcedure)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        msgProcessLog(strConnectionString, rtn)
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

    Public Sub ExecuteProcessWithParameter(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal strParameter As String, ByVal strType As String, ByRef queryString As String)

        On Error GoTo Err
        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        cmd.Parameters.AddWithValue(strParameter, strType).Value = queryString
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = con
        cmd.ExecuteNonQuery()
        con.Close()

        msgProcessLog(strConnectionString, strStoreProcedure, queryString, strParameter)

        Exit Sub
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        msgProcessLog(strConnectionString, rtn, queryString, strParameter)
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub ExecuteProcessWithParameter(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal strParameter As String, ByVal strType As String, ByRef queryString As String, picture As Byte(), pictureParameter As String)

        On Error GoTo Err
        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        cmd.Parameters.AddWithValue(strParameter, strType).Value = queryString
        Dim p As New SqlClient.SqlParameter(pictureParameter, SqlDbType.Image)
        p.Value = picture
        cmd.Parameters.Add(p)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = con
        cmd.ExecuteNonQuery()
        con.Close()

        msgProcessLog(strConnectionString, strStoreProcedure, queryString, strParameter)

        Exit Sub
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        msgProcessLog(strConnectionString, rtn, queryString, strParameter)
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

    Public Sub ExecuteProcessWithParameters(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal arrParameter As ArrayList, ByVal arrType As ArrayList, ByRef arrQueryString As ArrayList)

        On Error GoTo Err

        Dim arrCount As Integer = arrParameter.Count
        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        For r As Integer = 0 To arrCount - 1
            cmd.Parameters.AddWithValue(arrParameter(r), arrType(r)).Value = arrQueryString(r)
        Next
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = con
        cmd.ExecuteNonQuery()
        con.Close()

        msgProcessLogs(strConnectionString, strStoreProcedure, arrQueryString, arrParameter)

        Exit Sub
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        msgProcessLogs(strConnectionString, rtn, arrQueryString, arrParameter)
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub ExecuteProcessImageWithParameters(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal arrParameter As ArrayList, ByVal arrType As ArrayList, ByRef arrQueryString As ArrayList, picture As Byte(), pictureParameter As String)

        On Error GoTo Err

        Dim arrCount As Integer = arrParameter.Count
        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        For r As Integer = 0 To arrCount - 1
            cmd.Parameters.AddWithValue(arrParameter(r), arrType(r)).Value = arrQueryString(r)
        Next
        Dim p As New SqlClient.SqlParameter(pictureParameter, SqlDbType.Image)
        p.Value = picture
        cmd.Parameters.Add(p)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = con
        cmd.ExecuteNonQuery()
        con.Close()

        msgProcessLogs(strConnectionString, strStoreProcedure, arrQueryString, arrParameter)

        Exit Sub
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        msgProcessLogs(strConnectionString, rtn, arrQueryString, arrParameter)
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Function ExecuteProcessWithParametersReturnInteger(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal arrParameter As ArrayList, strParameterOutput As String, ByVal arrType As ArrayList, ByRef arrQueryString As ArrayList) As Integer

        On Error GoTo Err
        Dim arrCount As Integer = arrParameter.Count
        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = con

        For r As Integer = 0 To arrCount - 1
            Dim inParam As New SqlClient.SqlParameter()
            inParam.SqlDbType = arrType(r)
            inParam.ParameterName = arrParameter(r)
            inParam.Direction = ParameterDirection.Input
            inParam.Value = arrQueryString(r)
            cmd.Parameters.Add(inParam)
        Next

        Dim outParam As New SqlClient.SqlParameter()
        outParam.SqlDbType = SqlDbType.Int
        outParam.ParameterName = strParameterOutput
        outParam.Direction = ParameterDirection.Output
        outParam.Value = 0
        cmd.Parameters.Add(outParam)

        cmd.ExecuteNonQuery()

        Dim ID As Integer = CLng(cmd.Parameters(strParameterOutput).Value.ToString)
        con.Close()

        msgProcessLogs(strConnectionString, strStoreProcedure, arrQueryString, arrParameter)

        Return ID

        Exit Function
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + Err.Description + "."
        msgProcessLogs(strConnectionString, rtn, arrQueryString, arrParameter)
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
End Class
