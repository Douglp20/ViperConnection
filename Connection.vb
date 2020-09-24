Public Class Connection
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
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
        Return apt
        con.Close()


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
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
        Return apt

        con.Close()

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
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
        Return apt
        con.Close()

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
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
        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

    Public Sub ExecuteProcessWithParameter(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal strParameter As String, ByVal strType As String, ByRef queryString As String)


        Dim cmd As New SqlClient.SqlCommand()
        Dim con As New SqlClient.SqlConnection(strConnectionString)
        con.Open()
        cmd.CommandText = strStoreProcedure
        cmd.Parameters.AddWithValue(strParameter, strType).Value = queryString
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = con
        cmd.ExecuteNonQuery()
        con.Close()

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
        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Function ExecuteProcessWithParametersReturnInteger(ByVal strConnectionString As String, ByVal strStoreProcedure As String, ByVal arrParameter As ArrayList, strParameterOutput As String, ByVal arrType As ArrayList, ByRef arrQueryString As ArrayList) As Integer

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
        Return ID
        con.Close()

    End Function
End Class
