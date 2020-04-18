Imports System
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Security.Cryptography

Public Structure eTraceRole
    Public RoleID As String
    Public RoleDesc As String
    Public Status As String
    Public ChangedBy As String
    Public ChangedOn As String
    Public Remarks As String
End Structure

Public Structure RoleDetail
    Public RoleID As String
    Public RoleDesc As String
    Public Status As String
    Public ChangedBy As String
    Public ChangedOn As String
    Public Remarks As String
End Structure

Public Structure UserDetail
    Public UserID As String
    Public Name As String
    Public Division As String
    Public Dept As String
    Public EmpID As String
    Public Location As String
    Public Phone As String
    Public UserType As String
    Public Status As String
    Public User As String
    Public LastLogon As String
    Public Remarks As String
    Public ResetPWD As Boolean
End Structure

Public Class UserManagement
    Inherits PublicFunction

    Private key() As Byte = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24}
    Private iv() As Byte = {65, 110, 68, 26, 69, 178, 200, 219}

    'Encrypt the user data as byte before saving into SQL Server 2000
    Private Function Encrypt(ByVal plainText As String) As Byte()
        'Decalre UTF8Encoding object so we may use the GetByte method to transform
        'the plainText into Byte array
        Dim utf8encoder As UTF8Encoding = New UTF8Encoding
        Dim inputInBytes() As Byte = utf8encoder.GetBytes(plainText)
        'Create a new TripleDES service provider
        Dim tdesProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        'The ICryptTransform interface uses the TripleDes crypt provider along with
        'encryption key and init vector information
        Dim cryptoTransform As ICryptoTransform = tdesProvider.CreateEncryptor(Me.key, Me.iv)
        'All cryptographic functions need a stream to output the encrypted information.
        'Here we declare a memory stream for this purpose.
        Dim encryptedStream As MemoryStream = New MemoryStream
        Dim cryptStream As CryptoStream = New CryptoStream(encryptedStream, cryptoTransform, CryptoStreamMode.Write)
        'Write the encrypted information to the stream. Flush the information0
        'when done to ensure everything is out of the buffer.
        cryptStream.Write(inputInBytes, 0, inputInBytes.Length)
        cryptStream.FlushFinalBlock()
        encryptedStream.Position = 0
        'Read the stream back into a Byte array and return it to the calling method.
        Dim result(encryptedStream.Length - 1) As Byte
        encryptedStream.Read(result, 0, encryptedStream.Length)
        cryptStream.Close()
        Return result
    End Function
    'Decrypt the data from SQL Server 2000 before using it as string
    Private Function Decrypt(ByVal inputInBytes() As Byte) As String
        'UFTEncoding is used to transform the decrypted Byte Array information back into a string
        Dim utf8encoder As UTF8Encoding = New UTF8Encoding
        Dim tdesProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        'As before we must provide the encryption/decryption key along with the init vector
        Dim cryptoTransform As ICryptoTransform = tdesProvider.CreateDecryptor(Me.key, Me.iv)
        'Provider a memory stream to decrypt information into
        Dim decryptedStream As MemoryStream = New MemoryStream
        Dim cryptStream As CryptoStream = New CryptoStream(decryptedStream, cryptoTransform, CryptoStreamMode.Write)
        Dim i, j As Integer
        j = inputInBytes.Length
        For i = 0 To inputInBytes.Length - 1
            If inputInBytes(i).ToString = "0" Then
                j = j - 1
            End If
        Next
        cryptStream.Write(inputInBytes, 0, j)
        cryptStream.FlushFinalBlock()
        decryptedStream.Position = 0
        'Read the memory stream and convert it back into a string
        Dim result(decryptedStream.Length - 1) As Byte
        decryptedStream.Read(result, 0, decryptedStream.Length)
        cryptStream.Close()
        Dim myutf As UTF8Encoding = New UTF8Encoding
        Return myutf.GetString(result)
    End Function

    Public Function ReadUserData(ByVal UserID As String, ByVal User As String) As DataSet   'read with password
        Using da As DataAccess = GetDataAccess()
            ReadUserData = New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_Usr_UserData_Read  '{0}' ", UserID)   'Read Role Data

                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"dtUser", "dtPwd", "dtUserRole"}
                ReadUserData = da.ExecuteDataSet(sql, tables)
            Catch ex As Exception
                ErrorLogging("UserManagement-ReadUserData", User, ex.Message & ex.Source, "E")
            End Try
        End Using

    End Function

    Public Function ReadRoleData(ByVal User As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            ReadRoleData = New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_Usr_RoleData_Read ")   'Read Role Data

                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"dtRole", "dtTran", "dtRoleItem"}
                ReadRoleData = da.ExecuteDataSet(sql, tables)
            Catch ex As Exception
                ErrorLogging("UserManagement-ReadRoleData", User, ex.Message & ex.Source, "E")
            End Try
        End Using

    End Function

    Public Function PostUserData(ByVal UserData As UserDetail, ByVal dsUser As DataSet, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try

                Dim dsItem As New DataSet
                dsItem.DataSetName = "Items"
                dsItem.Merge(dsUser.Tables("dtUserRole"))

                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_Usr_UserData_Post '{0}','{1}', N'{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}', N'{11}','{12}','{13}' ", _
                           UserData.UserID, LoginData.User, UserData.Name, UserData.Division, UserData.Dept, UserData.EmpID, UserData.Location, UserData.Phone, UserData.UserType, UserData.Status, UserData.LastLogon, UserData.Remarks, UserData.ResetPWD, DStoXML(dsItem))
                PostUserData = da.ExecuteScalar(Sqlstr).ToString

            Catch ex As Exception
                ErrorLogging("UserManagement-PostUserData", LoginData.User, ex.Message & ex.Source, "E")
                Return "Error"
            End Try
        End Using

    End Function

    Public Function PostRoleData(ByVal RoleData As RoleDetail, ByVal dsTrans As DataSet, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr As String
                dsTrans.DataSetName = "Items"

                Sqlstr = String.Format("exec sp_Usr_RoleData_Post '{0}','{1}', N'{2}','{3}', N'{4}','{5}' ", _
                          RoleData.RoleID, LoginData.User, RoleData.RoleDesc, RoleData.Status, RoleData.Remarks, DStoXML(dsTrans))
                PostRoleData = da.ExecuteScalar(Sqlstr).ToString
            Catch ex As Exception
                ErrorLogging("UserManagement-PostRoleData", LoginData.User, ex.Message & ex.Source, "E")
                Return "Error"
            End Try
        End Using

    End Function

    Public Function PostUserData2(ByVal UserData As UserDetail, ByVal dsUserRole As DataSet, ByVal LoginData As ERPLogin) As String

        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim PostUserData As String = ""

        Try
            Dim EncPwd() As Byte = Encrypt("WELCOME1")

            Dim ResetPWD As Integer = 0
            If UserData.ResetPWD = True Then ResetPWD = 1

            Dim dsItem As New DataSet
            dsItem.DataSetName = "Items"
            dsItem.Merge(dsUserRole.Tables("dtUserRole"))

            myConn.Open()
            myCommand = New SqlClient.SqlCommand("sp_Usr_UserData_Post1", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.CommandTimeout = TimeOut_M5

            myCommand.Parameters.Add("@UserID", SqlDbType.VarChar, 50).Value = UserData.UserID
            myCommand.Parameters.Add("@User", SqlDbType.VarChar, 50).Value = LoginData.User
            myCommand.Parameters.Add("@Name", SqlDbType.VarChar, 50).Value = UserData.Name
            myCommand.Parameters.Add("@Division", SqlDbType.VarChar, 20).Value = UserData.Division
            myCommand.Parameters.Add("@Dept", SqlDbType.VarChar, 50).Value = UserData.Dept
            myCommand.Parameters.Add("@EmpID", SqlDbType.VarChar, 20).Value = UserData.EmpID
            myCommand.Parameters.Add("@Location", SqlDbType.VarChar, 100).Value = UserData.Location
            myCommand.Parameters.Add("@Phone", SqlDbType.VarChar, 50).Value = UserData.Phone
            myCommand.Parameters.Add("@UserType", SqlDbType.VarChar, 50).Value = UserData.UserType
            myCommand.Parameters.Add("@Status", SqlDbType.VarChar, 20).Value = UserData.Status
            myCommand.Parameters.Add("@LastLogon", SqlDbType.VarChar, 25).Value = UserData.LastLogon
            myCommand.Parameters.Add("@Remarks", SqlDbType.VarChar, 200).Value = UserData.Remarks
            myCommand.Parameters.Add("@ResetPWD", SqlDbType.Bit).Value = ResetPWD     'UserData.ResetPWD
            myCommand.Parameters.Add("@Password", SqlDbType.VarBinary, 50).Value = EncPwd
            myCommand.Parameters.Add("@strXml", SqlDbType.Xml, 8000).Value = DStoXML(dsItem)
            myCommand.Parameters.Add("@UserPost", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            ra = myCommand.ExecuteNonQuery()
            PostUserData = myCommand.Parameters("@UserPost").Value.ToString
            myConn.Close()
            Return PostUserData

        Catch ex As Exception
            ErrorLogging("UserManagement-PostUserData", LoginData.User, ex.Message & ex.Source, "E")
            Return "Error"
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function PostRoleData1(ByVal RoleData As RoleDetail, ByVal dsTrans As DataSet, ByVal LoginData As ERPLogin) As String

        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        Dim i, j, k As Integer
        Dim strSQL As String

        Try
            myConn.Open()
            'Dispose T_Transaction
            For j = 0 To dsTrans.Tables("dtTran").Rows.Count - 1
                If Not dsTrans.Tables("dtTran").Rows(j)("TranID") Is DBNull.Value Then
                    strSQL = String.Format("Select count(*) from T_RoleItem with (nolock) where RoleID='{0}' And TranID='{1}'", RoleData.RoleID, dsTrans.Tables("dtTran").Rows(j)("TranID"))
                    myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                    k = Convert.ToInt32(myCommand.ExecuteScalar)
                    If k = 0 Then
                        strSQL = String.Format("insert into T_RoleItem values('{0}','{1}') ", RoleData.RoleID, dsTrans.Tables("dtTran").Rows(j)("TranID"))
                        myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                        myCommand.ExecuteNonQuery()
                    End If
                    k = 0
                End If
            Next
            'Dispose T_RoleItem
            Dim DR() As DataRow
            Dim dtTept As DataTable = New DataTable
            Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("Select * from T_RoleItem with (nolock) where RoleID='{0}' ", RoleData.RoleID), myConn)
            da.Fill(dtTept)
            For i = 0 To dtTept.Rows.Count - 1
                DR = Nothing
                DR = dsTrans.Tables(0).Select(" TranID = '" & dtTept.Rows(i)("TranID") & "'")
                If DR.Length = 0 Then
                    strSQL = String.Format("delete from T_RoleItem where RoleID='{0}' And TranID='{1}'", dtTept.Rows(i)("RoleID"), dtTept.Rows(i)("TranID"))
                    myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                    myCommand.ExecuteNonQuery()
                End If
            Next
            'Dispose T_Role
            strSQL = String.Format("Select count(*) from T_Role with (nolock) where RoleID='{0}'", RoleData.RoleID)
            myCommand = New SqlClient.SqlCommand(strSQL, myConn)
            i = Convert.ToInt32(myCommand.ExecuteScalar)
            If i > 0 Then
                strSQL = String.Format("update T_Role set RoleDesc='{0}',Status='{1}',ChangedBy='{2}',ChangedOn=getDate(),Remarks=N'{3}' Where RoleID='{4}'", _
                              RoleData.RoleDesc, RoleData.Status, LoginData.User, RoleData.Remarks, RoleData.RoleID)
                myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                myCommand.ExecuteNonQuery()
                myConn.Close()
                Return "Update"
            Else
                strSQL = String.Format("insert into T_Role(RoleID,RoleDesc,Status,ChangedBy,ChangedOn,Remarks) " _
                & "values('{0}','{1}','{2}','{3}',getdate(),N'{4}')", _
                RoleData.RoleID, RoleData.RoleDesc, RoleData.Status, LoginData.User, RoleData.Remarks)
                myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                myCommand.ExecuteNonQuery()
                myConn.Close()
                Return "Insert"
            End If
        Catch ex As Exception
            ErrorLogging("UserManagement-PostRoleData", LoginData.User, ex.Message & ex.Source, "E")
            Return "Error"
        Finally
            If myConn.State = ConnectionState.Open Then myConn.Close()
        End Try
    End Function

    Public Function PostUserData1(ByVal UserData As UserDetail, ByVal dsUserRole As DataSet, ByVal LoginData As ERPLogin) As String

        Dim i, j, k As Integer
        Dim strSQL As String
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        Try
            myConn.Open()
            'Dispose insert and delete T_UserRole
            For j = 0 To dsUserRole.Tables("dtUserRole").Rows.Count - 1
                If Not dsUserRole.Tables("dtUserRole").Rows(j)("RoleID") Is DBNull.Value Then
                    strSQL = String.Format("Select count(*) from T_UserRole with (nolock) where UserID='{0}' And RoleID='{1}'", UserData.UserID, dsUserRole.Tables("dtUserRole").Rows(j)("RoleID"))
                    myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                    k = Convert.ToInt32(myCommand.ExecuteScalar)
                    If k = 0 Then
                        strSQL = String.Format("insert into T_UserRole values('{0}','{1}','{2}','{3}') ", UserData.UserID, dsUserRole.Tables("dtUserRole").Rows(j)("RoleID"), dsUserRole.Tables("dtUserRole").Rows(j)("ValidFrom"), dsUserRole.Tables("dtUserRole").Rows(j)("ValidTo"))
                        myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                        myCommand.ExecuteNonQuery()
                    End If
                    k = 0
                End If
            Next

            Dim DR() As DataRow
            Dim dtTept As DataTable = New DataTable
            Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("Select * from T_UserRole with (nolock) where UserID='{0}'", UserData.UserID), myConn)
            da.Fill(dtTept)

            For i = 0 To dtTept.Rows.Count - 1
                DR = Nothing
                DR = dsUserRole.Tables("dtUserRole").Select(" RoleID = '" & dtTept.Rows(i)("RoleID") & "'")
                If DR.Length = 0 Then
                    strSQL = String.Format("delete from T_UserRole where UserID='{0}' And RoleID='{1}'", dtTept.Rows(i)("UserID"), dtTept.Rows(i)("RoleID"))
                    myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                    myCommand.ExecuteNonQuery()
                End If
            Next

            'Dispose insert into T_UserPWD
            Dim ra As Integer
            Dim PwdCommand As SqlClient.SqlCommand
            If UserData.ResetPWD = True Then
                Dim EncPwd() As Byte = Encrypt("WELCOME1")
                PwdCommand = New SqlClient.SqlCommand("INSERT INTO T_UserPWD(UserID,Password,ValidFrom,ValidTo) values (@UserID,@Password,getdate(),getdate()+90)", myConn)
                PwdCommand.Parameters.Add("@UserID", SqlDbType.Char, 50, "UserID")
                PwdCommand.Parameters.Add("@Password", SqlDbType.VarBinary, 50, "Password")
                PwdCommand.Parameters("@UserID").Value = UserData.UserID
                PwdCommand.Parameters("@Password").Value = EncPwd
                PwdCommand.CommandType = CommandType.Text
                ra = PwdCommand.ExecuteNonQuery()
            End If

            'Dispose update and insert T_Users
            strSQL = String.Format("Select count(*) from T_Users with (nolock) where UserID='{0}'", UserData.UserID)
            myCommand = New SqlClient.SqlCommand(strSQL, myConn)
            i = Convert.ToInt32(myCommand.ExecuteScalar)
            If i > 0 Then
                strSQL = String.Format("update T_Users set Name=N'{0}',Division='{1}',Dept='{2}',EmpID='{3}'," _
              & " Location='{4}',Phone='{5}',UserType='{6}',Status='{7}',CreatedOn=getDate(),CreatedBy='{8}',ChangedOn=getDate(),ChangedBy='{9}',LastLogon='{10}',Remarks=N'{11}' Where UserID='{12}'", _
              UserData.Name, UserData.Division, UserData.Dept, UserData.EmpID, _
              UserData.Location, UserData.Phone, UserData.UserType, UserData.Status, LoginData.User, LoginData.User, UserData.LastLogon, UserData.Remarks, UserData.UserID)
                myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                myCommand.ExecuteNonQuery()
                myConn.Close()
                Return "Update"
            Else
                strSQL = String.Format("insert into T_Users(UserID,Name,Division,Dept,EmpID,Location,Phone,UserType,Status,CreatedOn,CreatedBy,ChangedOn,ChangedBy,LastLogon,Remarks) " _
                & "values('{0}',N'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',getdate(),'{9}',getdate(),'{10}','{11}','{12}')", _
                UserData.UserID, UserData.Name, UserData.Division, UserData.Dept, UserData.EmpID, UserData.Location, UserData.Phone, UserData.UserType, UserData.Status, LoginData.User, LoginData.User, UserData.LastLogon, UserData.Remarks)
                myCommand = New SqlClient.SqlCommand(strSQL, myConn)
                myCommand.ExecuteNonQuery()
                myConn.Close()
                Return "Insert"
            End If
        Catch ex As Exception
            ErrorLogging("UserManagement-PostRoleData", LoginData.User, ex.Message & ex.Source, "E")
            Return "Error"
        Finally
            If myConn.State = ConnectionState.Open Then myConn.Close()
        End Try
    End Function

End Class
