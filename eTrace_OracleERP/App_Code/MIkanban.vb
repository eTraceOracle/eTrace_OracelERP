Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Public Class MIkanban
    Inherits PublicFunction
    'Private ConnStr As String = "server=cnaecfuyapp30;database=eTrace;UID=etrace;PWD=etraceproduction@2012"  '  ConfigurationSettings.AppSettings.Item("eTraceDBConnString")      
    Private ConnStr As String = ConfigurationSettings.AppSettings.Item("eTraceDBConnString")
    Public Function getMatlByDJModel(ByVal DJ As String, ByVal PartNo As String) As DataSet
        Dim DS As DataSet = New DataSet()
        Dim strSql As StringBuilder = New StringBuilder()
        strSql.Append("     SELECT  M.CLID,M.Orgcode,M.SLOC,M.StorageBin,M.Manufacturer,M.ManufacturerPN,M.VendorID,")
        strSql.Append("     M.VendorName,M.RTLot RT,M.MaterialNo,M.MaterialDesc,ISNull(M.DateCode,'') as DateCode,ISNull(M.LotNo,'') as LotNo,POQ.CLIDQty,0 as KBQty")
        strSql.Append("       FROM ")
        strSql.Append("      ( SELECT A.CLID,SUM(B.CLIDQty) AS CLIDQty ")
        strSql.Append("          FROM  eTrace.dbo.T_CLMaster A INNER JOIN eTrace.dbo.T_PO_CLID B ")
        strSql.Append("            ON A.CLID = B.CLID ")
        strSql.Append("            WHERE   B.PO='" + FunStr(DJ) + "' AND ")
        strSql.Append("                    A.MaterialNo='" + FunStr(PartNo) + "'")
        strSql.Append("            GROUP BY A.CLID ")
        strSql.Append("        ) POQ ")
        strSql.Append("     LEFT JOIN eTrace.dbo.T_CLMaster M ")
        strSql.Append("       ON POQ.CLID=M.CLID ")

        Dim connection As SqlConnection = New SqlConnection(ConnStr)
        Try
            connection.Open()
            Dim command As SqlDataAdapter = New SqlDataAdapter(strSql.ToString, connection)
            command.Fill(DS, "Matl")
        Catch Ex As System.Data.SqlClient.SqlException
            Return Nothing
        End Try

        Return DS
    End Function

    Public Function GetMatlByCLIDDJ(ByVal CLID As String, ByVal DJ As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Try
                sql = String.Format("exec sp_121CLIDIssueData '{0}','{1}'", CLID, DJ)
                Return da.ExecuteDataSet(sql)
            Catch ex As Exception
                ErrorLogging("OTO-GetMatlByCLIDDJ", "", "CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function getMatlCLIDbyDJ(ByVal DJ As String) As DataSet
        Dim strSql As StringBuilder = New StringBuilder()
        strSql.Append("SELECT distinct A.MaterialNo,A.CLID FROM  eTrace.dbo.T_CLMaster A  ")
        strSql.Append("  INNER JOIN eTrace.dbo.T_PO_CLID B ")
        strSql.Append("   ON A.CLID = B.CLID where B.PO='" + FunStr(DJ) + "' order by A.CLID ")
        Dim ds As DataSet = New DataSet
        ' Dim ConnStr = "server=cnaecfuyapp20;database=eTrace;UID=etrace;PWD=etraceproduction@2010"
        Dim connection As SqlConnection = New SqlConnection(ConnStr)
        Try
            Dim command As SqlDataAdapter = New SqlDataAdapter(strSql.ToString, connection)
            command.Fill(ds, "MatlCLID")
        Catch ex As Exception
            Return Nothing
        End Try

        Return ds
    End Function
    Private Function FunStr(ByVal str As String) As String
        str = str.Replace("&", "&amp;")
        str = str.Replace("<", "&lt;")
        str = str.Replace(">", "&gt")
        str = str.Replace("'", "''")
        str = str.Replace("*", "")
        str = str.Replace("\n", "<br/>")
        str = str.Replace("\r\n", "<br/>")
        str = str.Replace("select", "")
        str = str.Replace("insert", "")
        str = str.Replace("update", "")
        str = str.Replace("delete", "")
        str = str.Replace("create", "")
        str = str.Replace("drop", "")
        str = str.Replace("delcare", "")
        str = str.Replace("   ", "&nbsp;")
        str = str.Trim()
        'If str.Trim().ToString() = "" Then
        '    str = "æ— "
        'End If
        Return str
    End Function

    Public Function PrintLabels(ByVal Printer As String, ByVal labelFile As String, ByVal strContent As String, Optional ByVal NoOfLabels As Integer = 1) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Try
                sql = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", labelFile, Printer, SQLString(strContent))
                da.ExecuteScalar(sql)
                PrintLabels = True
            Catch ex As Exception
                ErrorLogging("SFC-PrintLabels", "", "labelFile: " & labelFile & ", " & ex.Message & ex.Source, "E")
                PrintLabels = ex.Message & ex.Source
            End Try
        End Using
    End Function

    Public Function SaveEMCResult(ByVal sn As String, ByVal processname As String, ByVal result As String, ByVal userid As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Try
                sql = String.Format("exec sp_SaveEMCResult '{0}','{1}','{2}',{3}", sn, processname, result, userid)
                da.ExecuteScalar(sql)
                SaveEMCResult = "Pass"
            Catch ex As Exception
                ErrorLogging("OTO-SaveEMCResult", userid, ex.Message & ex.Source, "E")
                SaveEMCResult = "Fail"
            End Try
        End Using
    End Function
End Class
