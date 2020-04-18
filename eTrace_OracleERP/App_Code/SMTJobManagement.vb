Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data.Entity
Imports System.Data.Entity.Core.Objects
Imports System.Data.Entity.Infrastructure
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Web
Imports System.Transactions
Imports System.Configuration
Imports SMTJobManagementModel

Public Class SMTJobManagementRepository
    Inherits EntityBaseRepository

    Public Function GetDJInfo(sql As String, entityName As String, ByRef msg As String, xmlParameters As String) As String
        Return MyBase.SelectMutipleInString(Of T_DJInfo)(sql, entityName, msg, xmlParameters, False)
    End Function

    Public Function GetProductStructure(sql As String, entityName As String, ByRef msg As String, xmlParameters As String) As String
        Return MyBase.SelectMutipleInString(Of T_ProductStructure)(sql, entityName, msg, xmlParameters, False)
    End Function

    Public Function NewSeqNO(ByRef msg As String) As String
        Dim jobID As String = String.Empty
        Try
            Using db As New Entities(MyBase.ConnStr)
                Using ts As New TransactionScope(TransactionScopeOption.Required, New TransactionOptions() With {.IsolationLevel = IsolationLevel.ReadCommitted})
                    Try
                        Dim o As T_SEQNo = db.SEQNo.AsQueryable().[Single](Function(j) j.TYPE = Parameters.KeyName.SeqNO)
                        Dim lastNO As System.Nullable(Of Integer) = o.LastNumber + 1
                        Dim s1 As String = lastNO.ToString()
                        Dim len1 As Integer = 6 - s1.Length
                        While len1 > 0
                            len1 -= 1
                            s1 = Convert.ToString("0") & s1
                        End While
                        o.LastNumber = lastNO
                        jobID = o.Prefix + s1
                        db.SaveChanges()
                    Catch ex As Exception
                        Throw ex
                    Finally
                        ts.Complete()
                    End Try
                End Using
            End Using
        Catch ex As Exception
            msg = ex.Message
        End Try
        Return jobID
    End Function

    Public Function GetJobHeader(sql As String, entityName As String, ByRef msg As String) As String
        Return MyBase.SelectMutipleInString(Of T_SFJobHeader)(sql, entityName, msg, Nothing, False)
    End Function

    Public Function GetJobHeader2(sql As String, entityName As String, ByRef msg As String, xmlParameters As String) As String
        Return MyBase.SelectSingleInString(Of T_SFJobHeader)(sql, entityName, msg, xmlParameters, True)
    End Function

    Public Sub InsertJob(xmlObj As String, ByRef msg As String)
        MyBase.InsertObject(Of T_SFJobHeader)(xmlObj, msg)
    End Sub

    Public Sub UpdateJob(xmlObj As String, clearFirst As Boolean, ByRef msg As String)
        Dim original As T_SFJobHeader = xmlObj.DeSerializerObject(Of T_SFJobHeader)()
        MyBase.UpdateObject(Of T_SFJobHeader)(original, msg, True, clearFirst, New Object() {original.JobID})
    End Sub

    Public Sub UpdateJobStatus(xmlObj As String, ByRef msg As String)
        Dim original As T_SFJobHeader = xmlObj.DeSerializerObject(Of T_SFJobHeader)()
        MyBase.UpdateObject(Of T_SFJobHeader)(original, msg, New Object() {original.JobID})
    End Sub

    Public Function GetProdLineByLineType(sql As String, entityName As String, ByRef msg As String, xmlParameters As String) As String
        Return MyBase.SelectMutipleInString(Of T_SFProdLine)(sql, entityName, msg, xmlParameters, False)
    End Function

    Public Function GetProdLine(sql As String, entityName As String, ByRef msg As String) As String
        Return MyBase.SelectMutipleInString(Of T_SFProdLine_LD)(sql, entityName, msg, Nothing, False)
    End Function
    Public Function GetSMTMachines(sql As String, entityName As String, ByRef msg As String, xmlParameters As String) As String
        Return MyBase.SelectMutipleInString(Of T_SMEquipment)(sql, entityName, msg, xmlParameters, False)
    End Function
    Public Function GetMSLMachine(sql As String, entityName As String, ByRef msg As String, xmlParameters As String) As String
        Return MyBase.SelectMutipleInString(Of T_MSLMachine)(sql, entityName, msg, xmlParameters, True)
    End Function
    Public Function GetMSLModel(sql As String, entityName As String, ByRef msg As String, xmlParameters As String) As String
        Return MyBase.SelectMutipleInString(Of T_MSLModel)(sql, entityName, msg, xmlParameters, True)
    End Function
    Public Sub DeleteMsl(xmlObj As String, ByRef msg As String)
        MyBase.DeleteObject(Of T_SFJobMSL)(xmlObj, msg)
    End Sub
    Public Sub InsertMsl(xmlObj As String, ByRef msg As String)
        MyBase.InsertObject(Of T_SFJobMSL)(xmlObj, msg)
    End Sub
    Public Function GetJobMsl(sql As String, entityName As String, ByRef msg As String, xmlParameters As String) As String
        Return MyBase.SelectMutipleInString(Of JobMSL)(sql, entityName, msg, xmlParameters, False)
    End Function
    Public Function ExecSP(ByVal spname As String, ByVal msg As String, xmlParameters As String) As Integer
        Dim x As Integer

        Dim parms As List(Of MySqlParameter) = xmlParameters.DeSerializerCollection(Of MySqlParameter)(GetType(List(Of MySqlParameter)))
        Dim arr As Object() = parms.ToSqlParameter()
        MyBase.ExexStoreProcedue(spname, msg, arr)
        'MyBase.ExexStoreProcedue1(spname, msg, jobid, equipment, station)
        If String.IsNullOrEmpty(msg) Then
            x = 0
        Else
            x = -1
        End If
        Return x
    End Function
End Class



Public Class EntityBaseRepository
    Public ReadOnly Property ConnStr() As String
        Get
            Return ConfigurationManager.ConnectionStrings("eTraceDBConnection").ConnectionString
        End Get
    End Property

    Private Function SelectMutipleInString(Of T As BaseModel)(sql As String, entityName As String, ByRef msg As String, loadEntities As Boolean, ParamArray parameters As Object()) As String
        Dim temp As List(Of T) = Nothing
        Dim result As String = Nothing
        Using db As New Entities(Me.ConnStr)
            Using ts As New TransactionScope(TransactionScopeOption.Required, New TransactionOptions() With {.IsolationLevel = IsolationLevel.ReadUncommitted})
                Try
                    Dim query = db.context.ExecuteStoreQuery(Of T)(sql, entityName, MergeOption.AppendOnly, parameters).ToList()
                    If loadEntities Then
                        temp = query.ExplictiLoadingRelatedEntityCollection(db)
                    Else
                        temp = query
                    End If
                    result = temp.SerializerCollection()
                Catch ex As Exception
                    msg = ex.Message
                Finally
                    ts.Complete()
                End Try
            End Using
        End Using
        Return result
    End Function

    Public Function SelectMutipleInString(Of T As BaseModel)(sql As String, entityName As String, ByRef msg As String, xmlParameters As String, loadEntities As Boolean) As String
        Try
            If String.IsNullOrEmpty(xmlParameters) Then
                Return Me.SelectMutipleInString(Of T)(sql, entityName, msg, loadEntities)
            Else
                Dim parms As List(Of MySqlParameter) = xmlParameters.DeSerializerCollection(Of MySqlParameter)(GetType(List(Of MySqlParameter)))
                Dim arr As Object() = parms.ToSqlParameter()
                Return Me.SelectMutipleInString(Of T)(sql, entityName, msg, loadEntities, arr)
            End If
        Catch ex As Exception
            msg = ex.Message
        End Try
        Return Nothing
    End Function

    Private Function SelectSingleInString(Of T As BaseModel)(sql As String, entityName As String, ByRef msg As String, loadEntities As Boolean, ParamArray parameters As Object()) As String
        Dim temp As T = Nothing
        Dim result As String = Nothing
        Using db As New Entities(Me.ConnStr)
            Using ts As New TransactionScope(TransactionScopeOption.Required, New TransactionOptions() With {.IsolationLevel = IsolationLevel.ReadUncommitted})
                Try
                    Dim query = db.context.ExecuteStoreQuery(Of T)(sql, entityName, MergeOption.AppendOnly, parameters).FirstOrDefault()
                    If loadEntities Then
                        temp = query.ExplictiLoadingRelatedEntityObject(db)
                    Else
                        temp = query
                    End If
                    result = temp.SerializerObject()
                Catch ex As Exception
                    msg = ex.Message
                Finally
                    ts.Complete()
                End Try
            End Using
        End Using
        Return result
    End Function

    Public Function SelectSingleInString(Of T As BaseModel)(sql As String, entityName As String, ByRef msg As String, xmlParameters As String, loadEntities As Boolean) As String
        Try
            If String.IsNullOrEmpty(xmlParameters) Then
                Return Me.SelectSingleInString(Of T)(sql, entityName, msg, loadEntities)
            Else
                Dim parms As List(Of MySqlParameter) = xmlParameters.DeSerializerCollection(Of MySqlParameter)(GetType(List(Of MySqlParameter)))
                Dim arr As Object() = parms.ToSqlParameter()
                Return Me.SelectSingleInString(Of T)(sql, entityName, msg, loadEntities, arr)
            End If
        Catch ex As Exception
            msg = ex.Message
        End Try
        Return Nothing
    End Function

    Public Sub InsertCollection(Of T)(xmlObj As String, ByRef msg As String)
        Try
            Dim list As List(Of T) = xmlObj.DeSerializerObject(Of List(Of T))()
            If list Is Nothing Then
                Throw New Exception("No object to insert")
            End If
            Using db As New Entities(Me.ConnStr)
                Using ts As New TransactionScope(TransactionScopeOption.Required, New TransactionOptions() With {.IsolationLevel = IsolationLevel.ReadCommitted})
                    Try
                        Dim sets As DbSet = db.GetDbSet(Of T)()
                        For Each job As T In list
                            sets.Add(job)
                        Next
                        db.SaveChanges()
                        ts.Complete()
                    Catch ex As Exception
                        Throw ex
                    Finally
                        ts.Dispose()
                    End Try
                End Using
            End Using
        Catch ex As Exception
            msg = ex.Message
            Dim imsg1 As String = ""
            Dim imsg2 As String = ""
            If ex.InnerException IsNot Nothing Then
                imsg1 = ex.InnerException.Message
                If ex.InnerException.InnerException IsNot Nothing Then
                    imsg2 = ex.InnerException.InnerException.Message
                End If
            End If
            msg = imsg2 + " " + imsg1 + " " + ex.Message
            'ErrorLogging("SMTJobManagement", "", msg, "E")
        End Try
    End Sub

    Public Sub InsertObject(Of T)(xmlObj As String, ByRef msg As String)
        Try
            Dim obj As T = xmlObj.DeSerializerObject(Of T)()
            If obj Is Nothing Then
                Throw New Exception("No object to insert")
            End If
            Using db As New Entities(Me.ConnStr)
                Using ts As New TransactionScope(TransactionScopeOption.Required, New TransactionOptions() With {.IsolationLevel = IsolationLevel.ReadCommitted})
                    Try
                        Dim sets As DbSet = db.GetDbSet(Of T)()
                        sets.Add(obj)
                        db.SaveChanges()
                        ts.Complete()
                    Catch ex As Exception
                        Throw ex
                    Finally
                        ts.Dispose()
                    End Try
                End Using
            End Using
        Catch ex As Exception
            msg = ex.Message
            Dim imsg1 As String = ""
            Dim imsg2 As String = ""
            If ex.InnerException IsNot Nothing Then
                imsg1 = ex.InnerException.Message
                If ex.InnerException.InnerException IsNot Nothing Then
                    imsg2 = ex.InnerException.InnerException.Message
                End If
            End If
            msg = imsg2 + " " + imsg1 + " " + ex.Message
            'ErrorLogging("SMTJobManagement", "", msg, "E")
        End Try
    End Sub

    Public Sub UpdateObject(Of T As BaseModel)(obj As T, ByRef msg As String, ParamArray keys As Object())
        Try
            Using db As New Entities(Me.ConnStr)
                Using ts As New TransactionScope(TransactionScopeOption.Required, New TransactionOptions() With {.IsolationLevel = IsolationLevel.ReadCommitted})
                    Try
                        Dim targetObj As T = TryCast(db.GetDbSet(Of T)().Find(keys), T)
                        If targetObj Is Nothing Then
                            Throw New Exception("No object to update")
                        End If
                        targetObj.CopyPrimitiveProperites(obj)
                        db.SaveChanges()
                        ts.Complete()
                    Catch ex As Exception
                        Throw ex
                    Finally
                        ts.Dispose()
                    End Try
                End Using
            End Using
        Catch ex As Exception
            msg = ex.Message
            Dim imsg1 As String = ""
            Dim imsg2 As String = ""
            If ex.InnerException IsNot Nothing Then
                imsg1 = ex.InnerException.Message
                If ex.InnerException.InnerException IsNot Nothing Then
                    imsg2 = ex.InnerException.InnerException.Message
                End If
            End If
            msg = imsg2 + " " + imsg1 + " " + ex.Message
            'ErrorLogging("SMTJobManagement", "", msg, "E")
        End Try
    End Sub

    Public Sub UpdateObject(Of T As BaseModel)(obj As T, ByRef msg As String, loadEntities As Boolean, clearFirst As Boolean, ParamArray keys As Object())
        Try
            Using db As New Entities(Me.ConnStr)
                Using ts As New TransactionScope(TransactionScopeOption.Required, New TransactionOptions() With {.IsolationLevel = IsolationLevel.ReadCommitted})
                    Try
                        Dim targetObj As T = TryCast(db.GetDbSet(Of T)().Find(keys), T)
                        If targetObj Is Nothing Then
                            Throw New Exception("No object to update")
                        End If
                        targetObj.CopyPrimitiveProperites(obj)
                        If loadEntities Then
                            targetObj.ExplictiLoadingRelatedEntityObject(db)
                            targetObj.ReplaceRelatedEntityObject(obj, clearFirst)
                        End If
                        db.SaveChanges()
                        ts.Complete()
                    Catch ex As Exception
                        Throw ex
                    Finally
                        ts.Dispose()
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Dim imsg1 As String = ""
            Dim imsg2 As String = ""
            If ex.InnerException IsNot Nothing Then
                imsg1 = ex.InnerException.Message
                If ex.InnerException.InnerException IsNot Nothing Then
                    imsg2 = ex.InnerException.InnerException.Message
                End If
            End If
            msg = imsg2 + " " + imsg1 + " " + ex.Message
            'ErrorLogging("SMTJobManagement", "", msg, "E")
        End Try
    End Sub
    Public Sub DeleteObject(Of T As BaseModel)(xmlObj As String, ByRef msg As String)
        Try
            Dim obj As T = xmlObj.DeSerializerObject(Of T)()
            If obj Is Nothing Then
                Throw New Exception("No object to Delete")
            End If
            Using db As New Entities(Me.ConnStr)
                Using ts As New TransactionScope(TransactionScopeOption.Required, New TransactionOptions() With {.IsolationLevel = IsolationLevel.ReadCommitted})
                    Try
                        'Dim sets As DbSet = db.GetDbSet(Of T)()
                        'sets.Remove(obj)
                        db.Entry(obj).State = EntityState.Deleted
                        db.Set(Of T).Remove(obj)
                        db.SaveChanges()
                        ts.Complete()
                    Catch ex As Exception
                        Throw ex
                    Finally
                        ts.Dispose()
                    End Try
                End Using
            End Using
        Catch ex As Exception
            msg = ex.Message
        End Try
    End Sub

    Public Sub ExexStoreProcedue(ByVal spname As String, ByVal msg As String, ParamArray parms As Object())
        Try
            Using db As New Entities(Me.ConnStr)
                Using ts As New TransactionScope(TransactionScopeOption.Required, New TransactionOptions() With {.IsolationLevel = IsolationLevel.ReadCommitted})
                    Try
                        'db.Database.ExecuteSqlCommand(spname, parms.Cast(Of SqlParameter).ToArray(0).Value, parms.Cast(Of SqlParameter).ToArray(1).Value, parms.Cast(Of SqlParameter).ToArray(2).Value)
                        db.Database.ExecuteSqlCommand(spname, parms)
                        db.SaveChanges()
                        ts.Complete()
                    Catch ex As Exception
                        Throw ex
                    Finally
                        ts.Dispose()
                    End Try
                End Using
            End Using
        Catch ex As Exception
            msg = ex.Message
        End Try
    End Sub
End Class

