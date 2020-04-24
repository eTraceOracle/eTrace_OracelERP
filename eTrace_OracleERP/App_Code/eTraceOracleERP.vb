Imports System
Imports System.Web
Imports System.Collections
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports Oracle.DataAccess.Client
Imports System.Data.OracleClient

<WebService([Namespace]:="http://eTraceOracleERP.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
Public Class eTraceOracleERP    ' Publish target location: C:\Inetpub\wwwroot\eTrace_OracleERP       '  http://cnapgzhofs10/eTrace_OracleERP  
    Public Sub New()
    End Sub


#Region "UserManagement"
    <WebMethod()> _
    Public Function ReadUserData(ByVal UserID As String, ByVal User As String) As DataSet
        Dim _eTrace_UserManagement As New UserManagement()
        Return _eTrace_UserManagement.ReadUserData(UserID, User)
    End Function

    <WebMethod()> _
    Public Function ReadRoleData(ByVal User As String) As DataSet
        Dim _eTrace_UserManagement As New UserManagement()
        Return _eTrace_UserManagement.ReadRoleData(User)
    End Function

    <WebMethod()> _
    Public Function PostUserData(ByVal UserData As UserDetail, ByVal dsUser As DataSet, ByVal LoginData As ERPLogin) As String
        Dim _eTrace_UserManagement As New UserManagement()
        Return _eTrace_UserManagement.PostUserData(UserData, dsUser, LoginData)
    End Function

    <WebMethod()> _
    Public Function PostRoleData(ByVal RoleData As RoleDetail, ByVal dsTrans As DataSet, ByVal LoginData As ERPLogin) As String
        Dim _eTrace_UserManagement As New UserManagement()
        Return _eTrace_UserManagement.PostRoleData(RoleData, dsTrans, LoginData)
    End Function

#End Region

#Region "PublicFunction"

    ''' <summary>
    '''  Get DaughterBoardSN List by IntSN
    ''' </summary>
    ''' <param name="intSN"></param>
    ''' <returns></returns>
    <WebMethod()>
    Public Function InvMigCurrStatus() As Integer
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.InvMigCurrStatus()
    End Function

    <WebMethod()>
    Public Function InvMigUserCheck(ByVal UserName As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.InvMigUserCheck(UserName)
    End Function

    <WebMethod()>
    Public Function InvMigrationStatus(ByVal Status As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.InvMigrationStatus(Status)
    End Function
    <WebMethod()>
    Public Function GetDaughterBoardSN(intSN As String) As String()
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetDaughterBoardSN(intSN)
    End Function
    <WebMethod()>
    Public Function GetMailList(ByVal eTraceModule As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetMailList(eTraceModule)
    End Function

    <WebMethod()>
    Public Function AutoMail_SiplaceDataCheck() As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.AutoMail_SiplaceDataCheck()
    End Function

    <WebMethod()>
    Public Function AutoSendMail_PastDueDJ() As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.AutoSendMail_PastDueDJ()
    End Function
    <WebMethod()> _
    Public Function DeleteMOAllocated(ByVal delMOItemList As DataSet, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.DeleteMOAllocated(delMOItemList, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function RedoLedMOAllocated(ByVal EventID As String, ByVal ERPLoginData As ERPLogin) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.RedoLedMOAllocated(EventID, ERPLoginData)
    End Function

    ' --- PHUserAuthentication
    <WebMethod()> _
    Public Function ValidateUser(ByVal Uname As String, ByVal UPwd As String, ByVal CheckPWD As Boolean) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ValidateUser(Uname, UPwd, CheckPWD)
    End Function

    <WebMethod()> _
    Public Function UploadBJData_to_Oracle() As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.UploadBJData_to_Oracle()
    End Function

    <WebMethod()>
    Public Function UploadAutoEJIT_DJSum_to_Oracle() As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.UploadAutoEJIT_DJSum_to_Oracle()
    End Function
    <WebMethod()>
    Public Function UploadAutoEJIT_DJList_to_Oracle() As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.UploadAutoEJIT_DJList_to_Oracle()
    End Function
    <WebMethod()> _
    Public Function AutoCreatedEJIT() As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.AutoCreatedEJIT()
    End Function
    <WebMethod()> _
    Public Function AutoCreatedEJIT_ByDay() As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.AutoCreatedEJIT_ByDay()
    End Function
    <WebMethod()> _
    Public Function ErrorLogging(ByVal ModuleName As String, ByVal User As String, ByVal ErrMsg As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ErrorLogging(ModuleName, User, ErrMsg)
    End Function

    <WebMethod()> _
    Public Function ErrorLog(ByVal ModuleName As String, ByVal User As String, ByVal ErrMsg As String, ByVal Category As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ErrorLogging(ModuleName, User, ErrMsg, Category)
    End Function

    <WebMethod()> _
    Public Function ErrorLogOTO(ByVal ModuleName As String, ByVal User As String, ByVal ErrMsg As String, ByVal Category As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ErrorLogging(ModuleName, User, ErrMsg, Category)
    End Function

    <WebMethod()> _
    Public Function GetServerDate() As Date
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetServerDate()
    End Function

    <WebMethod()> _
    Public Function GetConfigValue(ByVal ConfigID As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetConfigValue(ConfigID)
    End Function

    <WebMethod()> _
    Public Function PrinterCheck(ByVal User As String, ByVal PrinterID As String, ByVal OutputType As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.PrinterCheck(User, PrinterID, OutputType)
    End Function

    <WebMethod()> _
    Public Function LoginCheck(ByVal ERPLoginData As ERPLogin) As UserData
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.LoginCheck(ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function ChangePassword(ByVal LoginData As ERPLogin) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ChangePassword(LoginData)
    End Function

    <WebMethod()> _
    Public Function ValidateItemRevision(ByVal OracleLoginData As ERPLogin, ByVal Item As String, ByVal Revision As String, ByVal MoveType As String) As ItemRevList
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ValidateItemRevision(OracleLoginData, Item, Revision, MoveType)
    End Function

    <WebMethod()> _
    Public Function ValidateItemType(ByVal OracleLoginData As ERPLogin, ByVal Item As String, ByVal Revision As String, ByVal MoveType As String) As ItemType
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ValidateItemType(OracleLoginData, Item, Revision, MoveType)
    End Function

    <WebMethod()> _
    Public Function ValidateSubLoc(ByVal LoginData As ERPLogin, ByVal SubInv As String, ByVal Locator As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ValidateSubLoc(LoginData, SubInv, Locator)
    End Function

    <WebMethod()> _
    Public Function Get_LocatorSP(ByVal LoginData As ERPLogin, ByVal SubInv As String, ByVal Locator As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.Get_LocatorSP(LoginData, SubInv, Locator)
    End Function

    <WebMethod()> _
    Public Function SlotCheck(ByVal LoginData As ERPLogin, ByRef myCLIDSlot As CLIDSlot) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.SlotCheck(LoginData, myCLIDSlot)
    End Function

    <WebMethod()> _
    Public Function SlotLightOn(ByVal SlotLists As String, ByVal LightOn As Boolean, ByVal User As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.SlotLightOn(SlotLists, LightOn, User)
    End Function

    <WebMethod()> _
    Public Function DockSlotLightOn(ByVal SlotLists As String, ByVal LightOn As Boolean, ByVal User As String, ByVal Interval As Integer) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.DockSlotLightOn(SlotLists, LightOn, User, Interval)
    End Function

    <WebMethod()> _
    Public Function Get_ReasonCode(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction
        Return _eTrace_PublicFunction.Get_ReasonCode(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function Get_SubinvLoc(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction
        Return _eTrace_PublicFunction.Get_SubinvLoc(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function get_iteminfo(ByVal ItemList As DataSet, ByVal oraclelogindata As ERPLogin, ByVal movetype As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction
        Return _eTrace_PublicFunction.get_iteminfo(ItemList, oraclelogindata, movetype)
    End Function

    <WebMethod()> _
    Public Function get_itemonhand(ByVal ItemList As DataSet, ByVal oraclelogindata As ERPLogin, ByVal movetype As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction
        Return _eTrace_PublicFunction.get_itemonhand(ItemList, oraclelogindata, movetype)
    End Function

    <WebMethod()> _
    Public Function Get_AccountAlias(ByVal OracleLoginData As ERPLogin, ByVal Type As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction
        Return _eTrace_PublicFunction.Get_AccountAlias(OracleLoginData, Type)
    End Function

    <WebMethod()> _
    Public Function Get_TransactionTypes(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction
        Return _eTrace_PublicFunction.Get_TransactionTypes(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetScreenElements(ByVal GroupName As String, ByVal Lang As String, ByVal MessageClass As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetScreenElements(GroupName, Lang, MessageClass)
    End Function

    <WebMethod()> _
    Public Function OpenPrintMatlLabel() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintMatlLabel()
    End Function

    <WebMethod()> _
    Public Function OpenPrintNoMatlLabel() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintNoMatlLabel()
    End Function

    <WebMethod()> _
    Public Function WritePrintLabel(ByVal LabelSeqNo As String, ByVal LabelFile As String, ByVal LabelPrinter As String, ByVal Content As String)
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.WritePrintLabel(LabelSeqNo, LabelFile, LabelPrinter, Content)
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabel() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabel()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabel_Production() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabel_Production()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabel_LAG() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabel_LAG()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabel_CVT() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabel_CVT()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabelForApps() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabelForApps()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabelForEtrace2() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabelForEtrace2()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabelForPrtSvr() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabelForPrtSvr()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabelForTemp01() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabelForTemp01()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabelForTemp02() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabelForTemp02()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabelForTemp03() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabelForTemp03()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabelForTemp04() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabelForTemp04()
    End Function

    <WebMethod()> _
    Public Function OpenPrintLabelForTemp05() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.OpenPrintLabelForTemp05()
    End Function

    <WebMethod()> _
    Public Function updatePrintLabel(ByVal SeqNo As String, ByVal ErrMsg As String)
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.updatePrintLabel(SeqNo, ErrMsg)
    End Function

    <WebMethod()> _
    Public Function PrintCLIDs(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.PrintCLIDs(CLIDs, Printer)
    End Function

    <WebMethod()> _
    Public Function GetProcessProperties(ByVal GroupName As String, ByVal Lang As String, ByVal TransactionID As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetProcessProperties(GroupName, Lang, TransactionID)
    End Function

    <WebMethod()> _
    Public Function GetAccessCardUserInfo(ByVal AccessCardID As String) As AccessCard
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetAccessCardUserInfo(AccessCardID)
    End Function

    <WebMethod()> _
    Public Function GetMGTraceLevel(ByVal MaterialNo As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetMGTraceLevel(MaterialNo)
    End Function

    <WebMethod()> _
    Public Function GetAML(ByVal ItemList() As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetAML(ItemList)
    End Function

    <WebMethod()> _
    Public Function GetAMLData(ByVal ItemList() As String, ByVal ModuleName As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetAMLData(ItemList, ModuleName)
    End Function

    <WebMethod()> _
    Public Function GetOrgLists(ByVal eTraceModule As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetOrgLists(eTraceModule)
    End Function

    <WebMethod()> _
    Public Function GetLoginData(ByVal eTraceModule As String, ByVal TransactionID As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetLoginData(eTraceModule, TransactionID)
    End Function

    <WebMethod()> _
    Public Function GetHHVersion(ByVal ConfigID As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetHHVersion(ConfigID)
    End Function

    '<WebMethod()> _
    'Public Function CleanBoxID(ByVal CLID As String, ByVal ERPLoginData As ERPLogin) As Boolean
    '    Dim _eTrace_PublicFunction As New PublicFunction()
    '    Return _eTrace_PublicFunction.CleanBoxID(CLID, ERPLoginData)
    'End Function

    <WebMethod()> _
    Public Function ClearBoxID(ByVal BoxID As String, ByVal ERPLoginData As ERPLogin) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ClearBoxID(BoxID, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function Get_Subinv_Restrict(ByVal OracleLoginData As ERPLogin, ByVal TransType As String, ByVal AcctAliasName As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.Get_Subinv_Restrict(OracleLoginData, TransType, AcctAliasName)
    End Function

    <WebMethod()> _
    Public Function GetCOOLists() As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetCOOLists()
    End Function

    <WebMethod()> _
    Public Function Post_SubInvTransfer(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.Post_SubInvTransfer(p_ds, OracleLoginData, MoveType, TransactionID)
    End Function

    <WebMethod()> _
    Public Function GetSysMessage(ByVal frequecy As String, ByVal device As String, ByVal eTraceModel As String) As DataSet
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.GetSysMessage(frequecy, device, eTraceModel)
    End Function

    <WebMethod()> _
    Public Function ShowMessage(ByVal MessageText As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ShowMessage(MessageText)
    End Function

    <WebMethod()> _
    Public Function ValidatePallet(ByVal LoginData As ERPLogin, ByVal PalletID As String, ByVal MatlStatus As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ValidatePallet(LoginData, PalletID, MatlStatus)
    End Function

    <WebMethod()> _
    Public Function ValidateBerth(ByVal LoginData As ERPLogin, ByVal BerthID As String) As String
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.ValidateBerth(LoginData, BerthID)
    End Function

    <WebMethod()> _
    Public Function PrintInterOrgCLIDs(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.PrintInterOrgCLIDs(CLIDs, Printer)
    End Function

    <WebMethod()> _
    Public Function PrintCH09Label(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.PrintCH09Label(CLIDs, Printer)
    End Function

#End Region

#Region "LabelBasicFunction"

    <WebMethod()> _
    Public Function ReadLabelIDs(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal TransactionType As String) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.ReadLabelIDs(LoginData, CLID, TransactionType)
    End Function

    '<WebMethod()> _
    'Public Function ReadLabelAML(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal StatusCheck As Boolean) As DataSet
    '    Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
    '    Return _eTrace_LabelBasicFunction.ReadLabelAML(LoginData, CLID, StatusCheck)
    'End Function

    <WebMethod()> _
    Public Function SplitLabelIDs(ByVal LoginData As ERPLogin, ByVal LabelID As String, ByVal LabelStatus As String, ByVal Items As DataSet) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.SplitLabelIDs(LoginData, LabelID, LabelStatus, Items)
    End Function

    <WebMethod()> _
    Public Function LabelDataUpdate(ByVal LoginData As ERPLogin, ByVal myCLID As LabelData) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.LabelDataUpdate(LoginData, myCLID)
    End Function

    <WebMethod()> _
    Public Function ReadPalletID(ByVal LoginData As ERPLogin, ByVal LabelID As String) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.ReadPalletID(LoginData, LabelID)
    End Function

    <WebMethod()> _
    Public Function PalletManagement(ByVal LoginData As ERPLogin, ByVal myCLIDs As DataSet) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.PalletManagement(LoginData, myCLIDs)
    End Function

    <WebMethod()> _
    Public Function GetCartonLabel(ByVal LoginData As ERPLogin, ByVal CartonID As String) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetCartonLabel(LoginData, CartonID)
    End Function

    <WebMethod()> _
    Public Function PrintCartonLabel(ByVal myCLID As DataSet, ByVal Printer As String) As Boolean
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.PrintCartonLabel(myCLID, Printer)
    End Function

    <WebMethod()> _
    Public Function LabelConversion(ByVal LoginData As ERPLogin, ByVal LabelPrinter As String, ByVal PrintLabels As Boolean, ByVal GRInfo As LabelData) As ConversionResult
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.LabelConversion(LoginData, LabelPrinter, PrintLabels, GRInfo)
    End Function

    <WebMethod()> _
    Public Function GetItemMaster(ByVal LoginData As ERPLogin, ByVal PartNo As String) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetItemMaster(LoginData, PartNo)
    End Function

    <WebMethod()> _
    Public Function GetItemDescription(ByVal LoginData As ERPLogin, ByVal PartNo As String) As ItemData
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetItemDescription(LoginData, PartNo)
    End Function

    <WebMethod()> _
    Public Function PartNoTransfer(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.PartNoTransfer(LoginData, Items)
    End Function

    <WebMethod()> _
    Public Function InterOrgTransfer(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.InterOrgTransfer(LoginData, Items)
    End Function

    <WebMethod()> _
    Public Function GetRTVLabels(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal CVMIFlag As String) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetRTVLabels(LoginData, CLID, CVMIFlag)
    End Function

    <WebMethod()> _
    Public Function MiscRTV(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.MiscRTV(LoginData, Items)
    End Function

    <WebMethod()> _
    Public Function CVMIRTV(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.CVMIRTV(LoginData, Items)
    End Function

    <WebMethod()> _
    Public Function GetItemRTLot(ByVal LoginData As ERPLogin, ByVal PartNo As String, ByVal RTLot As String) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetItemRTLot(LoginData, PartNo, RTLot)
    End Function

    <WebMethod()> _
    Public Function GetItemOnhand(ByVal LoginData As ERPLogin, ByVal ItemNo As String, ByVal Subinv As String, ByVal LocName As String, ByVal Revision As String, ByVal LotNum As String, ByVal Qty As String) As String()
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetItemOnhand(LoginData, ItemNo, Subinv, LocName, Revision, LotNum, Qty)
    End Function
    <WebMethod()> _
    Public Function ReadLabelIDInfo(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal StatusCheck As Boolean) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.ReadLabelIDInfo(LoginData, CLID, StatusCheck)
    End Function
    <WebMethod()> _
    Public Function LabelGeneration(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal QtyUnits As String, ByVal Type As String) As String()
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.LabelGeneration(LoginData, CLID, QtyUnits, Type)
    End Function
    <WebMethod()> _
    Public Function GetTypeID(ByVal CLID As String) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetTypeID(CLID)
    End Function

    <WebMethod()> _
    Public Function ReadLabelIDsForDJ(ByVal LoginData As ERPLogin, ByVal CLID As String) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.ReadLabelIDsForDJ(LoginData, CLID)
    End Function

    <WebMethod()> _
    Public Function Get_LabelInfoHasSameLot(ByVal CLID As String, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.Get_LabelInfoHasSameLot(CLID, LoginData)
    End Function

    <WebMethod()> _
    Public Function Update_LabelInfoHasSameLot(ByVal LoginData As ERPLogin, ByVal ItemNO As String, ByVal LotNum As String, ByVal ExpDate As Date) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.Update_LabelInfoHasSameLot(LoginData, ItemNO, LotNum, ExpDate)
    End Function

    <WebMethod()> _
    Public Function GetPalletList(ByVal PalletID As String) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetPalletList(PalletID)
    End Function

    <WebMethod()> _
    Public Function GetStatusList(ByVal TypeCode As String) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetStatusList(TypeCode)
    End Function

    <WebMethod()> _
    Public Function UpdatePalletList(ByVal ERPLogin As ERPLogin, ByVal myDashBoard As DashBoardData) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.UpdatePalletList(ERPLogin, myDashBoard)
    End Function

    <WebMethod()> _
    Public Function GetMaterialInfo(ByVal RTNo As String, ByVal Material As String) As DataSet
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.GetMaterialInfo(RTNo, Material)
    End Function

    <WebMethod()> _
    Public Function UpdateMaterial(ByVal ERPLogin As ERPLogin, ByVal myDashBoard As DashBoardData) As String
        Dim _eTrace_LabelBasicFunction As New LabelBasicFunction()
        Return _eTrace_LabelBasicFunction.UpdateMaterial(ERPLogin, myDashBoard)
    End Function

#End Region

#Region "Receiving"
    <WebMethod()>
    Public Function CheckItemMPN(ByVal Item As String, ByVal ManufacturerPN As String) As String
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.CheckItemMPN(Item, ManufacturerPN)
    End Function
    <WebMethod()>
    Public Function DownloadSuppData(ByVal ItemNo As String) As DataSet
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.DownloadSuppData(ItemNo)
    End Function
    <WebMethod()>
    Public Function SaveSuppData(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As CreateGRResponse
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.SaveSuppData(LoginData, Items)
    End Function
    <WebMethod()> _
    Public Function GetPOLineMPQ(ByVal p_ds As DataTable) As DataTable
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.GetPOLineMPQ(p_ds)
    End Function
    <WebMethod()> _
    Public Function GetItemMPQ(ByVal p_org_id As Int32, ByVal p_po_line_id As Integer) As Double
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.GetItemMPQ(p_org_id, p_po_line_id)
    End Function

    <WebMethod()> _
    Public Function GetRecData(ByVal eTraceModule As String, ByVal TransactionID As String) As DataSet
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.GetRecData(eTraceModule, TransactionID)
    End Function

    <WebMethod()> _
    Public Function GetItems(ByVal LoginData As ERPLogin, ByRef GRHeader As GRHeaderStructure) As DataSet
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.GetItems(LoginData, GRHeader)
    End Function

    <WebMethod()> _
    Public Function ProcessMatMovement(ByVal LoginData As ERPLogin, ByVal Header As GRHeaderStructure, ByVal Items As DataSet, ByVal PrintLabels As Boolean, ByVal LabelPrinter As String) As CreateGRResponse
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.ProcessMatMovement(LoginData, Header, Items, PrintLabels, LabelPrinter)
    End Function

    <WebMethod()> _
    Public Function SaveDCodeLotNo(ByVal LoginData As ERPLogin, ByVal CLIDs As DataSet) As DataSet
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.SaveDCodeLotNo(LoginData, CLIDs)
    End Function

    <WebMethod()> _
    Public Function PrintRTSlip(ByVal LoginData As ERPLogin, ByVal RTItems As DataSet, ByVal LabelPrinter As String) As Boolean
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.PrintRTSlip(LoginData, RTItems, LabelPrinter)
    End Function

    <WebMethod()> _
    Public Function PrintRECLabels(ByVal CLIDs As DataSet, ByVal LabelPrinter As String) As Boolean
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.PrintRECLabels(CLIDs, LabelPrinter)
    End Function

    <WebMethod()> _
    Public Function ReadShipmentData(ByVal LoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.ReadShipmentData(LoginData, OrderNo)
    End Function

    <WebMethod()> _
    Public Function ReadIRData(ByVal myIRData As IRData) As DataSet
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.ReadIRData(myIRData)
    End Function

    <WebMethod()> _
    Public Function UpdateIRRTNo(ByVal myIRData As IRData, ByVal Items As DataSet) As Boolean
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.UpdateIRRTNo(myIRData, Items)
    End Function

    <WebMethod()> _
    Public Function CleanBoxID(ByVal LoginData As ERPLogin, ByVal CLIDs As DataSet) As Boolean
        Dim _eTrace_Receiving As New Receiving()
        Return _eTrace_Receiving.CleanBoxID(LoginData, CLIDs)
    End Function

#End Region

#Region "eTrace-Putaway Module"
    <WebMethod()> _
    Public Function GetValidSource(ByVal OracleLoginData As ERPLogin, ByVal CLID As String, ByVal SubInv As String, ByVal Locator As String) As DataSet     ' ByVal StorageType As String
        Dim _eTrace_Putaway As New Putaway()
        Return _eTrace_Putaway.GetValidSource(OracleLoginData, CLID, SubInv, Locator)
    End Function

    <WebMethod()> _
    Public Function PutawayPost(ByVal OracleLoginData As ERPLogin, ByVal Items As DataSet) As DataSet
        Dim _eTrace_Putaway As New Putaway()
        Return _eTrace_Putaway.PutawayPost(OracleLoginData, Items)
    End Function

    <WebMethod()> _
    Public Function GetMRBSubInv(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_Putaway As New Putaway()
        Return _eTrace_Putaway.GetMRBSubInv(LoginData)
    End Function

    <WebMethod()> _
    Public Function ReadBlockDCLN(ByVal ds As DataSet) As DataSet
        Dim _eTrace_Putaway As New Putaway()
        Return _eTrace_Putaway.ReadBlockDCLN(ds)
    End Function

    <WebMethod()> _
    Public Function SaveBlockDCLN(ByVal LoginData As ERPLogin, ByVal ds As DataSet) As String
        Dim _eTrace_Putaway As New Putaway()
        Return _eTrace_Putaway.SaveBlockDCLN(LoginData, ds)
    End Function

    <WebMethod()> _
    Public Function SourceForCompToDJ(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_Putaway As New Putaway()
        Return _eTrace_Putaway.SourceForCompToDJ(CLID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function WS_CompToDJ(ByVal MatlList As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_Putaway As New Putaway()
        Return _eTrace_Putaway.WS_CompToDJ(MatlList, OracleLoginData)
    End Function
#End Region

#Region "WFOE U-Turn"
    <WebMethod()> _
    Public Function CheckUTurnSubinv(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckUTurnSubinv(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ReadCLID_UTurn(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ReadCLID_UTurn(CLID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function UTurnDelivery(ByVal OracleLoginData As ERPLogin, ByVal CLIDS As DataSet) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.UTurnDelivery(OracleLoginData, CLIDS)
    End Function

    <WebMethod()> _
    Public Function CheckUTurnCLIDFormat(ByVal ExcelData As DataSet, ByVal SubInv As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckUTurnCLIDFormat(ExcelData, SubInv, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function PostUTurnStatusChange(ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.PostUTurnStatusChange(ExcelData, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function CheckLocRTList(ByVal Item As String, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckLocRTList(Item, Subinv, Locator, OracleLoginData)
    End Function


#End Region

#Region "Batch Disable CLIDs"
    <WebMethod()>
    Public Function CheckCLIDDisableFlag(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckCLIDDisableFlag(OracleLoginData)
    End Function

    <WebMethod()>
    Public Function ValidateSubinv(ByVal OracleLoginData As ERPLogin, ByVal SubInv As String) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ValidateSubinv(OracleLoginData, SubInv)
    End Function

    <WebMethod()>
    Public Function CheckBatchDisableCLIDFormat(ByVal ExcelData As DataSet, ByVal SubInv As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckBatchDisableCLIDFormat(ExcelData, SubInv, OracleERPLogin)
    End Function

    <WebMethod()>
    Public Function PostBatchDisableCLID(ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.PostBatchDisableCLID(ExcelData, OracleERPLogin)
    End Function


#End Region

#Region "MaterialMovement"
    <WebMethod()> _
    Public Function UpdateSlotCheckOption(ByVal Options As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.UpdateSlotCheckOption(Options, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function PackingManagement(ByVal FirstScan As String, ByVal CLIDBoxID As String, ByVal PalletWeight As Decimal, ByVal ActionType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.PackingManagement(FirstScan, CLIDBoxID, PalletWeight, ActionType, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function SFGetDcodeLnforIntSN(ByVal IntSN As String, ByVal PCBA As String, ByVal Component As String) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.SFGetDcodeLnforIntSN(IntSN, PCBA, Component)
    End Function

    <WebMethod()> _
    Public Function SFGetDcodeLnforSN(ByVal SN As String, ByVal PCBA As String, ByVal Component As String) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.SFGetDcodeLnforSN(SN, PCBA, Component)
    End Function

    <WebMethod()> _
    Public Function Save_SubInvTransfer(ByVal p_ds As DataSet, ByVal UpdateTable As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.Save_SubInvTransfer(p_ds, UpdateTable, OracleLoginData, MoveType, TransactionID)
    End Function

    <WebMethod()> _
    Public Function MatSourceRead(ByVal CLID As String, ByVal OrgCode As String, ByVal SourceSubInv As String, ByVal SourceLocator As String, ByVal DestSubInv As String, ByVal DestLocator As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.MatSourceRead(CLID, OrgCode, SourceSubInv, SourceLocator, DestSubInv, DestLocator, MoveType, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ClearPalletID(ByVal CartonID As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ClearPalletID(CartonID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckCLID(ByVal CLID As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckCLID(CLID, MoveType, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function RepSourceRead(ByVal CLID As String, ByVal OrgCode As String, ByVal Item As String, ByVal Revision As String, ByVal SourceSubInv As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.RepSourceRead(CLID, OrgCode, Item, Revision, SourceSubInv, MoveType, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckDest(ByVal OrgCode As String, ByVal MaterialNo As String, ByVal RTLot As String, ByVal DestSubInv As String, ByVal DestLocator As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckDest(OrgCode, MaterialNo, RTLot, DestSubInv, DestLocator, MoveType, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function Check_Dest(ByVal OrgCode As String, ByVal MaterialNo As String, ByVal RTLot As String, ByVal DestSubInv As String, ByVal DestLocator As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.Check_Dest(OrgCode, MaterialNo, RTLot, DestSubInv, DestLocator, MoveType, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function Check_Return_Option(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.Check_Return_Option(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function account_alias_batch_receipt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Double, ByVal MiscType As String) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.account_alias_batch_receipt(p_ds, OracleLoginData, MoveType, TransactionID, MiscType)
    End Function

    <WebMethod()> _
    Public Function account_alias_batch_issue(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Double, ByVal MiscType As String) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.account_alias_batch_issue(p_ds, OracleLoginData, MoveType, TransactionID, MiscType)
    End Function

    <WebMethod()> _
    Public Function account_receipt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.account_receipt(p_ds, OracleLoginData, MoveType, TransactionID)
    End Function

    <WebMethod()> _
    Public Function save_account_receipt(ByVal p_ds As DataSet, ByVal UpdateTable As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.save_account_receipt(p_ds, UpdateTable, OracleLoginData, MoveType, TransactionID)
    End Function

    <WebMethod()> _
    Public Function account_alias_receipt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer, ByVal MiscType As String) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.account_alias_receipt(p_ds, OracleLoginData, MoveType, TransactionID, MiscType)
    End Function

    <WebMethod()> _
    Public Function account_issue(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.account_issue(p_ds, OracleLoginData, MoveType, TransactionID)
    End Function

    <WebMethod()> _
    Public Function post_pull_return(ByVal MoveOracle As DataSet, ByVal TransactionID As Long, ByVal MiscType As String, ByVal CLID As String, ByVal Qty As Decimal, ByVal DestSub As String, ByVal DestLoc As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin, ByVal SlotNo As String) As Result
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.post_pull_return(MoveOracle, TransactionID, MiscType, CLID, Qty, DestSub, DestLoc, MoveType, OracleLoginData, SlotNo)
    End Function

    <WebMethod()> _
    Public Function account_issue_rcpt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer, ByVal MiscType As String) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.account_issue_rcpt(p_ds, OracleLoginData, MoveType, TransactionID, MiscType)
    End Function

    <WebMethod()> _
    Public Function account_alias_issue(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer, ByVal MiscType As String) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.account_alias_issue(p_ds, OracleLoginData, MoveType, TransactionID, MiscType)
    End Function

    <WebMethod()> _
    Public Function save_account_issue(ByVal p_ds As DataSet, ByVal UpdateTable As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.save_account_issue(p_ds, UpdateTable, OracleLoginData, MoveType, TransactionID)
    End Function

    <WebMethod()> _
    Public Function post_misc_rcpt(ByVal MoveOracle As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Long, ByVal LabelInfo As LabelData, ByVal Pkg As Integer, ByVal Printer As String) As ConversionResult
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.post_misc_rcpt(MoveOracle, OracleLoginData, MoveType, TransactionID, LabelInfo, Pkg, Printer)
    End Function

    <WebMethod()> _
    Public Function LabelForMiscReceipt(ByVal LoginData As ERPLogin, ByVal LabelPrinter As String, ByVal PrintLabels As Boolean, ByVal LabelInfo As LabelData) As ConversionResult
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.LabelForMiscReceipt(LoginData, LabelPrinter, PrintLabels, LabelInfo)
    End Function

    <WebMethod()> _
    Public Function ExportHWDataInfo(ByVal User As String, ByVal exportdata As String) As HW_ExportDataInfo
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ExportHWDataInfo(User, exportdata)
    End Function

    <WebMethod()> _
    Public Function eTrace_Update(ByVal UpdtTable As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String) As Boolean
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.eTrace_Update(UpdtTable, OracleLoginData, MoveType)
    End Function

    <WebMethod()> _
    Public Function Misc_issue_rcpt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.Misc_issue_rcpt(p_ds, OracleLoginData, MoveType)
    End Function

    <WebMethod()> _
    Public Function post_push_return(ByVal MoveOracle As DataSet, ByVal CLID As String, ByVal Qty As Decimal, ByVal DestSub As String, ByVal DestLoc As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin, ByVal SlotNo As String) As Result
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.post_push_return(MoveOracle, CLID, Qty, DestSub, DestLoc, MoveType, OracleLoginData, SlotNo)
    End Function

    <WebMethod()> _
    Public Function component_return(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.component_return(p_ds, OracleLoginData, MoveType)
    End Function

    <WebMethod()> _
    Public Function ChangeCLID(ByVal CLID As String, ByVal Qty As Double, ByVal DestSub As String, ByVal DestLoc As String, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal SlotNo As String) As Boolean
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ChangeCLID(CLID, Qty, DestSub, DestLoc, OracleLoginData, MoveType, SlotNo)
    End Function

    <WebMethod()> _
    Public Function GetRTNo(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.GetRTNo(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetRTNo_MiscRcpt(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.GetRTNo_MiscRcpt(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetDate(ByVal OracleLoginData As ERPLogin) As Date
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.GetDate(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetNextInvID(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.GetNextInvID(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckFormat(ByVal InvID As String, ByVal MRListData As DataSet, ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckFormat(InvID, MRListData, ExcelData, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function CheckBatchFormat(ByVal InvID As String, ByVal BatchListData As DataSet, ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckBatchFormat(InvID, BatchListData, ExcelData, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function CheckPIFormat(ByVal InvID As String, ByVal PIListData As DataSet, ByVal ExcelData As DataSet, ByVal Type As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckPIFormat(InvID, PIListData, ExcelData, Type, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function ValidMRData(ByVal MRListData As DataSet, ByVal ItemListData As DataSet, ByVal dsAML As DataSet, ByVal SubinvLoc As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ValidMRData(MRListData, ItemListData, dsAML, SubinvLoc, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ValidPIData(ByVal PIListData As DataSet, ByVal ItemListData As DataSet, ByVal ItemOnhand As DataSet, ByVal SubinvLoc As DataSet, ByVal Type As String, ByVal TransactionID As Long, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ValidPIData(PIListData, ItemListData, ItemOnhand, SubinvLoc, Type, TransactionID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ValidINVNo(ByVal INVNo As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ValidINVNo(INVNo, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function ValidBatchNo(ByVal BatchNo As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ValidBatchNo(BatchNo, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function PostMR(ByVal MRListData As DataSet, ByVal MoveType As String, ByVal Printer As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.PostMR(MRListData, MoveType, Printer, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function PrintCLIDforMR(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.PrintCLIDforMR(CLIDs, Printer)
    End Function

    <WebMethod()> _
      Public Function CS_GetCLIDInfo(ByVal clid As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CS_GetCLIDInfo(clid, OracleLoginData)
    End Function

    <WebMethod()> _
  Public Function CS_UpdateCLID(ByVal clid As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CS_UpdateCLID(clid, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetBerth() As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.GetBerth()
    End Function

    <WebMethod()> _
    Public Function GetDashboardData() As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.GetDashboardData()
    End Function

    <WebMethod()> _
    Public Function UploadEMCfile(ByVal DS As DataSet) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.UploadEMCfile(DS)
    End Function
    <WebMethod()> _
    Public Function ValidBatchData(ByVal MRListData As DataSet, ByVal dsAML As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.ValidBatchData(MRListData, dsAML, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function PostBatchList(ByVal BatchListData As DataSet, ByVal MoveType As String, ByVal Print As Boolean, ByVal Printer As String, ByVal OracleLoginData As ERPLogin) As PostBatchRslt
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.PostBatchList(BatchListData, MoveType, Print, Printer, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function BJ_GetDJInfo(ByVal DJ As String, ByVal PWC As String, ByVal OracleLoginData As ERPLogin) As BJ_Rs
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.BJ_GetDJInfo(DJ, PWC, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function BJ_Creation(ByVal DJInfo As DataSet, ByVal BJds As DataSet, ByVal PWC As String, ByVal OracleLoginData As ERPLogin) As BJ_Rs
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.BJ_Creation(DJInfo, BJds, PWC, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function BJ_SaveChange(ByVal BJInfo As DataSet, ByVal BJInitial As DataSet, ByVal txtDJ As String, ByVal CreateFlag As Boolean, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.BJ_SaveChange(BJInfo, BJInitial, txtDJ, CreateFlag, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckBJInfo(ByVal DJ As String, ByVal PWC As String, ByVal BJds As DataSet, ByVal OracleLoginData As ERPLogin) As BJ_Rs
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.CheckBJInfo(DJ, PWC, BJds, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function BJ_GetBJ(ByVal DJ As String, ByVal PWC As String, ByVal BJds As DataSet, ByVal OnlyOpen As Boolean, ByVal OracleLoginData As ERPLogin) As BJ_Rs
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.BJ_GetBJ(DJ, PWC, BJds, OnlyOpen, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function BJ_GenMO(ByVal PWC As String, ByVal DJ As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.BJ_GenMO(PWC, DJ, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function LEDDashBoardByRack() As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.LEDDashBoardByRack()
    End Function
    <WebMethod()> _
    Public Function LEDDashBoardPCB(ByVal PCBWarehouse As String) As DataSet
        Dim _eTraceOracleMatMove As New MaterialMoveMent()
        Return _eTraceOracleMatMove.LEDDashBoardPCB(PCBWarehouse)
    End Function
#End Region

#Region "Material Replemenishment"
    <WebMethod()> _
    Public Function Get_SrcSubInv(ByVal ERPLoginData As ERPLogin) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Get_SrcSubInv(ERPLoginData)
    End Function
    <WebMethod()> _
    Public Function Get_AvlQty(ByVal Material As String, ByVal Rev As String, ByVal Manufacturer As String, ByVal ManufacturerPN As String, ByVal Header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As Decimal
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Get_AvlQty(Material, Rev, Manufacturer, ManufacturerPN, Header, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function GetMSBSource(ByVal mydataset As DataSet, ByVal DJItems As DataSet, ByVal Header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetMSBSource(mydataset, DJItems, Header, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function Save_Template(ByVal TOMydataSet As DataSet, ByVal TemplateStr As StrTemplate, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Save_Template(TOMydataSet, TemplateStr, ERPLoginData, Header)
    End Function

    <WebMethod()> _
    Public Function SaveProdPicking(ByVal TOMydataSet As DataSet, ByVal TemplateStr As StrTemplate, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.SaveProdPicking(TOMydataSet, TemplateStr, ERPLoginData, Header)
    End Function

    <WebMethod()> _
    Public Function GetTemplateItem(ByVal Template As String, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetTemplateItem(Template, ERPLoginData)
    End Function
    <WebMethod()> _
    Public Function GetDJInfo(ByVal ID As String, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetDJInfo(ID, ERPLoginData)
    End Function
    <WebMethod()> _
    Public Function UpdateCLID(ByVal ID As String, ByVal ERPLoginData As ERPLogin) As Boolean
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.UpdateCLID(ID, ERPLoginData)
    End Function
    <WebMethod()> _
    Public Function Update_CLID(ByVal ID As String, ByVal TransType As String, ByVal ERPLoginData As ERPLogin) As Boolean
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.Update_CLID(ID, TransType, ERPLoginData)
    End Function
    <WebMethod()> _
    Public Function Post_DJ_Reversal(ByVal ERPLoginData As ERPLogin, ByVal DJ As String, ByVal orgcode As String, ByVal return_qty As Integer, ByVal uom As String, ByVal SubInventory As String, ByVal Locator As String, ByVal CLID As String) As dj_response
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.Post_DJ_Reversal(ERPLoginData, DJ, orgcode, return_qty, uom, SubInventory, Locator, CLID)
    End Function
    <WebMethod()> _
    Public Function GetTOItem(ByVal TOID As String, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetTOItem(TOID, ERPLoginData)
    End Function
    <WebMethod()> _
    Public Function Search_Template(ByVal Template As String) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Search_Template(Template)
    End Function

    <WebMethod()> _
    Public Function Check_Template(ByVal Template As String, ByVal ERPLoginData As ERPLogin) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Check_Template(Template, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function Search_LockTO(ByVal TOVALUE As String) As Boolean
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Search_LockTO(TOVALUE)
    End Function

    <WebMethod()> _
    Public Function Search_OpenTO(ByVal TOVALUE As String) As Boolean
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Search_OpenTO(TOVALUE)
    End Function

    <WebMethod()> _
    Public Function Unlock_PickOrder(ByVal PickOrder As String, ByVal erplogindata As ERPLogin) As Boolean
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Unlock_PickOrder(PickOrder, erplogindata)
    End Function

    <WebMethod()> _
    Public Function ClosePickOrder(ByVal mydataset As DataSet, ByVal PickOrder As String, ByVal ERPLoginData As ERPLogin) As Boolean
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.ClosePickOrder(mydataset, PickOrder, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function Get_PickOrder(ByVal header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Get_PickOrder(header, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function GetTemplate(ByVal Header As ProdPickingStructure) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetTemplate(Header)
    End Function

    <WebMethod()> _
    Public Function Check_PickList(ByVal mydataset As DataSet, ByVal Header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Check_PickList(mydataset, Header, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function PostProdPicking(ByVal TOMydataSet As DataSet, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure) As ProdPickingResponse
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.PostProdPicking(TOMydataSet, ERPLoginData, Header)
    End Function
    <WebMethod()> _
    Public Function DeleteTemplate(ByVal Template As String) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.DeleteTemplate(Template)
    End Function

    <WebMethod()> _
    Public Function GetBOMFromERP(ByVal mydataset As DataSet, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure, ByVal OrderNo As String) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetBOMFromERP(mydataset, ERPLoginData, Header, OrderNo)
    End Function

    <WebMethod()> _
    Public Function GetListFromExcel(ByVal mydataset As DataSet, ByVal ExcelData As DataSet, ByVal Header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As PPDataRst
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetListFromExcel(mydataset, ExcelData, Header, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function Get_DestSubLoc(ByVal DJ As String, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure) As ProdPickingStructure
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Get_DestSubLoc(DJ, ERPLoginData, Header)
    End Function

    <WebMethod()> _
    Public Function Get_ChangeExcelFlag(ByVal ERPLoginData As ERPLogin) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Get_ChangeExcelFlag(ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function GetOrderInfoFromERP(ByVal SAPLoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetOrderInfoFromERP(SAPLoginData, OrderNo)
    End Function

    <WebMethod()> _
    Public Function GetDJInfoFromERP(ByVal SAPLoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetDJInfoFromERP(SAPLoginData, OrderNo)
    End Function

    <WebMethod()> _
    Public Function GetOrderInfoFromETRACE(ByVal SAPLoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetOrderInfoFromETRACE(SAPLoginData, OrderNo)
    End Function

    <WebMethod()> _
    Public Function GetIDInfo(ByVal SAPLoginData As ERPLogin, ByVal ID As String) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetIDInfo(SAPLoginData, ID)
    End Function

    <WebMethod()> _
    Public Function GetDJ(ByVal SAPLoginData As ERPLogin, ByVal ID As String) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetDJ(SAPLoginData, ID)
    End Function

    <WebMethod()> _
    Public Function Check_HW_Flag(ByVal ERPLoginData As ERPLogin) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Check_HW_Flag(ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function Flag_AutoGetBOMRev(ByVal ERPLoginData As ERPLogin) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Flag_AutoGetBOMRev(ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function Post_DJ_Completion(ByVal ERPLoginData As ERPLogin, ByVal DJ As String, ByVal com_qty As Integer, ByVal Inventory As String, ByVal Locator As String, ByVal ds As DataSet, ByVal Productlabel As LabelData, ByVal Flag As Integer) As dj_response
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.Post_DJ_Completion(ERPLoginData, DJ, com_qty, Inventory, Locator, ds, Productlabel, Flag)
    End Function

    <WebMethod()> _
    Public Function get_uploadinfor(ByVal upload As DataSet, ByVal mydataset As DataSet, ByVal header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.get_uploadinfor(upload, mydataset, header, ERPLoginData)
    End Function

    <WebMethod()> _
    Public Function DJ_Completion(ByVal ERPLoginData As ERPLogin, ByVal DJ As String, ByVal com_qty As Integer, ByVal Inventory As String, ByVal Locator As String) As dj_response
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.DJ_Completion(ERPLoginData, DJ, com_qty, Inventory, Locator)
    End Function

    <WebMethod()> _
    Public Function DJ_Completion_BoxID(ByVal ERPLoginData As ERPLogin, ByVal BoxID As String, ByVal com_qty As Integer, ByVal Inventory As String, ByVal Locator As String) As dj_response
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.DJ_Completion_BoxID(ERPLoginData, BoxID, com_qty, Inventory, Locator)
    End Function

    <WebMethod()> _
    Public Function DJCompletion_For_PMJob(ByVal ERPLoginData As ERPLogin, ByVal DJ As String, ByVal dj_rev As String, ByVal Assembly As String, ByVal com_qty As String, ByVal uom As String, ByVal SubInventory As String, ByVal Locator As String, ByVal ROHS As String, ByVal MSB As String, ByVal item_desc As String, ByVal LabelPrinter As String) As String
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.DJCompletion_For_PMJob(ERPLoginData, DJ, dj_rev, Assembly, com_qty, uom, SubInventory, Locator, ROHS, MSB, item_desc, LabelPrinter)
    End Function

    <WebMethod()> _
    Public Function CountOfBoxID(ByVal SAPLoginData As ERPLogin, ByVal ID As String) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.CountOfBoxID(SAPLoginData, ID)
    End Function

    <WebMethod()> _
    Public Function GetSOLine(ByVal LoginData As ERPLogin, ByVal SONo As String, ByVal SOLine As String) As dj_response
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.GetSOLine(LoginData, SONo, SOLine)
    End Function

    <WebMethod()> _
    Public Function Getrelease_lines(ByVal ERPLoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.Getrelease_lines(ERPLoginData, OrderNo)
    End Function

    <WebMethod()> _
    Public Function writeClidRefPo(ByVal listItem As DataSet, ByVal DjNo As String, ByVal userName As String) As Boolean
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.writeClidRefPo(listItem, DjNo, userName)
    End Function

    <WebMethod()> _
    Public Function AllowEditDJ(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.AllowEditDJ(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function AllowDiffDJ(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.AllowDiffDJ(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function DJCmp_GetBoxInfo(ByVal ID As String, ByVal OracleLoginData As ERPLogin) As dj_response
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.DJCmp_GetBoxInfo(ID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckDJforCLID(ByVal CLID As String, ByVal DJ As String, ByVal Qty As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_materialreplenishment As New MaterialReplemenishment()
        Return _eTrace_materialreplenishment.CheckDJforCLID(CLID, DJ, Qty, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CreateCLIDforPo(ByVal intUnit As Integer) As DataSet
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.CreateCLIDforPo(intUnit)
    End Function

    <WebMethod()> _
    Public Function CountOfPoCLID(ByVal PO As String) As Double
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.CountOfPoCLID(PO)
    End Function

    <WebMethod()> _
    Public Function SaveCreateCLID(ByVal ds As DataSet, ByVal SAPLoginData As ERPLogin, ByVal productlabel As LabelData, ByVal Flag As Integer)
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.SaveCreateCLID(ds, SAPLoginData, productlabel, Flag)
    End Function

    <WebMethod()> _
    Public Function SaveCreatCLIDforPo(ByVal ds As DataSet, ByVal SAPLoginData As ERPLogin, ByVal productlabel As LabelData, ByVal Flag As Integer)
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.SaveCreatCLIDforPo(ds, SAPLoginData, productlabel, Flag)
    End Function

    <WebMethod()> _
    Public Function RePrintCreateCLIDforPOCLIDs(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean
        Dim _eTrace_MaterialReplenishment As New MaterialReplemenishment
        Return _eTrace_MaterialReplenishment.RePrintCreateCLIDforPOCLIDs(CLIDs, Printer)
    End Function

#End Region


#Region "ProdPicking"
    <WebMethod()>
    Public Function CheckCHDJ(ByVal DJ As String, ByVal CLID As String) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.CheckCHDJ(DJ, CLID)
    End Function

    <WebMethod()> _
    Public Function ReadCLIDs(ByVal CLID As String, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ReadCLIDs(CLID, LoginData)
    End Function

    <WebMethod()> _
    Public Function ReadCLIDData(ByVal CLID As String, ByVal TransactionType As String, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ReadCLIDData(CLID, TransactionType, LoginData)
    End Function
    <WebMethod()> _
    Public Function GetMOFromDJ(ByVal DJ As String, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.GetMOFromDJ(DJ, LoginData)
    End Function

    <WebMethod()> _
    Public Function GetDJMOLines(ByVal DJ As String, ByVal MO As String, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.GetDJMOLines(DJ, MO, LoginData)
    End Function

    <WebMethod()> _
    Public Function MOrderPost(ByVal LoginData As ERPLogin, ByRef myMOData As MOData) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.MOrderPost(LoginData, myMOData)
    End Function

    <WebMethod()> _
    Public Function MOEarlyLotPost(ByVal LoginData As ERPLogin, ByVal myMOData As MOData) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.MOEarlyLotPost(LoginData, myMOData)
    End Function

    <WebMethod()> _
    Public Function MOSpecialPick(ByVal LoginData As ERPLogin, ByRef myMOData As MOData) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.MOSpecialPick(LoginData, myMOData)
    End Function

    <WebMethod()> _
    Public Function MOSpecialPost(ByVal LoginData As ERPLogin, ByVal myMOData As MOData) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.MOSpecialPost(LoginData, myMOData)
    End Function

    <WebMethod()> _
    Public Function ReadPOrderList(ByVal ProdFloor As String, ByVal Open As Boolean, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ReadPOrderList(ProdFloor, Open, LoginData)
    End Function

    <WebMethod()> _
    Public Function ReadPOrderItems(ByVal PickOrder As String, ByVal ShowAllMatls As Boolean, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ReadPOrderItems(PickOrder, ShowAllMatls, LoginData)
    End Function

    <WebMethod()> _
    Public Function UpdatePOrderHeader(ByVal PickOrder As String, ByVal Status As String, ByVal LoginData As ERPLogin) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.UpdatePOrderHeader(PickOrder, Status, LoginData)
    End Function

    <WebMethod()> _
    Public Function PickOrderPost(ByVal Items As DataSet, ByVal LoginData As ERPLogin) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.PickOrderPost(Items, LoginData)
    End Function

    <WebMethod()> _
    Public Function GetMOFromDN(ByVal DN As String, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.GetMOFromDN(DN, LoginData)
    End Function

    <WebMethod()> _
    Public Function GetDNMOLines(ByVal DN As String, ByVal MO As String, ByVal PickStatus As String, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.GetDNMOLines(DN, MO, PickStatus, LoginData)
    End Function

    <WebMethod()> _
    Public Function ShipmentPost(ByVal Items As DataSet, ByVal LoginData As ERPLogin) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ShipmentPost(Items, LoginData)
    End Function

    <WebMethod()> _
    Public Function GetDNPickedLists(ByVal LoginData As ERPLogin, ByVal DN As String) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.GetDNPickedLists(LoginData, DN)
    End Function

    <WebMethod()> _
    Public Function GetDNPickedCLIDs(ByVal LoginData As ERPLogin, ByVal DN As String, ByVal CLID As String) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.GetDNPickedCLIDs(LoginData, DN, CLID)
    End Function

    <WebMethod()> _
    Public Function SOReversalPost(ByVal Items As DataSet, ByVal LoginData As ERPLogin) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.SOReversalPost(Items, LoginData)
    End Function

    <WebMethod()> _
    Public Function PrintProcessLabel(ByVal LoginData As ERPLogin, ByVal ds As DataSet, ByVal Printer As String) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.PrintProcessLabel(LoginData, ds, Printer)
    End Function

    <WebMethod()> _
    Public Function ReadProcessLabel(ByVal LoginData As ERPLogin, ByVal LabelID As String) As DataSet
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ReadProcessLabel(LoginData, LabelID)
    End Function

    <WebMethod()> _
    Public Function PCBSlotLightOn(ByVal LoginData As ERPLogin, ByRef myMOData As MOData) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.PCBSlotLightOn(LoginData, myMOData)
    End Function

#End Region

#Region "eTrace - SLED"
    <WebMethod()> _
    Public Function GetMatList(ByVal Selection As SelCriteria) As DataSet
        Dim _eTraceOracleSLED As New SLED()
        Return _eTraceOracleSLED.GetMatList(Selection)
    End Function

    <WebMethod()> _
    Public Function SLEDSaveChanges(ByVal OracleLoginData As ERPLogin, ByVal CLIDs As DataSet) As PostSLEDResponse
        Dim _eTraceOracleSLED As New SLED()
        Return _eTraceOracleSLED.SLEDSaveChanges(OracleLoginData, CLIDs)
    End Function

    <WebMethod()> _
    Public Function SLED_ReviewUpd(ByVal p_ds As DataSet, ByVal OracleLoginData As String, ByVal TransactionID As Integer) As DataSet
        Dim _eTraceOracleSLED As New SLED()
        Return _eTraceOracleSLED.SLED_ReviewUpd(p_ds, OracleLoginData, TransactionID)
    End Function

    <WebMethod()> _
    Public Function SLED_FailUpd(ByVal p_ds_fail As DataSet, ByVal OracleLoginData As String, ByVal TransactionID As Integer) As DataSet
        Dim _eTraceOracleSLED As New SLED()
        Return _eTraceOracleSLED.SLED_FailUpd(p_ds_fail, OracleLoginData, TransactionID)
    End Function

    <WebMethod()> _
    Public Function eTrace_Upd(ByVal UpdtTable As DataSet, ByVal OracleLoginData As String) As PostSLEDResponse
        Dim _eTraceOracleSLED As New SLED()
        Return _eTraceOracleSLED.eTrace_Upd(UpdtTable, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function SLEDSaveTrace(ByVal UpdtTable As DataSet, ByVal OracleLoginData As String) As PostSLEDResponse
        Dim _eTraceOracleSLED As New SLED()
        Return _eTraceOracleSLED.SLEDSaveTrace(UpdtTable, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function eTraceUpd_Review(ByVal UpdtTable As DataSet, ByVal OracleLoginData As String) As String
        Dim _eTraceOracleSLED As New SLED()
        Return _eTraceOracleSLED.eTraceUpd_Review(UpdtTable, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function eTraceUpd_Fail(ByVal UpdtTable As DataSet, ByVal OracleLoginData As String) As String
        Dim _eTraceOracleSLED As New SLED()
        Return _eTraceOracleSLED.eTraceUpd_Fail(UpdtTable, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function eTraceUpd_Pass(ByVal UpdtTable As DataSet, ByVal OracleLoginData As String) As String
        Dim _eTraceOracleSLED As New SLED()
        Return _eTraceOracleSLED.eTraceUpd_Pass(UpdtTable, OracleLoginData)
    End Function
#End Region

#Region "Shipment"

    <WebMethod()> _
    Public Function GetModelRevision(ByVal PalletID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleShipment As New Shipment()
        Return _eTraceOracleShipment.GetModelRevision(PalletID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetModelRevDJ_WithType(ByVal PalletID As String, ByVal Type As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleShipment As New Shipment()
        Return _eTraceOracleShipment.GetModelRevDJ_WithType(PalletID, Type, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetModelRevDJ(ByVal PalletID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleShipment As New Shipment()
        Return _eTraceOracleShipment.GetModelRevDJ(PalletID, OracleLoginData)
    End Function


    '<WebMethod()> _
    'Public Function GetModelRevForTE(ByVal IntSN As String) As String
    '    Dim _eTraceOracleShipment As New Shipment()
    '    Return _eTraceOracleShipment.GetModelRevForTE(IntSN)
    'End Function

    '<WebMethod()> _
    'Public Function GetProdLineAndResultForTE(ByVal IntSN As String, ByVal Process As String) As String
    '    Dim _eTraceOracleShipment As New Shipment()
    '    Return _eTraceOracleShipment.GetProdLineAndResultForTE(IntSN, Process)
    'End Function

    <WebMethod()> _
    Public Function GetSNList(ByVal PalletID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleShipment As New Shipment()
        Return _eTraceOracleShipment.GetSNList(PalletID, OracleLoginData)
    End Function
    <WebMethod()> _
    Public Function ChangeModel(ByVal BoxID() As String, ByVal NewModel As String, ByVal NewRev As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleShipment As New Shipment()
        Return _eTraceOracleShipment.ChangeModel(BoxID, NewModel, NewRev, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function Change_Model(ByVal BoxID() As String, ByVal NewModel As String, ByVal NewRev As String, ByVal NewDJ As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleShipment As New Shipment()
        Return _eTraceOracleShipment.Change_Model(BoxID, NewModel, NewRev, NewDJ, OracleLoginData)
    End Function
#End Region

#Region "Cycle Count"
    <WebMethod()> _
    Public Function GetSMTCCName(ByVal OrgCode As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.GetSMTCCName(OrgCode, OracleLoginData)
    End Function
    <WebMethod()>
    Public Function SMTCycleCountHH(ByVal EventID As String, ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.SMTCycleCountHH(EventID, CLID, OracleLoginData)
    End Function
    <WebMethod()> _
    Public Function SMTCycleCountCLIDValid(ByVal EventID As String, ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.SMTCycleCountCLIDValid(EventID, CLID, OracleLoginData)
    End Function
    <WebMethod()> _
    Public Function SMTCCSave(ByVal EventID As String, ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.SMTCCSave(EventID, p_ds, OracleLoginData)
    End Function
    <WebMethod()> _
    Public Function SMTCLIDStatusChanged(ByRef EventID As String, ByRef CLID As String, ByRef Action As String, ByRef ResultFlag As String) As String
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.SMTCLIDStatusChanged(EventID, CLID, Action, ResultFlag)
    End Function
    <WebMethod()> _
    Public Function SMTCLIDStatusChangedByPC(ByRef EventID As String, ByRef CLID As String, ByRef Action As String, ByRef ResultFlag As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.SMTCLIDStatusChangedByPC(EventID, CLID, Action, ResultFlag, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetSMTCCList(ByVal OrgCode As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.GetSMTCCList(OrgCode, OracleLoginData)
    End Function
    <WebMethod()> _
    Public Function GetSMTScanedCLID(ByVal EventID As String, ByVal Seq As Integer, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.GetSMTScanedCLID(EventID, Seq, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetCyDate(ByVal cc_name As String, ByVal seq As Integer, ByVal OracleLoginData As ERPLogin) As GetCycleCount
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.GetCyDate(cc_name, seq, OracleLoginData)
    End Function

    ''<WebMethod()> _
    ''    Public Function PostCycle(ByVal cc_name As String, ByVal seq As Integer) As DataSet
    ''    Dim _eTraceOracleCycleCount As New CycleCount()
    ''    Return _eTraceOracleCycleCount.PostCycle(cc_name, seq)
    ''End Function

    <WebMethod()> _
    Public Function UpdateCycle(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.UpdateCycle(p_ds, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetCLIDate(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.GetCLIDate(CLID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function TBoxData(ByVal BoxID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.TBoxData(BoxID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function Testupdata(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.UpdateCycle(p_ds, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function UpdateReason(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTraceOracleCycleCount As New CycleCount()
        Return _eTraceOracleCycleCount.UpdateReason(p_ds, OracleLoginData)
    End Function

#End Region

#Region "InventoryOptimization"

    <WebMethod()> _
    Public Function GetMRPData(ByVal myInputData As InputData) As DataSet
        Dim _eTrace_InventoryOptimization As New InventoryOptimization()
        Return _eTrace_InventoryOptimization.GetMRPData(myInputData)
    End Function

#End Region

#Region "SIT_MassPrintCLIDs"
    <WebMethod()> _
    Public Function SIT_MassPrintCLIDs(ByVal Printer As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.SIT_MassPrintCLIDs(Printer)
    End Function

    <WebMethod()> _
    Public Function SIT_DeleteCLIDs() As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.SIT_DeleteCLIDs()
    End Function

    <WebMethod()> _
    Public Function SIT_Print_OnHand(ByVal Printer As String) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.SIT_Print_OnHand(Printer)
    End Function

#End Region

#Region "TestDataCollection"
    '<WebMethod()> _
    'Public Function OpenPrintLabels(ByVal boxid As String, ByVal printer As String) As Boolean
    '    Dim _eTrace_TDCFunction As New TestDataCollection()
    '    Return _eTrace_TDCFunction.OpenPrintLabels(boxid, printer)
    'End Function
    <WebMethod()>
    Public Function UploadProductRoutingLog(ByVal productRouting As DataTable, ByVal userID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadProductRoutingLog(productRouting, userID)
    End Function
    <WebMethod()>
    Public Function UploadProductChangeLog(ByVal LogTable As DataTable, ByVal userID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadProductChangeLog(LogTable, userID)
    End Function
    <WebMethod()> _
    Public Function CopyDJInfo(ByVal orgcode As String, ByVal isdelOldData As Integer) As Boolean
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.CopyDJInfo(orgcode, isdelOldData)
    End Function
    <WebMethod()> _
    Public Function CopyDJinfoUseServerName() As Boolean
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.CopyDJinfoUseServerName()
    End Function

    <WebMethod()> _
    Public Function isValidDJ(ByVal dj_name As String, ByVal Status As String, ByVal OracleERPLogin As ERPLogin) As Boolean
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.isValidDJ(dj_name, Status, OracleERPLogin)
    End Function
    <WebMethod()> _
    Public Function insertTDHeader(ByVal start As Integer, ByVal endindex As Integer) As Boolean
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.insertTDHeader(start, endindex)
    End Function

    <WebMethod()> _
    Public Function ReadRuleFromDB() As DataSet
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.ReadRuleFromDB()
    End Function

    <WebMethod()> _
    Public Function postProdMaster(ByVal pordMaster As RuleDetail, ByVal ProdStructure As DataSet, ByVal SAPLoginData As ERPLogin) As String
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.postProdMaster(pordMaster, ProdStructure, SAPLoginData)
    End Function

    '<WebMethod()> _
    'Private Function postProdStructure(ByVal ProdStructure As DataSet, ByVal Model As String) As String
    '    Dim _eTrace_TDCFunction As New TestDataCollection()
    '    Return _eTrace_TDCFunction.postProdStructure(ProdStructure, Model)

    'End Function

    <WebMethod()> _
    Public Function DelProdMaster(ByVal Model As String) As Boolean
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.DelProdMaster(Model)

    End Function
    '    <WebMethod()> _
    'Public Function CheckBuildPanFormat(ByVal InvID As String, ByVal MRListData As DataSet, ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
    '        Dim _eTrace_TDCFunction As New TestDataCollection()
    '        Return _eTrace_TDCFunction.CheckFormat(InvID, MRListData, ExcelData, OracleERPLogin)
    '    End Function
    '    <WebMethod()> _
    '    Public Function CommitBuildPlan(ByVal MRListData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
    '        Dim _eTrace_TDCFunction As New TestDataCollection()
    '        Return _eTrace_TDCFunction.CommitBuildPlan(MRListData, OracleERPLogin)
    '    End Function

    <WebMethod()> _
    Public Function Getwipflowdata(ByVal intSN As String, ByVal ProcessName As String) As DataSet
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.Getwipflowdata(intSN, ProcessName)
    End Function

    <WebMethod()> _
    Public Function GetModelRevForTE(ByVal IntSN As String) As String
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.GetModelRevForTE(IntSN)
    End Function

    <WebMethod()> _
    Public Function GetProdLineAndResultForTE(ByVal IntSN As String, ByVal Process As String) As String
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.GetProdLineAndResultForTE(IntSN, Process)
    End Function

    <WebMethod()>
    Public Function CheckFixutureID(ByVal FixtureID As String) As String
        Dim _eTrace_TDCFunction As New TestDataCollection
        Return _eTrace_TDCFunction.CheckFixtureID(FixtureID)
    End Function

    <WebMethod()>
    Public Function GetWipHeaderByIntSN(ByVal IntSN As String) As DataSet
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.GetWipHeaderByIntSN(IntSN)
    End Function
    <WebMethod()>
    Public Function GetPanelIDByIntSN(ByVal intSN As String) As String
        Dim _eTrace_TDCFunction As New TestDataCollection
        Return _eTrace_TDCFunction.GetPanelIDByIntSN(intSN)
    End Function
    ''' <summary>
    ''' if success returns true,fail return error message
    ''' </summary>
    ''' <param name="FixtureID"></param>
    ''' <returns></returns>
    <WebMethod()>
    Public Function ClearFixtureID(ByVal FixtureID As String) As String
        Dim _eTrace_TDCFunction As New TestDataCollection
        Return _eTrace_TDCFunction.ClearFixtureID(FixtureID)
    End Function
    ''' <summary>
    ''' Save the fixtureMount info
    ''' </summary>
    ''' <param name="DsFixture">dataset</param>
    ''' <returns></returns>
    <WebMethod()>
    Public Function FixtureMount(ByVal DsFixture As DataSet, ByVal FixuteID As String, ByVal user As String) As String
        Dim _eTrace_TDCFunction As New TestDataCollection
        Return _eTrace_TDCFunction.FixtureMount(DsFixture, FixuteID, user)
    End Function

    <WebMethod()>
    Public Function TE_IntSNListReadbyFixtureID(ByVal FixtureID As String) As String
        Dim _eTrace_TDCFunction As New TestDataCollection
        Return _eTrace_TDCFunction.TE_IntSNListReadbyFixtureID(FixtureID)
    End Function

    <WebMethod()>
    Public Function TE_ReworkUnitFlag(ByVal IntSN As String) As String
        Dim _eTrace_TDCFunction As New TestDataCollection
        Return _eTrace_TDCFunction.TE_ReworkUnitFlag(IntSN)
    End Function
    <WebMethod()>
    Public Function TE_UnbindFixtureID(ByVal FixtureID As String) As String
        Dim _eTrace_TDCFunction As New TestDataCollection
        Return _eTrace_TDCFunction.TE_UnbindFixtureID(FixtureID)
    End Function

#End Region

#Region "Kanban"

    <WebMethod()> _
    Public Function ChangeExpdateUser(ByVal UserName As String) As Boolean
        Dim objKanban As New Kanban()
        Return objKanban.ChangeExpdateUser(UserName)
    End Function

    <WebMethod()> _
    Public Function BatchChangeExpdate(ByVal BatchID As String, ByVal OracleERPLogin As ERPLogin) As Boolean
        Dim objKanban As New Kanban()
        Return objKanban.BatchChangeExpdate(BatchID, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function CheckCLIDExpFormat(ByVal BatchID As String, ByVal mydsCLID As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim objKanban As New Kanban()
        Return objKanban.CheckCLIDExpFormat(BatchID, mydsCLID, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function getKanbanIDinfo(ByVal LoginData As ERPLogin, ByVal iKanbanID As String) As DataSet
        Dim objKanban As New Kanban()
        Return objKanban.getKanbanIDinfo(LoginData, iKanbanID)
    End Function

    <WebMethod()> _
    Public Function UpdateKanbanCodeDS(ByVal p_OracleLoginData As ERPLogin, ByVal iKanbanID As String, ByVal p_needbydate As String) As String
        Dim objKanban As New Kanban()
        Return objKanban.UpdateKanbanCodeDS(p_OracleLoginData, iKanbanID, p_needbydate)
    End Function

    <WebMethod()> _
    Public Function CheckPNFormat(ByVal InvID As String, ByVal MRListData As DataSet, ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.CheckPNFormat(InvID, MRListData, ExcelData, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function CheckBuildPanFormat(ByVal InvID As String, ByVal MRListData As DataSet, ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.CheckFormat(InvID, MRListData, ExcelData, OracleERPLogin)
    End Function
    <WebMethod()> _
    Public Function CommitBuildPlan(ByVal MRListData As DataSet, ByVal OracleERPLogin As ERPLogin, ByVal uploadType As Integer, ByVal saType As Integer, ByVal triggerType As String) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.CommitBuildPlan(MRListData, OracleERPLogin, uploadType, saType, triggerType)
    End Function
    <WebMethod()> _
    Public Function SaveKanbanData(ByVal MRListData As DataSet, ByVal OracleERPLogin As ERPLogin) As String
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.SaveKanbanData(MRListData, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function uploadQuotaSharing(ByVal exceldata As DataSet) As String
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.uploadQuotaSharing(exceldata)
    End Function
    <WebMethod()> _
    Public Function uploadTransmission(ByVal exceldata As DataSet) As String
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.uploadTransmission(exceldata)
    End Function
    <WebMethod()> _
    Public Function getQuotaSharing(ByVal LoginData As ERPLogin, ByVal ItemNo As String, ByVal BuyerName As String) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getQuotaSharing(LoginData, ItemNo, BuyerName)
    End Function

    <WebMethod()> _
    Public Function getTransmission(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getTransmission(LoginData)
    End Function

    <WebMethod()> _
    Public Function GetNextBPID(ByVal LoginData As ERPLogin) As String
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.GetNextBPID(LoginData)
    End Function

    <WebMethod()> _
    Public Function getBuildPlanList(ByVal startDate As String, ByVal endDate As String, ByVal productionfloor As String, ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getBuildPlanList(startDate, endDate, productionfloor, p_ds, LoginData)
    End Function

    <WebMethod()> _
    Public Function getBuildPlandetail(ByVal buildplanid As String, ByVal ischecked As Boolean, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getBuildPlanDetail(buildplanid, ischecked, LoginData)
    End Function

    <WebMethod()> _
    Public Function getLocatorsPB(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getLocatorsPB(LoginData)
    End Function

    <WebMethod()> _
    Public Function getNow() As Date
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getNow()
    End Function

    <WebMethod()> _
    Public Function getPONumber(ByVal buildplanDetail As DataSet, ByVal ERPLogin As ERPLogin) As Boolean
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getPONumber(buildplanDetail, ERPLogin)
    End Function

    <WebMethod()> _
    Public Function checkPermission(ByVal ERPLogin As ERPLogin) As Boolean
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.checkPermission(ERPLogin)
    End Function

    <WebMethod()> _
    Public Function checkTMFormat(ByVal MRListData As DataSet, ByVal ExcelMRData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.checkTMFormat(MRListData, ExcelMRData, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function checkQSFormat(ByVal MRListData As DataSet, ByVal ExcelMRData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.checkQSFormat(MRListData, ExcelMRData, OracleERPLogin)
    End Function

    <WebMethod()> _
    Public Function getExceptionReport(ByVal LoginData As ERPLogin, ByVal report As DataSet) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getExceptionReport(LoginData, report)
    End Function

    <WebMethod()> _
    Public Function SaveBuildPlanData(ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As Boolean
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.SaveData(p_ds, LoginData)
    End Function

    <WebMethod()> _
    Public Function SkipLot(ByVal oracleLoginData As ERPLogin, ByVal buildPlanId As Int64) As Boolean
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.SkipLot(oracleLoginData, buildPlanId)
    End Function
    <WebMethod()> _
    Public Function checkIPPFormat(ByVal MRListData As DataSet, ByVal ExcelMRData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.checkIPPFormat(MRListData, ExcelMRData, OracleERPLogin)
    End Function
    <WebMethod()> _
    Public Function Upload_EJITIPP(ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As Boolean
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.Upload_EJITIPP(p_ds, LoginData)
    End Function
    <WebMethod()> _
    Public Function getProdFloor(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getProdFloor(LoginData)
    End Function
    <WebMethod()> _
    Public Function getEjitIPPList(ByVal startDate As String, ByVal endDate As String, ByVal productionfloor As String, ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getEjitIPPList(startDate, endDate, productionfloor, p_ds, LoginData)
    End Function
    <WebMethod()> _
    Public Function getIPPDetail(ByVal IPPId As String, ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_KanbanFunction As New Kanban()
        Return _eTrace_KanbanFunction.getIPPDetail(IPPId, LoginData)
    End Function


#End Region

#Region "Report"
    <WebMethod()> _
    Public Function Report_TCLIDData(ByVal CLID As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.TCLIDData(CLID)
    End Function

    <WebMethod()> _
    Public Function Report_TItemData(ByVal Item As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.TItemData(Item)
    End Function

    <WebMethod()> _
    Public Function Report_TAging(ByVal Org As String, ByVal SubINV As String, ByVal Item As String, ByVal Comparison As ComparisonSign, ByVal AgingDays As Integer) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.TAging(Org, SubINV, Item, Comparison, AgingDays)
    End Function

    <WebMethod()> _
    Public Function Report_TDJData(ByVal Org As String, ByVal CLID As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.TDJData(Org, CLID)
    End Function

    <WebMethod()> _
    Public Function Report_TCLIDIssueData(ByVal CLID As String, ByVal DJ As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.TCLIDIssueData(CLID, DJ)
    End Function

    <WebMethod()> _
    Public Function Report_TCLIDInfo(ByVal CLID As String) As String
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.TCLIDInfo(CLID)
    End Function

    <WebMethod()> _
    Public Function Report_AMLIProVSeTrace(ByVal Org As String, ByVal SubInv As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.IProAMLVSeTrace(Org, SubInv)
    End Function

    <WebMethod()> _
    Public Function Report_AMLIProVSeTrace2(ByVal strOrgCode As String, ByVal strSubINV As String, ByVal rtDateFrom As String, ByVal rtDateTo As String, ByVal amlStatus As String, ByVal clidStatus As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.IProAMLVSeTrace2(strOrgCode, strSubINV, rtDateFrom, rtDateTo, amlStatus, clidStatus)
    End Function

    <WebMethod()> _
    Public Function Report_TCLIDMSLData(ByVal CLId As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.TCLIDMSLData(CLId)
    End Function

    <WebMethod()> _
    Public Function Report_MPNOnHand(ByVal OrgId As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.MPNOnHand(OrgId)
    End Function

    <WebMethod()> _
    Public Function Report_OnHandMFGMPN(ByVal Org As String, ByVal Material As String, ByVal SubInv As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.OnHandMFGMPN(Org, Material, SubInv)
    End Function

    <WebMethod()> _
    Public Function Report_TIssueCompare(ByVal DJ As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.TIssueCompare(DJ)
    End Function

    <WebMethod()> _
    Public Function Report_TMaterialTransfer(ByVal Org As String, ByVal MaterialNo As String, ByVal MPN As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.TMaterialTransfer(Org, MaterialNo, MPN)
    End Function

    <WebMethod()> _
    Public Function Report_GeteTraceOH(ByVal Org As String, ByVal ItemNo As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GeteTraceOH(Org, ItemNo)
    End Function

    <WebMethod()> _
    Public Function Report_StandardTime() As String
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.StandardTime
    End Function

    <WebMethod()> _
    Public Function Report_GetOHQTYWithMPNList(ByVal Org As String, ByVal MPNlist As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GetOHQTYWithMPNList(Org, MPNlist)
    End Function

    <WebMethod()> _
    Public Function Report_GeteTraceItemOHMPQ(ByVal Org As String, ByVal Itemlist As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GeteTraceItemOHMPQ(Org, Itemlist)
    End Function
    <WebMethod()> _
    Public Function Report_GetConnectString(ByVal Type As String) As String
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GetConnectString(Type)
    End Function


    'External Function Called by SMT and other detartments
    <WebMethod()> _
    Public Function GetMaterialInfoByCLID(ByVal CLID As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GetMaterialInfoByCLID(CLID)
    End Function

    <WebMethod()> _
    Public Function GetMaterialToXMLByCLID(ByVal From_MR As DateTime, ByVal To_MR As DateTime, ByVal pn As String) As String
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GetMaterialToXMLByCLID(From_MR, To_MR, pn)
    End Function

    <WebMethod()> _
    Public Function GetMaterialToXMLByCLID2(ByVal From_MR As DateTime, ByVal To_MR As DateTime, ByVal pn As String) As String
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GetMaterialToXMLByCLID2(From_MR, To_MR, pn)
    End Function

    <WebMethod()> _
    Public Function GetDJToXML() As String
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GetDJToXML()
    End Function

    <WebMethod()> _
    Public Function GetMaterialByCLIDToDataSet(ByVal From_MR As DateTime, ByVal To_MR As DateTime, ByVal pn As String) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GetMaterialByCLIDToDataSet(From_MR, To_MR, pn)
    End Function

    <WebMethod()> _
    Public Function GetReturnCLIDByDate(ByVal From_RTN As DateTime, ByVal To_RTN As DateTime) As DataSet
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.GetReturnCLIDByDate(From_RTN, To_RTN)
    End Function

    <WebMethod()> _
    Public Function ValidReqlineStatus(ByVal p_orgcode As String, ByVal p_dnpo As String, ByVal p_input_type As String, ByVal p_ejit_id As Integer) As String
        Dim _eTraceOracle_Report As New Report
        Return _eTraceOracle_Report.ValidReqlineStatus(p_orgcode, p_dnpo, p_input_type, p_ejit_id)
    End Function
#End Region

#Region "MI Kanban"
    <WebMethod()> _
    Public Function getMatlByDJModel(ByVal DJ As String, ByVal PartNo As String) As DataSet
        Dim MIKBS As MIkanban = New MIkanban()
        Return MIKBS.getMatlByDJModel(DJ, PartNo)
    End Function

    <WebMethod()> _
    Public Function GetMatlByCLIDDJ(ByVal CLID As String, ByVal DJ As String) As DataSet
        Dim MIKBS As MIkanban = New MIkanban()
        Return MIKBS.GetMatlByCLIDDJ(CLID, DJ)
    End Function

    <WebMethod()> _
    Public Function getMatlCLIDbyDJ(ByVal DJ As String) As DataSet
        Dim MIKBS As MIkanban = New MIkanban()
        Return MIKBS.getMatlCLIDbyDJ(DJ)
    End Function

    <WebMethod()> _
    Public Function PrintLabels(ByVal Printer As String, ByVal labelFile As String, ByVal strContent As String) As String
        Dim MIKBS As New MIkanban
        Return MIKBS.PrintLabels(Printer, labelFile, strContent)
    End Function

    <WebMethod()> _
    Public Function SaveEMCResult(ByVal sn As String, ByVal processname As String, ByVal result As String, ByVal userid As String) As String
        Dim MIKBS As New MIkanban
        Return MIKBS.SaveEMCResult(sn, processname, result, userid)
    End Function

#End Region

#Region "Mac Address"

    <WebMethod()> _
    Public Function PrintMacAddress(ByVal macType As String, ByVal addressTotal As Integer, _
            ByVal user As String, ByVal printerName As String) As String
        Dim _eTranceOracle_Mac As New TestDataCollection
        Return _eTranceOracle_Mac.PrintMacAddress(macType, addressTotal, user, printerName)
    End Function

    <WebMethod()> _
    Public Function ReprintMacAddress(ByVal macType As String, ByVal startAddress As String, _
            ByVal endAddress As String, ByVal user As String, ByVal printerName As String) As String
        Dim _eTranceOracle_Mac As New TestDataCollection
        Return _eTranceOracle_Mac.ReprintMacAddress(macType, startAddress, endAddress, user, printerName)
    End Function

    <WebMethod()> _
    Public Function GetPreMacAddress(ByVal AddressType As String) As String
        Dim _eTranceOracle_Mac As New TestDataCollection
        Return _eTranceOracle_Mac.GetPreMacAddress(AddressType)
    End Function

    <WebMethod()> _
    Public Function GetLengthMacAddress(ByVal Category As String) As String
        Dim _eTranceOracle_Mac As New TestDataCollection
        Return _eTranceOracle_Mac.GetLengthMacAddress(Category)
    End Function



#End Region

#Region "Traveller Printing"

    <WebMethod()> _
    Public Function GetTraveller(ByVal model As String) As DataTable
        Dim traveller As New TestDataCollection
        Return traveller.GetTraveller(model)
    End Function

    <WebMethod()> _
    Public Function GetAllTravellerData() As DataTable
        Dim traveller As New TestDataCollection
        Return traveller.GetAllTravellerData()
    End Function


    <WebMethod()> _
    Public Function UpdateTravellerInfo(ByVal ds As DataSet) As String
        Dim traveller As New TestDataCollection
        Return traveller.UpdateTravellerInfo(ds)
    End Function

    <WebMethod()> _
    Public Function GetAllModels() As DataTable
        Dim traveller As New TestDataCollection
        Return traveller.GetAllModels()
    End Function


    <WebMethod()> _
    Public Function DeleteModel(ByVal model As String)
        Dim traveller As New TestDataCollection
        Return traveller.DeleteModel(model)
    End Function

    <WebMethod()> _
    Public Function GetPrinterServer() As String
        Dim traveller As New TestDataCollection
        Return traveller.GetPrinterServer()
    End Function


#End Region

#Region "QC Data Collection"
    <WebMethod()> _
    Public Function GetDaughterBdByDJ(ByVal DJ As String, ByVal Flag As String) As DataSet
        Dim _TestDataCollection As New TestDataCollection()
        Return _TestDataCollection.GetDaughterBdByDJ(DJ, Flag)
    End Function

    <WebMethod()> _
    Public Function GetSMTQCByDJ(ByVal DJ As String) As DataTable
        Dim serviceObject As New TestDataCollection
        Return serviceObject.GetSMTQCByDJ(DJ)
    End Function

    <WebMethod()> _
    Public Function GetSMTQCModels(ByVal lineType As String) As DataTable
        Dim serviceObject As New TestDataCollection
        Return serviceObject.GetSMTQCModels(lineType)
    End Function

    <WebMethod()> _
    Public Function SaveSMTQCData(ByVal dtSMTQC As DataTable, ByVal ParamArray paras() As String) As String
        Dim serviceObject As New TestDataCollection
        Return serviceObject.SaveSMTQCData(dtSMTQC, paras)
    End Function

    <WebMethod()> _
    Public Function GetAllSMTQCData(ByVal ParamArray para() As String) As DataTable
        Dim serviceObject As New TestDataCollection
        Return serviceObject.GetAllSMTQCData(para)
    End Function

    <WebMethod()> _
    Public Function DeleteSMTQCData(ByVal ParamArray para() As String) As String
        Dim serviceObject As New TestDataCollection
        Return serviceObject.DeleteSMTQCData(para)
    End Function

#End Region

#Region "Shop Floor Control"

    <WebMethod()> _
    Public Function LookUpProductInfo() As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.LookUpProductInfo()
    End Function

    <WebMethod()> _
    Public Function RevinBox(ByVal BoxID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.RevinBox(BoxID)
    End Function
    <WebMethod()>
    Public Function CloneProductInfo(ByVal newModel As String, ByVal originalModel As String, ByVal userName As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CloneProductInfo(newModel, originalModel, userName)
    End Function

    <WebMethod()>
    Public Function UploadProductInfo(ByVal productModel As TestDataCollection.ProductModel) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadProductInfo(productModel)
    End Function

    <WebMethod()> _
    Public Function LookUpMSLInfo(ByVal model As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.LookUpMSLInfo(model)
    End Function

    <WebMethod()> _
    Public Function UploadProductCPN(ByVal productCPN As DataTable, ByVal model As String, ByVal changedBy As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadProductCPN(productCPN, model, changedBy)

    End Function


    <WebMethod()> _
    Public Function LookUpLabelPara(ByVal labelID As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.LookUpLabelPara(labelID)
    End Function

    <WebMethod()> _
    Public Function LookUpTransPara(ByVal transId As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.LookUpTransPara(transId)
    End Function

    <WebMethod()> _
    Public Function LookUpSFCInfoByModel(ByVal model As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.LookUpSFCInfoByModel(model)
    End Function

    <WebMethod()> _
    Public Function UploadProductStructure(ByVal productStructure As DataTable, ByVal model As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadProductStructure(productStructure, model)
    End Function

    <WebMethod()> _
    Public Function UploadProductRouting(ByVal productRouting As DataTable, ByVal model As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadProductRouting(productRouting, model)
    End Function

    <WebMethod()> _
    Public Function UploadProductProperties(ByVal productProperties As DataTable, ByVal model As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadProductProperties(productProperties, model)
    End Function

    <WebMethod()> _
    Public Function UploadLabelInfo(ByVal labelModel As TestDataCollection.LabelModel, ByVal dtLabelPara As DataTable) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadLabelInfo(labelModel, dtLabelPara)
    End Function

    <WebMethod()> _
    Public Function UploadProcessInfo(ByVal processEntity As DataTable, ByVal userName As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadProcessInfo(processEntity, userName)
    End Function

    <WebMethod()> _
    Public Function UploadTransInfo(ByVal labelArray As String(), ByVal transEntity As DataTable) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UploadTransInfo(labelArray, transEntity)
    End Function

    <WebMethod()> _
    Public Function IsAuthenticateUser(ByVal userId As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IsAuthenticateUser(userId)
    End Function

    <WebMethod()> _
    Public Function LookUpRoutingInfoBy(ByVal routingId As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.LookUpRoutingInfoBy(routingId)
    End Function

    <WebMethod()> _
    Public Function UpdateRoutingInfo(ByVal dsRouting As DataSet, ByVal ParamArray list As Object()) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UpdateRoutingInfo(dsRouting, list)
    End Function


    <WebMethod()> _
    Public Function GetProductLineInfo() As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetProductLineInfo()
    End Function

    <WebMethod()> _
    Public Function IsReworkUnit(ByVal WIPID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IsReworkUnit(WIPID)
    End Function

    <WebMethod()> _
    Public Function IsLastProcess(ByVal Model As String, ByVal PCBA As String, ByVal Process As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IsLastProcess(Model, PCBA, Process)
    End Function

    <WebMethod()> _
    Public Function IsBottomPCBA(ByVal Model As String, ByVal PCBA As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IsBottomPCBA(Model, PCBA)
    End Function

    <WebMethod()> _
    Public Function GetProdLineBy(ByVal prodLine As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetProductLineInfo(prodLine)
    End Function

    <WebMethod()> _
    Public Function UpdateProductLine(ByVal dt As DataTable, ByVal orgCode As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UpdateProductLine(dt, orgCode)
    End Function

    <WebMethod()> _
    Public Function GetEquipmentInfo() As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetEquipmentInfo()
    End Function

    <WebMethod()> _
    Public Function UpdateEquipmentInfo(ByVal dt As DataTable, ByVal ParamArray para() As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UpdateEquipmentInfo(dt, para)
    End Function

    <WebMethod()> _
    Public Function GetEquipmentDetailInfo(ByVal eqptId As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetEquipmentInfo(eqptId)
    End Function

    <WebMethod()> _
    Public Function UpdatePMdata(ByVal productModel As TestDataCollection.ProductModel, ByVal ds As DataSet) As String()
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UpdatePMdata(productModel, ds)
    End Function

    <WebMethod()> _
    Public Function ModelDefined(ByVal ModelNo As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ModelDefined(ModelNo)
    End Function

    <WebMethod()> _
    Public Function ModelStructure(ByVal ModelNo As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ModelStructure(ModelNo)
    End Function

    <WebMethod()> _
    Public Function ProdQty(ByVal DJ As String, ByVal PCBA As String, ByVal OrgCode As String) As Integer
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ProdQty(DJ, PCBA, OrgCode)
    End Function

    <WebMethod()> _
    Public Function GetDJMatchedQty(ByVal DJ As String, ByVal PCBA As String, ByVal OrgCode As String) As Integer
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetDJMatchedQty(DJ, PCBA, OrgCode)
    End Function

    <WebMethod()> _
    Public Function InsertPoQty(ByVal OrgCode As String, ByVal DJ As String, ByVal Model As String, ByVal PCBA As String, ByVal Qty As Integer) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.InsertPoQty(OrgCode, DJ, Model, PCBA, Qty)
    End Function

    <WebMethod()> _
    Public Function CountPoQty(ByVal OrgCode As String, ByVal DJ As String, ByVal Model As String, ByVal PCBA As String, ByVal Qty As Integer, ByVal ModelRev As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CountPoQty(OrgCode, DJ, Model, PCBA, Qty, ModelRev)
    End Function

    <WebMethod()> _
    Public Function CountPoQtyII(ByVal OrgCode As String, ByVal DJ As String, ByVal Model As String, ByVal PCBA As String, ByVal Qty As Integer, ByVal ModelRev As String, ByVal DJType As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CountPoQtyII(OrgCode, DJ, Model, PCBA, Qty, ModelRev, DJType)
    End Function

    <WebMethod()> _
    Public Function IntSNIsValid(ByVal IntSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IntSNIsValid(IntSN)
    End Function

    <WebMethod()> _
    Public Function ReadPOQtyByPOAndPCBA(ByVal PO As String, ByVal PCBA As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ReadPOQtyByPOAndPCBA(PO, PCBA)
    End Function

    <WebMethod()> _
    Public Function PCBARouting(ByVal ModelNo As String, ByVal PCBA As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.PCBARouting(ModelNo, PCBA)
    End Function

    <WebMethod()> _
    Public Function PCBListOfRework(ByVal WIPID As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.PCBListOfRework(WIPID)
    End Function

    <WebMethod()> _
    Public Function WIPMatching1(ByVal DSWIP As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPMatching1(DSWIP)
    End Function

    <WebMethod()> _
    Public Function WIPVisualInspection(ByVal DSWIP As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPVisualInspection(DSWIP)
    End Function

    <WebMethod()> _
    Public Function WIPBurnIn(ByVal DSWIP As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPBurnIn(DSWIP)
    End Function

    <WebMethod()> _
    Public Function WIPIDSwop(ByVal DSWIP As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPIDSwop(DSWIP)
    End Function

    <WebMethod()> _
    Public Function WIPModelSwop(ByVal DSWIP As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPModelSwop(DSWIP)
    End Function

    <WebMethod()> _
    Public Function getPCBAinWIPHeader(ByVal IntSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.getPCBAinWIPHeader(IntSN)
    End Function

    <WebMethod()> _
    Public Function MI_getPCBAinWIPHeader(ByVal IntSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.MI_getPCBAinWIPHeader(IntSN)
    End Function

    <WebMethod()> _
    Public Function DBoardIsValid(ByVal IntSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.DBoardIsValid(IntSN)
    End Function

    <WebMethod()> _
    Public Function ModelConfiguratorSNValid(ByVal SN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ModelConfiguratorSNValid(SN)
    End Function

    <WebMethod()> _
    Public Function IntSNPattern(ByVal Model As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IntSNPattern(Model)
    End Function

    <WebMethod()> _
    Public Function GetOrderInfoFromOracle(ByVal OrderNo As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetOrderInfoFromOracle(OrderNo)


    End Function
    <WebMethod()> _
    Public Function ComponentReplacement(ByVal DSWIP As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ComponentReplacement(DSWIP)
    End Function

    <WebMethod()> _
    Public Function CheckCompIssueToDJ(ByVal DJ As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CheckCompIssueToDJ(DJ)
    End Function

    <WebMethod()> _
    Public Function MatchingAccount(ByVal DJ As String, ByVal OrgCode As String, ByVal PCBA As String, ByVal Allow As Integer, ByVal PanelSize As Integer) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.MatchingAccount(DJ, OrgCode, PCBA, Allow, PanelSize)
    End Function

    <WebMethod()> _
    Public Function GetResult(ByVal SCID As String, ByVal Process As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetResult(SCID, Process)
    End Function

    <WebMethod()> _
    Public Function DJinBox(ByVal DJ As String, ByVal BoxID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.DJinBox(DJ, BoxID)
    End Function

    <WebMethod()> _
    Public Function CheckPrevResult(ByVal IntSerialNo As String, ByVal Process As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CheckPrevResult(IntSerialNo, Process)
    End Function
    <WebMethod()> _
    Public Function LargeThanMaxTest(ByVal IntSerialNo As String, ByVal Process As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.LargeThanMaxTest(IntSerialNo, Process)
    End Function

    <WebMethod()> _
    Public Function WIPIn(ByVal IntSerialNo As String, ByVal Process As String, ByVal user As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPIn(IntSerialNo, Process, user)
    End Function

    <WebMethod()> _
    Public Function WIPOut(ByVal header As StatusHeaderStructure) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPOut(header)
    End Function
    <WebMethod()> _
    Public Function checkSamplingTest(ByVal IntSN, ByVal CurrProcess) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.checkSamplingTest(IntSN, CurrProcess)
    End Function

    <WebMethod()> _
    Public Function ReadDBoards(ByVal header As StatusHeaderStructure) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ReadDBoards(header)
    End Function

    <WebMethod()> _
    Public Function WIPOutMatchingN(ByVal header As StatusHeaderStructure, ByVal ds As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPOutMatchingN(header, ds)
    End Function
    <WebMethod()> _
    Public Function GetDataByIntSN(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetDataByIntSN(IntSN)
    End Function

    <WebMethod()> _
    Public Function RDCBoardSNValid(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.RDCBoardSNValid(IntSN)
    End Function

    <WebMethod()> _
    Public Function MatListOnPCBA(ByVal WIPID As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.MatListOnPCBA(WIPID)
    End Function

    <WebMethod()> _
    Public Function IDSwop(ByVal header As StatusHeaderStructure, ByVal type As Integer) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IDSwop(header, type)
    End Function

    <WebMethod()> _
    Public Function IntSNRecycle(ByVal header As StatusHeaderStructure) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IntSNRecycle(header)
    End Function

    <WebMethod()>
    Public Function IntSNRecycleII(ByVal header As StatusHeaderStructure) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IntSNRecycleII(header)
    End Function

    <WebMethod()> _
    Public Function GetProductCPN(ByVal CPN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetProductCPN(CPN)
    End Function

    <WebMethod()> _
    Public Function GetResultList(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetResultList(IntSN)
    End Function

    <WebMethod()> _
    Public Function GetNextProcess(ByVal IntSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetNextProcess(IntSN)
    End Function

    <WebMethod()> _
    Public Function GetResultAndAttributesList(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetResultAndAttributesList(IntSN)
    End Function


    <WebMethod()> _
    Public Function GetBoxInfo(ByVal BoxID As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetBoxInfo(BoxID)
    End Function

    <WebMethod()> _
    Public Function GetLabel1(ByVal Model As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetLabel1(Model)
    End Function

    <WebMethod()> _
    Public Function GetPackingListLabel(ByVal Model As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetPackingListLabel(Model)
    End Function

    <WebMethod()> _
    Public Function GetBoxQtyInPallet(ByVal PalletID As String) As Integer
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetBoxQtyInPallet(PalletID)
    End Function

    <WebMethod()> _
    Public Function ATEWIPIn(ByVal IntSerialNo As String, ByVal Process As String, ByVal user As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATEWIPIn(IntSerialNo, Process, user)
    End Function

    <WebMethod()> _
    Public Function ATEWIPout(ByVal xml As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATEWIPout(xml)
    End Function

    <WebMethod()> _
    Public Function ATEWIPOutDirect(ByVal xml As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATEWIPOutDirect(xml)
    End Function

    <WebMethod()> _
    Public Function IsWipIn(ByVal IntSerial As String, ByVal CurrProcess As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IsWipIn(IntSerial, CurrProcess)
    End Function


    <WebMethod()> _
    Public Function BurnInWipOut(ByVal Header As StatusHeaderStructure, ByVal checked As Boolean) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.BurnInWipOut(Header, checked)
    End Function


    <WebMethod()> _
    Public Function WIPPacking(ByVal DSWIP As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPPacking(DSWIP)
    End Function

    <WebMethod()> _
    Public Function PrintSNLabel(ByVal prodcutSN As String, ByVal labelId As String, ByVal printerName As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.PrintSNLabel(prodcutSN, labelId, printerName)
    End Function


    <WebMethod()> _
    Public Function PrintPaking(ByVal CartonID As String, ByVal labelid As String, ByVal printer As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.PrintPaking(CartonID, labelid, printer)
    End Function

    <WebMethod()> _
    Public Function GetBurnInTime(ByVal IntSerial As String, ByVal CurrProcess As String) As Integer
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetBurnInTime(IntSerial, CurrProcess)

    End Function

    <WebMethod()> _
    Public Function GetShipInfoBySN(ByVal SerialNo As String) As ShipInfo
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetShipInfoBySN(SerialNo)

    End Function

    <WebMethod()> _
    Public Function ChangeBox(ByVal SerialNo As String, ByVal oldboxid As String, ByVal newBoxID As String, ByVal user As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ChangeBox(SerialNo, oldboxid, newBoxID, user)
    End Function

    <WebMethod()> _
    Public Function GetShipInfoByBoxID(ByVal BoxID As String, ByVal user As String) As ShipInfo
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetShipInfoByBoxID(BoxID, user)
    End Function

    <WebMethod()> _
    Public Function GetShipInfoByPalletID(ByVal PalletID As String, ByVal user As String) As ShipInfo
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetShipInfoByPalletID(PalletID, user)
    End Function

    <WebMethod()> _
    Public Function ChangePallet(ByVal BoxID As String, ByVal PalletID As String, ByVal user As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ChangePallet(BoxID, PalletID, user)
    End Function

    <WebMethod()> _
    Public Function OQAWipIn(ByVal ExtNo As String, ByVal OperatorName As String, ByVal InvOrg As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.OQAWipIn(ExtNo, OperatorName, InvOrg)
    End Function

    <WebMethod()> _
    Public Function ExistsFunctionalTest(ByVal ExtNo As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ExistsFunctionalTest(ExtNo)
    End Function

    <WebMethod()> _
    Public Function TLAFlow(ByVal Model As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.TLAFlow(Model)
    End Function

    <WebMethod()> _
    Public Function OQACosmetic(ByVal ExtSN As String, ByVal Model As String, ByVal RetestNo As String, ByVal Result As String, ByVal ERPLoginData As ERPLogin, ByVal dsFlow As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.OQACosmetic(ExtSN, Model, RetestNo, Result, ERPLoginData, dsFlow)
    End Function

    <WebMethod()> _
    Public Function getLabels(ByVal Model As String, ByVal PCBA As String, ByVal Process As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.getLabels(Model, PCBA, Process)
    End Function

    <WebMethod()> _
    Public Function Rework(ByVal ExtSN As String, ByVal DJ As String, ByVal Model As String, ByVal RetestNo As String, ByVal check As Boolean, ByVal ERPLoginData As ERPLogin, ByVal dsFlow As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Rework(ExtSN, DJ, Model, RetestNo, check, ERPLoginData, dsFlow)
    End Function

    <WebMethod()> _
    Public Function Rework_New(ByVal dsFlow As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Rework_New(dsFlow)
    End Function
    <WebMethod()> _
    Public Function TraceLevel(ByVal Model As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.TraceLevel(Model)
    End Function

    <WebMethod()> _
    Public Function GetPanelSize(ByVal Model As String, ByVal PCBA As String) As Integer
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetPanelSize(Model, PCBA)
    End Function

    <WebMethod()> _
    Public Function IntSNIsValidByPanel(ByVal IntSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IntSNIsValidByPanel(IntSN)
    End Function

    <WebMethod()> _
    Public Function WIPMatchingByPanel(ByVal DSWIP As DataSet, ByVal PanelSize As Integer) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPMatchingByPanel(DSWIP, PanelSize)
    End Function

    <WebMethod()> _
    Public Function CheckPanelID(ByVal PanelID As String, ByVal Model As String, ByVal Process As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CheckPanelID(PanelID, Model, Process)
    End Function

    <WebMethod()> _
    Public Function WIPRework(ByVal IntSN As String, ByVal ERPLoginData As ERPLogin, ByVal dsFlow As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPRework(IntSN, ERPLoginData, dsFlow)
    End Function

    <WebMethod()> _
    Public Function ReadMIData(ByVal model As String, ByVal pcba As String, ByVal process As String, ByVal status As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ReadMIData(model, pcba, process, status)
    End Function
    <WebMethod()> _
    Public Function getMIFileData(ByVal model As String, ByVal PCBA As String, ByVal Process As String) As Object
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.getMIFileData(model, PCBA, Process)
    End Function

    <WebMethod()> _
    Public Function SaveMIRecord(ByVal MIRecord As DataSet, ByVal username As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.SaveMIRecord(MIRecord, username)
    End Function

    <WebMethod()> _
    Public Function GetConfig(ByVal eTraceModule As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetConfig(eTraceModule)
    End Function

    <WebMethod()> _
    Public Function GetModelByIntSN(ByVal intSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetModelByIntSN(intSN)
    End Function

    <WebMethod()> _
    Public Function GetPCBAByIntSN(ByVal intSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetPCBAByIntSN(intSN)
    End Function

    <WebMethod()> _
    Public Function GetLastTestResult(ByVal intSN As String, ByVal ProcessName As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetLastTestResult(intSN, ProcessName)
    End Function

    <WebMethod()> _
    Public Function GetPropertiesName(ByVal intSN As String, ByVal processname As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetPropertiesName(intSN, processname)
    End Function
    <WebMethod()> _
    Public Function GetMaxTestRound(ByVal intSN As String, ByVal ProcessName As String) As Integer
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetMaxTestRound(intSN, ProcessName)
    End Function
    <WebMethod()> _
    Public Function GetMaxFailure(ByVal intSN As String, ByVal ProcessName As String) As Integer
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetMaxFailure(intSN, ProcessName)
    End Function

    ' Temp Funvtion just for TE to test
    <WebMethod()> _
    Public Function CleanTestResult(ByVal intSN As String, ByVal ProcessName As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CleanTestResult(intSN, ProcessName)
    End Function
    <WebMethod()> _
    Public Function WIPInOQA(ByVal ExtSN As String, ByVal OperatorName As String, ByVal OrgCode As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPInOQA(ExtSN, OperatorName, OrgCode)
    End Function

    <WebMethod()> _
    Public Function CheckExtSNSameIntSNByModel(ByVal Model As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CheckExtSNSameIntSNByModel(Model)
    End Function

    <WebMethod()> _
    Public Function Employee_Login(ByVal AccessCardID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Employee_Login(AccessCardID)
    End Function

    <WebMethod()> _
    Public Function Employee_Certify(ByVal AccessCardID As String, ByVal CourseCode As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Employee_Certify(AccessCardID, CourseCode)
    End Function

    <WebMethod()> _
    Public Function Loading_CheckModel(ByVal Model As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Loading_CheckModel(Model)
    End Function

    <WebMethod()> _
    Public Function AutoStopLine(ByVal TestData As String, ByVal Type As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.AutoStopLine(TestData, Type)
    End Function

    <WebMethod()> _
    Public Function Laser_VerifyUnit(ByVal Model As String, ByVal IntSN As String, ByVal Computer As String, ByVal User As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Laser_VerifyUnit(Model, IntSN, Computer, User)
    End Function

    <WebMethod()> _
    Public Function Laser_VerifyUnitTest(ByVal Model As String, ByVal IntSN As String, ByVal Computer As String, ByVal User As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Laser_VerifyUnitTest(Model, IntSN, Computer, User)
    End Function

    <WebMethod()> _
    Function Fixture_Register(ByVal Type As String, ByVal FixtureID As String, ByVal MaxUse As Integer, ByVal Description As String, ByVal User As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_Register(Type, FixtureID, MaxUse, Description, User)
    End Function

    <WebMethod()> _
    Function Fixture_Type(ByVal Type As String, ByVal Slot As Integer, ByVal Warning As Integer, ByVal Maintenance As Integer, ByVal Repair As Integer, ByVal DefaultMax As Integer, ByVal Description As String, ByVal User As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_Type(Type, Slot, Warning, Maintenance, Repair, DefaultMax, Description, User)
    End Function

    <WebMethod()> _
    Function Fixture_Maintain(ByVal FixtureID As String, ByVal Description As String, ByVal User As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_Maintain(FixtureID, Description, User)
    End Function

    <WebMethod()> _
    Function Fixture_Repair(ByVal FixtureID As String, ByVal Description As String, ByVal User As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_Repair(FixtureID, Description, User)
    End Function

    <WebMethod()> _
    Function Fixture_InActive(ByVal FixtureID As String, ByVal Description As String, ByVal User As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_InActive(FixtureID, Description, User)
    End Function

    <WebMethod()> _
    Function Fixture_RegisterView(ByVal FixtureID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_RegisterView(FixtureID)
    End Function

    <WebMethod()> _
    Function Fixture_InActiveLog() As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_InActiveLog
    End Function

    <WebMethod()> _
    Function Fixture_InActiveLogByFixture(ByVal FixtureID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_InActiveLogByFixture(FixtureID)
    End Function

    <WebMethod()> _
    Function Fixture_MaintainLog(ByVal FixtureID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_MaintainLog(FixtureID)
    End Function

    <WebMethod()> _
    Function Fixture_RepairLog(ByVal FixtureID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_RepairLog(FixtureID)
    End Function

    <WebMethod()> _
    Function Fixture_TypeView() As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_TypeView()
    End Function

    <WebMethod()> _
    Function Fixture_Update(ByVal FixtureID As String, ByVal Type As String, ByVal MaxUse As Integer, ByVal Description As String, ByVal User As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Fixture_Update(FixtureID, Type, MaxUse, Description, User)
    End Function

    <WebMethod()> _
    Public Function ATE_FixtureVerify(ByVal FixtureID As String, ByVal Type As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATE_FixtureVerify(FixtureID, Type)
    End Function

    <WebMethod()> _
    Public Function ATE_CreateRelation(ByVal FixtureID As String, ByVal Slot As Integer, ByVal Model As String, ByVal IntSN As String, ByVal Process As String, ByVal User As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATE_CreateRelation(FixtureID, Slot, Model, IntSN, Process, User)
    End Function

    <WebMethod()> _
    Public Function ATE_FixtureSign(ByVal FixtureID As String, ByVal User As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATE_FixtureSign(FixtureID, User)
    End Function

    <WebMethod()> _
    Public Function ATE_ReturnSNbySlot(ByVal FixtureID As String, ByVal sLot As Integer) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATE_ReturnSNbySlot(FixtureID, sLot)
    End Function

    <WebMethod()> _
    Public Function ATE_ReturnSNbyFixture(ByVal FixtureID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATE_ReturnSNbyFixture(FixtureID)
    End Function

    <WebMethod()> _
    Public Function ATE_ReleaseRelationbySlot(ByVal FixtureID As String, ByVal Slot As Integer, ByVal User As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATE_ReleaseRelationbySlot(FixtureID, Slot, User)
    End Function

    <WebMethod()> _
    Public Function ATE_ReleaseRelationbyFixture(ByVal FixtureID As String, ByVal User As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATE_ReleaseRelationbyFixture(FixtureID, User)
    End Function

    <WebMethod()> _
    Public Function ATE_IntSlotReview(ByVal IntSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ATE_IntSlotReview(IntSN)
    End Function

    <WebMethod()> _
    Public Function Function_ProcessVerify(ByVal IntSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Function_ProcessVerify(IntSN)
    End Function

    <WebMethod()> _
    Public Function Depanel_VerifyMatching1(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Depanel_VerifyMatching1(IntSN)
    End Function

    <WebMethod()> _
    Public Function Test_Reflow(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Test_Reflow(IntSN)
    End Function

    <WebMethod()> _
    Public Function Test_ItemData(ByVal IntSN As String, ByVal Process As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Test_ItemData(IntSN, Process)
    End Function

    <WebMethod()> _
    Public Function GetSystemGMTDateTime(ByVal TimeZone As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetSystemGMTDateTime(TimeZone)
    End Function

    <WebMethod()> _
    Public Function WIP_UpdateStatus(ByVal IntSN As String, ByVal Process As String, ByVal OperatorName As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIP_UpdateStatus(IntSN, Process, OperatorName)
    End Function

    <WebMethod()> _
    Public Function Test_ReflowProcess(ByVal IntSN As String, ByVal Process As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Test_ReflowProcess(IntSN, Process)
    End Function

    <WebMethod()> _
    Public Function Depanel_VerifyLastTest(ByVal IntSN As DataSet) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Depanel_VerifyLastTest(IntSN)
    End Function

    <WebMethod()> _
    Public Function Get_WIPFGSN(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Get_WIPFGSN(IntSN)
    End Function

    <WebMethod()> _
    Public Function Get_WIPTestData(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetWIPTestData(IntSN)
    End Function

    <WebMethod()> _
    Public Function Get_MaintainExpireLog(ByVal Expire As Integer, ByVal FixtureID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetMaintainExpireLog("Maintenance", Expire, FixtureID)
    End Function

    <WebMethod()> _
    Public Function Get_RepairExpireLog(ByVal Expire As Integer, ByVal FixtureID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetMaintainExpireLog("Repair", Expire, FixtureID)
    End Function

    <WebMethod()> _
    Public Function WIP_UpdateParameter(ByVal IntSN As String, ByVal Process As String, ByVal TestStep As String, ByVal TestName As String, ByVal LowerLimit As Double, ByVal UperLimit As Double) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIP_UpdateParameter(IntSN, Process, TestStep, TestName, LowerLimit, UperLimit)
    End Function

    <WebMethod()> _
    Public Function Temp_InquiryProcess(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Temp_InquiryProcess(IntSN)
    End Function

    <WebMethod()> _
    Public Function Temp_UpdateProcess(ByVal IntSN As String, ByVal SeqNo As Integer, ByVal Process As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Temp_UpdateProcess(IntSN, SeqNo, Process)
    End Function

    <WebMethod()> _
    Public Function SaveDJ(ByVal DJ As DataSet) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.SaveDJ(DJ)
    End Function

    <WebMethod()> _
    Public Function ReadOrg(ByVal ServerName As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ReadOrg(ServerName)
    End Function

    <WebMethod()> _
    Public Function getFlatFileProperties(ByVal type As String, ByVal attribute As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.getFlatFileProperties(type, attribute)
    End Function

    <WebMethod()> _
    Public Function setFlatFileProperties(ByVal type As String, ByVal attribute As String, ByVal value As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.setFlatFileProperties(type, attribute, value)
    End Function

    <WebMethod()> _
    Public Function SaveFlatFileMessage(ByVal model As String, ByVal boxid As String, ByVal palletid As String, ByVal serialno As String, ByVal flatfile As String, ByVal sentby As String) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.SaveFlatFileMessage(model, boxid, palletid, serialno, flatfile, sentby)
    End Function

    <WebMethod()> _
    Public Function saveFlatFileSN(ByVal FlatfileDS As DataSet, ByVal ParamArray parmsArray As String()) As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.saveFlatFileSN(FlatfileDS, parmsArray)
    End Function

    <WebMethod()> _
    Public Function GetFlatfile(ByVal model As String, ByVal BoxID As String, ByVal PalletID As String, ByVal SerialNo As String) As Data.DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetFlatfile(model, BoxID, PalletID, SerialNo)
    End Function

    <WebMethod()> _
    Public Function CheckFlatFileSN(ByVal FlatfileDS As DataSet, ByVal username As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CheckFlatFileSN(FlatfileDS, username)
    End Function

    <WebMethod()> _
    Public Function GetMacAddress() As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetMacAddress()
    End Function

    <WebMethod()> _
    Public Function GetResultAndPCBAList(ByVal IntSN As String, ByVal Proc As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetResultAndPCBAList(IntSN, Proc)
    End Function

    <WebMethod()> _
    Public Function IDUpdate(ByVal IntSN As String, ByVal DJ As String, ByVal Model As String, ByVal TVANo As String, ByVal OrgCode As String, ByVal user As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.IDUpdate(IntSN, DJ, Model, TVANo, OrgCode, user)
    End Function

    <WebMethod()> _
    Public Function WIPMatchingN(ByVal DSWIP As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.WIPMatchingN(DSWIP)
    End Function
    <WebMethod()> _
    Public Function SNListChangeBox(ByVal SNList As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.SNListChangeBox(SNList)
    End Function

    <WebMethod()> _
    Public Function ChangeRevision(ByVal SNList As DataSet) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.ChangeRevision(SNList)
    End Function
    <WebMethod()> _
    Public Function getCartonInfo(ByVal boxid As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.getCartonInfo(boxid)
    End Function

    <WebMethod()> _
    Public Function GetShipInfoByBoxIDSN(ByVal BoxIDSN As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetShipInfoByBoxIDSN(BoxIDSN)
    End Function

    <WebMethod()> _
    Public Function StructureReadByPCBA(ByVal Model As String, ByVal PCBA As String, ByVal mode As Integer) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.StructureReadByPCBA(Model, PCBA, mode)
    End Function

    <WebMethod()> _
    Public Function BackToEeprom(ByVal IntSN As String, ByVal ExtSN As String, ByVal Attribute As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.BackToEeprom(IntSN, ExtSN, Attribute)
    End Function

    <WebMethod()> _
    Public Function GetProductCPNbyModel(ByVal Model As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetProductCPNbyModel(Model)
    End Function

    <WebMethod()> _
    Public Function GetModelByExtSN(ByVal ExtSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetModelByExtSN(ExtSN)
    End Function

    <WebMethod()> _
    Public Function getShipmentByBoxid(ByVal boxid As String, ByVal user As String) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.getShipmentByBoxid(boxid, user)
    End Function

    <WebMethod()> _
    Public Function SFCDBoardIsValid(ByVal intSN As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.SFCDBoardIsValid(intSN)
    End Function

    <WebMethod()> _
    Public Function SaveCLID(ByVal Items As DataSet) As Boolean
        Dim _eTrace_PublicFunction As New PublicFunction()
        Return _eTrace_PublicFunction.SaveCLID(Items)
    End Function

    <WebMethod()> _
    Public Function Get_CusTableLists(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.Get_CusTableLists(LoginData)
    End Function

    <WebMethod()> _
    Public Function GetEmplogeeID(ByVal AccessCardID As String) As String
        Return GetAccessCardUserInfo(AccessCardID).EmplogeeID
    End Function

    <WebMethod()> _
    Public Function GetDept(ByVal AccessCardID As String) As String
        Return GetAccessCardUserInfo(AccessCardID).Dept
    End Function

    <WebMethod()> _
    Public Function CopyCardInfo() As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CopyCardInfo
    End Function

    <WebMethod()> _
    Public Function CopyCardInfoZS() As Boolean
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.CopyCardInfoZS
    End Function


    <WebMethod()> _
    Public Function GetHRTrainingInfo(ByVal ATEMachine As String, ByVal employeeID As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetHRTrainingInfo(ATEMachine, employeeID)
    End Function

    <WebMethod()> _
    Public Function GetLocks() As DataSet
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.GetAllLocks()
    End Function
    <WebMethod()> _
    Public Function UpdateLockByID(ByVal lockId As String, ByVal symptom As String, ByVal te As String, ByVal pe As String, ByVal pqe As String, ByVal other As String, ByVal pbr As String, ByVal remarks As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UpdateLockByID(lockId, symptom, te, pe, pqe, other, pbr, remarks)
    End Function
    <WebMethod()> _
    Public Function UnlockdByID(ByVal lockId As String, ByVal user As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Return _eTraceOracle_SFC.UnlockdByID(lockId, user)
    End Function
    <WebMethod()> _
    Public Function SaveAndUnlockdByID(ByVal lockId As String, ByVal symptom As String, ByVal te As String, ByVal pe As String, ByVal pqe As String, ByVal other As String, ByVal pbr As String, ByVal remarks As String, ByVal user As String) As String
        Dim _eTraceOracle_SFC As New TestDataCollection
        Dim updateResult As String = _eTraceOracle_SFC.UpdateLockByID(lockId, symptom, te, pe, pqe, other, pbr, remarks)
        If updateResult = "true" Then
            Return _eTraceOracle_SFC.UnlockdByID(lockId, user)
        Else
            Return updateResult
        End If
    End Function
#End Region

#Region "Smart Card Validation"

    <WebMethod()> _
    Public Function SaveSmartCardHistory(ByVal ParamArray cardParams As String()) As DataTable
        Dim _eTraceOracle_TDC As New TestDataCollection
        Return _eTraceOracle_TDC.SaveSmartCardHistory(cardParams)
    End Function

    <WebMethod()> _
    Public Function GetSmartCardHistory(ByVal IntSN As String) As DataTable
        Dim _eTraceOracle_TDC As New TestDataCollection
        Return _eTraceOracle_TDC.GetSmartCardHistory(IntSN)
    End Function

#End Region

#Region "RDC"

    <WebMethod()> _
    Public Function ReadFlow(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadFlow(IntSN)
    End Function

    <WebMethod()> _
    Public Function ReadRepairCode() As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadRepairCode()
    End Function

    <WebMethod()> _
    Public Function ReadRepCodesByCategory(ByVal Category As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadRepCodesByCategory(Category)
    End Function

    <WebMethod()> _
    Public Function ReadRepDefectCodeByCause(ByVal Cause As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadRepDefectCodeByCause(Cause)
    End Function

    <WebMethod()> _
    Public Function ReadRepairData(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadRepairData(IntSN)
    End Function

    <WebMethod()> _
    Public Function ReadFailData(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadFailData(IntSN)
    End Function

    <WebMethod()> _
    Public Function GetCLIDData(ByVal CLID As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.GetCLIDData(CLID)
    End Function

    <WebMethod()> _
    Public Function CheckOTO(ByVal SN As String, ByVal Type As String) As Boolean
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.CheckOTO(SN, Type)
    End Function

    <WebMethod()> _
    Public Function SaveRepairDetailData(ByVal RepData As DataSet, ByVal OTO As Boolean, ByVal IntSN As String) As Boolean
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.SaveRepairDetailData(RepData, OTO, IntSN)
    End Function

    <WebMethod()> _
    Public Function ReadStructure(ByVal Model As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadStructure(Model)
    End Function

    <WebMethod()> _
    Public Function ReadConfig(ByVal ConfigID As String) As Boolean
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadConfig(ConfigID)
    End Function

    <WebMethod()> _
    Public Function StandardTime() As String
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.StandardTime()
    End Function

    <WebMethod()> _
    Public Function ServerTime() As DateTime
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ServerTime()
    End Function

    <WebMethod()> _
    Public Function ReadWIPFlow(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadWIPFlow(IntSN)
    End Function

    <WebMethod()> _
    Public Function RepScrap(ByVal IntSN As String) As Boolean
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.RepScrap(IntSN)
    End Function

    <WebMethod()> _
    Public Function RDCScrap(ByVal IntSN As String) As String
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.RDCScrap(IntSN)
    End Function

    <WebMethod()> _
    Public Function RDCSaveII(ByVal DS As DataSet) As String
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.RDCSaveII(DS)
    End Function

    <WebMethod()> _
    Public Function FailRecordII(ByVal SN As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.FailRecordII(SN)
    End Function

    <WebMethod()> _
    Public Function RDCScrapII(ByVal IntSN As String, ByVal ChangedBy As String) As String
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.RDCScrapII(IntSN, ChangedBy)
    End Function

    <WebMethod()> _
    Public Function RDC121MatlInfo(ByVal PCBID As String, ByVal Model As String, ByVal PCBA As String, ByVal Circuit As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.RDC121MatlInfo(PCBID, Model, PCBA, Circuit)
    End Function


    <WebMethod()> _
    Public Function RDCMatInfo(ByVal DJ As String, ByVal MaterialNo As String, ByVal PCBA As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.RDCMatInfo(DJ, MaterialNo, PCBA)
    End Function

    <WebMethod()> _
    Public Function RepRaiseTimes(ByVal ds As DataSet) As String
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.RepRaiseTimes(ds)
    End Function

    <WebMethod()> _
    Public Function SaveRepairRecordData(ByVal RepData As DataSet, ByVal OTO As Boolean, ByVal IntSN As String) As Boolean
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.SaveRepairRecordData(RepData, OTO, IntSN)
    End Function
    <WebMethod()> _
    Public Function SkipBI(ByVal IntSN As String) As String
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.SkipBI(IntSN)
    End Function

    <WebMethod()> _
    Public Function ReadFailItem(ByVal IntSN As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadFailItem(IntSN)
    End Function

    <WebMethod()> _
    Public Function ReadNDFData(ByVal IntSN As String, ByVal OperatorName As String, ByVal Type As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ReadNDFData(IntSN, OperatorName, Type)
    End Function

    <WebMethod()> _
    Public Function FailRecord(ByVal SN As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.FailRecord(SN)
    End Function

    <WebMethod()> _
    Public Function NewFailData(ByVal SN As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.NewFailData(SN)
    End Function

    <WebMethod()> _
    Public Function RDCWIPFLow(ByVal WIPID As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.RDCWIPFLow(WIPID)
    End Function

    <WebMethod()> _
    Public Function RDCSave(ByVal DS As DataSet) As String
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.RDCSave(DS)
    End Function

    <WebMethod()> _
    Public Function GetSpecifySeatItem(ByVal PCB As String, ByVal RefD As String, ByVal DJ As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.GetSpecifySeatItem(PCB, RefD, DJ)
    End Function
    <WebMethod()> _
    Public Function GetSpecifySeatItemByOrg(ByVal PCB As String, ByVal RefD As String, ByVal DJ As String, ByVal Org As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.GetSpecifySeatItemByOrg(PCB, RefD, DJ, Org)
    End Function

    <WebMethod()> _
    Public Function GetPdControlByDJ(ByVal p_DJ As String) As DataSet
        Dim _eTraceOracle_RDC As New RDC()
        Return _eTraceOracle_RDC.GetPdControlByDJ(p_DJ)
    End Function

    <WebMethod()> _
    Public Function UpdatePOQty(ByVal PO As String, ByVal OrgCode As String, ByVal Model As String, ByVal ModelRev As String, ByVal POQty As Integer, ByVal TVA As String, ByVal AllowMatching As Boolean, ByVal AllowPacking As Boolean, ByVal ChangedBy As String, ByVal Remarks As String) As String
        Dim _eTraceOracle_RDC As New RDC()
        Return _eTraceOracle_RDC.UpdatePOQty(PO, OrgCode, Model, ModelRev, POQty, TVA, AllowMatching, AllowPacking, ChangedBy, Remarks)
    End Function

    <WebMethod()> _
    Public Function ATELockingRDCWIPIN(ByVal IntSN As String) As String
        Dim _eTraceOracle_RDC As New RDC
        Return _eTraceOracle_RDC.ATELockingRDCWIPIN(IntSN)
    End Function
#End Region

#Region "SAPMigration"
    '<WebMethod()> _
    'Public Function UpdateOraItem(ByVal Logindata As ERPLogin) As DataSet
    '    Dim _eTrace_SAPMigration As New SAPMigration()
    '    Return _eTrace_SAPMigration.UpdateOraItem(Logindata)
    'End Function

    <WebMethod()> _
    Public Function UpdateOraItem(ByVal Logindata As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.UpdateOraItem(Logindata)
    End Function

    <WebMethod()> _
    Public Function ArchiveCLID(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.ArchiveCLID(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function UpdateSTypeBin(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.UpdateSTypeBin(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function DisableManualItems(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.DisableManualItems(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckNoMapping(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.CheckNoMapping(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CountNoMapping(ByVal OracleLoginData As ERPLogin) As Integer
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.CountNoMapping(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckMigrateStatus(ByVal OracleLoginData As ERPLogin) As Integer
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.CheckMigrateStatus(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function UpdateCLMaster(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.UpdateCLMaster(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function SumSAPIM(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.SumSAPIM(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckQtyMatch(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.CheckQtyMatch(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function AssignClientID(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.AssignClientID(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function UploadToOracle(ByVal ClientID As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.UploadToOracle(ClientID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function IssueFmOracle(ByVal ClientID As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.IssueFmOracle(ClientID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckSAPPN(ByVal SAPPN As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.CheckSAPPN(SAPPN, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function Check_SAPPN(ByVal SAPPN As String, ByVal OracleLoginData As ERPLogin) As SAPPN_Check
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.Check_SAPPN(SAPPN, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function RollbackCLIDInfo(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_SAPMigration As New SAPMigration()
        Return _eTrace_SAPMigration.RollbackCLIDInfo(OracleLoginData)
    End Function
#End Region

#Region "StockTake"

#Region "Physical Inventory"
    <WebMethod()> _
    Public Function Get_STCtrlList(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.Get_STCtrlList(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ST_CheckAction(ByVal Action As String, ByVal OracleLogindata As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ST_CheckAction(Action, OracleLogindata)
    End Function

    <WebMethod()> _
    Public Function ST_CompAction(ByVal Action As String, ByVal OracleLogindata As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ST_CompAction(Action, OracleLogindata)
    End Function

    <WebMethod()> _
    Public Function UpdateActionStatus(ByVal Action As String, ByVal Enabled As Boolean, ByVal Done As Boolean, ByVal OracleLogindata As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.UpdateActionStatus(Action, Enabled, Done, OracleLogindata)
    End Function

    <WebMethod()> _
    Public Function Lock_eTrace(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.Lock_eTrace(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function UnLock_eTrace(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.UnLock_eTrace(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ClearPIName(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ClearPIName(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetOrgList_StockTake(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.GetOrgList_StockTake(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function AddPIName(ByVal OrgCode As String, ByVal PIName As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.AddPIName(OrgCode, PIName, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function StockTake_CpySubLoc(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.StockTake_CpySubLoc(OracleLoginData)
    End Function
    <WebMethod()> _
    Public Function CopyPI(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.CopyPI(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CopyCLMaster(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.CopyCLMaster(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ExtCnt2(ByVal OracleLogindata As ERPLogin) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ExtCnt2(OracleLogindata)
    End Function

    <WebMethod()> _
    Public Function SetCount1(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.SetCount1(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function SetCount2(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.SetCount2(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckCountOption(ByVal OracleLogindata As ERPLogin) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.CheckCountOption(OracleLogindata)
    End Function

    <WebMethod()> _
    Public Function StockTake_ValidateSubLoc(ByVal OracleLoginData As ERPLogin, ByVal Subinv As String, ByVal Locator As String) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.StockTake_ValidateSubLoc(OracleLoginData, Subinv, Locator)
    End Function

    <WebMethod()> _
    Public Function CheckInSubLocList(ByVal LoginData As ERPLogin, ByVal SubInv As String, ByVal Locator As String) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.CheckInSubLocList(LoginData, SubInv, Locator)
    End Function

    <WebMethod()> _
    Public Function CheckBFSubinv(ByVal Subinv As String, ByVal OracleLogindata As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.CheckBFSubinv(Subinv, OracleLogindata)
    End Function

    <WebMethod()> _
    Public Function GetCLIDInfo(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.GetCLIDInfo(CLID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckStopFlag(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.CheckStopFlag(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function Save_STChange(ByVal pstInfo As DataSet, ByVal CountSeq As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.Save_STChange(pstInfo, CountSeq, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function CheckNotFound(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.CheckNotFound(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GenVarRpt(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.GenVarRpt(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function EnableScan(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.EnableScan(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function StopScan(ByVal SeqNo As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.StopScan(SeqNo, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function NoValidate(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.NoValidate(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function WithValidate(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.WithValidate(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function BkpRmv(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.BkpRmv(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function SumQtyForPI(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.SumQtyForPI(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function StockTake_AssignExpDate(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.StockTake_AssignExpDate(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function UpdateDiffLocator(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.UpdateDiffLocator(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function GetPIResult(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.GetPIResult(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function StockTake_BkpBFAdjust(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.StockTake_BkpBFAdjust(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ForLoss(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ForLoss(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ForGain(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ForGain(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ForNotFound(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ForNotFound(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ForNewFind(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ForNewFind(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ForDiffLocator(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ForDiffLocator(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ForLossDiff(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ForLossDiff(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function ForGainDiff(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.ForGainDiff(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function StockTake_BkpAfterAdjust(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.StockTake_BkpAfterAdjust(OracleLoginData)
    End Function
#End Region

#Region "BF Count"

    <WebMethod()> _
    Public Function BFC_GetWeekCode(ByVal OracleLoginData As ERPLogin) As String
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.BFC_GetWeekCode(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function BFC_GetBFSubinv(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.BFC_GetBFSubinv(OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function BFC_DelOldBFCount(ByVal WeekCode As String, ByVal BFSubinv As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.BFC_DelOldBFCount(WeekCode, BFSubinv, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function BFC_GetCLIDInfo(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.BFC_GetCLIDInfo(CLID, OracleLoginData)
    End Function

    <WebMethod()> _
    Public Function Save_BFCount(ByVal WeekCode As String, ByVal lblInfo As DataSet, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim _eTrace_StockTake As New StockTake()
        Return _eTrace_StockTake.Save_BFCount(WeekCode, lblInfo, OracleLoginData)
    End Function

#End Region

#End Region

#Region "Customs"

    <WebMethod()> _
    Public Function Get_Secondary_Inventory(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Get_Secondary_Inventory(LoginData)
    End Function

    <WebMethod()> _
    Public Function Get_Sub_Locator(ByVal LoginData As ERPLogin, ByVal subinventory_code As String) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Get_Sub_Locator(LoginData, subinventory_code)
    End Function

    <WebMethod()> _
    Public Function Get_List_Sub_Locator(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Get_List_Sub_Locator(LoginData)
    End Function

    <WebMethod()> _
    Public Function Find_List_Sub_Locator(ByVal LoginData As ERPLogin, ByVal Sub_type As String, ByVal Subinventory As String, ByVal Locator As String, ByVal Indecator As String) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Find_List_Sub_Locator(LoginData, Sub_type, Subinventory, Locator, Indecator)
    End Function

    <WebMethod()> _
    Public Function Add_Sub_Locator(ByVal LoginData As ERPLogin, ByVal SUBType As String, ByVal SUBINV As String, ByVal Locator As String, ByVal Indecator As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Add_Sub_Locator(LoginData, SUBType, SUBINV, Locator, Indecator)
    End Function

    <WebMethod()> _
    Public Function Delete_Sub_Locator(ByVal RowID As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Delete_Sub_Locator(RowID)
    End Function

    <WebMethod()> _
    Public Function Validate_item_number(ByVal LoginData As ERPLogin, ByVal ItemNo As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Validate_item_number(LoginData, ItemNo)
    End Function

    <WebMethod()> _
    Public Function add_floor_stock_material(ByVal LoginData As ERPLogin, ByVal ItemNo As String, ByVal usage1 As String, ByVal usage2 As String, ByVal usage3 As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.add_floor_stock_material(LoginData, ItemNo, usage1, usage2, usage3)
    End Function

    <WebMethod()> _
    Public Function Get_List_Floor_Stock(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Get_List_Floor_Stock(LoginData)
    End Function

    <WebMethod()> _
    Public Function Find_List_Floor_Stock(ByVal LoginData As ERPLogin, ByVal Item As String, ByVal Usage1 As String, ByVal Usage2 As String, ByVal Usage3 As String) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Find_List_Floor_Stock(LoginData, Item, Usage1, Usage2, Usage3)
    End Function

    <WebMethod()> _
    Public Function Delete_Floor_Stock(ByVal RowID As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Delete_Floor_Stock(RowID)
    End Function

    <WebMethod()> _
    Public Function Update_Floor_Stock(ByVal RowID As String, ByVal Usage1 As String, ByVal Usage2 As String, ByVal Usage3 As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Update_Floor_Stock(RowID, Usage1, Usage2, Usage3)
    End Function

    <WebMethod()> _
    Public Function Get_Generic_Disposition(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Get_Generic_Disposition(LoginData)
    End Function

    <WebMethod()> _
    Public Function Get_List_Generic_Disposition(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Get_List_Generic_Disposition(LoginData)
    End Function

    <WebMethod()> _
    Public Function add_generic_dispositions(ByVal LoginData As ERPLogin, ByVal DispName As String, ByVal DispDESC As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.add_generic_dispositions(LoginData, DispName, DispDESC)
    End Function

    <WebMethod()> _
    Public Function Delete_Generic_Dispositions(ByVal RowID As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Delete_Generic_Dispositions(RowID)
    End Function

    <WebMethod()> _
    Public Function Get_List_Item_Mapping(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Get_List_Item_Mapping(LoginData)
    End Function

    <WebMethod()> _
    Public Function add_List_Item_Mapping(ByVal LoginData As ERPLogin, ByVal Ora_segment1 As String, ByVal Iems_segment1 As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.add_List_Item_Mapping(LoginData, Ora_segment1, Iems_segment1)
    End Function

    <WebMethod()> _
    Public Function Delete_List_Item_Mapping(ByVal RowID As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Delete_List_Item_Mapping(RowID)
    End Function

    <WebMethod()> _
    Public Function LoadCustomReportData(ByVal dataType As String, ByVal pannelVisible As Boolean, ByVal checkBox As String, ByVal selectedTable As String, ByVal erpLogin As ERPLogin, ByVal dsInput As DataSet) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.LoadCustomReportData(dataType, pannelVisible, checkBox, selectedTable, erpLogin, dsInput)
    End Function
    <WebMethod()> _
    Public Function Get_item_wastage(ByVal LoginData As ERPLogin) As DataSet
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Get_item_wastage(LoginData)
    End Function
    <WebMethod()> _
    Public Function Insert_item_wastage(ByVal LoginData As ERPLogin, ByVal Itemnum As String, ByVal Wastage As Decimal, ByVal Description As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Insert_item_wastage(LoginData, Itemnum, Wastage, Description)
    End Function

    <WebMethod()> _
    Public Function delete_xxetr_item_wastage(ByVal RowID As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.delete_xxetr_item_wastage(RowID)
    End Function

    <WebMethod()> _
    Public Function Update_xxetr_item_wastage(ByVal RowID As String, ByVal wastage As Decimal, ByVal description As String) As String
        Dim _eTrace_Customs As New Customs()
        Return _eTrace_Customs.Update_xxetr_item_wastage(RowID, wastage, description)
    End Function

#End Region

#Region "QCcode"

    <WebMethod()> _
    Public Function SaveTDCRepairCode(ByVal RepairCode As DataSet) As Boolean
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.SaveTDCRepairCode(RepairCode)
    End Function

    <WebMethod()> _
    Public Function ReadRepairCodeForQCcode(ByVal Category As String, ByVal Site As String) As DataSet
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.ReadRepairCodeForQCcode(Category, Site)
    End Function

    <WebMethod()> _
    Public Function IsProductProcess(ByVal Process As String) As Boolean
        Dim _eTrace_TDCFunction As New TestDataCollection()
        Return _eTrace_TDCFunction.IsProductProcess(Process)
    End Function
#End Region

#Region "WMS-Allocation"
    <WebMethod()>
    Public Function GetActiveEvent(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetActiveEvent(OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetActiveJob(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetActiveJob(OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetActiveEvent_ActiveJob(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetActiveEvent_ActiveJob(OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetActiveEvent_ActiveJob_LD(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetActiveEvent_ActiveJob_LD(OracleLoginData)
    End Function

    <WebMethod()>
    Public Function EventLightOff(ByVal EventID As String, ByVal EventType As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.EventLightOff(EventID, EventType, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetLocConfig(ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.GetLocConfig(OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetWMSConfig(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetWMSConfig(OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetItemUsage(ByVal Job As String, ByVal PCBA As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetItemUsage(Job, PCBA, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetItemUsage_LD(ByVal Job As String, ByVal PCBA As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetItemUsage_LD(Job, PCBA, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function WMS_Check_PickedFlag(ByVal Job As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.WMS_Check_PickedFlag(Job, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function Get_MO_Information(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal Job As String, ByVal ProdLine As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String, ByVal TransType As String) As MO_Information
        Dim _WMS As New WMS()
        Return _WMS.Get_MO_Information(OracleLoginData, DN, MO, Job, ProdLine, SubInv, Locator, Item, TransType)
    End Function

    <WebMethod()>
    Public Function Get_MO_Information_SQL(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal Job As String, ByVal ProdLine As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String, ByVal TransType As String) As MO_Information
        Dim _WMS As New WMS()
        Return _WMS.Get_MO_Information_SQL(OracleLoginData, DN, MO, Job, ProdLine, SubInv, Locator, Item, TransType)
    End Function

    <WebMethod()>
    Public Function Get_MO_Information_LD(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal Job As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String, ByVal TransType As String) As MO_Information
        Dim _WMS As New WMS()
        Return _WMS.Get_MO_Information_LD(OracleLoginData, DN, MO, Job, SubInv, Locator, Item, TransType)
    End Function

    <WebMethod()>
    Public Function Get_SubinvLoc_for_CS(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.Get_SubinvLoc_for_CS(OracleLoginData)
    End Function

    <WebMethod()>
    Public Function Get_MO_Information_CS_SQL(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal Job As String, ByVal ProdLine As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String, ByVal TransType As String) As MO_Information
        Dim _WMS As New WMS()
        Return _WMS.Get_MO_Information_CS_SQL(OracleLoginData, DN, MO, Job, ProdLine, SubInv, Locator, Item, TransType)
    End Function

    <WebMethod()>
    Public Function Get_MO_For_LED(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String) As MO_List
        Dim _WMS As New WMS()
        Return _WMS.Get_MO_For_LED(OracleLoginData, DN, MO, SubInv, Locator, Item)
    End Function

    <WebMethod()>
    Public Function Get_RefQty(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.Get_RefQty(OracleLoginData)
    End Function

    <WebMethod()>
    Public Function Get_JobInform_MO(ByVal OracleLoginData As ERPLogin, ByVal DJ As String) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.Get_JobInform_MO(OracleLoginData, DJ)
    End Function

    <WebMethod()>
    Public Function WMS_CheckSlotShortage(ByVal dsMOList As DataSet, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As SlotShortageList
        Dim _WMS As New WMS()
        Return _WMS.WMS_CheckSlotShortage(dsMOList, Subinv, Locator, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function WMS_CheckSlotAvl(ByVal dsMOList As DataSet, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As SlotShortageList
        Dim _WMS As New WMS()
        Return _WMS.WMS_CheckSlotAvl(dsMOList, Subinv, Locator, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetAllCLIDInfo_LED(ByVal dsMOList As DataSet, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetAllCLIDInfo_LED(dsMOList, Subinv, Locator, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetCLIDInfo_LED(ByVal Component As String, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetCLIDInfo_LED(Component, Subinv, Locator, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetCLIDInfo_LED_ByID(ByVal CLIDList As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetCLIDInfo_LED_ByID(CLIDList, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetCLIDInfo_RTLot(ByVal Component As String, ByVal RTLot As String, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetCLIDInfo_RTLot(Component, RTLot, Subinv, Locator, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetCLIDCombination(ByVal srcCLIDs As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.GetCLIDCombination(srcCLIDs, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function WMS_Save_Table(ByVal dsCLID As DataSet, ByVal dsItem As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.WMS_Save_Table(dsCLID, dsItem, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function WMS_Save_Table2(ByVal dsOracle As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.WMS_Save_Table2(dsOracle, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function WMS_Save_Table3(ByVal dsComb As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.WMS_Save_Table3(dsComb, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function WMS_Post_MO_Allocation(ByVal dsMOList As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.WMS_Post_MO_Allocation(dsMOList, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function WMS_Save_Allocation(ByVal dsHeader As DataSet, ByVal dsItem As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.WMS_Save_Allocation(dsHeader, dsItem, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetCycleCountData(ByVal cc_name As String, ByVal seq As Integer, ByVal OracleLoginData As ERPLogin) As GetCycleCount
        Dim _WMS As New WMS()
        Return _WMS.GetCycleCountData(cc_name, seq, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function GetCycleCountList(ByVal cc_name As String, ByVal OracleLoginData As ERPLogin) As GetCycleCount
        Dim _WMS As New WMS()
        Return _WMS.GetCycleCountList(cc_name, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function PostCycleCountAllocation(ByVal dsCC As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.PostCycleCountAllocation(dsCC, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function PostOccupiedAllocation(ByVal Rack As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.PostOccupiedAllocation(Rack, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function PostEmptyAllocation(ByVal Rack As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.PostEmptyAllocation(Rack, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function PostConditionalAllocation(ByVal Item As String, ByVal CLID As String, ByVal Rack As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.PostConditionalAllocation(Item, CLID, Rack, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function UpdateCLIDMissing(ByVal EventID As String, ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As String
        Dim _WMS As New WMS()
        Return _WMS.UpdateCLIDMissing(EventID, CLID, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function WMS_Check_EventID(ByVal EventID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.WMS_Check_EventID(EventID, OracleLoginData)
    End Function

    <WebMethod()>
    Public Function WMS_Check_Rack(ByVal Rack As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim _WMS As New WMS()
        Return _WMS.WMS_Check_Rack(Rack, OracleLoginData)
    End Function

#End Region



#Region "OneToOne"

    <WebMethod()> _
    Public Function IsAuthorizedPCB(ByVal model As String, ByVal pcb As String, ByVal userName As String) As String
        Dim oto As New OneToOne()
        Return oto.IsAuthorizedPCB(model, pcb, userName)
    End Function

    <WebMethod()> _
    Public Function UpdateKanbanLabel(ByVal userName As String, ByVal kbCode As String, ByVal subType As String, ByVal qty As Double) As String
        Dim oto As New OneToOne()
        Return oto.UpdateKanbanLabel(userName, kbCode, subType, qty)
    End Function

    <WebMethod()> _
    Public Function ErrorLoggingOTO(ByVal ModuleName As String, ByVal User As String, ByVal ErrMsg As String) As Boolean
        Dim oto As New OneToOne()
        Return oto.ErrorLoggingOTO(ModuleName, User, ErrMsg)
    End Function

    ''    <WebMethod()> _
    ''Public Function GetPTImplementDJInfo(ByVal userName As String, ByVal dj As String) As DataTable
    ''        Dim oto As New OneToOne()
    ''        Return oto.GetPTImplementDJInfo(userName, dj)
    ''    End Function

    <WebMethod()> _
    Public Function EmployeeIDLoginOTO(ByVal EmployeeID As String) As DataSet
        Dim oto As New OneToOne()
        Return oto.EmployeeIDLoginOTO(EmployeeID)
    End Function



#End Region

#Region "Andon"

    <WebMethod()> _
    Public Function AndonInput(ByVal Model As String, ByVal Line As String, ByVal Station As String, ByVal status As String) As String
        Dim _eTraceOracle_Andon As New Andon()
        Return _eTraceOracle_Andon.AndonInput(Model, Line, Station, status)

    End Function
    <WebMethod()> _
    Public Function ExistsModel(ByVal Model As String) As Integer
        Dim _eTraceOracle_Andon As New Andon()
        Return _eTraceOracle_Andon.ExistsModel(Model)
    End Function

    <WebMethod()> _
    Public Function AndonActualQty(ByVal Model As String, ByVal Line As String) As String
        Dim _eTraceOracle_Andon As New Andon()
        Return _eTraceOracle_Andon.AndonActualQty(Model, Line)

    End Function

    <WebMethod()> _
    Public Function AndonActualQtyofProcess(ByVal Model As String, ByVal Line As String, ByVal process As String) As String
        Dim _eTraceOracle_Andon As New Andon()
        Return _eTraceOracle_Andon.AndonActualQtyofProcess(Model, Line, process)
    End Function

    <WebMethod()> _
    Public Function AndonFailedPercent(ByVal Model As String, ByVal Line As String) As String
        Dim _eTraceOracle_Andon As New Andon()
        Return _eTraceOracle_Andon.AndonFailedPercent(Model, Line)
    End Function

    <WebMethod()> _
    Public Function AndonFailedQty(ByVal Model As String, ByVal Line As String) As String
        Dim _eTraceOracle_Andon As New Andon()
        Return _eTraceOracle_Andon.AndonFailedQty(Model, Line)
    End Function

    <WebMethod()> _
    Public Function AndonLineStopTime(ByVal Model As String, ByVal Line As String) As String
        Dim _eTraceOracle_Andon As New Andon()
        Return _eTraceOracle_Andon.AndonLineStopTime(Model, Line)
    End Function

    <WebMethod()> _
    Public Function AndonLineStopFreq(ByVal Model As String, ByVal Line As String) As String
        Dim _eTraceOracle_Andon As New Andon()
        Return _eTraceOracle_Andon.AndonLineStopFreq(Model, Line)
    End Function

    <WebMethod()>
    Public Function GetAndonProdSche(ByVal Line As String) As DataSet
        Dim i As Integer
        Dim dsAndonProdSche As DataSet = New DataSet("dsAndonProdSche")
        Dim dtAndonProdSche As DataTable = New DataTable("dtAndonProdSche")
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("eTraceDBConnString"))
        Try
            Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("select StartTime, EndTime from T_AndonProdSche with (nolock) where Line='{0}' order by StartTime", Line), myConn)
            da.Fill(dtAndonProdSche)
            dsAndonProdSche.Tables.Add(dtAndonProdSche)

        Catch ex As Exception
            ErrorLogging("TDC-GetAndonProdSche", "", ex.Message)
        End Try
        Return dsAndonProdSche
    End Function

    <WebMethod()>
    Public Function Test(ByVal inputTest As String) As String
        Return inputTest
    End Function

    <WebMethod()> _
    Public Function UpdateAndonProdSche(ByVal ds As DataSet, ByVal Line As String) As String
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("eTraceDBConnString"))
        Dim strSQL As String
        Try
            myConn.Open()
            strSQL = String.Format("delete T_AndonProdSche where Line='{0}'", Line)
            cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
            cmdSQL.ExecuteNonQuery()
            Dim i As Integer
            For i = 0 To ds.Tables(0).Rows.Count - 1
                strSQL = String.Format("exec sp_AndonProdScheInsert '{0}', '{1}', '{2}'", Line, ds.Tables(0).Rows(i)("StartTime"), ds.Tables(0).Rows(i)("EndTime"))
                cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                cmdSQL.ExecuteNonQuery()
            Next
            myConn.Close()
            UpdateAndonProdSche = ""
        Catch ex As Exception
            UpdateAndonProdSche = "Error: " & ex.Message & " Please contact Admin"
            ErrorLogging("TDC-UpdateAndonProdSche", "", ex.Message)
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
    End Function

    <WebMethod()> _
    Public Function AndonProjectedQty(ByVal Line As String, ByVal Target As Integer) As Integer
        Dim _eTraceOracle_Andon As New Andon()
        Return _eTraceOracle_Andon.AndonProjectedQty(Line, Target)
    End Function
#End Region

#Region "LED Hardware System Control"
    <WebMethod()> _
    Public Function LEDControlByRack(ByVal RackID As String, ByVal Code As Integer) As Boolean
        Dim _eTrace_WMS As New WMS
        Return _eTrace_WMS.LEDControlByRack(RackID, Code)
    End Function

    <WebMethod()> _
    Public Function LEDControlBySlot(ByVal slotList As DataSet, ByVal Code As Integer, ByVal Interval As Integer) As Boolean
        Dim _eTrace_WMS As New WMS
        Return _eTrace_WMS.LEDControlBySlot(slotList, Code, Interval)
    End Function
#End Region

#Region "WMS-SMTMO"

    <WebMethod()>
    Public Function CheckBoxID(ByRef mySMOData As SMTData) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.CheckBoxID(mySMOData)
    End Function
    <WebMethod()> _
    Public Function CheckProdLine(ByRef mySMOData As SMTData) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.CheckProdLine(mySMOData)
    End Function

    <WebMethod()> _
    Public Function ConfirmCLID(ByRef mySMOData As SMTData) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ConfirmCLID(mySMOData)
    End Function

    <WebMethod()> _
    Public Function ConfirmMO(ByVal LoginData As ERPLogin, ByRef mySMOData As SMTData) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ConfirmMO(LoginData, mySMOData)
    End Function

    <WebMethod()> _
    Public Function ReadEventID(ByRef mySMOData As SMTData) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ReadEventID(mySMOData)
    End Function

    <WebMethod()> _
    Public Function PrintSMTDJLabel(ByRef mySMOData As SMTData, ByVal Printer As String) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.PrintSMTDJLabel(mySMOData, Printer)
    End Function

    <WebMethod()> _
    Public Function CancelMO(ByRef mySMOData As SMTData) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.CancelMO(mySMOData)
    End Function

    <WebMethod()> _
    Public Function ProcessDocking(ByRef mySMOData As SMTData) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ProcessDocking(mySMOData)
    End Function

#End Region

#Region "LD-WMS-SMTMO"

    <WebMethod()> _
    Public Function CheckProdLine_LD(ByRef mySMOData As SMTData) As String
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.CheckProdLine_LD(mySMOData)
    End Function

    <WebMethod()> _
    Public Function ConfirmCLID_LD(ByRef mySMOData As SMTData) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ConfirmCLID_LD(mySMOData)
    End Function

    <WebMethod()> _
    Public Function ReadEventID_LD(ByRef mySMOData As SMTData) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ReadEventID_LD(mySMOData)
    End Function

    <WebMethod()> _
    Public Function PrintSMTDJLabel_LD(ByRef mySMOData As SMTData, ByVal Printer As String) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.PrintSMTDJLabel_LD(mySMOData, Printer)
    End Function

    <WebMethod()> _
    Public Function ProcessDocking_LD(ByRef mySMOData As SMTData) As Boolean
        Dim _eTrace_ProdPicking As New ProdPicking
        Return _eTrace_ProdPicking.ProcessDocking_LD(mySMOData)
    End Function

#End Region

#Region "SMT Jobmanagement"

    <WebMethod()> _
    Public Function GetDJInfoSMT(ByVal sql As String, ByVal entityName As String, ByRef msg As String, ByVal xmlParameters As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetDJInfo(sql, entityName, msg, xmlParameters)
    End Function

    <WebMethod()> _
    Public Function GetProductStructure(ByVal sql As String, ByVal entityName As String, ByRef msg As String, ByVal xmlParameters As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetProductStructure(sql, entityName, msg, xmlParameters)
    End Function

    <WebMethod()> _
    Public Function NewSeqNO(ByRef msg As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.NewSeqNO(msg)
    End Function

    <WebMethod()> _
    Public Function GetJobHeader(ByVal sql As String, ByVal entityName As String, ByRef msg As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetJobHeader(sql, entityName, msg)
    End Function

    <WebMethod()> _
    Public Function GetJobHeader2(ByVal sql As String, ByVal entityName As String, ByRef msg As String, ByVal xmlParameters As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetJobHeader2(sql, entityName, msg, xmlParameters)
    End Function

    <WebMethod()> _
    Public Sub InsertJob(ByVal xmlObj As String, ByRef msg As String)
        Dim rep = New SMTJobManagementRepository()
        rep.InsertJob(xmlObj, msg)
    End Sub

    <WebMethod()> _
    Public Sub UpdateJob(ByVal xmljob As String, ByVal clearFirst As Boolean, ByRef msg As String)
        Dim pijm = New SMTJobManagementRepository()
        pijm.UpdateJob(xmljob, clearFirst, msg)
    End Sub

    <WebMethod()> _
    Public Sub UpdateJobStatus(ByVal xmljob As String, ByRef msg As String)
        Dim pijm = New SMTJobManagementRepository()
        pijm.UpdateJobStatus(xmljob, msg)
    End Sub

    <WebMethod()> _
    Public Function GetProdLineByLineType(ByVal sql As String, ByVal entityName As String, ByRef msg As String, ByVal xmlParameters As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetProdLineByLineType(sql, entityName, msg, xmlParameters)
    End Function

    <WebMethod()> _
    Public Function GetSMTMachines(ByVal sql As String, ByVal entityName As String, ByRef msg As String, ByVal xmlParameters As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetSMTMachines(sql, entityName, msg, xmlParameters)
    End Function
    <WebMethod()>
    Public Function GetProdLine(ByVal sql As String, ByVal entityName As String, ByRef msg As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetProdLine(sql, entityName, msg)
    End Function

    <WebMethod()>
    Public Function GetMSLMachine(ByVal sql As String, ByVal entityName As String, ByRef msg As String, ByVal xmlParameters As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetMSLMachine(sql, entityName, msg, xmlParameters)
    End Function
    <WebMethod()>
    Public Function GetMSLModel(ByVal sql As String, ByVal entityName As String, ByRef msg As String, ByVal xmlParameters As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetMSLModel(sql, entityName, msg, xmlParameters)
    End Function
    <WebMethod()>
    Public Sub DeleteMsl(ByVal xmlObj As String, ByRef msg As String)
        Dim rep = New SMTJobManagementRepository()
        rep.DeleteMsl(xmlObj, msg)
    End Sub
    <WebMethod()>
    Public Sub InsertMsl(ByVal xmlObj As String, ByRef msg As String)
        Dim rep = New SMTJobManagementRepository()
        rep.InsertMsl(xmlObj, msg)
    End Sub
    <WebMethod()>
    Public Function GetJobMsl(ByVal sql As String, ByVal entityName As String, ByRef msg As String, ByVal xmlParameters As String) As String
        Dim rep = New SMTJobManagementRepository()
        Return rep.GetJobMsl(sql, entityName, msg, xmlParameters)
    End Function
    <WebMethod()>
    Public Function ExecSP(ByVal spname As String, ByVal msg As String, ByVal xmlParameters As String) As Integer
        Dim rep = New SMTJobManagementRepository()
        Return rep.ExecSP(spname, msg, xmlParameters)
    End Function
#End Region


End Class