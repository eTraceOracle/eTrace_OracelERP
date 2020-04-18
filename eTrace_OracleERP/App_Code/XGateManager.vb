Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml.Serialization
Imports eTrace.Entities
Imports Microsoft.VisualBasic
Imports Ptl.Device
Imports Ptl.Device.Communication.Command
Imports Ptl.Device.Communication.XGateConfig

Class PtlTeraManager
    Inherits PublicFunction
    Private Shared _instance As PtlTeraManager

    Private _wms As WMS = New WMS()

    Private _allLedIndexs() As Integer = New Integer() {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99}

    Public Shared ReadOnly Property Instance() As PtlTeraManager
        Get
            If _instance Is Nothing Then
                _instance = New PtlTeraManager()
            End If
            Return _instance
        End Get
    End Property

    Private _installProject As InstallProject = Nothing

    Private _ptlTeraExs As List(Of PtlTeraEx)

    Private _ptlMxp1O5Exs As List(Of PtlMXP1O5Ex)

    Private _xGates As List(Of XGateEx)

    Private Sub New()
        _ptlTeraExs = New List(Of PtlTeraEx)()
        _ptlMxp1O5Exs = New List(Of PtlMXP1O5Ex)()
        _xGates = New List(Of XGateEx)
        Init()
        Log.Logger.Instance = New CustomLogger()
    End Sub

    Private Sub Init()
        Try
            Dim racks As List(Of T_InvRack) = _wms.GetAllRacks()
            If (racks IsNot Nothing) Then
                For Each rack In racks
                    Dim ledBus As Integer = -1, index As Integer = -1, device As Integer = -1
                    Try
                        Dim LEDAddr() As String
                        Dim RackAddr() As String
                        Dim doIndex As Integer = 0
                        Dim listDevices = rack.T_InvSlot.GroupBy(Function(o) o.LEDAddr.Substring(0, 5)).ToList()
                        If (listDevices Is Nothing OrElse listDevices.Count = 0) Then
                            Dim msg = "Slot still not set up in system"
                            ErrorLogging("WS-WMS-LightByRack", rack.Rack, msg, "I")
                            ''Throw New NotSupportedException(msg)
                        End If

                        If (String.IsNullOrEmpty(rack.LEDAddr)) Then
                            Dim msg = "Big light still not set up in system"
                            ErrorLogging("WS-WMS-LightByRack", rack.Rack, msg, "I")
                            ''Throw New NotSupportedException(msg)
                        End If

                        '' 实例化XGate对象
                        Dim gate = GetXGate(rack.ControlPC)

                        For Each item In listDevices
                            LEDAddr = item.Key.Split("-")
                            ledBus = Integer.Parse(LEDAddr(0))
                            device = Integer.Parse(LEDAddr(1))

                            '' 实例化灯条对象
                            Dim tera = GetTera(rack.ControlPC, ledBus, device, rack.Rack, 0)
                            If Not gate.Buses(ledBus).Devices.Contains(tera) Then
                                gate.Buses(ledBus).Devices.AddOrUpdate(tera)
                            End If
                        Next
                        '' TO be comment
                        RackAddr = rack.LEDAddr.Split("-")
                        Dim mxp1O5Address = Integer.Parse(RackAddr(1))
                        If (RackAddr.Length = 3) Then
                            doIndex = Integer.Parse(RackAddr(2))
                        End If

                        '' 实例化大灯设备
                        Dim mxp1O5 = GetMXP1O5(rack.ControlPC, ledBus, mxp1O5Address, rack.Rack, doIndex)

                        If Not gate.Buses(ledBus).Devices.Contains(mxp1O5) Then
                            gate.Buses(ledBus).Devices.AddOrUpdate(mxp1O5)
                        End If

                    Catch ex As Exception
                        ErrorLogging("WS-WMS-Init", rack.Rack, ex.Message & ex.Source, "E")
                        [Stop]()
                    End Try

                Next

                '' 启动XGate任务队列
                For Each item In _xGates
                    If (Not item.IsStart) Then
                        If (item.XGate.CommandCount <> 0) Then
                            item.XGate.ClearCommandQueue()
                        End If
                        item.XGate.StartUnicastCommandQueue()
                        item.IsStart = True
                    End If
                Next
            End If

        Catch ex As Exception
            ''Dim msg = rack.Rack + ";" + rack.ControlPC + ";"
            ErrorLogging("WS-WMS-Init", "", ex.Message & ex.Source, "E")
            [Stop]()
        End Try
    End Sub

    ''' <summary>
    ''' 结束料架控制
    ''' </summary>
    Public Sub [Stop]()

        For Each item In _ptlTeraExs
            item.StopThread()
        Next
        ''For Each item In _xGates
        ''    item.IsStart = False
        ''    item.XGate.StopUnicastCommandQueue()
        ''    item.XGate = Nothing
        ''Next
    End Sub

    Private Function GetXGate(ByVal controlPC As String) As XGate
        Dim gate = Nothing
        Try
            Dim existedGate = _xGates.FirstOrDefault(Function(o) String.Equals(o.XGate.IPEndPoint.Address.ToString, controlPC))
            If (existedGate Is Nothing) Then
                gate = New XGate(controlPC)
                _xGates.Add(New XGateEx(gate))
            Else
                gate = existedGate.XGate
            End If
        Catch ex As Exception
            ErrorLogging("WS-WMS-GetXGate", controlPC, ex.Message & ex.Source, "E")
        End Try

        Return gate
    End Function

    Private Function GetTera(ByVal controlPC As String, ByVal ledBus As String, ByVal device As String, ByVal rack As String, ByVal action As Integer, Optional indexArray As List(Of Integer) = Nothing) As PtlTera
        Try
            Dim existedTera = _ptlTeraExs.FirstOrDefault(Function(x) x.Ip = controlPC And x.BusIndex = ledBus And x.BusAddress = device)
            If (existedTera Is Nothing) Then
                '' 实例化灯条对象
                Dim tera = New PtlTera() With {.Address = device, .MinorType = PtlTeraType.A_B04002}
                existedTera = New PtlTeraEx(controlPC, ledBus, device, tera) With {.Rack = rack}
                If (indexArray IsNot Nothing) Then
                    existedTera.LightLedIndex = indexArray.ToList()
                End If
                If (action = 2) Then
                    For Each item In indexArray
                        If (Not existedTera.FlashLedIndex.Contains(item)) Then
                            existedTera.FlashLedIndex.Add(item)
                        End If

                    Next
                End If
                _ptlTeraExs.Add(existedTera)
                Return tera
            Else
                If (indexArray IsNot Nothing) Then
                    existedTera.LightLedIndex = indexArray.ToList()
                End If
                If (action = 2) Then
                    For Each item In indexArray
                        If (Not existedTera.FlashLedIndex.Contains(item)) Then
                            existedTera.FlashLedIndex.Add(item)
                        End If

                    Next
                End If
                Return existedTera.PtlTera
            End If
        Catch ex As Exception
            Dim msg = ledBus + ";" + device + ";" + rack
            ErrorLogging("WS-WMS-GetTera", controlPC, msg + ex.Message & ex.Source, "E")
        End Try
        Return Nothing
    End Function

    Private Function GetMXP1O5(ByVal controlPC As String, ByVal ledBus As String, ByVal device As String, ByVal rack As String, ByVal doIndex As Integer) As PtlMXP1O5
        Try
            Dim existedMXP1O5 = _ptlMxp1O5Exs.FirstOrDefault(Function(o) o.Ip = controlPC And o.BusIndex = ledBus And o.BusAddress = device And o.DoIndex = doIndex)
            If (existedMXP1O5 Is Nothing) Then
                '' 实例化大灯设备
                existedMXP1O5 = _ptlMxp1O5Exs.FirstOrDefault(Function(o) o.Ip = controlPC And o.BusIndex = ledBus And o.BusAddress = device)
                Dim mxp1O5 As PtlMXP1O5 = Nothing
                If (existedMXP1O5 Is Nothing) Then
                    mxp1O5 = New PtlMXP1O5() With {.Address = device, .MinorType = PtlMXP1O5Type.M3}
                Else
                    mxp1O5 = existedMXP1O5.PtlMXP1O5
                End If

                Dim mxp1O5Ex = New PtlMXP1O5Ex(controlPC, ledBus, device, mxp1O5) With {.Rack = rack, .DoIndex = doIndex}
                ''mxp1O5Ex.Racks.Add(doIndex, rack)
                _ptlMxp1O5Exs.Add(mxp1O5Ex)
                Return mxp1O5
            End If
            Return existedMXP1O5.PtlMXP1O5
        Catch ex As Exception
            Dim msg = ledBus + ";" + device + ";" + doIndex + ";" + rack
            ErrorLogging("WS-WMS-GetMXP1O5", controlPC, msg + ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Sub LightByRack(action As Integer, rack As T_InvRack)
        Dim ledBus As Integer = -1, index As Integer = -1, device As Integer = -1
        Try
            ''Dim LEDAddr() As String
            Dim listDevices = rack.T_InvSlot.GroupBy(Function(o) o.LEDAddr.Substring(0, 5)).ToList()
            If (listDevices Is Nothing OrElse listDevices.Count = 0) Then
                Dim msg = "Slot still not set up in system"
                ErrorLogging("WS-WMS-LightByRack", rack.Rack, msg, "I")
                Throw New NotSupportedException(msg)
            End If

            If (String.IsNullOrEmpty(rack.LEDAddr)) Then
                Dim msg = "Big light still not set up in system"
                ErrorLogging("WS-WMS-LightByRack", rack.Rack, msg, "I")
                Throw New NotSupportedException(msg)
            End If

            Dim mXP1O5ByRack = _ptlMxp1O5Exs.FirstOrDefault(Function(o) o.Rack = rack.Rack)
            If (mXP1O5ByRack IsNot Nothing) Then
                mXP1O5ByRack.Light(action)
            End If

            Dim terasByRack = _ptlTeraExs.Where(Function(o) o.Rack = rack.Rack).ToList()
            For Each item In terasByRack
                item.Light(action, _allLedIndexs)
            Next
        Catch ex As Exception
            ErrorLogging("WS-WMS-LightByRack", rack.Rack, ex.Message & ex.Source, "E")
            [Stop]()
        End Try
    End Sub

    Public Sub LightAllRacks(action As Integer, racks As List(Of T_InvRack))
        Dim rack As T_InvRack = Nothing
        Try
            For Each rack In racks
                ErrorLogging("WS-WMS-LightAllRacks", action, rack.Rack, "I")
                LightByRack(action, rack)
            Next
        Catch ex As Exception
            Dim msg = rack.Rack + ";" + rack.ControlPC + ";"
            ErrorLogging("WS-WMS-LightAllRacks", msg, ex.Message & ex.Source, "E")
            [Stop]()
        End Try
    End Sub

    Public Sub LightBySlot(action As Integer, racks As List(Of T_InvRack), slots As List(Of T_InvSlot), interval As Integer)
        Dim msg = String.Empty
        Try
            Dim indexArray = New List(Of Integer)()
            Dim gate As XGate = Nothing

            '' 获取第一个Slot总线，设备号，用于循环比较
            Dim ledAddrArray = slots(0).LEDAddr.Split("-")
            Dim busIndex = Integer.Parse(ledAddrArray(0))
            Dim address = -1
            Dim ledIndex = Integer.Parse(ledAddrArray(2))
            Dim tera As PtlTera = Nothing
            Dim teraAddr = Integer.Parse(slots(0).LEDAddr.Split("-")(1))
            Dim controlPC As String = String.Empty
            Dim nextControlPC As String = String.Empty

            For i = 0 To slots.Count - 1
                ledAddrArray = slots(i).LEDAddr.Split("-")
                busIndex = Integer.Parse(ledAddrArray(0))
                address = Integer.Parse(ledAddrArray(1))
                ledIndex = Integer.Parse(ledAddrArray(2))

                Dim rackID = slots(i).Rack
                controlPC = racks.Find(Function(o) o.Rack = rackID).ControlPC
                If (String.IsNullOrEmpty(controlPC)) Then
                    ErrorLogging("WS-WMS-LightBySlot", slots(i).Slot + ";" + slots(i).LEDAddr + ";" + slots(i).Rack, msg, "I")
                    Throw New NotSupportedException(msg)
                End If

                '' 遍历到最后一个Slot做特殊处理
                If (i + 1 < slots.Count) Then
                    teraAddr = Integer.Parse(slots(i + 1).LEDAddr.Split("-")(1))
                    rackID = slots(i + 1).Rack
                    nextControlPC = racks.Find(Function(o) o.Rack = rackID).ControlPC
                Else
                    teraAddr = -1
                End If

                '' 对比每个Slot如果属于同一个设备，记录led对于的索引位置
                If (teraAddr = address And String.Equals(controlPC, nextControlPC)) Then
                    indexArray.Add(ledIndex)
                    Continue For
                End If
                '' 保存led索引位置
                indexArray.Add(ledIndex)
                '' 实例化XGate对象
                ''gate = GetXGate(controlPC)
                tera = GetTera(controlPC, busIndex, address, rackID, action, indexArray)
                ''If Not gate.Buses(busIndex).Devices.Contains(tera) Then
                ''    gate.Buses(busIndex).Devices.AddOrUpdate(tera)
                ''End If

                '' 清空，为记录下一个灯条的索引位置做准备
                indexArray.Clear()
            Next

            Dim terasByRack = _ptlTeraExs.Where(Function(o) racks.Exists(Function(y) y.Rack = o.Rack)).ToList()
            If (action = 0 Or action = 1) Then
                For Each item In terasByRack
                    If item.LightLedIndex IsNot Nothing Then
                        '' 灯组亮灭控制
                        If item.LightLedIndex.Count > 0 Then
                            ''ErrorLogging("WS-XGateManager-LightBySlot", item.Rack, action, "I")
                            item.Light(action, item.LightLedIndex.ToArray(), interval)
                        End If
                    End If


                Next
            ElseIf (action = 2) Then
                For Each item In terasByRack
                    '' 灯组闪亮控制
                    If item.LightLedIndex IsNot Nothing Then
                        If item.LightLedIndex.Count > 0 Then
                            item.Flash(action, item.LightLedIndex, interval)
                        End If
                    End If
                Next
            End If


            Dim mxp1O5Exs = _ptlMxp1O5Exs.Where(Function(o) racks.Exists(Function(y) y.Rack = o.Rack)).ToList()
            If (interval = 0) Then
                '' 实例化一个WMS对象

                For Each item In mxp1O5Exs
                    '' 大灯亮灭控制
                    If (action = 0) Then
                        '' 检查是否有小灯亮着，如果没有熄灭大灯
                        msg = action.ToString() + ";" + interval.ToString()
                        ''ErrorLogging("WS-XGateManager-LightBySlot", msg, item.Rack, "I")
                        If (Not _wms.ExistLedLightOnByRack(item.Rack)) Then
                            ''ErrorLogging("WS-XGateManager-LightBySlot", msg, item.Rack, "I")
                            item.Light(action)
                        End If
                    Else
                        item.Light(action)
                    End If
                Next
            End If

        Catch ex As Exception
            ''Dim msg As String = String.Empty
            msg = String.Empty
            For Each item In slots
                msg += item.Slot
            Next

            ErrorLogging("WS-XGateManager-LightBySlot", msg, ex.Message & ex.Source, "E")
            Me.Stop()
        End Try
    End Sub

    ''' <summary>
    ''' 检查连接状态
    ''' </summary>
    ''' <param name="xgateIp"></param>
    ''' <param name="busIndex"></param>
    ''' <returns></returns>
    Public Function CheckXgate(xgateIp As String, busIndex As Integer) As Boolean

        Try
            If (String.IsNullOrEmpty(xgateIp)) Then
                ErrorLogging("WS-WMS-CheckXgate", xgateIp + busIndex.ToString(), "IP can't be empty", "I")
                Return False
            End If

            If (String.IsNullOrEmpty(busIndex)) Then
                ErrorLogging("WS-WMS-CheckXgate", xgateIp + busIndex.ToString(), "Bus index can't be empty", "I")
                Return False
            End If

            Dim ip As IPAddress = Nothing
            If (Not IPAddress.TryParse(xgateIp.Trim(), ip)) Then
                ErrorLogging("WS-WMS-CheckXgate", xgateIp + busIndex.ToString(), "IP format is not correct", "I")
                Return False
            End If

            Dim xGate = New XGate(xgateIp)
            If xGate IsNot Nothing Then
                If xGate.Buses(CByte(busIndex)).CommunicationClient.Connected = True Then
                    Return True
                End If
            End If
        Catch ex As Exception
            ErrorLogging("WS-WMS-CheckXgate", xgateIp + busIndex.ToString(), ex.Message & ex.Source, "E")
        End Try

        Return False
    End Function

    Public Sub Reboot(ByVal xgateIP As String)
        Try
            Dim gate = GetXGate(xgateIP)
            If gate IsNot Nothing Then
                gate.SettingsExecutor.Reboot()
            End If
        Catch ex As Exception
            ErrorLogging("WS-WMS-xgateIP", xgateIP, ex.Message & ex.Source, "E")
        End Try

    End Sub
End Class

Class PtlTeraEx
    Inherits PublicFunction
    ''Private ReadOnly _playColors As LightColor() = New LightColor(99) {}
    Private _playModes As IList(Of LightMode) = New List(Of LightMode)(New LightMode() {New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(), New LightMode(),
                                                                       New LightMode()})
    Private _isBeginFlash As Boolean = False
    Public Sub New(ip__1 As String, busIndex__2 As Byte, busAddress__3 As Byte, ptlTera__4 As PtlTera)
        Ip = ip__1
        BusIndex = busIndex__2
        BusAddress = busAddress__3
        PtlTera = ptlTera__4
        FlashLedIndex = New List(Of Integer)
    End Sub

    ''' <summary>
    ''' 所属控制器IP
    ''' </summary>
    Public Property Ip() As String
        Get
            Return m_Ip
        End Get
        Private Set(value As String)
            m_Ip = value
        End Set
    End Property
    Private m_Ip As String
    ''' <summary>
    ''' 所属控制器总线序号
    ''' </summary>
    Public Property BusIndex() As Byte
        Get
            Return m_BusIndex
        End Get
        Private Set(value As Byte)
            m_BusIndex = value
        End Set
    End Property
    Private m_BusIndex As Byte
    ''' <summary>
    ''' 总线地址
    ''' </summary>
    Public Property BusAddress() As Byte
        Get
            Return m_BusAddress
        End Get
        Private Set(value As Byte)
            m_BusAddress = value
        End Set
    End Property
    Private m_BusAddress As Byte
    Public Property Rack() As String
        Get
            Return m_Rack
        End Get
        Set(value As String)
            m_Rack = value
        End Set
    End Property
    Private m_Rack As String
    Public Property LightLedIndex() As List(Of Integer)
        Get
            Return m_LightLedIndex
        End Get
        Set(value As List(Of Integer))
            m_LightLedIndex = value
        End Set
    End Property
    Private m_LightLedIndex As List(Of Integer)

    Public Property FlashLedIndex() As List(Of Integer)
        Get
            Return m_FlashLedIndex
        End Get
        Set(value As List(Of Integer))
            m_FlashLedIndex = value
        End Set
    End Property
    Private m_FlashLedIndex As List(Of Integer)

    ''' <summary>
    ''' 料架灯
    ''' </summary>
    Public Property PtlTera() As PtlTera
        Get
            Return m_PtlTera
        End Get
        Private Set(value As PtlTera)
            m_PtlTera = value
        End Set
    End Property
    Private m_PtlTera As PtlTera

    Public Property IsAdded() As Boolean
        Get
            Return m_IsAdded
        End Get
        Set(value As Boolean)
            m_IsAdded = value
        End Set
    End Property
    Private m_IsAdded As Boolean

    Public Sub Light(action As Integer, indexArray As Integer(), Optional second As Integer = 0)
        Try

            ''ErrorLogging("WS-XGateManager-Light", action, indexArray.Count, "I")
            For Each index As Integer In indexArray
                If index < 0 OrElse index > 100 Then
                    Throw New NotSupportedException("led index " + index + " is error!")
                End If
                If action = 1 Then
                    '亮
                    ''_playColors(index) = LightColor.Red
                    _playModes(index).Color = LightColor.Red
                    _playModes(index).Period = Nothing
                    _playModes(index).Ratio = Nothing
                ElseIf action = 0 Then
                    '灭
                    ''_playColors(index) = LightColor.Off
                    _playModes(index).Color = LightColor.Off
                    _playModes(index).Period = Nothing
                    _playModes(index).Ratio = Nothing
                End If
                FlashLedIndex.Remove(index)

                ''ErrorLogging("WS-XGateManager-Light", action, index, "I")
            Next

            PtlTera.Display(_playModes)
            If (second = 0) Then

                If (LightLedIndex IsNot Nothing) Then
                    LightLedIndex.Clear()
                End If
            End If

            If (FlashLedIndex IsNot Nothing AndAlso FlashLedIndex.Count = 0) Then
                _isBeginFlash = False
            End If

            '' 亮second后把等熄灭
            If second > 0 Then
                LightSomeTime(action, indexArray, second)
            End If

        Catch ex As Exception

            ErrorLogging("WS-XGateManager-Light", "", ex.Message & ex.Source, "E")
        End Try


    End Sub

    Public Sub LightSomeTime(action As Integer, indexArray As Integer(), second As Integer)
        If (action = 1 And second > 0) Then
            Task.Factory.StartNew(Function()
                                      Thread.Sleep(TimeSpan.FromSeconds(second))

                                      For Each index As Integer In indexArray
                                          _playModes(index).Color = LightColor.Off
                                          _playModes(index).Period = Nothing
                                          _playModes(index).Ratio = Nothing
                                      Next
                                      PtlTera.Display(_playModes)
                                      If (LightLedIndex IsNot Nothing) Then
                                          LightLedIndex.Clear()
                                      End If
                                      Return 0
                                  End Function)
        End If
    End Sub

    Public Sub Flash(action As Integer, indexArray As List(Of Integer), Optional second As Integer = 0)
        Try
            For Each item In indexArray
                If item < 0 OrElse item > 100 Then
                    Throw New NotSupportedException("led index is error!")
                End If
            Next

            If (action = 2) Then
                If (second = 0) Then
                    LightLedIndex.Clear()
                    If FlashLedIndex IsNot Nothing AndAlso FlashLedIndex.Count > 0 Then
                        For Each item In FlashLedIndex
                            _playModes(item).Color = LightColor.Red
                            _playModes(item).Period = LightOnOffPeriod.Period500
                            _playModes(item).Ratio = LightOnOffRatio.RatioP1V1
                        Next
                        PtlTera.Display(_playModes)
                    End If

                ElseIf (second > 0) Then
                    Dim cnt = 0
                    Task.Factory.StartNew(Function()

                                              Try
                                                  If FlashLedIndex IsNot Nothing AndAlso FlashLedIndex.Count > 0 Then
                                                      For Each item In indexArray
                                                          FlashLedIndex.Remove(item)
                                                      Next
                                                  End If

                                                  ''For Each index As Integer In indexArray
                                                  ''    _playModes(index).Color = LightColor.Red
                                                  ''    _playModes(index).Period = LightOnOffPeriod.Period1000
                                                  ''    _playModes(index).Ratio = LightOnOffRatio.RatioP1V1
                                                  ''Next
                                                  ''PtlTera.Display(_playModes)

                                                  While cnt < second

                                                      For Each index As Integer In indexArray
                                                          _playModes(index).Color = LightColor.Red
                                                          _playModes(index).Period = Nothing
                                                          _playModes(index).Ratio = Nothing
                                                      Next
                                                      PtlTera.Display(_playModes)
                                                      Thread.Sleep(500)

                                                      For Each index As Integer In indexArray
                                                          _playModes(index).Color = LightColor.Off
                                                          _playModes(index).Period = Nothing
                                                          _playModes(index).Ratio = Nothing
                                                      Next
                                                      PtlTera.Display(_playModes)
                                                      Thread.Sleep(500)
                                                      cnt += 1
                                                  End While
                                                  ''Thread.Sleep(TimeSpan.FromSeconds(second))
                                                  ''For Each index As Integer In indexArray
                                                  ''    _playModes(index).Color = LightColor.Off
                                                  ''    _playModes(index).Period = Nothing
                                                  ''    _playModes(index).Ratio = Nothing
                                                  ''Next
                                                  ''PtlTera.Display(_playModes)

                                                  LightLedIndex.Clear()
                                              Catch

                                              End Try
                                              Return 0
                                          End Function)

                End If
            End If
        Catch ex As Exception
            ErrorLogging("WS-XGateManager-Flash", "", ex.Message & ex.Source, "E")
        End Try
    End Sub
    Public Sub StopThread()
        Me._isBeginFlash = False
    End Sub

End Class

Class PtlMXP1O5Ex
    Public Sub New(ip__1 As String, busIndex__2 As Byte, busAddress__3 As Byte, ptlMxp1O5__4 As PtlMXP1O5)
        Ip = ip__1
        BusIndex = busIndex__2
        BusAddress = busAddress__3
        PtlMXP1O5 = ptlMxp1O5__4
    End Sub

    ''' <summary>
    ''' 所属控制器IP
    ''' </summary>
    Public Property Ip() As String
        Get
            Return m_Ip
        End Get
        Private Set(value As String)
            m_Ip = value
        End Set
    End Property
    Private m_Ip As String
    ''' <summary>
    ''' 所属控制器总线序号
    ''' </summary>
    Public Property BusIndex() As Byte
        Get
            Return m_BusIndex
        End Get
        Private Set(value As Byte)
            m_BusIndex = value
        End Set
    End Property
    Private m_BusIndex As Byte
    ''' <summary>
    ''' 总线地址
    ''' </summary>
    Public Property BusAddress() As Byte
        Get
            Return m_BusAddress
        End Get
        Private Set(value As Byte)
            m_BusAddress = value
        End Set
    End Property
    Private m_BusAddress As Byte
    ''' <summary>
    ''' Rack地址
    ''' </summary>
    Public Property Rack() As String
        Get
            Return m_Rack
        End Get
        Set(value As String)
            m_Rack = value
        End Set
    End Property
    Private m_Rack As String

    Public Property IsAdded() As Boolean
        Get
            Return m_IsAdded
        End Get
        Set(value As Boolean)
            m_IsAdded = value
        End Set
    End Property
    Private m_IsAdded As Boolean

    Public Property DoIndex() As Byte
        Get
            Return m_DoAddress
        End Get
        Set(value As Byte)
            m_DoAddress = value
        End Set
    End Property
    Private m_DoAddress As Byte

    ''' <summary>
    ''' 料架灯
    ''' </summary>
    Public Property PtlMXP1O5() As PtlMXP1O5
        Get
            Return m_PtlMXP1O5
        End Get
        Private Set(value As PtlMXP1O5)
            m_PtlMXP1O5 = value
        End Set
    End Property
    Private m_PtlMXP1O5 As PtlMXP1O5

    Public Sub Light(action As Integer)
        If PtlMXP1O5 IsNot Nothing Then
            If action = 1 Or action = 2 Then
                PtlMXP1O5.Lighthouses(DoIndex).Display()
                ''PtlMXP1O5.Lighthouses(0).Display()
                ''PtlMXP1O5.Lighthouses(1).Display()
                ''PtlMXP1O5.Lighthouses(2).Display()
                ''PtlMXP1O5.Lighthouses(3).Display()
                ''PtlMXP1O5.Lighthouses(4).Display()

            ElseIf action = 0 Then
                PtlMXP1O5.Lighthouses(DoIndex).Clear()

                ''PtlMXP1O5.Lighthouses(0).Clear()
                ''PtlMXP1O5.Lighthouses(1).Clear()
                ''PtlMXP1O5.Lighthouses(2).Clear()
                ''PtlMXP1O5.Lighthouses(3).Clear()
                ''PtlMXP1O5.Lighthouses(4).Clear()
            End If
        End If
    End Sub
End Class

Class XGateEx

    Public Sub New(_xgate As XGate)
        XGate = _xgate
        IsStart = False
    End Sub
    Public Property XGate() As XGate
        Get
            Return _xgate
        End Get
        Set(value As XGate)
            _xgate = value
        End Set
    End Property
    Private _xgate As XGate

    Public Property IsStart() As Boolean
        Get
            Return m_IsStart
        End Get
        Set(value As Boolean)
            m_IsStart = value
        End Set
    End Property
    Private m_IsStart As Boolean
End Class

Class CustomLogger
    Inherits PublicFunction
    Implements Log.ILogger

    Public Sub Write(source As String, message As String, level As Log.LogLevel) Implements Log.ILogger.Write
        Dim msg = String.Format("{0}:{1}", source, message)
        Dim category As String = "I"
        If (level > 1) Then
            category = "E"
            'ErrorLogging("WS-WMS-XGate-Write", level, msg, category)
        End If
    End Sub

    Public Sub Delete(expiryDate As DateTime) Implements Log.ILogger.Delete

    End Sub

    Public Property IsInformationEnabled() As Boolean Implements Log.ILogger.IsInformationEnabled
        Get
            Return m_IsInformationEnabled
        End Get
        Set(value As Boolean)
            m_IsInformationEnabled = value
        End Set
    End Property
    Private m_IsInformationEnabled As Boolean
    Public Property IsWarningEnabled() As Boolean Implements Log.ILogger.IsWarningEnabled
        Get
            Return m_IsWarningEnabled
        End Get
        Set(value As Boolean)
            m_IsWarningEnabled = value
        End Set
    End Property
    Private m_IsWarningEnabled As Boolean
    Public Property IsErrorEnabled() As Boolean Implements Log.ILogger.IsErrorEnabled
        Get
            Return m_IsErrorEnabled
        End Get
        Set(value As Boolean)
            m_IsErrorEnabled = value
        End Set
    End Property
    Private m_IsErrorEnabled As Boolean
End Class

