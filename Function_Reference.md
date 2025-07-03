# TBIOCP 函数参考手册

## 目录

1. [核心服务器函数](#核心服务器函数)
2. [连接管理函数](#连接管理函数)
3. [数据传输函数](#数据传输函数)
4. [配置和属性函数](#配置和属性函数)
5. [事件处理函数](#事件处理函数)
6. [错误处理函数](#错误处理函数)
7. [性能监控函数](#性能监控函数)
8. [实用工具函数](#实用工具函数)

---

## 核心服务器函数

### StartServer

启动 IOCP 服务器并开始监听指定端口。

**函数签名**:
```vb
Public Function StartServer(ByVal Port As Long, Optional ByVal MaxConnections As Long = 1000) As Boolean
```

**参数**:
| 参数名 | 类型 | 描述 | 默认值 |
|--------|------|------|---------|
| Port | Long | 监听端口号 (1-65535) | 必需 |
| MaxConnections | Long | 最大同时连接数 | 1000 |

**返回值**:
- `True`: 服务器启动成功
- `False`: 服务器启动失败

**可能的错误**:
- 端口已被占用
- 权限不足
- 网络初始化失败

**使用示例**:
```vb
Dim server As TBIOCPServer
Set server = New TBIOCPServer

' 基础用法
If server.StartServer(8080) Then
    Debug.Print "服务器在端口 8080 启动成功"
End If

' 指定最大连接数
If server.StartServer(9999, 2000) Then
    Debug.Print "服务器启动成功，最大连接数: 2000"
End If

' 错误处理
If Not server.StartServer(80) Then
    Select Case server.LastError
        Case -1
            MsgBox "端口 80 已被占用"
        Case -2
            MsgBox "需要管理员权限"
        Case Else
            MsgBox "启动失败: " & server.GetErrorDescription()
    End Select
End If
```

---

### StopServer

停止 IOCP 服务器并清理所有资源。

**函数签名**:
```vb
Public Sub StopServer()
```

**参数**: 无

**返回值**: 无

**行为**:
- 停止接受新连接
- 断开所有现有连接
- 释放所有工作线程
- 清理内存资源

**使用示例**:
```vb
' 优雅停止服务器
If server.IsRunning() Then
    Debug.Print "正在停止服务器..."
    server.StopServer()
    Debug.Print "服务器已停止"
End If

' 在程序退出时清理
Private Sub Form_Unload(Cancel As Integer)
    If Not server Is Nothing Then
        server.StopServer()
        Set server = Nothing
    End If
End Sub
```

---

### IsRunning

检查服务器是否正在运行。

**函数签名**:
```vb
Public Function IsRunning() As Boolean
```

**参数**: 无

**返回值**:
- `True`: 服务器正在运行
- `False`: 服务器已停止

**使用示例**:
```vb
' 检查服务器状态
If server.IsRunning() Then
    lblStatus.Caption = "服务器运行中"
    btnStart.Enabled = False
    btnStop.Enabled = True
Else
    lblStatus.Caption = "服务器已停止"
    btnStart.Enabled = True
    btnStop.Enabled = False
End If

' 条件性重启
If Not server.IsRunning() Then
    server.StartServer(8080)
End If
```

---

## 连接管理函数

### GetConnectionCount

获取当前活跃连接数。

**函数签名**:
```vb
Public Function GetConnectionCount() As Long
```

**参数**: 无

**返回值**: 当前活跃连接数 (Long)

**使用示例**:
```vb
' 显示连接统计
Private Sub Timer1_Timer()
    Dim count As Long
    count = server.GetConnectionCount()
    lblConnections.Caption = "连接数: " & count
    
    ' 连接数报警
    If count > 1500 Then
        lblConnections.ForeColor = vbRed
        LogWarning "连接数过高: " & count
    Else
        lblConnections.ForeColor = vbBlack
    End If
End Sub

' 负载均衡决策
If server.GetConnectionCount() > MaxConnectionsPerServer Then
    RedirectNewConnections()
End If
```

---

### DisconnectClient

断开指定客户端的连接。

**函数签名**:
```vb
Public Sub DisconnectClient(ByVal ConnectionID As Long)
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| ConnectionID | Long | 要断开的连接 ID |

**返回值**: 无

**行为**:
- 优雅关闭连接
- 清理连接相关资源
- 触发 OnClientDisconnect 事件

**使用示例**:
```vb
' 踢出特定用户
Private Sub KickUser(ByVal UserID As Long)
    server.DisconnectClient(UserID)
    LogAction "用户 " & UserID & " 已被踢出"
End Sub

' 批量断开连接
Private Sub DisconnectAllClients()
    Dim connections As Collection
    Set connections = server.GetConnectionList()
    
    Dim i As Integer
    For i = 1 To connections.Count
        server.DisconnectClient(connections(i))
    Next i
End Sub

' 超时连接清理
Private Sub CleanupInactiveConnections()
    Dim connections As Collection
    Set connections = server.GetInactiveConnections(300) ' 5分钟超时
    
    Dim connID As Long
    For Each connID In connections
        server.DisconnectClient(connID)
        LogInfo "断开非活跃连接: " & connID
    Next
End Sub
```

---

### GetConnectionInfo

获取指定连接的详细信息。

**函数签名**:
```vb
Public Function GetConnectionInfo(ByVal ConnectionID As Long) As ConnectionInfo
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| ConnectionID | Long | 连接 ID |

**返回值**: ConnectionInfo 结构体

**ConnectionInfo 结构**:
```vb
Public Type ConnectionInfo
    ConnectionID As Long        ' 连接 ID
    RemoteIP As String         ' 客户端 IP
    RemotePort As Long         ' 客户端端口
    ConnectTime As Date        ' 连接时间
    LastActiveTime As Date     ' 最后活跃时间
    BytesSent As Long          ' 已发送字节数
    BytesReceived As Long      ' 已接收字节数
    State As ConnectionState   ' 连接状态
End Type
```

**使用示例**:
```vb
' 显示连接详情
Private Sub ShowConnectionDetails(ByVal ConnectionID As Long)
    Dim info As ConnectionInfo
    info = server.GetConnectionInfo(ConnectionID)
    
    If info.ConnectionID <> 0 Then
        Dim details As String
        details = "连接 ID: " & info.ConnectionID & vbCrLf
        details = details & "IP 地址: " & info.RemoteIP & ":" & info.RemotePort & vbCrLf
        details = details & "连接时间: " & info.ConnectTime & vbCrLf
        details = details & "最后活跃: " & info.LastActiveTime & vbCrLf
        details = details & "已发送: " & FormatBytes(info.BytesSent) & vbCrLf
        details = details & "已接收: " & FormatBytes(info.BytesReceived)
        
        MsgBox details, vbInformation, "连接信息"
    Else
        MsgBox "连接 " & ConnectionID & " 不存在", vbExclamation
    End If
End Sub

' 监控连接状态
Private Sub MonitorConnection(ByVal ConnectionID As Long)
    Dim info As ConnectionInfo
    info = server.GetConnectionInfo(ConnectionID)
    
    ' 检查连接健康状态
    Dim inactiveTime As Long
    inactiveTime = DateDiff("s", info.LastActiveTime, Now)
    
    If inactiveTime > 300 Then ' 5分钟无活动
        LogWarning "连接 " & ConnectionID & " 可能已死连接"
        server.DisconnectClient(ConnectionID)
    End If
End Sub
```

---

## 数据传输函数

### SendData

向指定客户端发送数据。

**函数签名**:
```vb
Public Function SendData(ByVal ConnectionID As Long, ByVal Data As String) As Boolean
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| ConnectionID | Long | 目标连接 ID |
| Data | String | 要发送的数据 |

**返回值**:
- `True`: 数据发送成功
- `False`: 数据发送失败

**使用示例**:
```vb
' 基础数据发送
Private Sub SendMessage(ByVal ConnectionID As Long, ByVal Message As String)
    If server.SendData(ConnectionID, Message) Then
        LogInfo "消息发送成功: " & Message
    Else
        LogError "消息发送失败: " & server.GetLastError()
    End If
End Sub

' 发送 JSON 数据
Private Sub SendJSON(ByVal ConnectionID As Long, ByVal JsonData As String)
    Dim packet As String
    packet = "Content-Type: application/json" & vbCrLf
    packet = packet & "Content-Length: " & Len(JsonData) & vbCrLf & vbCrLf
    packet = packet & JsonData
    
    server.SendData ConnectionID, packet
End Sub

' 文件传输
Private Sub SendFile(ByVal ConnectionID As Long, ByVal FilePath As String)
    Dim fileData As String
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open FilePath For Binary As #fileNum
    fileData = Space$(LOF(fileNum))
    Get #fileNum, , fileData
    Close #fileNum
    
    ' 分块发送大文件
    Dim chunkSize As Long: chunkSize = 8192
    Dim offset As Long: offset = 1
    
    While offset <= Len(fileData)
        Dim chunk As String
        chunk = Mid$(fileData, offset, chunkSize)
        
        If Not server.SendData(ConnectionID, chunk) Then
            LogError "文件传输失败在偏移: " & offset
            Exit Sub
        End If
        
        offset = offset + chunkSize
    Wend
    
    LogInfo "文件发送完成: " & FilePath
End Sub
```

---

### SendDataToAll

向所有连接的客户端广播数据。

**函数签名**:
```vb
Public Function SendDataToAll(ByVal Data As String) As Long
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| Data | String | 要广播的数据 |

**返回值**: 成功发送的客户端数量 (Long)

**使用示例**:
```vb
' 系统广播
Private Sub BroadcastSystemMessage(ByVal Message As String)
    Dim timestamp As String
    timestamp = Format(Now, "hh:mm:ss")
    
    Dim broadcastData As String
    broadcastData = "[系统消息 " & timestamp & "] " & Message
    
    Dim successCount As Long
    successCount = server.SendDataToAll(broadcastData)
    
    LogInfo "系统消息已广播给 " & successCount & " 个客户端"
End Sub

' 服务器公告
Private Sub SendServerAnnouncement()
    Dim announcement As String
    announcement = "服务器将在 5 分钟后重启进行维护"
    
    Dim count As Long
    count = server.SendDataToAll(announcement)
    
    If count > 0 Then
        LogInfo "维护公告已发送给 " & count & " 个用户"
        ' 设置5分钟后重启计时器
        Timer2.Interval = 5 * 60 * 1000
        Timer2.Enabled = True
    End If
End Sub

' 实时数据推送
Private Sub PushRealtimeData()
    Dim marketData As String
    marketData = GetCurrentMarketData() ' 获取市场数据
    
    Dim jsonData As String
    jsonData = "{""type"":""market_update"",""data"":" & marketData & "}"
    
    server.SendDataToAll(jsonData)
End Sub
```

---

### SendDataBinary

发送二进制数据（用于非文本数据传输）。

**函数签名**:
```vb
Public Function SendDataBinary(ByVal ConnectionID As Long, ByVal Data() As Byte) As Boolean
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| ConnectionID | Long | 目标连接 ID |
| Data() | Byte | 二进制数据数组 |

**返回值**:
- `True`: 数据发送成功
- `False`: 数据发送失败

**使用示例**:
```vb
' 发送图片数据
Private Sub SendImage(ByVal ConnectionID As Long, ByVal ImagePath As String)
    Dim imageData() As Byte
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open ImagePath For Binary As #fileNum
    ReDim imageData(LOF(fileNum) - 1)
    Get #fileNum, , imageData
    Close #fileNum
    
    ' 发送文件头信息
    Dim header As String
    header = "BINARY_FILE|" & UBound(imageData) + 1 & "|" & GetFileExtension(ImagePath)
    server.SendData ConnectionID, header
    
    ' 发送二进制数据
    If server.SendDataBinary(ConnectionID, imageData) Then
        LogInfo "图片发送成功: " & ImagePath
    Else
        LogError "图片发送失败"
    End If
End Sub

' 发送自定义协议数据包
Private Sub SendProtocolPacket(ByVal ConnectionID As Long, ByVal Command As Integer, ByVal Payload() As Byte)
    Dim packet() As Byte
    ReDim packet(3 + UBound(Payload) + 1)
    
    ' 构建数据包: [命令2字节][长度2字节][负载数据]
    packet(0) = Command And &HFF
    packet(1) = (Command >> 8) And &HFF
    
    Dim payloadLen As Integer
    payloadLen = UBound(Payload) + 1
    packet(2) = payloadLen And &HFF
    packet(3) = (payloadLen >> 8) And &HFF
    
    ' 复制负载数据
    Dim i As Integer
    For i = 0 To UBound(Payload)
        packet(4 + i) = Payload(i)
    Next i
    
    server.SendDataBinary ConnectionID, packet
End Sub
```

---

## 配置和属性函数

### SetMaxConnections

设置服务器最大连接数。

**函数签名**:
```vb
Public Sub SetMaxConnections(ByVal MaxConnections As Long)
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| MaxConnections | Long | 最大连接数 (1-10000) |

**返回值**: 无

**注意**: 只能在服务器停止状态下修改

**使用示例**:
```vb
' 动态调整连接数限制
Private Sub AdjustMaxConnections()
    If Not server.IsRunning() Then
        Dim newMax As Long
        newMax = CLng(txtMaxConnections.Text)
        
        If newMax >= 1 And newMax <= 10000 Then
            server.SetMaxConnections(newMax)
            LogInfo "最大连接数已设置为: " & newMax
        Else
            MsgBox "连接数必须在 1-10000 之间", vbExclamation
        End If
    Else
        MsgBox "请先停止服务器", vbExclamation
    End If
End Sub

' 根据系统资源自动调整
Private Sub AutoConfigureMaxConnections()
    Dim availableMemoryMB As Long
    availableMemoryMB = GetAvailableMemoryMB()
    
    ' 每个连接大约占用 10KB 内存
    Dim recommendedMax As Long
    recommendedMax = (availableMemoryMB * 1024) \ 10
    
    ' 限制在合理范围内
    If recommendedMax > 5000 Then recommendedMax = 5000
    If recommendedMax < 100 Then recommendedMax = 100
    
    server.SetMaxConnections(recommendedMax)
    LogInfo "自动设置最大连接数: " & recommendedMax
End Sub
```

---

### SetWorkerThreads

设置 IOCP 工作线程数。

**函数签名**:
```vb
Public Sub SetWorkerThreads(ByVal ThreadCount As Long)
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| ThreadCount | Long | 工作线程数 (1-64) |

**返回值**: 无

**推荐值**: CPU 核心数 × 2

**使用示例**:
```vb
' 根据 CPU 核心数配置线程
Private Sub ConfigureOptimalThreads()
    Dim cpuCores As Long
    cpuCores = GetCPUCoreCount()
    
    Dim optimalThreads As Long
    optimalThreads = cpuCores * 2
    
    ' 确保在合理范围内
    If optimalThreads < 2 Then optimalThreads = 2
    If optimalThreads > 16 Then optimalThreads = 16
    
    server.SetWorkerThreads(optimalThreads)
    LogInfo "工作线程数设置为: " & optimalThreads
End Sub

' 性能调优配置
Private Sub TunePerformance()
    Dim threads As Long
    
    ' 根据服务器负载类型调整
    Select Case GetServerWorkloadType()
        Case "high_concurrent"  ' 高并发
            threads = GetCPUCoreCount() * 4
        Case "cpu_intensive"    ' CPU 密集
            threads = GetCPUCoreCount()
        Case "io_intensive"     ' IO 密集
            threads = GetCPUCoreCount() * 3
        Case Else              ' 默认
            threads = GetCPUCoreCount() * 2
    End Select
    
    server.SetWorkerThreads(threads)
End Sub
```

---

## 事件处理函数

### RegisterEventCallback

注册事件回调函数。

**函数签名**:
```vb
Public Sub RegisterEventCallback(ByVal EventType As EventType, ByVal CallbackFunction As Long)
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| EventType | EventType | 事件类型 |
| CallbackFunction | Long | 回调函数指针 |

**事件类型**:
- `EVENT_CLIENT_CONNECT`
- `EVENT_CLIENT_DISCONNECT`
- `EVENT_DATA_RECEIVED`
- `EVENT_ERROR_OCCURRED`

**使用示例**:
```vb
' 注册连接事件处理
Private Sub SetupEventHandlers()
    server.RegisterEventCallback EVENT_CLIENT_CONNECT, AddressOf OnClientConnect
    server.RegisterEventCallback EVENT_CLIENT_DISCONNECT, AddressOf OnClientDisconnect
    server.RegisterEventCallback EVENT_DATA_RECEIVED, AddressOf OnDataReceived
    server.RegisterEventCallback EVENT_ERROR_OCCURRED, AddressOf OnServerError
End Sub

' 连接事件处理函数
Private Sub OnClientConnect(ByVal ConnectionID As Long, ByVal RemoteIP As String, ByVal RemotePort As Long)
    LogInfo "新客户端连接: " & RemoteIP & ":" & RemotePort & " (ID: " & ConnectionID & ")"
    
    ' 发送欢迎消息
    server.SendData ConnectionID, "欢迎连接到服务器!"
    
    ' 更新统计信息
    UpdateConnectionStats()
    
    ' 检查连接限制
    If server.GetConnectionCount() > 0.9 * server.GetMaxConnections() Then
        BroadcastSystemMessage "服务器连接数接近上限"
    End If
End Sub

' 断开连接事件处理
Private Sub OnClientDisconnect(ByVal ConnectionID As Long, ByVal Reason As String)
    LogInfo "客户端断开: ID " & ConnectionID & ", 原因: " & Reason
    
    ' 清理用户相关数据
    CleanupUserData(ConnectionID)
    
    ' 通知其他用户
    NotifyUserLeft(ConnectionID)
    
    ' 更新统计信息
    UpdateConnectionStats()
End Sub
```

---

### UnregisterEventCallback

取消注册事件回调函数。

**函数签名**:
```vb
Public Sub UnregisterEventCallback(ByVal EventType As EventType)
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| EventType | EventType | 要取消的事件类型 |

**使用示例**:
```vb
Private Sub CleanupEventHandlers()
    server.UnregisterEventCallback EVENT_CLIENT_CONNECT
    server.UnregisterEventCallback EVENT_CLIENT_DISCONNECT
    server.UnregisterEventCallback EVENT_DATA_RECEIVED
    server.UnregisterEventCallback EVENT_ERROR_OCCURRED
End Sub
```

---

## 错误处理函数

### GetLastError

获取最后一次操作的错误代码。

**函数签名**:
```vb
Public Function GetLastError() As Long
```

**参数**: 无

**返回值**: 错误代码 (Long)

**常见错误代码**:
| 代码 | 描述 |
|------|------|
| 0 | 无错误 |
| -1 | 端口已被占用 |
| -2 | 权限不足 |
| -3 | 网络初始化失败 |
| -4 | 内存不足 |
| -5 | 创建完成端口失败 |

**使用示例**:
```vb
Private Sub HandleServerError()
    Dim errorCode As Long
    errorCode = server.GetLastError()
    
    Select Case errorCode
        Case 0
            ' 无错误
        Case -1
            MsgBox "端口已被其他程序占用，请选择其他端口", vbCritical
        Case -2
            MsgBox "需要管理员权限才能启动服务器", vbCritical
        Case -3
            MsgBox "网络初始化失败，请检查网络配置", vbCritical
        Case -4
            MsgBox "系统内存不足，请关闭其他程序", vbCritical
        Case -5
            MsgBox "创建 IOCP 失败，请重试", vbCritical
        Case Else
            MsgBox "未知错误: " & errorCode, vbCritical
    End Select
End Sub
```

---

### GetErrorDescription

获取错误代码的文本描述。

**函数签名**:
```vb
Public Function GetErrorDescription(Optional ByVal ErrorCode As Long = -1) As String
```

**参数**:
| 参数名 | 类型 | 描述 | 默认值 |
|--------|------|------|---------|
| ErrorCode | Long | 错误代码 | -1 (使用最后错误) |

**返回值**: 错误描述字符串

**使用示例**:
```vb
Private Sub LogLastError()
    Dim errorCode As Long
    Dim errorDesc As String
    
    errorCode = server.GetLastError()
    errorDesc = server.GetErrorDescription(errorCode)
    
    If errorCode <> 0 Then
        LogError "操作失败: [" & errorCode & "] " & errorDesc
    End If
End Sub

Private Sub ShowDetailedError()
    Dim errorMsg As String
    errorMsg = "操作失败!" & vbCrLf & vbCrLf
    errorMsg = errorMsg & "错误代码: " & server.GetLastError() & vbCrLf
    errorMsg = errorMsg & "错误描述: " & server.GetErrorDescription() & vbCrLf
    errorMsg = errorMsg & "时间: " & Now
    
    MsgBox errorMsg, vbCritical, "错误详情"
End Sub
```

---

## 性能监控函数

### GetPerformanceStats

获取服务器性能统计信息。

**函数签名**:
```vb
Public Function GetPerformanceStats() As PerformanceStats
```

**参数**: 无

**返回值**: PerformanceStats 结构体

**PerformanceStats 结构**:
```vb
Public Type PerformanceStats
    TotalConnections As Long        ' 总连接数
    ActiveConnections As Long       ' 活跃连接数
    TotalBytesReceived As Long      ' 总接收字节数
    TotalBytesSent As Long          ' 总发送字节数
    OperationsPerSecond As Double   ' 每秒操作数
    AverageResponseTime As Double   ' 平均响应时间
    CPUUsage As Double              ' CPU 使用率
    MemoryUsage As Long             ' 内存使用量
    UptimeSeconds As Long           ' 运行时间（秒）
End Type
```

**使用示例**:
```vb
Private Sub DisplayPerformanceInfo()
    Dim stats As PerformanceStats
    stats = server.GetPerformanceStats()
    
    Dim info As String
    info = "=== 服务器性能统计 ===" & vbCrLf
    info = info & "活跃连接数: " & stats.ActiveConnections & vbCrLf
    info = info & "总连接数: " & stats.TotalConnections & vbCrLf
    info = info & "接收数据: " & FormatBytes(stats.TotalBytesReceived) & vbCrLf
    info = info & "发送数据: " & FormatBytes(stats.TotalBytesSent) & vbCrLf
    info = info & "操作/秒: " & Format(stats.OperationsPerSecond, "0.00") & vbCrLf
    info = info & "响应时间: " & Format(stats.AverageResponseTime, "0.00") & "ms" & vbCrLf
    info = info & "CPU 使用: " & Format(stats.CPUUsage, "0.0") & "%" & vbCrLf
    info = info & "内存使用: " & FormatBytes(stats.MemoryUsage) & vbCrLf
    info = info & "运行时间: " & FormatUptime(stats.UptimeSeconds)
    
    txtPerformance.Text = info
End Sub

Private Sub MonitorPerformance()
    Dim stats As PerformanceStats
    stats = server.GetPerformanceStats()
    
    ' 检查性能阈值
    If stats.CPUUsage > 80 Then
        LogWarning "CPU 使用率过高: " & stats.CPUUsage & "%"
    End If
    
    If stats.AverageResponseTime > 1000 Then
        LogWarning "响应时间过长: " & stats.AverageResponseTime & "ms"
    End If
    
    If stats.ActiveConnections > 0.9 * server.GetMaxConnections() Then
        LogWarning "连接数接近上限: " & stats.ActiveConnections
    End If
End Sub
```

---

## 实用工具函数

### FormatBytes

格式化字节数为可读字符串。

**函数签名**:
```vb
Public Function FormatBytes(ByVal Bytes As Long) As String
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| Bytes | Long | 字节数 |

**返回值**: 格式化的字符串 (如 "1.5 MB")

**使用示例**:
```vb
Private Function FormatBytes(ByVal Bytes As Long) As String
    Dim units As Variant
    units = Array("B", "KB", "MB", "GB", "TB")
    
    Dim size As Double: size = Bytes
    Dim unitIndex As Integer: unitIndex = 0
    
    While size >= 1024 And unitIndex < UBound(units)
        size = size / 1024
        unitIndex = unitIndex + 1
    Wend
    
    If unitIndex = 0 Then
        FormatBytes = size & " " & units(unitIndex)
    Else
        FormatBytes = Format(size, "0.0") & " " & units(unitIndex)
    End If
End Function

' 使用示例
Debug.Print FormatBytes(1024)      ' "1.0 KB"
Debug.Print FormatBytes(1536)      ' "1.5 KB"
Debug.Print FormatBytes(2097152)   ' "2.0 MB"
```

---

### FormatUptime

格式化运行时间为可读字符串。

**函数签名**:
```vb
Public Function FormatUptime(ByVal Seconds As Long) As String
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| Seconds | Long | 秒数 |

**返回值**: 格式化的时间字符串

**使用示例**:
```vb
Private Function FormatUptime(ByVal Seconds As Long) As String
    Dim days As Long: days = Seconds \ 86400
    Dim hours As Long: hours = (Seconds Mod 86400) \ 3600
    Dim minutes As Long: minutes = (Seconds Mod 3600) \ 60
    Dim secs As Long: secs = Seconds Mod 60
    
    Dim result As String
    If days > 0 Then
        result = days & "天 "
    End If
    
    result = result & Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(secs, "00")
    FormatUptime = result
End Function

' 使用示例
Debug.Print FormatUptime(90)       ' "00:01:30"
Debug.Print FormatUptime(3661)     ' "01:01:01"
Debug.Print FormatUptime(90061)    ' "1天 01:01:01"
```

---

### ValidateIPAddress

验证 IP 地址格式。

**函数签名**:
```vb
Public Function ValidateIPAddress(ByVal IPAddress As String) As Boolean
```

**参数**:
| 参数名 | 类型 | 描述 |
|--------|------|------|
| IPAddress | String | 要验证的 IP 地址 |

**返回值**:
- `True`: IP 地址格式正确
- `False`: IP 地址格式错误

**使用示例**:
```vb
Private Function ValidateIPAddress(ByVal IPAddress As String) As Boolean
    Dim parts As Variant
    parts = Split(IPAddress, ".")
    
    If UBound(parts) <> 3 Then
        ValidateIPAddress = False
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To 3
        If Not IsNumeric(parts(i)) Then
            ValidateIPAddress = False
            Exit Function
        End If
        
        Dim num As Integer
        num = CInt(parts(i))
        If num < 0 Or num > 255 Then
            ValidateIPAddress = False
            Exit Function
        End If
    Next i
    
    ValidateIPAddress = True
End Function

' 使用示例
Private Sub TestIPValidation()
    Debug.Print ValidateIPAddress("192.168.1.1")   ' True
    Debug.Print ValidateIPAddress("256.1.1.1")     ' False
    Debug.Print ValidateIPAddress("192.168.1")     ' False
    Debug.Print ValidateIPAddress("abc.def.ghi.jkl") ' False
End Sub
```

---

### GetSystemInfo

获取系统信息。

**函数签名**:
```vb
Public Function GetSystemInfo() As SystemInfo
```

**参数**: 无

**返回值**: SystemInfo 结构体

**SystemInfo 结构**:
```vb
Public Type SystemInfo
    OSVersion As String             ' 操作系统版本
    CPUCount As Long                ' CPU 核心数
    TotalMemoryMB As Long           ' 总内存 (MB)
    AvailableMemoryMB As Long       ' 可用内存 (MB)
    NetworkAdapters As String       ' 网络适配器信息
End Type
```

**使用示例**:
```vb
Private Sub ShowSystemInfo()
    Dim sysInfo As SystemInfo
    sysInfo = GetSystemInfo()
    
    Dim info As String
    info = "=== 系统信息 ===" & vbCrLf
    info = info & "操作系统: " & sysInfo.OSVersion & vbCrLf
    info = info & "CPU 核心数: " & sysInfo.CPUCount & vbCrLf
    info = info & "总内存: " & sysInfo.TotalMemoryMB & " MB" & vbCrLf
    info = info & "可用内存: " & sysInfo.AvailableMemoryMB & " MB" & vbCrLf
    info = info & "网络适配器: " & sysInfo.NetworkAdapters
    
    MsgBox info, vbInformation, "系统信息"
End Sub
```

---

## 完整使用示例

### 综合聊天服务器实现

```vb
' 主服务器类的完整实现
Private WithEvents server As TBIOCPServer
Private connectedUsers As Collection
Private chatRooms As Collection
Private performanceTimer As Timer

Private Sub Form_Load()
    ' 初始化集合
    Set connectedUsers = New Collection
    Set chatRooms = New Collection
    
    ' 创建服务器实例
    Set server = New TBIOCPServer
    
    ' 配置服务器参数
    ConfigureServer()
    
    ' 设置定时器
    Set performanceTimer = New Timer
    performanceTimer.Interval = 5000 ' 5秒更新一次
    performanceTimer.Enabled = True
    
    ' 启动服务器
    StartChatServer()
End Sub

Private Sub ConfigureServer()
    ' 根据系统配置优化参数
    Dim sysInfo As SystemInfo
    sysInfo = GetSystemInfo()
    
    ' 设置最大连接数
    Dim maxConn As Long
    maxConn = sysInfo.AvailableMemoryMB \ 2 ' 每连接约2MB
    If maxConn > 1000 Then maxConn = 1000
    If maxConn < 100 Then maxConn = 100
    server.SetMaxConnections(maxConn)
    
    ' 设置工作线程数
    server.SetWorkerThreads(sysInfo.CPUCount * 2)
    
    LogInfo "服务器配置: 最大连接=" & maxConn & ", 线程数=" & (sysInfo.CPUCount * 2)
End Sub

Private Sub StartChatServer()
    If server.StartServer(9999) Then
        lblStatus.Caption = "聊天服务器运行中 - 端口 9999"
        lblStatus.ForeColor = vbGreen
        btnStart.Enabled = False
        btnStop.Enabled = True
        LogInfo "聊天服务器启动成功"
    Else
        lblStatus.Caption = "服务器启动失败"
        lblStatus.ForeColor = vbRed
        LogError "启动失败: " & server.GetErrorDescription()
    End If
End Sub

' 事件处理
Private Sub server_OnClientConnect(ByVal ConnectionID As Long, ByVal RemoteIP As String, ByVal RemotePort As Long)
    LogInfo "用户连接: " & RemoteIP & ":" & RemotePort & " (ID: " & ConnectionID & ")"
    
    ' 发送欢迎消息
    server.SendData ConnectionID, "WELCOME|请输入您的用户名"
    
    ' 更新界面
    UpdateConnectionDisplay()
End Sub

Private Sub server_OnDataReceived(ByVal ConnectionID As Long, ByVal Data As String, ByVal DataLength As Long)
    ProcessChatCommand ConnectionID, Data
End Sub

Private Sub server_OnClientDisconnect(ByVal ConnectionID As Long, ByVal Reason As String)
    LogInfo "用户断开: ID " & ConnectionID & " (" & Reason & ")"
    
    ' 清理用户数据
    CleanupUser ConnectionID
    
    ' 更新界面
    UpdateConnectionDisplay()
End Sub

' 聊天命令处理
Private Sub ProcessChatCommand(ByVal ConnectionID As Long, ByVal Command As String)
    Dim parts As Variant
    parts = Split(Command, "|")
    
    If UBound(parts) < 0 Then Exit Sub
    
    Select Case UCase(parts(0))
        Case "LOGIN"
            HandleUserLogin ConnectionID, parts(1)
        Case "MESSAGE"
            HandleChatMessage ConnectionID, parts(1)
        Case "JOIN_ROOM"
            HandleJoinRoom ConnectionID, parts(1)
        Case "LEAVE_ROOM"
            HandleLeaveRoom ConnectionID, parts(1)
        Case "LIST_USERS"
            HandleListUsers ConnectionID
        Case "PRIVATE_MSG"
            HandlePrivateMessage ConnectionID, parts(1), parts(2)
        Case Else
            server.SendData ConnectionID, "ERROR|未知命令: " & parts(0)
    End Select
End Sub

Private Sub HandleUserLogin(ByVal ConnectionID As Long, ByVal Username As String)
    ' 验证用户名
    If Len(Username) < 3 Or Len(Username) > 20 Then
        server.SendData ConnectionID, "ERROR|用户名长度必须在3-20字符之间"
        Exit Sub
    End If
    
    ' 检查用户名是否已存在
    If IsUsernameExists(Username) Then
        server.SendData ConnectionID, "ERROR|用户名已存在"
        Exit Sub
    End If
    
    ' 添加用户到列表
    connectedUsers.Add Username, CStr(ConnectionID)
    
    ' 发送登录成功消息
    server.SendData ConnectionID, "LOGIN_SUCCESS|欢迎, " & Username & "!"
    
    ' 通知其他用户
    BroadcastUserJoined Username
    
    LogInfo "用户登录: " & Username & " (ID: " & ConnectionID & ")"
End Sub

' 性能监控
Private Sub performanceTimer_Timer()
    UpdatePerformanceDisplay()
    MonitorSystemHealth()
End Sub

Private Sub UpdatePerformanceDisplay()
    Dim stats As PerformanceStats
    stats = server.GetPerformanceStats()
    
    lblConnections.Caption = "连接数: " & stats.ActiveConnections
    lblBytesReceived.Caption = "接收: " & FormatBytes(stats.TotalBytesReceived)
    lblBytesSent.Caption = "发送: " & FormatBytes(stats.TotalBytesSent)
    lblUptime.Caption = "运行时间: " & FormatUptime(stats.UptimeSeconds)
    
    ' 更新性能图表
    UpdatePerformanceChart(stats)
End Sub

Private Sub MonitorSystemHealth()
    Dim stats As PerformanceStats
    stats = server.GetPerformanceStats()
    
    ' 检查警告条件
    If stats.CPUUsage > 85 Then
        LogWarning "CPU使用率过高: " & stats.CPUUsage & "%"
    End If
    
    If stats.AverageResponseTime > 2000 Then
        LogWarning "响应时间过长: " & stats.AverageResponseTime & "ms"
    End If
    
    If stats.ActiveConnections > 0.95 * server.GetMaxConnections() Then
        LogWarning "连接数接近上限"
        BroadcastSystemMessage "服务器连接数接近上限，请稍后重试"
    End If
End Sub
```

---

*本函数参考手册提供了 TBIOCP 模块所有公共函数的详细说明和使用示例。通过这些函数，您可以构建功能完整的高性能网络服务应用。*