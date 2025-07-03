# TBIOCP API 文档

## 概述

TBIOCP 是一个基于 Windows I/O 完成端口（IOCP）的高性能网络通信模块，提供稳定的异步网络服务功能。

### 版本信息
- **版本**: 2025.6.5-9
- **开发者**: 静咫流云
- **邮箱**: 8226932@qq.com
- **创建时间**: 2025.5.29

### 主要特点

1. **内存管理优化**: 完全释放内存和句柄，无内存泄漏
2. **高并发稳定性**: 支持客户端高频连接和断开，经长时间暴力测试稳定运行
3. **跨平台支持**: 完美支持 x86 和 x64 架构
4. **依赖项**: 基于 fafalone WinDevLib >= 8.12.550.0 包
5. **基础框架**: 目前提供稳定的 ECHO 框架流程

---

## 公共 API 接口

### 核心 IOCP 函数

#### 服务器管理

##### `StartServer`
启动 IOCP 服务器

**语法**:
```vb
Public Function StartServer(ByVal Port As Long, Optional ByVal MaxConnections As Long = 1000) As Boolean
```

**参数**:
- `Port` (Long): 监听端口号
- `MaxConnections` (Long, 可选): 最大连接数，默认 1000

**返回值**:
- `Boolean`: 成功返回 True，失败返回 False

**示例**:
```vb
Dim server As TBIOCPServer
Set server = New TBIOCPServer

' 启动服务器在端口 8080，最大 500 连接
If server.StartServer(8080, 500) Then
    Debug.Print "服务器启动成功"
Else
    Debug.Print "服务器启动失败"
End If
```

##### `StopServer`
停止 IOCP 服务器

**语法**:
```vb
Public Sub StopServer()
```

**示例**:
```vb
server.StopServer()
Debug.Print "服务器已停止"
```

##### `IsRunning`
检查服务器运行状态

**语法**:
```vb
Public Function IsRunning() As Boolean
```

**返回值**:
- `Boolean`: 服务器运行中返回 True，否则返回 False

**示例**:
```vb
If server.IsRunning() Then
    Debug.Print "服务器正在运行"
End If
```

#### 连接管理

##### `GetConnectionCount`
获取当前连接数

**语法**:
```vb
Public Function GetConnectionCount() As Long
```

**返回值**:
- `Long`: 当前活跃连接数

**示例**:
```vb
Dim connectionCount As Long
connectionCount = server.GetConnectionCount()
Debug.Print "当前连接数: " & connectionCount
```

##### `DisconnectClient`
断开指定客户端连接

**语法**:
```vb
Public Sub DisconnectClient(ByVal ConnectionID As Long)
```

**参数**:
- `ConnectionID` (Long): 连接 ID

**示例**:
```vb
' 断开连接 ID 为 12345 的客户端
server.DisconnectClient(12345)
```

#### 数据传输

##### `SendData`
向指定客户端发送数据

**语法**:
```vb
Public Function SendData(ByVal ConnectionID As Long, ByVal Data As String) As Boolean
```

**参数**:
- `ConnectionID` (Long): 目标连接 ID
- `Data` (String): 要发送的数据

**返回值**:
- `Boolean`: 发送成功返回 True，失败返回 False

**示例**:
```vb
Dim message As String
message = "Hello Client!"

If server.SendData(12345, message) Then
    Debug.Print "数据发送成功"
Else
    Debug.Print "数据发送失败"
End If
```

##### `SendDataToAll`
向所有连接的客户端广播数据

**语法**:
```vb
Public Function SendDataToAll(ByVal Data As String) As Long
```

**参数**:
- `Data` (String): 要广播的数据

**返回值**:
- `Long`: 成功发送的客户端数量

**示例**:
```vb
Dim broadcastMessage As String
broadcastMessage = "服务器公告：系统将在5分钟后重启"

Dim successCount As Long
successCount = server.SendDataToAll(broadcastMessage)
Debug.Print "广播成功发送给 " & successCount & " 个客户端"
```

---

## 事件处理接口

### 连接事件

#### `OnClientConnect`
客户端连接事件

**语法**:
```vb
Public Event OnClientConnect(ByVal ConnectionID As Long, ByVal RemoteIP As String, ByVal RemotePort As Long)
```

**参数**:
- `ConnectionID` (Long): 新连接的 ID
- `RemoteIP` (String): 客户端 IP 地址
- `RemotePort` (Long): 客户端端口

**示例**:
```vb
Private Sub server_OnClientConnect(ByVal ConnectionID As Long, ByVal RemoteIP As String, ByVal RemotePort As Long)
    Debug.Print "新客户端连接: ID=" & ConnectionID & ", IP=" & RemoteIP & ":" & RemotePort
    
    ' 发送欢迎消息
    server.SendData ConnectionID, "欢迎连接到 TBIOCP 服务器!"
End Sub
```

#### `OnClientDisconnect`
客户端断开连接事件

**语法**:
```vb
Public Event OnClientDisconnect(ByVal ConnectionID As Long, ByVal Reason As String)
```

**参数**:
- `ConnectionID` (Long): 断开连接的 ID
- `Reason` (String): 断开原因

**示例**:
```vb
Private Sub server_OnClientDisconnect(ByVal ConnectionID As Long, ByVal Reason As String)
    Debug.Print "客户端断开连接: ID=" & ConnectionID & ", 原因=" & Reason
    
    ' 清理与该连接相关的资源
    CleanupClientResources(ConnectionID)
End Sub
```

### 数据事件

#### `OnDataReceived`
接收到数据事件

**语法**:
```vb
Public Event OnDataReceived(ByVal ConnectionID As Long, ByVal Data As String, ByVal DataLength As Long)
```

**参数**:
- `ConnectionID` (Long): 发送数据的连接 ID
- `Data` (String): 接收到的数据
- `DataLength` (Long): 数据长度

**示例**:
```vb
Private Sub server_OnDataReceived(ByVal ConnectionID As Long, ByVal Data As String, ByVal DataLength As Long)
    Debug.Print "收到数据: 连接ID=" & ConnectionID & ", 长度=" & DataLength
    
    ' ECHO 服务器 - 将收到的数据原样返回
    server.SendData ConnectionID, Data
    
    ' 或者处理特定命令
    ProcessClientCommand ConnectionID, Data
End Sub

Private Sub ProcessClientCommand(ByVal ConnectionID As Long, ByVal Command As String)
    Select Case UCase(Trim(Command))
        Case "TIME"
            server.SendData ConnectionID, "当前时间: " & Now
        Case "QUIT"
            server.DisconnectClient ConnectionID
        Case Else
            server.SendData ConnectionID, "未知命令: " & Command
    End Select
End Sub
```

---

## 配置和属性

### 服务器属性

#### `MaxConnections`
最大连接数属性

**语法**:
```vb
Public Property Let MaxConnections(ByVal Value As Long)
Public Property Get MaxConnections() As Long
```

**示例**:
```vb
' 设置最大连接数
server.MaxConnections = 2000

' 获取最大连接数
Debug.Print "最大连接数: " & server.MaxConnections
```

#### `Port`
服务器端口属性

**语法**:
```vb
Public Property Get Port() As Long
```

**示例**:
```vb
Debug.Print "服务器端口: " & server.Port
```

#### `WorkerThreads`
工作线程数属性

**语法**:
```vb
Public Property Let WorkerThreads(ByVal Value As Long)
Public Property Get WorkerThreads() As Long
```

**示例**:
```vb
' 设置工作线程数（通常设为 CPU 核心数的 2 倍）
server.WorkerThreads = 8

Debug.Print "工作线程数: " & server.WorkerThreads
```

---

## 使用示例

### 基础 ECHO 服务器

```vb
' 声明服务器对象
Private WithEvents server As TBIOCPServer

Private Sub Form_Load()
    ' 创建服务器实例
    Set server = New TBIOCPServer
    
    ' 配置服务器
    server.MaxConnections = 1000
    server.WorkerThreads = 4
    
    ' 启动服务器
    If server.StartServer(8080) Then
        lblStatus.Caption = "服务器已启动在端口 8080"
    Else
        lblStatus.Caption = "服务器启动失败"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' 停止服务器
    If Not server Is Nothing Then
        server.StopServer
        Set server = Nothing
    End If
End Sub

Private Sub server_OnClientConnect(ByVal ConnectionID As Long, ByVal RemoteIP As String, ByVal RemotePort As Long)
    UpdateLog "客户端连接: " & RemoteIP & ":" & RemotePort & " (ID: " & ConnectionID & ")"
    UpdateConnectionCount
End Sub

Private Sub server_OnClientDisconnect(ByVal ConnectionID As Long, ByVal Reason As String)
    UpdateLog "客户端断开: ID " & ConnectionID & " (" & Reason & ")"
    UpdateConnectionCount
End Sub

Private Sub server_OnDataReceived(ByVal ConnectionID As Long, ByVal Data As String, ByVal DataLength As Long)
    UpdateLog "收到数据: [" & ConnectionID & "] " & Data
    
    ' ECHO - 原样返回数据
    server.SendData ConnectionID, Data
End Sub

Private Sub UpdateConnectionCount()
    lblConnections.Caption = "当前连接数: " & server.GetConnectionCount()
End Sub

Private Sub UpdateLog(ByVal Message As String)
    txtLog.Text = txtLog.Text & Format(Now, "hh:mm:ss") & " " & Message & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
```

### 聊天服务器示例

```vb
Private WithEvents chatServer As TBIOCPServer
Private clientList As Collection

Private Sub Form_Load()
    Set chatServer = New TBIOCPServer
    Set clientList = New Collection
    
    chatServer.MaxConnections = 500
    
    If chatServer.StartServer(9999) Then
        UpdateStatus "聊天服务器已启动"
    End If
End Sub

Private Sub chatServer_OnClientConnect(ByVal ConnectionID As Long, ByVal RemoteIP As String, ByVal RemotePort As Long)
    ' 添加到客户端列表
    clientList.Add ConnectionID, CStr(ConnectionID)
    
    ' 欢迎消息
    chatServer.SendData ConnectionID, "欢迎加入聊天室! 输入 /help 查看命令"
    
    ' 通知其他用户
    BroadcastMessage "用户 " & ConnectionID & " 加入了聊天室", ConnectionID
End Sub

Private Sub chatServer_OnClientDisconnect(ByVal ConnectionID As Long, ByVal Reason As String)
    ' 从客户端列表移除
    On Error Resume Next
    clientList.Remove CStr(ConnectionID)
    On Error GoTo 0
    
    ' 通知其他用户
    BroadcastMessage "用户 " & ConnectionID & " 离开了聊天室", -1
End Sub

Private Sub chatServer_OnDataReceived(ByVal ConnectionID As Long, ByVal Data As String, ByVal DataLength As Long)
    Dim message As String
    message = Trim(Data)
    
    If Left(message, 1) = "/" Then
        ' 处理命令
        ProcessChatCommand ConnectionID, message
    Else
        ' 广播聊天消息
        BroadcastMessage "[" & ConnectionID & "] " & message, -1
    End If
End Sub

Private Sub ProcessChatCommand(ByVal ConnectionID As Long, ByVal Command As String)
    Select Case UCase(Left(Command, 5))
        Case "/HELP"
            chatServer.SendData ConnectionID, "可用命令: /help, /list, /quit"
            
        Case "/LIST"
            chatServer.SendData ConnectionID, "在线用户数: " & clientList.Count
            
        Case "/QUIT"
            chatServer.SendData ConnectionID, "再见!"
            chatServer.DisconnectClient ConnectionID
            
        Case Else
            chatServer.SendData ConnectionID, "未知命令: " & Command
    End Select
End Sub

Private Sub BroadcastMessage(ByVal Message As String, ByVal ExcludeConnectionID As Long)
    Dim i As Integer
    Dim connectionID As Long
    
    For i = 1 To clientList.Count
        connectionID = clientList(i)
        If connectionID <> ExcludeConnectionID Then
            chatServer.SendData connectionID, Message
        End If
    Next i
End Sub
```

---

## 错误处理

### 常见错误代码

| 错误代码 | 描述 | 解决方案 |
|---------|------|----------|
| -1 | 端口已被占用 | 更换端口或关闭占用端口的程序 |
| -2 | 权限不足 | 以管理员权限运行程序 |
| -3 | 网络初始化失败 | 检查网络配置和防火墙设置 |
| -4 | 内存不足 | 减少最大连接数或增加系统内存 |
| -5 | 创建完成端口失败 | 检查系统资源和权限 |

### 错误处理示例

```vb
Private Sub StartServerWithErrorHandling()
    On Error GoTo ErrorHandler
    
    If Not server.StartServer(8080) Then
        Select Case server.LastError
            Case -1
                MsgBox "端口 8080 已被占用，请选择其他端口", vbCritical
            Case -2
                MsgBox "权限不足，请以管理员身份运行", vbCritical
            Case -3
                MsgBox "网络初始化失败，请检查网络配置", vbCritical
            Case Else
                MsgBox "启动服务器失败，错误代码: " & server.LastError, vbCritical
        End Select
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "发生未预期的错误: " & Err.Description, vbCritical
End Sub
```

---

## 性能优化建议

### 1. 线程配置
```vb
' 根据 CPU 核心数配置工作线程
Dim coreCount As Long
coreCount = GetSystemCoreCount()
server.WorkerThreads = coreCount * 2
```

### 2. 连接数限制
```vb
' 根据系统内存设置合理的最大连接数
' 每个连接大约占用 8-16KB 内存
server.MaxConnections = Min(AvailableMemoryMB * 64, 10000)
```

### 3. 数据缓冲
```vb
' 启用数据缓冲以减少系统调用
server.EnableBuffering = True
server.BufferSize = 8192  ' 8KB 缓冲区
```

---

## 依赖和要求

### 系统要求
- Windows XP SP3 或更高版本
- .NET Framework 2.0 或更高版本（如果使用 .NET 互操作）
- 最小 1GB RAM（推荐 2GB 或更多）

### 开发依赖
- fafalone WinDevLib >= 8.12.550.0
- Visual Basic 6.0 或兼容的开发环境
- Windows SDK（用于 IOCP API）

### 运行时依赖
- msvcrt.dll
- ws2_32.dll
- kernel32.dll

---

## 版本历史

### 2025.6.5-9 (当前版本)
- 基础 ECHO 框架稳定运行
- 完善内存管理，无内存泄漏
- 支持高频连接测试
- 支持 x86/x64 架构

### 未来计划
- 实现 AcceptEx 优化
- 添加心跳包机制
- 添加类似 WinSock 的事件函数
- 支持 HTTP 协议
- 支持 WebSocket 协议
- 优化代码结构，减少冗余

---

## 技术支持

**开发者**: 静咫流云  
**邮箱**: 8226932@qq.com  
**项目地址**: [待补充]

### 常见问题

1. **Q: 如何提高并发连接数？**
   A: 增加工作线程数，优化系统内核参数，使用更大的连接池。

2. **Q: 如何处理大数据传输？**
   A: 实现数据分片传输，使用压缩算法，设置合适的缓冲区大小。

3. **Q: 如何添加 SSL/TLS 支持？**
   A: 未来版本将集成 SSL/TLS 功能，当前版本可在应用层实现加密。

---

*本文档基于 TBIOCP 2025.6.5-9 版本编写，如有更新请关注版本变更说明。*