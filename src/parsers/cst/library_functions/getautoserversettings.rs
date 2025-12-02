//! `GetAutoServerSettings` Function
//!
//! Returns information about the security settings for a `DCOM` (`Distributed Component Object Model`) server.
//!
//! # Syntax
//!
//! ```vb
//! GetAutoServerSettings(progid, clsid, machine)
//! ```
//!
//! # Parameters
//!
//! - `progid` - Required. String expression that specifies the programmatic identifier (`ProgID`) of the server.
//! - `clsid` - Required. String expression that specifies the class identifier (`CLSID`) of the server.
//! - `machine` - Required. String expression that specifies the name of the machine where the server is located.
//!
//! # Return Value
//!
//! Returns a `Long` value containing security settings information for the specified `DCOM` server.
//!
//! # Remarks
//!
//! - This function is specific to `DCOM` (`Distributed Component Object Model`) automation servers.
//! - Used primarily in distributed computing scenarios.
//! - Returns security configuration information from the Windows registry.
//! - The function is part of VB6's `DCOM` support infrastructure.
//! - Typically used in enterprise applications with distributed components.
//! - Requires appropriate `DCOM` permissions on the target machine.
//! - The progid and clsid must correspond to a registered `COM`/`DCOM` server.
//! - Machine name can be a `NetBIOS` name, `DNS` name, or `IP` address.
//! - Returns 0 if the server settings cannot be retrieved.
//!
//! # Typical Uses
//!
//! - Querying DCOM server security configurations
//! - Validating remote server accessibility
//! - Auditing distributed component settings
//! - Troubleshooting DCOM connection issues
//! - Enterprise application deployment verification
//! - Remote component diagnostics
//!
//! # Basic Usage Examples
//!
//! ```vb
//! ' Check DCOM server settings
//! Dim settings As Long
//! settings = GetAutoServerSettings("MyServer.Application", _
//!                                   "{12345678-1234-1234-1234-123456789012}", _
//!                                   "SERVER01")
//!
//! If settings <> 0 Then
//!     Debug.Print "Server settings retrieved: " & settings
//! Else
//!     Debug.Print "Unable to retrieve server settings"
//! End If
//!
//! ' Verify remote component availability
//! Dim result As Long
//! result = GetAutoServerSettings("Excel.Application", _
//!                                "{00024500-0000-0000-C000-000000000046}", _
//!                                "REMOTE-PC")
//!
//! If result <> 0 Then
//!     MsgBox "Remote Excel server is configured"
//! End If
//!
//! ' Query local server settings
//! Dim localSettings As Long
//! localSettings = GetAutoServerSettings("MyApp.Server", _
//!                                       "{ABCDEF01-2345-6789-ABCD-EF0123456789}", _
//!                                       ".")
//! ```
//!
//! # Common Patterns
//!
//! ## 1. DCOM Server Validation
//!
//! ```vb
//! Function ValidateDCOMServer(progID As String, _
//!                             clsID As String, _
//!                             serverName As String) As Boolean
//!     Dim settings As Long
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     settings = GetAutoServerSettings(progID, clsID, serverName)
//!     
//!     If settings <> 0 Then
//!         Debug.Print "DCOM server validated on " & serverName
//!         ValidateDCOMServer = True
//!     Else
//!         Debug.Print "DCOM server not accessible on " & serverName
//!         ValidateDCOMServer = False
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     Debug.Print "Error validating DCOM server: " & Err.Description
//!     ValidateDCOMServer = False
//! End Function
//! ```
//!
//! ## 2. Multi-Server Configuration Check
//!
//! ```vb
//! Sub CheckServersConfiguration()
//!     Dim servers() As String
//!     Dim i As Long
//!     Dim settings As Long
//!     
//!     servers = Array("SERVER01", "SERVER02", "SERVER03")
//!     
//!     For i = LBound(servers) To UBound(servers)
//!         settings = GetAutoServerSettings("MyApp.DataServer", _
//!                                          "{11111111-2222-3333-4444-555555555555}", _
//!                                          servers(i))
//!         
//!         If settings <> 0 Then
//!             Debug.Print servers(i) & " - Configured: " & settings
//!         Else
//!             Debug.Print servers(i) & " - Not configured"
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ## 3. Server Discovery and Verification
//!
//! ```vb
//! Function FindAvailableServer(progID As String, _
//!                              clsID As String, _
//!                              servers As Collection) As String
//!     Dim server As Variant
//!     Dim settings As Long
//!     
//!     For Each server In servers
//!         On Error Resume Next
//!         settings = GetAutoServerSettings(progID, clsID, CStr(server))
//!         On Error GoTo 0
//!         
//!         If settings <> 0 Then
//!             FindAvailableServer = CStr(server)
//!             Exit Function
//!         End If
//!     Next server
//!     
//!     FindAvailableServer = ""
//! End Function
//!
//! ' Usage
//! Dim servers As New Collection
//! servers.Add "PRIMARY-SERVER"
//! servers.Add "BACKUP-SERVER"
//! servers.Add "FAILOVER-SERVER"
//!
//! Dim activeServer As String
//! activeServer = FindAvailableServer("MyApp.Service", _
//!                                    "{FEDCBA98-7654-3210-FEDC-BA9876543210}", _
//!                                    servers)
//!
//! If activeServer <> "" Then
//!     Debug.Print "Using server: " & activeServer
//! End If
//! ```
//!
//! ## 4. Settings Comparison Across Servers
//!
//! ```vb
//! Sub CompareServerSettings(progID As String, _
//!                           clsID As String, _
//!                           server1 As String, _
//!                           server2 As String)
//!     Dim settings1 As Long
//!     Dim settings2 As Long
//!     
//!     settings1 = GetAutoServerSettings(progID, clsID, server1)
//!     settings2 = GetAutoServerSettings(progID, clsID, server2)
//!     
//!     Debug.Print "Server: " & server1 & " - Settings: " & settings1
//!     Debug.Print "Server: " & server2 & " - Settings: " & settings2
//!     
//!     If settings1 = settings2 Then
//!         Debug.Print "Servers have identical settings"
//!     Else
//!         Debug.Print "Warning: Server settings differ"
//!     End If
//! End Sub
//! ```
//!
//! ## 5. Dynamic Server Connection
//!
//! ```vb
//! Function ConnectToServer(progID As String, _
//!                          clsID As String, _
//!                          preferredServer As String, _
//!                          fallbackServer As String) As Object
//!     Dim settings As Long
//!     Dim targetServer As String
//!     
//!     ' Try preferred server first
//!     settings = GetAutoServerSettings(progID, clsID, preferredServer)
//!     
//!     If settings <> 0 Then
//!         targetServer = preferredServer
//!     Else
//!         ' Fall back to alternate server
//!         settings = GetAutoServerSettings(progID, clsID, fallbackServer)
//!         
//!         If settings <> 0 Then
//!             targetServer = fallbackServer
//!         Else
//!             Err.Raise vbObjectError + 1000, , "No available servers"
//!         End If
//!     End If
//!     
//!     Debug.Print "Connecting to: " & targetServer
//!     Set ConnectToServer = CreateObject(progID, targetServer)
//! End Function
//! ```
//!
//! ## 6. Server Health Monitoring
//!
//! ```vb
//! Type ServerStatus
//!     ServerName As String
//!     Settings As Long
//!     LastChecked As Date
//!     IsAvailable As Boolean
//! End Type
//!
//! Function CheckServerHealth(progID As String, _
//!                           clsID As String, _
//!                           serverName As String) As ServerStatus
//!     Dim status As ServerStatus
//!     
//!     status.ServerName = serverName
//!     status.LastChecked = Now
//!     
//!     On Error Resume Next
//!     status.Settings = GetAutoServerSettings(progID, clsID, serverName)
//!     On Error GoTo 0
//!     
//!     status.IsAvailable = (status.Settings <> 0)
//!     
//!     CheckServerHealth = status
//! End Function
//!
//! Sub MonitorServers()
//!     Dim servers() As String
//!     Dim i As Long
//!     Dim status As ServerStatus
//!     
//!     servers = Array("SERVER-A", "SERVER-B", "SERVER-C")
//!     
//!     For i = LBound(servers) To UBound(servers)
//!         status = CheckServerHealth("MyApp.Service", _
//!                                   "{12345678-ABCD-EFGH-IJKL-123456789ABC}", _
//!                                   servers(i))
//!         
//!         Debug.Print status.ServerName & ": " & _
//!                     IIf(status.IsAvailable, "Online", "Offline") & _
//!                     " (" & status.Settings & ")"
//!     Next i
//! End Sub
//! ```
//!
//! ## 7. Configuration Auditing
//!
//! ```vb
//! Sub AuditDCOMConfiguration(progID As String, clsID As String)
//!     Dim servers() As String
//!     Dim i As Long
//!     Dim settings As Long
//!     Dim fileNum As Integer
//!     
//!     servers = Array("PROD-01", "PROD-02", "TEST-01", "DEV-01")
//!     
//!     fileNum = FreeFile
//!     Open "C:\Audit\DCOM_Audit.txt" For Output As #fileNum
//!     
//!     Print #fileNum, "DCOM Configuration Audit Report"
//!     Print #fileNum, "Date: " & Now
//!     Print #fileNum, "ProgID: " & progID
//!     Print #fileNum, "CLSID: " & clsID
//!     Print #fileNum, String(50, "=")
//!     
//!     For i = LBound(servers) To UBound(servers)
//!         settings = GetAutoServerSettings(progID, clsID, servers(i))
//!         
//!         Print #fileNum, "Server: " & servers(i)
//!         Print #fileNum, "  Settings Value: " & settings
//!         Print #fileNum, "  Status: " & IIf(settings <> 0, "Available", "Unavailable")
//!         Print #fileNum, ""
//!     Next i
//!     
//!     Close #fileNum
//!     Debug.Print "Audit complete"
//! End Sub
//! ```
//!
//! ## 8. Load Balancing Server Selection
//!
//! ```vb
//! Function SelectLeastLoadedServer(progID As String, _
//!                                  clsID As String, _
//!                                  servers As Collection) As String
//!     Dim server As Variant
//!     Dim settings As Long
//!     Dim minSettings As Long
//!     Dim selectedServer As String
//!     
//!     minSettings = 2147483647  ' Max Long value
//!     
//!     For Each server In servers
//!         On Error Resume Next
//!         settings = GetAutoServerSettings(progID, clsID, CStr(server))
//!         On Error GoTo 0
//!         
//!         If settings <> 0 And settings < minSettings Then
//!             minSettings = settings
//!             selectedServer = CStr(server)
//!         End If
//!     Next server
//!     
//!     SelectLeastLoadedServer = selectedServer
//! End Function
//! ```
//!
//! ## 9. Deployment Verification
//!
//! ```vb
//! Function VerifyDeployment(progID As String, _
//!                          clsID As String, _
//!                          targetServers As Variant) As Boolean
//!     Dim i As Long
//!     Dim settings As Long
//!     Dim allConfigured As Boolean
//!     
//!     allConfigured = True
//!     
//!     For i = LBound(targetServers) To UBound(targetServers)
//!         settings = GetAutoServerSettings(progID, clsID, targetServers(i))
//!         
//!         If settings = 0 Then
//!             Debug.Print "Deployment failed on: " & targetServers(i)
//!             allConfigured = False
//!         Else
//!             Debug.Print "Deployment verified on: " & targetServers(i)
//!         End If
//!     Next i
//!     
//!     VerifyDeployment = allConfigured
//! End Function
//!
//! ' Usage in deployment script
//! Sub DeploymentCheck()
//!     Dim productionServers As Variant
//!     
//!     productionServers = Array("WEB-01", "WEB-02", "APP-01", "APP-02")
//!     
//!     If VerifyDeployment("MyApp.BusinessLogic", _
//!                        "{AAAABBBB-CCCC-DDDD-EEEE-FFFF00001111}", _
//!                        productionServers) Then
//!         MsgBox "Deployment successful on all servers"
//!     Else
//!         MsgBox "Deployment incomplete - check logs"
//!     End If
//! End Sub
//! ```
//!
//! ## 10. Regional Server Discovery
//!
//! ```vb
//! Type RegionalServer
//!     Region As String
//!     ServerName As String
//!     Settings As Long
//! End Type
//!
//! Function GetRegionalServer(progID As String, _
//!                           clsID As String, _
//!                           region As String) As String
//!     Dim regionalServers(1 To 3) As RegionalServer
//!     Dim i As Long
//!     
//!     ' Define regional servers
//!     regionalServers(1).Region = "US-EAST"
//!     regionalServers(1).ServerName = "US-EAST-SVR01"
//!     
//!     regionalServers(2).Region = "US-WEST"
//!     regionalServers(2).ServerName = "US-WEST-SVR01"
//!     
//!     regionalServers(3).Region = "EUROPE"
//!     regionalServers(3).ServerName = "EU-SVR01"
//!     
//!     ' Check settings for each regional server
//!     For i = 1 To 3
//!         regionalServers(i).Settings = GetAutoServerSettings(progID, _
//!                                                              clsID, _
//!                                                              regionalServers(i).ServerName)
//!     Next i
//!     
//!     ' Find matching region
//!     For i = 1 To 3
//!         If regionalServers(i).Region = region And regionalServers(i).Settings <> 0 Then
//!             GetRegionalServer = regionalServers(i).ServerName
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     GetRegionalServer = ""
//! End Function
//! ```
//!
//! # Advanced Usage
//!
//! ## 1. DCOM Server Manager Class
//!
//! ```vb
//! ' Class: DCOMServerManager
//! Private m_ProgID As String
//! Private m_CLSID As String
//! Private m_Servers As Collection
//! Private m_CurrentServer As String
//!
//! Public Sub Initialize(progID As String, clsID As String)
//!     m_ProgID = progID
//!     m_CLSID = clsID
//!     Set m_Servers = New Collection
//! End Sub
//!
//! Public Sub AddServer(serverName As String)
//!     m_Servers.Add serverName
//! End Sub
//!
//! Public Function FindActiveServer() As String
//!     Dim server As Variant
//!     Dim settings As Long
//!     
//!     For Each server In m_Servers
//!         On Error Resume Next
//!         settings = GetAutoServerSettings(m_ProgID, m_CLSID, CStr(server))
//!         On Error GoTo 0
//!         
//!         If settings <> 0 Then
//!             m_CurrentServer = CStr(server)
//!             FindActiveServer = m_CurrentServer
//!             Exit Function
//!         End If
//!     Next server
//!     
//!     FindActiveServer = ""
//! End Function
//!
//! Public Function GetServerSettings(serverName As String) As Long
//!     GetServerSettings = GetAutoServerSettings(m_ProgID, m_CLSID, serverName)
//! End Function
//!
//! Public Function ValidateAllServers() As Collection
//!     Dim results As New Collection
//!     Dim server As Variant
//!     Dim settings As Long
//!     Dim result As String
//!     
//!     For Each server In m_Servers
//!         settings = GetAutoServerSettings(m_ProgID, m_CLSID, CStr(server))
//!         result = CStr(server) & ":" & CStr(settings)
//!         results.Add result
//!     Next server
//!     
//!     Set ValidateAllServers = results
//! End Function
//! ```
//!
//! ## 2. Failover Connection Handler
//!
//! ```vb
//! Type FailoverConfig
//!     PrimaryServer As String
//!     SecondaryServer As String
//!     TertiaryServer As String
//!     RetryCount As Integer
//!     RetryDelay As Long
//! End Type
//!
//! Function ConnectWithFailover(progID As String, _
//!                              clsID As String, _
//!                              config As FailoverConfig) As Object
//!     Dim servers(1 To 3) As String
//!     Dim i As Integer
//!     Dim attempt As Integer
//!     Dim settings As Long
//!     
//!     servers(1) = config.PrimaryServer
//!     servers(2) = config.SecondaryServer
//!     servers(3) = config.TertiaryServer
//!     
//!     For i = 1 To 3
//!         For attempt = 1 To config.RetryCount
//!             On Error Resume Next
//!             settings = GetAutoServerSettings(progID, clsID, servers(i))
//!             On Error GoTo 0
//!             
//!             If settings <> 0 Then
//!                 Debug.Print "Connected to: " & servers(i)
//!                 Set ConnectWithFailover = CreateObject(progID, servers(i))
//!                 Exit Function
//!             End If
//!             
//!             If attempt < config.RetryCount Then
//!                 Sleep config.RetryDelay
//!             End If
//!         Next attempt
//!     Next i
//!     
//!     Err.Raise vbObjectError + 1001, , "All servers unavailable"
//! End Function
//! ```
//!
//! ## 3. Configuration Cache Manager
//!
//! ```vb
//! Type CachedServerInfo
//!     ServerName As String
//!     Settings As Long
//!     CacheTime As Date
//!     TTL As Long  ' Time to live in seconds
//! End Type
//!
//! Private m_Cache As Collection
//!
//! Sub InitializeCache()
//!     Set m_Cache = New Collection
//! End Sub
//!
//! Function GetCachedServerSettings(progID As String, _
//!                                  clsID As String, _
//!                                  serverName As String, _
//!                                  Optional cacheTTL As Long = 300) As Long
//!     Dim cacheKey As String
//!     Dim cached As CachedServerInfo
//!     Dim i As Long
//!     Dim found As Boolean
//!     
//!     cacheKey = serverName
//!     
//!     ' Check cache
//!     For i = 1 To m_Cache.Count
//!         cached = m_Cache(i)
//!         If cached.ServerName = cacheKey Then
//!             If DateDiff("s", cached.CacheTime, Now) < cached.TTL Then
//!                 GetCachedServerSettings = cached.Settings
//!                 Exit Function
//!             Else
//!                 m_Cache.Remove i
//!                 Exit For
//!             End If
//!         End If
//!     Next i
//!     
//!     ' Not in cache or expired, fetch fresh
//!     cached.ServerName = serverName
//!     cached.Settings = GetAutoServerSettings(progID, clsID, serverName)
//!     cached.CacheTime = Now
//!     cached.TTL = cacheTTL
//!     
//!     m_Cache.Add cached
//!     GetCachedServerSettings = cached.Settings
//! End Function
//! ```
//!
//! ## 4. Server Pool Manager
//!
//! ```vb
//! Type ServerPool
//!     PoolName As String
//!     Servers() As String
//!     ProgID As String
//!     CLSID As String
//! End Type
//!
//! Function GetHealthyServersFromPool(pool As ServerPool) As Collection
//!     Dim healthyServers As New Collection
//!     Dim i As Long
//!     Dim settings As Long
//!     
//!     For i = LBound(pool.Servers) To UBound(pool.Servers)
//!         On Error Resume Next
//!         settings = GetAutoServerSettings(pool.ProgID, pool.CLSID, pool.Servers(i))
//!         On Error GoTo 0
//!         
//!         If settings <> 0 Then
//!             healthyServers.Add pool.Servers(i)
//!         End If
//!     Next i
//!     
//!     Set GetHealthyServersFromPool = healthyServers
//! End Function
//!
//! Function GetPoolStatistics(pool As ServerPool) As String
//!     Dim total As Long
//!     Dim healthy As Long
//!     Dim i As Long
//!     Dim settings As Long
//!     
//!     total = UBound(pool.Servers) - LBound(pool.Servers) + 1
//!     healthy = 0
//!     
//!     For i = LBound(pool.Servers) To UBound(pool.Servers)
//!         settings = GetAutoServerSettings(pool.ProgID, pool.CLSID, pool.Servers(i))
//!         If settings <> 0 Then healthy = healthy + 1
//!     Next i
//!     
//!     GetPoolStatistics = pool.PoolName & ": " & healthy & "/" & total & " servers available"
//! End Function
//! ```
//!
//! # Error Handling
//!
//! ```vb
//! Function SafeGetAutoServerSettings(progID As String, _
//!                                    clsID As String, _
//!                                    serverName As String) As Long
//!     On Error GoTo ErrorHandler
//!     
//!     SafeGetAutoServerSettings = GetAutoServerSettings(progID, clsID, serverName)
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 429  ' ActiveX component can't create object
//!             Debug.Print "Server not available: " & serverName
//!         Case 462  ' Remote server machine does not exist
//!             Debug.Print "Machine not found: " & serverName
//!         Case 70   ' Permission denied
//!             Debug.Print "Access denied to server: " & serverName
//!         Case Else
//!             Debug.Print "Error " & Err.Number & ": " & Err.Description
//!     End Select
//!     
//!     SafeGetAutoServerSettings = 0
//! End Function
//! ```
//!
//! Common errors:
//! - **Error 429**: `ActiveX` component can't create object - server not registered or accessible.
//! - **Error 462**: Remote server machine does not exist or is unavailable.
//! - **Error 70**: Permission denied - insufficient `DCOM` permissions.
//! - **Error 5**: Invalid procedure call - invalid `ProgID` or `CLSID` format.
//!
//! # Performance Considerations
//!
//! - Network latency affects remote server queries
//! - Consider caching results for frequently checked servers
//! - Use timeouts for unresponsive servers
//! - Parallel checks may improve performance for multiple servers
//! - DCOM configuration on both client and server affects response time
//! - Firewall settings can cause delays or failures
//!
//! # Best Practices
//!
//! 1. **Always use error handling** - network and `DCOM` issues are common
//! 2. **Validate `ProgID` and `CLSID` format** before calling
//! 3. **Use descriptive server names** for better diagnostics
//! 4. **Implement retry logic** for transient failures
//! 5. **Cache results** to reduce network overhead
//! 6. **Log all queries** for auditing and troubleshooting
//! 7. **Test connectivity** before production deployment
//! 8. **Configure `DCOM` security** appropriately on all servers
//!
//! # Comparison with Other Functions
//!
//! ## `GetAutoServerSettings` vs `CreateObject`
//!
//! ```vb
//! ' GetAutoServerSettings - Query server settings
//! settings = GetAutoServerSettings(progID, clsID, serverName)
//!
//! ' CreateObject - Actually create server instance
//! Set obj = CreateObject(progID, serverName)
//! ```
//!
//! ## `GetAutoServerSettings` vs `GetObject`
//!
//! ```vb
//! ' GetAutoServerSettings - Check DCOM configuration
//! settings = GetAutoServerSettings(progID, clsID, serverName)
//!
//! ' GetObject - Connect to existing instance
//! Set obj = GetObject(, progID)
//! ```
//!
//! # Limitations
//!
//! - Windows-specific functionality (DCOM is Windows-only)
//! - Requires DCOM to be enabled and properly configured
//! - Network connectivity required for remote servers
//! - Security settings may block access
//! - Return value interpretation is not well documented
//! - Limited to COM/DCOM servers
//! - May not work with modern .NET components
//! - Deprecated in favor of newer technologies (WCF, REST APIs)
//!
//! # DCOM Configuration
//!
//! For this function to work properly:
//!
//! 1. **Component Services** (dcomcnfg) must be configured
//! 2. **DCOM permissions** must allow remote access
//! 3. **Firewall rules** must permit DCOM traffic
//! 4. **Authentication level** must be set appropriately
//! 5. **Launch and activation permissions** must be granted
//!
//! # Related Functions
//!
//! - `CreateObject` - Creates an instance of a `COM` object
//! - `GetObject` - Returns a reference to an `ActiveX` object
//! - `CallByName` - Executes methods on objects dynamically
//! - `TypeName` - Returns type information about an object
//! - `GetSetting` - Retrieves application settings from registry

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_getautoserversettings_basic() {
        let source = r#"settings = GetAutoServerSettings("MyServer.Application", "{12345678-1234-1234-1234-123456789012}", "SERVER01")"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_with_variables() {
        let source = r#"result = GetAutoServerSettings(progID, clsID, serverName)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_in_if() {
        let source =
            r#"If GetAutoServerSettings(progID, clsID, "SERVER01") <> 0 Then MsgBox "Available""#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_in_function() {
        let source = r#"Function ValidateServer() As Boolean
    ValidateServer = (GetAutoServerSettings(m_ProgID, m_CLSID, m_Server) <> 0)
End Function"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_assignment() {
        let source = r#"Dim settings As Long
settings = GetAutoServerSettings("App.Server", "{AAAABBBB-CCCC-DDDD-EEEE-FFFF00001111}", "REMOTE-PC")"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_comparison() {
        let source = r#"If GetAutoServerSettings(progID, clsID, server1) = GetAutoServerSettings(progID, clsID, server2) Then
    Debug.Print "Same settings"
End If"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_for_loop() {
        let source = r#"For i = 1 To serverCount
    settings = GetAutoServerSettings(progID, clsID, servers(i))
Next i"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_select_case() {
        let source = r#"Select Case GetAutoServerSettings(progID, clsID, serverName)
    Case Is > 0
        Debug.Print "Configured"
End Select"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_error_handling() {
        let source = r#"On Error GoTo ErrorHandler
settings = GetAutoServerSettings(progID, clsID, serverName)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_debug_print() {
        let source =
            r#"Debug.Print "Settings: " & GetAutoServerSettings(progID, clsID, serverName)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_array_element() {
        let source = r#"settings(i) = GetAutoServerSettings(progID, clsID, servers(i))"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_class_member() {
        let source = r#"m_Settings = GetAutoServerSettings(m_ProgID, m_CLSID, m_ServerName)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_type_field() {
        let source =
            r#"serverInfo.Settings = GetAutoServerSettings(progID, clsID, serverInfo.Name)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_do_loop() {
        let source = r#"Do While retry < maxRetries
    settings = GetAutoServerSettings(progID, clsID, serverName)
    If settings <> 0 Then Exit Do
Loop"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_local_server() {
        let source = r#"localSettings = GetAutoServerSettings(progID, clsID, ".")"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_concatenation() {
        let source =
            r#"result = "Settings: " & CStr(GetAutoServerSettings(progID, clsID, serverName))"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_msgbox() {
        let source = r#"MsgBox "Settings: " & GetAutoServerSettings(progID, clsID, serverName)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_collection() {
        let source = r#"results.Add GetAutoServerSettings(progID, clsID, serverName)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_file_output() {
        let source = r#"Print #fileNum, "Server: " & serverName & " Settings: " & GetAutoServerSettings(progID, clsID, serverName)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_iif() {
        let source = r#"status = IIf(GetAutoServerSettings(progID, clsID, serverName) <> 0, "Online", "Offline")"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_with_statement() {
        let source = r#"With serverConfig
    .Settings = GetAutoServerSettings(.ProgID, .CLSID, .ServerName)
End With"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_on_error_resume() {
        let source = r#"On Error Resume Next
settings = GetAutoServerSettings(progID, clsID, serverName)
On Error GoTo 0"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_min_comparison() {
        let source = r#"If GetAutoServerSettings(progID, clsID, serverName) < minSettings Then
    minSettings = GetAutoServerSettings(progID, clsID, serverName)
End If"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_for_each() {
        let source = r#"For Each server In servers
    settings = GetAutoServerSettings(progID, clsID, CStr(server))
Next server"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_property() {
        let source = r#"Property Get ServerSettings() As Long
    ServerSettings = GetAutoServerSettings(m_ProgID, m_CLSID, m_ServerName)
End Property"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_exit_condition() {
        let source =
            r#"If GetAutoServerSettings(progID, clsID, serverName) = 0 Then Exit Function"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getautoserversettings_listbox() {
        let source = r#"lstServers.AddItem serverName & " - " & GetAutoServerSettings(progID, clsID, serverName)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAutoServerSettings"));
        assert!(debug.contains("Identifier"));
    }
}
