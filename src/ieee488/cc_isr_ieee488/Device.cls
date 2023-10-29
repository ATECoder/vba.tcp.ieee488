VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Device"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type this_
    
    ' session settings
    Session As TcpSession
    GpibLanControllerPort As Long
    ReadAfterWriteDelay As Long
    Termination As String

    ' device settings
    DeviceErrorReader As IDeviceErrorReader
    Identity As String
    Options As String

End Type

Private This As this_

Private WithEvents m_socket As IPv4StreamSocket
Attribute m_socket.VB_VarHelpID = -1

''' <summary>   Handles the Class Initialize event; constructs the <see cref="Session"/>. </summary>
Private Sub Class_Initialize()
    
    This.GpibLanControllerPort = 1234
    This.Termination = VBA.vbLf
    This.ReadAfterWriteDelay = 1
    
    This.Identity = VBA.vbNullString
    This.Options = VBA.vbNullString
    Set This.Session = New TcpSession
    Set This.DeviceErrorReader = New DeviceErrorReader
    
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="Session"/>. </summary>
Private Sub Class_Terminate()
    
    Me.Dispose

End Sub

''' <summary>   Initializes and returns <see cref="Device"/>. </summary>
''' <returns>   [<see cref="Device"/>]. </returns>
Public Function Initialize() As Device

    This.Session.Initialize cc_isr_Winsock.Factory.NewIPv4StreamSocket()
    This.Session.GpibLanControllerPort = This.GpibLanControllerPort
    This.Session.Termination = This.Termination
    This.Session.ReadAfterWriteDelay = This.ReadAfterWriteDelay
    
    ' set the event handler after initializing the Tcp Session to ensure that the
    ' Tcp Session handles the connection events first.
    Set m_socket = This.Session.Socket

    This.DeviceErrorReader.Initialize This.Session
    
    Set Initialize = Me
   
End Function

''' <summary>   Closes the connection and releases the <see cref="Session">Tcp Session</see>. </summary>
Public Sub Dispose()
    
    If Not Me.Connectable Is Nothing Then
        If Me.Connectable.Connected Then
            Me.Connectable.CloseConnection
        End If
    End If

    If Not Me.Connectable Is Nothing Then
        Me.Connectable.Dispose
    End If
    
    Set m_socket = Nothing
    Set This.Session = Nothing
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Session Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the Session. </summary>
''' <value>   [<see cref="Session"/>]. </returns>
Public Property Get Session() As TcpSession
    Set Session = This.Session
End Property

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Long]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = This.GpibLanControllerPort
End Property

''' <summary>   Sets the port number of the GPIB Lan controller. </summary>
''' <param name="a_value">   [Long] the GPIB Lan controller port. </param>
Public Property Let GpibLanControllerPort(ByVal a_value As Long)
    This.GpibLanControllerPort = a_value
    If Not This.Session Is Nothing Then _
        This.Session.GpibLanControllerPort = a_value
End Property

''' <summary>   Gets the Read After Write Delay in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReadAfterWriteDelay() As Integer
    ReadAfterWriteDelay = This.ReadAfterWriteDelay
End Property

''' <summary>   Sets the Read After Write Delay Delay in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Read After Write Delay Delay in milliseconds. </param>
Public Property Let ReadAfterWriteDelay(ByVal a_value As Integer)
    This.ReadAfterWriteDelay = a_value
    If Not This.Session Is Nothing Then _
        This.Session.ReadAfterWriteDelay = a_value
End Property

''' <summary>   Gets the termination to append to output messages. </summary>
''' <value>   [String]. </value>
Public Property Get Termination() As String
    Termination = This.Termination
End Property

''' <summary>   Sets the termination to append to output messages. </summary>
''' <param name="a_value">   [String] the termination. </param>
Public Property Let Termination(ByVal a_value As String)
    This.Termination = a_value
    If Not This.Session Is Nothing Then _
        This.Session.Termination = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Device Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the <see cref="IDeviceErrorReader"/> subsystem. </summary>
''' <value>   An <see cref="IDeviceErrorReader"/>. </returns>
Public Property Get DeviceErrorReader() As IDeviceErrorReader
    Set DeviceErrorReader = This.DeviceErrorReader
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Device Methods
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Clears Status (*CLS) and returns the operation completion reply of '1'. </summary>
''' <remarks>   This method does not await for the message available flag. Thus, if not awaiting operation
''' completion, the method returns control to the calling method immediately where some waiting needs to be
''' done to allow the instrument to clear. The method does nto check for device error (EAV).  </remarks>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*CLS</c> command by querying <c>*CLS;*WAI;*OPC?</c></param>
Public Sub ClearExecutionState(Optional ByVal a_awaitOpc As Boolean = True)
    
    Const p_procedureName As String = "ClearExecutionState"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    p_success = Me.TryClearExecutionState(p_details, a_awaitOpc)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & VBA.TypeName(Me) & "." & p_procedureName, _
            " Failed clearing execition state; " & p_details
    
End Sub

''' <summary>   Clears Status (*CLS) and returns true if the instrument returned the
''' operation completion reply of '1'. </summary>
''' <param name="a_details">    [Out, String] Details why this method failed. </param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*CLS</c> command by querying <c>*CLS;*WAI;*OPC?</c></param>
''' <returns>   [Boolean]. </returns>
Public Function TryClearExecutionState(ByRef a_details As String, Optional ByVal a_awaitOpc As Boolean = True) As Boolean

    Const p_procedureName As String = "TryClearExecutionState"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_command As String: p_command = cc_isr_Ieee488.Syntax.ClearExecutionStateCommand
    Dim p_expectedReply As String: p_expectedReply = "1"
    Dim p_reply As String: p_reply = p_expectedReply
    
    If a_awaitOpc Then
        
        p_command = p_command & _
            ";" & cc_isr_Ieee488.Syntax.WaitCommand & _
            ";" & cc_isr_Ieee488.Syntax.OperationCompletedQueryCommand
        p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
        If Not p_success Then _
            p_details = "Failed querying the '" & p_command & "' query; " & p_details
    Else
        p_length = This.Session.TryWriteLine(p_command, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
        If Not p_success Then _
            p_details = "Failed sending the '" & p_command & "' command; " & p_details
    End If

    If p_success Then
    
        p_success = p_expectedReply = p_reply
    
        If Not p_success Then _
            p_details = "Failed clearing the device execution state; The instrument returned '" & _
                p_reply & "' instead of '" & "" & _
                "' in response to the '" & p_command & "' query."
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryClearExecutionState = p_success

End Function

''' <summary>   Returns the identity. </summary>
''' <value>   [String]. </value>
Public Property Get Identity() As String
    If StringExtensions.IsNullOrEmpty(This.Identity) And This.Session.Connected Then
        This.Identity = Me.QueryIdentity
    End If
    Identity = This.Identity
End Property

''' <summary>   Returns the instrument identity using the *IDN? query command. </summary>
''' <returns>   [String]. </returns>
Public Function QueryIdentity() As String
    This.Identity = This.Session.QueryLine(cc_isr_Ieee488.Syntax.IdentityQueryCommand)
    QueryIdentity = This.Identity
End Function

''' <summary>   Operation Complete (*OPC) command. </summary>
''' <returns>   [Long] The number of characters that were sent. </returns>
Public Function OperationComplete() As Long
    
    OperationComplete = This.Session.WriteLine(cc_isr_Ieee488.Syntax.OperationCompleteCommand)

End Function

''' <summary>   Queries and waits for operation completions returning '1' if operation was completed;
''' otherwise, '0'. </summary>
''' <returns>   A string. 1 if completed; otherwise 0. </returns>
Public Function QueryOperationCompleted() As String
    
    Const p_procedureName As String = "QueryOperationCompleted"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_reply As String: p_reply = VBA.vbNullString
    p_success = Me.TryQueryOperationCompleted(p_reply, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
            ThisWorkbook.VBProject.Name & "." & VBA.TypeName(Me) & "." & p_procedureName, _
            " Failed querying operation completed; " & p_details
    
    QueryOperationCompleted = p_reply

End Function

''' <summary>   Queries and waits for operation completions returning true if the instrument replied with the
''' expected operation completion reply value of '1'. </summary>
''' <param name="a_reply">     [Out, String] Teh query reply. </param>
''' <param name="a_details">   [Out, String] Details why this method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryQueryOperationCompleted(ByRef a_reply As String, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryQueryOperationCompleted"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_command As String: p_command = cc_isr_Ieee488.Syntax.OperationCompletedQueryCommand
    Dim p_expectedReply As String: p_expectedReply = "1"
    Dim p_reply As String: p_reply = p_expectedReply: a_reply = VBA.vbNullString
    
    p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If Not p_success Then _
        p_details = "Failed querying the '" & p_command & "' query; " & p_details

    If p_success Then
        p_success = p_expectedReply = p_reply
        If p_success Then
            a_reply = p_reply
        Else
            p_details = "Failed waiting for operation completion; The instrument returned '" & _
                p_reply & "' instead of '" & p_expectedReply & _
                "' in response to the '" & p_command & "' query."
        End If
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryQueryOperationCompleted = p_success

End Function

''' <summary>   Applies a command, waits for implementation and queries operation completion. </summary>
''' <param name="a_command">   [String] the command to apply. </param>
''' <param name="a_details">   [Out, String] Details why this method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryApplyCommandAwaitOperationCompleted(ByVal a_command As String, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryApplyCommandAwaitOperationCompleted"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    ' append wait and operation completion to the command.
    a_command = a_command & _
        ";" & cc_isr_Ieee488.Syntax.WaitCommand & _
        ";" & cc_isr_Ieee488.Syntax.OperationCompletedQueryCommand
    
    Dim p_expectedReply As String: p_expectedReply = "1"
    Dim p_reply As String
    p_length = This.Session.TryQueryLine(a_command, p_reply, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If p_success Then
        p_success = p_expectedReply = p_reply
        If Not p_success Then _
            p_details = "Failed applying command; The instrument returned '" & _
                p_reply & "' instead of '" & p_expectedReply & _
                "' in response to the '" & a_command & "' query."
    Else
        p_details = "Failed applying the '" & a_command & "' command; " & p_details
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryApplyCommandAwaitOperationCompleted = p_success

End Function


''' <summary>   Returns the outcome of the <see cref="QueryOptions"/> method. </summary>
''' <value>   [String]. </value>
Public Property Get Options() As String
    
    Options = This.Options

End Property

''' <summary>   Returns option (instrument specific). </summary>
''' <returns>   [String]. </returns>
Public Function QueryOptions() As String
    
    This.Options = This.Session.QueryLine(cc_isr_Ieee488.Syntax.OptionsQueryCommand)
    QueryOptions = This.Options
    
End Function

''' <summary>   Reads the status byte. </summary>
''' <param name="a_canQuery">   [Optional, Boolean, False] True to send <c>*STB?</c>
'''                             if not <see cref="TcpSession.GpibLanControllerAttached"/>. </param>
''' <returns>   [Integer].
''' </returns>
Public Function ReadStatusByte(Optional ByVal a_canQuery As Boolean = False) As Integer
    
    Const p_procedureName As String = "ReadStatusByte"
    
    Dim p_details As String
    Dim p_statusByte As Integer
    p_statusByte = Me.TryReadStatusByte(p_details, a_canQuery)
    If 0 > p_statusByte Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
            ThisWorkbook.VBProject.Name & "." & VBA.TypeName(Me) & "." & p_procedureName, _
            p_details
    
    ReadStatusByte = p_statusByte
    
End Function

''' <summary>   Tries to read the status byte either using <see cref="TcpSession.TrySerialPoll{String}"/>
''' or <see cref="TcpSession.TrySerialPoll{String}"/>. </summary>
''' <param name="a_details">    [Out, String] details the failure reason. </param>
''' <param name="a_canQuery">   [Optional, Boolean, False] True to send <c>*STB?</c>
'''                             if not <see cref="TcpSession.GpibLanControllerAttached"/>. </param>
''' <returns>   [Integer] The status byte or RECEIVE_ERROR (-1) if failed receiving a reading or
''' failed parsing the reading to an integer.
''' </returns>
Public Function TryReadStatusByte(ByRef a_details As String, _
    Optional ByVal a_canQuery As Boolean = False) As Integer
    
    Const p_procedureName As String = "TryReadStatusByte"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_statusByte As Integer: p_statusByte = 0
    If This.Session.GpibLanControllerAttached Then
        p_statusByte = This.Session.TrySerialPoll(p_details)
    ElseIf a_canQuery Then
        p_statusByte = Me.TryQueryServiceRequestStatus(p_details)
    Else
        ' querying the service request status could cause
        ' a query unterminated error.
        p_statusByte = 0
    End If
    
    If 0 > p_statusByte Then _
        a_details = ThisWorkbook.VBProject.Name & "." & VBA.TypeName(Me) & ".TryReadStatusByte " & p_details
    
    TryReadStatusByte = p_statusByte
    
End Function

''' <summary>   Checks if a service is requested. </summary>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <param name="canQuery">    [Optional, Boolean, False] True if the <c>*STB?</c>
'''                            can be used to query the status byte in
'''                            case <see cref="TcpSession.GpibLanControllerAttached"/>
'''                            is false. </param>
''' <returns>   [Integer] <c>1</c> if the SRQ signal is asserted (low),
''' <c>0</c> if the signal is not asserted (high), or -1 if the query failed. </returns>
Public Function TryQueryServiceRequested(ByRef a_details As String, _
    Optional ByVal a_canQuery As Boolean = False) As Integer
    
    Const p_procedureName As String = "TryQueryServiceRequested"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_result As Integer
    If This.Session.GpibLanControllerAttached Then
        p_result = This.Session.TryQueryServiceRequested(p_details)
    ElseIf a_canQuery Then
        Dim p_statusByte As Integer
        p_statusByte = Me.TryQueryServiceRequestStatus(p_details)
        If p_statusByte = -1 Then
            p_result = -1
        ElseIf cc_isr_Ieee488.ServiceRequestFlags.RequestingService = _
            (cc_isr_Ieee488.ServiceRequestFlags.RequestingService And p_statusByte) Then
            p_result = 1
        Else
            p_result = 0
        End If
    Else
        p_result = 0
    End If
    
    p_success = p_result <> -1
    If p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
        "; " & p_details

    TryQueryServiceRequested = p_result
    
End Function

''' <summary>   Issues a wait (*WAI) command. </summary>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                            after issuing the <c>*WAI</c> command by querying <c>*WAI; *OPC?</c></param>
Public Sub Wait(Optional ByVal a_awaitOpc As Boolean = True)
    
    Const p_procedureName As String = "Wait"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    p_success = Me.TryWait(p_details, a_awaitOpc)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed waiting; " & p_details

End Sub

''' <summary>   Issues a wait (*WAI) command. </summary>
''' <param name="a_details">    [Out, String] details the failure reason. </param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*WAI</c> command by querying <c>*WAI; *OPC?</c></param>
''' <returns>   [String] The operation completion reply; defaults to "1". </returns>
Public Function TryWait(ByRef a_details As String, Optional ByVal a_awaitOpc As Boolean = True) As String
    
    Const p_procedureName As String = "TryWait"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_command As String: p_command = cc_isr_Ieee488.Syntax.WaitCommand
    Dim p_expectedReply As String: p_expectedReply = "1"
    Dim p_reply As String: p_reply = p_expectedReply
    
    If a_awaitOpc Then
        
        p_command = p_command & _
            ";" & cc_isr_Ieee488.Syntax.OperationCompletedQueryCommand
        p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
        If Not p_success Then _
            p_details = "Failed querying the '" & p_command & "' query; " & p_details
    Else
        
        p_length = This.Session.TryWriteLine(p_command, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
        If Not p_success Then _
            p_details = "Failed sending the '" & p_command & "' command; " & p_details
    End If
    
    If p_success Then
        p_success = p_expectedReply = p_reply
        If Not p_success Then _
            p_details = "Failed waiting; The instrument returned '" & _
                p_reply & "' instead of '" & p_expectedReply & _
                "' in response to the '" & p_command & "' query."
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryWait = p_reply

End Function

''' <summary>   Enables standard events using the (*ESE {0}) command. </summary>
''' <param name="a_bitMask">    [Integer] Defines the bits corresponding to the standard events
'''                             to enable. <param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*ESE #</c> command by querying <c>*ESE #; *OPC?</c></param>
Public Sub EnableStandardEvents(ByVal a_bitMask As Integer, _
        Optional ByVal a_awaitOpc As Boolean = True)
    
    Const p_procedureName As String = "EnableStandardEvents"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    p_success = TryEnableStandardEvents(a_bitMask, p_details, a_awaitOpc)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed enabling standard events; " & p_details
    
End Sub

''' <summary>   Enables standard events using the (*ESE {0}) command. </summary>
''' <param name="a_bitMask">    [Integer] Defines the bits corresponding to the standard events
'''                             to enable. <param>
''' <param name="a_details">    [Out, String] details the failure reason. </param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*ESE #</c> command by querying <c>*ESE #; *OPC?</c></param>
''' <returns>   [String] The operation completion reply; defaults to "1". </returns>
Public Function TryEnableStandardEvents(ByVal a_bitMask As Integer, _
        ByRef a_details As String, Optional ByVal a_awaitOpc As Boolean = True) As Boolean
    
    Const p_procedureName As String = "TryEnableStandardEvents"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_expectedReply As String: p_expectedReply = "1"
    Dim p_reply As String: p_reply = p_expectedReply
    
    Dim p_command As String
    p_command = StringExtensions.StringFormat(cc_isr_Ieee488.Syntax.StandardEventEnableCommand, a_bitMask)
    
    If a_awaitOpc Then
        p_command = p_command & _
            ";" & cc_isr_Ieee488.Syntax.WaitCommand & _
            ";" & cc_isr_Ieee488.Syntax.OperationCompletedQueryCommand
        p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
        If Not p_success Then _
            p_details = "Failed querying the '" & p_command & "' query; " & p_details
    Else
        p_length = This.Session.TryWriteLine(p_command, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
        If Not p_success Then _
            p_details = "Failed sending the '" & p_command & "' command; " & p_details
    End If
    
    If p_success Then
        p_success = p_expectedReply = p_reply
        If Not p_success Then _
            p_details = "Failed enabling standard events; The instrument returned '" & _
                p_reply & "' instead of '" & p_expectedReply & _
                "' in response to the '" & p_command & "' query."
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
   
    TryEnableStandardEvents = p_success
    
End Function

''' <summary>   Returns the standard events enable byte using the *ESE? query command. </summary>
''' <returns>   [Integer] The standard event register enable status value. </returns>
Public Function QueryStandardEventsEnable() As Integer
    
    Const p_procedureName As String = "QueryStandardEventsEnable"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_value As Integer
    p_value = Me.TryQueryStandardEventsEnable(p_details)
    p_success = p_value >= 0
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed querying standard events enable status; " & p_details

    QueryStandardEventsEnable = p_value

End Function

''' <summary>   Returns the standard events enable byte using the *ESE? query command. </summary>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Integer] The standard event register enable status value or -1 if failed. </returns>
Public Function TryQueryStandardEventsEnable(ByRef a_details As String) As Integer
    
    Const p_procedureName As String = "TryQueryStandardEventsEnable"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_command As String: p_command = cc_isr_Ieee488.Syntax.StandardEventEnableQueryCommand
    Dim p_reply As String: p_reply = VBA.vbNullString
    Dim p_value As Integer: p_value = -1
    
    p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If Not p_success Then _
        p_details = "Failed querying the '" & p_command & "' query; " & p_details
    
    If p_success Then _
        p_success = cc_isr_Core.StringExtensions.TryParseInteger(p_reply, p_value, p_details)
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryQueryStandardEventsEnable = p_value

End Function

''' <summary>   Returns the standard events status byte using the *ESR? query command. </summary>
''' <returns>   [Integer] The standard event register status value or -1 if failed. </returns>
Public Function QueryStandardEventsStatus() As Integer
    
    Const p_procedureName As String = "QueryStandardEventsStatus"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_value As Integer
    p_value = Me.TryQueryStandardEventsStatus(p_details)
    p_success = p_value >= 0
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed querying standard events status; " & p_details

    QueryStandardEventsStatus = p_value
    
End Function

''' <summary>   Returns the standard events status byte using the *ESR? query command. </summary>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Integer] The standard event register status value or -1 if failed. </returns>
Public Function TryQueryStandardEventsStatus(ByRef a_details As String) As Integer
    
    Const p_procedureName As String = "TryQueryStandardEventsStatus"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_command As String: p_command = cc_isr_Ieee488.Syntax.StandardEventStatusQueryCommand
    Dim p_reply As String: p_reply = VBA.vbNullString
    Dim p_value As Integer: p_value = -1
    
    p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If Not p_success Then _
        p_details = "Failed querying the '" & p_command & "' query; " & p_details
    
    If p_success Then
        p_success = cc_isr_Core.StringExtensions.TryParseInteger(p_reply, p_value, p_details)
        If p_success Then _
            p_details = "Failed parsing standard event status byte '" & p_reply & _
                "' using the '" & p_command & "'; " & p_details
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryQueryStandardEventsStatus = p_value
    
End Function

''' <summary>   Enables Service Request using the (*SRE {0}) command. </summary>
''' <param name="a_bitMask">    [Integer] Defines the bits corresponding to the service request
'''                             events to enable. <param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*SRE #</c> command by querying
'''                             <c>*SRE #; *OPC?</c></param>
''' <returns>   [String] The operation completion reply; defaults to "1". </returns>
Public Sub EnableServiceRequest(ByVal a_bitMask As Integer, _
        Optional ByVal a_awaitOpc As Boolean = True)
    
    Const p_procedureName As String = "EnableServiceRequest"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    p_success = Me.TryEnableServiceRequest(a_bitMask, p_details, a_awaitOpc)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed enabling service requests; " & p_details
    
End Sub

''' <summary>   Enables Service Request using the (*SRE {0}) command. </summary>
''' <param name="a_bitMask">    [Integer] Defines the bits corresponding to the service request
'''                             events to enable. <param>
''' <param name="a_details">    [Out, String] details the failure reason. </param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*SRE #</c> command by querying
'''                             <c>*SRE #; *OPC?</c></param>
''' <returns>   [String] The operation completion reply; defaults to "1". </returns>
Public Function TryEnableServiceRequest(ByVal a_bitMask As Integer, _
        ByRef a_details As String, Optional ByVal a_awaitOpc As Boolean = True) As String
    
    Const p_procedureName As String = "TryEnableServiceRequest"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_expectedReply As String: p_expectedReply = "1"
    Dim p_reply As String: p_reply = p_expectedReply
    Dim p_command As String
    p_command = StringExtensions.StringFormat(cc_isr_Ieee488.Syntax.ServiceRequestEnableCommand, a_bitMask)
    If a_awaitOpc Then
        
        p_command = p_command & _
            ";" & cc_isr_Ieee488.Syntax.WaitCommand & _
            ";" & cc_isr_Ieee488.Syntax.OperationCompletedQueryCommand
        p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
        
        If Not p_success Then _
            p_details = "Failed querying the '" & p_command & "' query; " & p_details
    Else
        p_length = This.Session.TryWriteLine(p_command, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
        If Not p_success Then _
            p_details = "Failed sending the '" & p_command & "' command; " & p_details
    End If
    
    If p_success Then
        p_success = p_expectedReply = p_reply
    
        If Not p_success Then _
            p_details = "Failed enabling service request; The instrument returned '" & _
                p_reply & "' instead of '" & "" & _
                "' in response to the '" & p_command & "' query."
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryEnableServiceRequest = p_reply
    
    
End Function

''' <summary>   Returns the Service Request enable byte using the *SRE? query command. </summary>
''' <returns>   [Integer] The service request event register enable status value
''' or -1 if failed. </returns>
Public Function QueryServiceRequestEnable() As Integer
    
    Const p_procedureName As String = "QueryServiceRequestEnable"
    
    Dim p_value As Integer: p_value = -1
    Dim p_details As String: p_details = VBA.vbNullString
    p_value = Me.TryQueryServiceRequestEnable(p_details)
    If 0 > p_value Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed querying service request enable; " & p_details
    
    QueryServiceRequestEnable = p_value
    
End Function

''' <summary>   Returns the Service Request enable byte using the *SRE? query command. </summary>
''' <param name="a_details">    [Out, String] details the failure reason. </param>
''' <returns>   [Integer] The service request event register enable status value
''' or -1 if failed. </returns>
Public Function TryQueryServiceRequestEnable(ByRef a_details As String) As Integer
    
    Const p_procedureName As String = "TryQueryServiceRequestEnable"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_command As String: p_command = cc_isr_Ieee488.Syntax.ServiceRequestEnableQueryCommand
    Dim p_reply As String: p_reply = VBA.vbNullString
    Dim p_value As Integer: p_value = -1
    
    p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If p_success Then _
        p_success = cc_isr_Core.StringExtensions.TryParseInteger(p_reply, p_value, p_details)
            
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryQueryServiceRequestEnable = p_value
    
End Function

''' <summary>   Queries the service request status using the <c>*STB?</c> query command. </summary>
''' <param name="a_details">    [Out, String] Details why this method failed. </param>
''' <returns>   [Integer] The status byte or RECEIVE_ERROR (-1) if failed receiving a reading or
''' failed parsing the reading to an integer.
''' </returns>
Public Function TryQueryServiceRequestStatus(ByRef a_details As String) As Integer

    Const p_procedureName As String = "TryQueryServiceRequestStatus"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long

    Dim p_value As Integer: p_value = -1
    Dim p_command As String: p_command = cc_isr_Ieee488.Syntax.ServiceRequestQueryCommand
    Dim p_reply As String: p_reply = VBA.vbNullString
    
    p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If p_success Then
        p_success = cc_isr_Core.StringExtensions.TryParseInteger(p_reply, p_value, p_details)
        If Not p_success Then
            p_value = -1
            p_details = "Failed parsing service request status byte using the '" & p_command & "'; " & p_details
        End If
    Else
        p_value = -1
        p_details = "faild querying service request status byte using the '" & p_command & "'; " & p_details
    End If
    
    If 0 > p_value Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryQueryServiceRequestStatus = p_value

End Function


''' <summary>   Returns the Service Request status byte using the *STB? query command. </summary>
''' <returns>   The service request event register status value. </returns>
Public Function QueryServiceRequestStatus() As Integer
    
    Const p_procedureName As String = "QueryServiceRequestStatus"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_value As Integer: p_value = Me.TryQueryServiceRequestStatus(p_details)
    If 0 > p_value Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed querying service request status; " & p_details
    End If
    
    QueryServiceRequestStatus = p_value
    
End Function

''' <summary>   Reset to known state (*RST) and return the operation completion reply of '1'. </summary>
''' <remarks>   This method does not await for the message available flag. Thus, if not awaiting operation
''' completion, the method returns control to the calling method immediately where some waiting needs to be
''' done to allow the instrument to reset. The method does nto check for device error (EAV).  </remarks>
''' <param name="awaitOpc">    [Optional, Boolean, True] True to wait for operation completion
'''                            after issuing the <c>*SRE</c> command by querying <c>*SRE; *OPC?</c></param>
Public Sub ResetKnownState(Optional ByVal a_awaitOpc As Boolean = True)
    
    Const p_procedureName As String = "ResetKnownState"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    p_success = Me.TryResetKnownState(p_details, a_awaitOpc)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & VBA.TypeName(Me) & "." & p_procedureName, _
            " Failed resetting known state; " & p_details
    
End Sub

''' <summary>   Reset to known state (*RST) and return true if the instrument returned the
''' the operation completion reply of '1'. </summary>
''' <param name="a_details">    [Out, String] Details why this method failed. </param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*CLS</c> command by querying <c>*CLS;*WAI;*OPC?</c></param>
''' <returns>   [Boolean]. </returns>
Public Function TryResetKnownState(ByRef a_details As String, Optional ByVal a_awaitOpc As Boolean = True) As Boolean

    Const p_procedureName As String = "TryResetKnownState"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_command As String: p_command = cc_isr_Ieee488.Syntax.ResetKnownStateCommand
    Dim p_expectedReply As String: p_expectedReply = "1"
    Dim p_reply As String: p_reply = p_expectedReply
    
    If a_awaitOpc Then
        
        p_command = p_command & _
            ";" & cc_isr_Ieee488.Syntax.WaitCommand & _
            ";" & cc_isr_Ieee488.Syntax.OperationCompletedQueryCommand
        p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
        If Not p_success Then _
            p_details = "Failed querying the '" & p_command & "' query; " & p_details
    Else
        p_length = This.Session.TryWriteLine(p_command, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
        If Not p_success Then _
            p_details = "Failed sending the '" & p_command & "' command; " & p_details
    End If

    If p_success Then
    
        p_success = p_expectedReply = p_reply
    
        If Not p_success Then _
            p_details = "Failed resetting the device known state; The instrument returned '" & _
                p_reply & "' instead of '" & "" & _
                "' in response to the '" & p_command & "' query."
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryResetKnownState = p_success

End Function

''' <summary>   Enumerates device, queued and archived errors. </summary>
''' <para name="a_delimiter">   [String] a delimiter for concatenating error messages
'''                             from the error queue. </param>
Public Function EnumerateErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As String

    Const p_procedureName As String = "EnumerateErrors"

    Dim p_archivedErrorCount As Integer
    p_archivedErrorCount = cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount
    
    Dim p_queuedErrorCount As Integer
    p_queuedErrorCount = cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount

    Dim p_deviceErrors As String
    Dim p_deviceErrorsCount As Integer
    If This.DeviceErrorReader Is Nothing Or This.Session Is Nothing Then
        p_deviceErrorsCount = 0
        p_deviceErrors = VBA.vbNullString
    ElseIf Me.Connected Then
        p_deviceErrorsCount = This.DeviceErrorReader.TryDequeueErrors(p_deviceErrors)
    End If
    
    ' if has device errors, enqueue the device error(s) to be asserted below.
    If p_deviceErrorsCount > 0 Then
        cc_isr_Core_IO.UserDefinedErrors.EnqueueError _
            cc_isr_Core_IO.UserDefinedErrors.Ieee488DeviceError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " The 2700 instrument error(s): " & p_deviceErrors
    End If
    
    Dim p_outcome As String
    p_outcome = VBA.vbNullString
    If p_archivedErrorCount + p_queuedErrorCount + p_deviceErrorsCount > 0 Then
        
        ' if we have leftover errors, move the queued errors to the archive
        Dim p_queuedErrors As String
        p_queuedErrors = cc_isr_Core_IO.UserDefinedErrors.ArchiveQueuedErrors()
        
        ' build the errors to report from the archive.
        Dim p_archivedErrors As String
        p_archivedErrors = cc_isr_Core_IO.UserDefinedErrors.PopArchivedErrors()
        
        p_outcome = VBA.vbNullString
        If p_deviceErrorsCount > 0 Then _
            p_outcome = "Found instrument error(s) (" & VBA.CStr(p_deviceErrorsCount) & ")"
            
        If p_queuedErrorCount > 0 Then _
            p_outcome = p_outcome & IIf(p_deviceErrorsCount > 0, " and ", "Found ") & _
                "Enqueued errors (" & VBA.CStr(p_queuedErrorCount) & ")"

        If p_archivedErrorCount > 0 Then _
            p_outcome = p_outcome & IIf(p_deviceErrorsCount + p_queuedErrorCount > 0, " and ", "Found ") & _
                "Archived errors (" & VBA.CStr(p_queuedErrorCount) & ")"
        
        p_outcome = p_outcome & ":" & VBA.vbCrLf & p_archivedErrors
        
    End If
    
    EnumerateErrors = p_outcome

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Restoring state
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Clears the instrument output and then clears the instrument execution state. </summary>
''' <remarks>   This method catches and hides any execution errors. </remarks>
''' <param name="a_command">   [Option, String, <c>*CLS;*WAI;*OPC?</c>] the command. </param>
''' <returns>   [String]. </returns>
Public Function QueryClearExecutionState(Optional ByVal a_command As String = "*CLS;*WAI;*OPC?") As String
    
    QueryClearExecutionState = This.Session.QueryClearExecutionState(a_command)

End Function

''' <summary>   Configures the IEEE 488 Device initial state. </summary>
''' <remarks>   Call this method upon connecting and restoring Initial State. </remarks>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryConfigureInitialState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryConfigureInitialState"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' check if the socket is connected.
    p_success = This.Session.Socket.Connected
    If p_success Then
        Dim p_tries As Integer: p_tries = 0
        ' reset and clear the instrument.
        p_success = Me.TryResetClears(p_details, p_tries, 3)
    Else
        p_details = "Socket is not connected to '" & This.Session.SocketAddress & "'."
    End If
    
    If p_success Then _
        p_success = This.Session.TryEstablishConnection(p_success, p_details)
    
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryConfigureInitialState = p_success

End Function

''' <summary>   Tries to reset and clear the device. </summary>
''' <param name="a_details">    [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryResetClear(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryResetClear"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_reply As String: p_reply = VBA.vbNullString
    Dim p_length As Long
    
    ' selective device clear
    p_success = This.Session.TrySelectiveDeviceClear(p_details, False, 200)
    
    ' Reset known state
    p_length = This.Session.TryWriteLine(cc_isr_Ieee488.Syntax.ResetKnownStateCommand, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    ' Clear execution state
    p_length = This.Session.TryWriteLine(cc_isr_Ieee488.Syntax.ClearExecutionStateCommand, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    ' disable all bits in the service request event register.
    p_length = This.Session.TryWriteLine(cc_isr_Ieee488.Syntax.ServiceRequestEnableCommand & " 0", p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    ' disable all bits in the standard event enable register.
    p_length = This.Session.TryWriteLine(cc_isr_Ieee488.Syntax.StandardEventEnableCommand & " 0", p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    ' query operation completion; success means a connection was establised.
    p_success = Me.TryQueryOperationCompleted(p_reply, p_details)
        
    If Not p_success Then _
        p_details = "Failed to reset clear the instrument; " & p_details
    
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details

    TryResetClear = p_success

End Function

''' <summary>   Tries multiple reset clears of the device. </summary>
''' <param name="a_details">    [Out, String] Details why the method failed. </param>
''' <param name="a_count">      [Out, Integer] gets the actual try count. </param>
''' <param name="a_tryCount">   [Optional, Integer, 3] sets how many times to retry. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryResetClears(ByRef a_details As String, _
    ByRef a_count As Integer, Optional ByVal a_tryCount As Integer = 3) As Boolean

    Const p_procedureName As String = "TryResetClear"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_reply As String: p_reply = VBA.vbNullString
    Dim p_length As Long
    
    p_success = False
    a_count = 0
    
    ' loop throught the following systems a few times
    Do Until p_success Or (a_tryCount <= a_count)
    
        ' increment the actual count.
        a_count = a_count + 1
        
        ' try anotehr eeset and clear.
        p_success = Me.TryResetClear(p_details)
        
    Loop
        
    If Not p_success Then _
        p_details = "Failed resetting and clearing the instrument after " & VBA.CStr(a_count) & "tries; " & p_details
    
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details

    TryResetClears = p_success

End Function

''' <summary>   Checks if the device needs to restore to its initial state. </summary>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            should restore state. </param>
''' <returns>   [Boolean]. </returns>
Public Function ShouldRestoreInitialState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "ShouldRestoreInitialState"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    If This.Session Is Nothing Then
        a_details = "The Tcp Session instance is nothing."
        ShouldRestoreInitialState = True
    Else
        ShouldRestoreInitialState = This.Session.ShouldRestoreInitialState(p_details)
    End If

    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details

End Function

''' <summary>   Restores the instrument to its initial state in case the instrument power
''' was toggled or a recovery from error is needed. </summary>
''' <remarks>   This method attempts to restore the instrument to a Initial State in case it is
''' powered off while the program is active as follows:
''' - The user defined errors are cleared;
''' - The class is initialized if the session is nothing;
''' - The device is connected if it is not connected;
''' - The GPIB-Lan device is restored to its Initial State;
''' - Execution state is cleared.
''' </remarks>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean] true if restored. </returns>
Public Function TryRestoreInitialState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryRestoreInitialState"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' clear the error state
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState

    If p_success And This.Session Is Nothing Then
    
        p_details = "Unable to restore state because the Tcp Session device was disposed. Please restart the program."
        p_success = False
    
    End If
    
    If p_success Then
    
        If This.Session.Socket Is Nothing Then Me.Initialize
        
    End If
    
    If p_success And Not (Me.Connected Or This.Session.Socket.CanConnect) Then
        
        ' if not connected and not 'can connect' it means that the
        ' socket was disposed and needs to be rebuilt.
        Me.Initialize
            
        p_success = This.Session.Socket.CanConnect
        
        If Not p_success Then
            p_details = "Failed initializing the socket; the socket id is invalid ('" & _
                VBA.CStr(cc_isr_Winsock.ws32_INVALID_SOCKET) & "')."
        End If
    
    End If
    
    If p_success Then
    
        If This.Session.TryRestoreInitialState(p_details) Then
        
            ' remove any accumulated errors
            
            p_details = Me.EnumerateErrors()
            
            p_details = VBA.vbNullString
            
            ' clear the instrument execution state to remove any lingering errors
            Dim p_reply As String: p_reply = VBA.vbNullString
            p_success = Me.TryQueryOperationCompleted(p_reply, p_details)
            
        Else
        
            p_details = "Failed restoring Tcp Session state; " & p_details
                
        End If
    
    End If
    
    ' configure the initial state
    If p_success Then _
        p_success = Me.TryConfigureInitialState(p_details)

    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryRestoreInitialState = p_success

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Connectable implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets a reference to the <see cref="cc_isr_Winsock.IConnectable"/>
''' <see cref="Session"/see>.<see cref="TcpSession.Socket"/see>. </summary>
''' <value>   [<see cref="cc_isr_Winsock.IConnectable"/>]. </value>
Public Property Get Connectable() As cc_isr_Winsock.IConnectable
    If This.Session Is Nothing Then
        Set Connectable = Nothing
    ElseIf This.Session.Connectable Is Nothing Then
        Set Connectable = Nothing
    Else
        Set Connectable = This.Session.Connectable
    End If
End Property

''' <summary>   Returns the connection state of the <see cref="Session">Tcp Session</see>. </summary>
''' <value>   [Boolean] True if the <see cref="Session">Tcp Session</see> is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
        
    If Me.Connectable Is Nothing Then
        Connected = False
    Else
        Connected = Me.Connectable.Connected
    End If

End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean
    
    If Me.Connectable Is Nothing Then
        CanConnect = False
    Else
        CanConnect = Me.Connectable.CanConnect
    End If

End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Tcp Session Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="IPv4StreeamSocket.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_socket_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
    
    Const p_procedureName As String = "m_socket_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_eventArgs.Connected Then
                                                
        Dim p_details As String
         
        ' configure the initial state
        If Not Me.TryConfigureInitialState(p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.EnqueueError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Failed configuring the initial state; " & p_details
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the <see cref="IPv4StreamSocket.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_socket_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

    Const p_procedureName As String = "m_socket_ConnectionChanging"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' no longer rasing connection events here.
    ' OnConnectionChanging a_eventArgs
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub



