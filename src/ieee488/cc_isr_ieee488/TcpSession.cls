VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A TCP Session capable of communicating with a GPIB-Lan controller. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    DelayStopper As cc_isr_Core_IO.Stopwatch
    
    ' socket settins
    Socket As cc_isr_Winsock.IPv4StreamSocket
    SocketAddress As String
    SocketTimeout As Long
    Established As Boolean
    
    ' Session settings
    ConnectionQueryMessage As String
    ConnectionQueryReply As String
    SynchronizationQueryMessage As String
    SynchronizationQueryReply As String
    GpibLanControllerPort As Long
    ReadAfterWriteDelay As Long
    Termination As String
    MinimumTimeout As Integer
    
    ' Controller settings
    ReadTimeoutMinimum As Integer
    ReadTimeoutMaximum As Integer
    
    ' Cached GPIB-Lan Controller state
    AppendTermination As AppendTerminationOption
    AutoAssertTalk As Boolean
    EndOrIdentify As Boolean
    ReadTimeout As Integer
    GpibAddressReading As String
    PrimaryGpibAddress As Integer
    SecondaryGpibAddress As Integer
    
End Type

Private This As this_

Private WithEvents m_socket As cc_isr_Winsock.IPv4StreamSocket
Attribute m_socket.VB_VarHelpID = -1

''' <summary>   Handles the Class Initialize event. </summary>
Private Sub Class_Initialize()

    Set This.DelayStopper = cc_isr_Core_IO.Factory.NewStopwatch
    This.GpibLanControllerPort = 1234
    This.ReadAfterWriteDelay = 1
    This.Termination = VBA.vbLf
    This.MinimumTimeout = 1000
    This.ReadTimeoutMaximum = 3000
    This.ReadTimeoutMinimum = 1
    This.ConnectionQueryMessage = _
            cc_isr_Ieee488.Syntax.ClearExecutionStateCommand & ";" & _
            cc_isr_Ieee488.Syntax.WaitCommand & ";" & _
            cc_isr_Ieee488.Syntax.OperationCompletedQueryCommand
    This.ConnectionQueryReply = "1"
    This.SynchronizationQueryMessage = _
            cc_isr_Ieee488.Syntax.OperationCompletedQueryCommand
    This.SynchronizationQueryReply = "1"
    
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="Socket"/>. </summary>
Private Sub Class_Terminate()

    Set This.Socket = Nothing
    Set m_socket = Nothing
    Set This.DelayStopper = Nothing

End Sub

''' <summary>   Returns this unconnected Tcp Client instance. No actual initialization is done
''' at this time. </summary>
''' <remarks>   Raises an invalid operation exception is a socket is already connected. </remarks>
''' <param name="a_socket ">   [See cref="cc_isr_Winsock.IPv4StreamSocket"/>] An INET stream socket. </param>
''' <returns>   [<see cref="TcpSession"/>]. </returns>
Public Function Initialize(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket) As TcpSession
    
    Set This.Socket = a_socket
    Set m_socket = a_socket
    Set Initialize = Me
    
End Function

''' <summary>   Gets the Tcp Client socket. </summary>
''' <value>   [<see cref="IPv4StreamSocket"/>]. </value>
Public Property Get Socket() As cc_isr_Winsock.IPv4StreamSocket
    Set Socket = This.Socket
End Property

''' <summary>   Gets a reference to the <see cref="cc_isr_Winsock.IConnectable"/>
''' <see cref="ViSession"/see>.<see cref="TcpSession.Socket"/see>. </summary>
''' <value>   [<see cref="cc_isr_Winsock.IConnectable"/>]. </value>
Public Property Get Connectable() As cc_isr_Winsock.IConnectable
    If This.Socket Is Nothing Then
        Set Connectable = Nothing
    Else
        Set Connectable = This.Socket
    End If
End Property

''' <summary>   Returns the connection state of the socket the attached
''' instrumet. </summary>
''' <value>   [Boolean] True if the socket is connected and the connection with
''' the instrument was established; otherwise, False. </value>
Public Property Get Connected() As Boolean
        
    If This.Socket Is Nothing Then
        Connected = False
    Else
        Connected = This.Socket.Connected And This.Established
    End If

End Property

''' <summary>   Gets the connection query message. </summary>
''' <value>   [String]. </value>
Public Property Get ConnectionQueryMessage() As String
    ConnectionQueryMessage = This.ConnectionQueryMessage
End Property

''' <summary>   Sets the connection query message. </summary>
''' <param name="a_value">   [String] The connection query message. </param>
Public Property Let ConnectionQueryMessage(ByVal a_value As String)
    This.ConnectionQueryMessage = a_value
End Property

''' <summary>   Gets the connection query Reply. </summary>
''' <value>   [String]. </value>
Public Property Get ConnectionQueryReply() As String
    ConnectionQueryReply = This.ConnectionQueryReply
End Property

''' <summary>   Sets the connection query Reply. </summary>
''' <param name="a_value">   [String] The connection query Reply. </param>
Public Property Let ConnectionQueryReply(ByVal a_value As String)
    This.ConnectionQueryReply = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  I/O
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Address of the last connected socket. </summary>
''' <value>   [String]. </value>
Public Property Get SocketAddress() As String
    SocketAddress = This.SocketAddress
End Property

''' <summary>   Gets the Socket Timeout in milliseconds. </summary>
''' <remarks>
''' TCP Session timeout considerations:
''' The TCP Session has two timeouts: socket and read timeouts.
''' The socket timeout is the timeout assigned to the socket from reading from the Lan terminal.
''' The read timeout is the timeout for the serial poll and read command for the
''' GPIB-Lan device to read from the instrument.
''' The socket timeout is set when the socket is connected.
''' The read timeout is set when the session state is initialized upon connecting.
''' The socket timeout reflects the time it might take to receive data from the instrument.
''' While the GPIB-Lan is connected to the instrument, its timeout for reading from the
''' instrument essentially depends on the same limitation as the socket timeout, namely,
''' the time it might take the instrument to respond to a query.
''' Thus, the Tcp Session sets these two timeout to the same value as follows:
''' (1) The socket timeout is set when the connection is made.
''' (2) This timeout is cached when the socket connection event is handled.
''' (3) The read timeout is then set as part of the initialization of the GPIB-Lan settings.
''' </remarks>
''' <value>   [Long]. </value>
Public Property Get SocketTimeout() As Long
    SocketTimeout = This.SocketTimeout
End Property

''' <summary>   Sets the Socket Timeout Delay in milliseconds. </summary>
''' <param name="a_value">   [Long] the Socket Timeout Delay in milliseconds. </param>
Public Property Let SocketTimeout(ByVal a_value As Long)
    This.SocketTimeout = a_value
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
End Property

''' <summary>   Gets the Minimum Session Timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get MinimumTimeout() As Integer
    MinimumTimeout = This.MinimumTimeout
End Property

''' <summary>   Sets the Minimum Session Timeout Delay in milliseconds. </summary>
''' <param name="a_value">   [Integer] the MinimumSession Timeout Delay in milliseconds. </param>
Public Property Let MinimumTimeout(ByVal a_value As Integer)
    This.MinimumTimeout = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Instrument send and receive
'  use socket send and receive to communicate with the controller/
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sends an unterminated message to the instrument. </summary>
''' <param name="a_message">   [String] The unterminated message. </param>
''' <param name="a_delay">     [Optional, Integer, 0] the millisecond duration to add to the
'''                            <see cref="ReadAfterWriteDelay"/>. </param>
''' <returns>   [Long] The number of bytes that were sent to the server. </returns>
Public Function WriteLine(ByVal a_message As String, Optional ByVal a_delay As Integer = 0) As Long
        
    Const p_procedureName As String = "WriteLine"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Integer
    
    p_length = Me.TryWriteLine(a_message, p_details, a_delay)
    p_success = (p_length > 0) And (0 = VBA.Len(p_details))
    
    If Not p_success Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    End If
    
    WriteLine = p_length
    
End Function

''' <summary>   Sends an unterminated message to the instrument. </summary>
''' <param name="a_message">   [String] The unterminated message. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <param name="a_delay">     [Optional, Integer, 0] the millisecond duration to add to the
'''                            <see cref="ReadAfterWriteDelay"/>. </param>
''' <returns>   [Long]. 0 if nothing to send, negative (-1) if error,
''' positive if the entire message was sent. </returns>
Public Function TryWriteLine(ByVal a_message As String, _
    ByRef a_details As String, _
    Optional ByVal a_delay As Integer = 0) As Long

    Const p_procedureName As String = "TryWriteLine"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    p_length = This.Socket.TrySendMessage(a_message & This.Termination, _
        p_details, This.ReadAfterWriteDelay + a_delay)
    
    p_success = (p_length > 0) And (0 = VBA.Len(p_details))
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
            
    TryWriteLine = p_length
    
End Function

''' <summary>   This command can be used to read data from an instrument. </summary>
''' <remarks>   Data is read from the instrument until:
''' - EOI is detected or timeout expires, or
''' - A specified character is read or timeout expires, or
''' - Timeout expires
''' Timeout is set using the <c>read_tmo_ms</c> command and applies to inter-character delay, i.e.,
''' the delay since the last character was read. Timeout is not be confused with the total
''' time for which data is read.
''' Syntax: <c>++read [eoi|<char>]</c> where <c>char</c> is a decimal value less than 256, e.g.,
''' <c>++read</c> Read until timeout.
''' <c>++read eoi</c> Read until EOI detected or timeout.
''' <c>++read 10</c> Read until LF (ASCII 10) is received or timeout.
''' </remarks>
Public Property Get ReadUntilTimeoutCommand() As String
    ReadUntilTimeoutCommand = "++read"
End Property

''' <summary>   This command can be used to read data from an instrument. </summary>
''' <remarks>   Data is read from the instrument until:
''' - EOI is detected or timeout expires, or
''' - A specified character is read or timeout expires, or
''' - Timeout expires
''' Timeout is set using the <c>read_tmo_ms</c> command and applies to inter-character delay, i.e.,
''' the delay since the last character was read. Timeout is not be confused with the total
''' time for which data is read.
''' Syntax: <c>++read [eoi|<char>]</c> where <c>char</c> is a decimal value less than 256, e.g.,
''' <c>++read</c> Read until timeout.
''' <c>++read eoi</c> Read until EOI detected or timeout.
''' <c>++read 10</c> Read until LF (ASCII 10) is received or timeout.
''' </remarks>
Public Property Get ReadEOICommand() As String
    ReadEOICommand = "++read eoi"
End Property

''' <summary>   Receives a message from the instrument until reaching the specified termination. </summary>
''' <remarks>   The <see cref="TcpSession.Socket."/>.<see cref="ReceiveTimeout"/> must be set
''' long enough the exceed any instrument or controller delays. </remarks>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [String] The received message. </returns>
Public Function Read(Optional ByVal a_trimEnd As Boolean = True, _
    Optional ByVal a_termination As String = vbLf) As String

    Const p_procedureName As String = "Read"

    Dim p_buffer As String
    Dim p_details As String
    
    Dim p_length As Long
    p_length = Me.TryRead(p_buffer, p_details, a_trimEnd, a_termination)
    
    If 0 > p_length Or 0 < VBA.Len(p_details) Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Function

''' <summary>   Tries to receive a message from the instrument until reaching the specified termination. </summary>
''' <remarks>   The <see cref="ReceiveTimeout"/> must be set long enough the exceed any instrument
''' or controller delays. </remarks>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [Long] If no error, the number of bytes that were received.
''' Otherwise, RECEIVE_ERROR (-1).
''' </returns>
Public Function TryRead(ByRef a_buffer As String, _
                        ByRef a_details As String, _
                        Optional ByVal a_trimEnd As Boolean = True, _
                        Optional ByVal a_termination As String = vbLf) As Long
                           
    Const p_procedureName As String = "TryRead"
                           
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
                           
    If Me.GpibLanControllerAttached Then
        p_length = Me.Socket.TrySendMessage(Me.ReadEOICommand & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    End If
    
    If p_success Then
        p_length = This.Socket.TryReceive(a_buffer, p_details, a_trimEnd, a_termination)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    End If
                           
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
                        
    TryRead = p_length
                        
End Function

''' <summary>   Sends an unterminated query message and reads back from the
''' server until reaching the specified
''' termination. </summary>
''' <param name="a_message">       [String] The unterminated query message. </param>
''' <param name="a_delay">         [Optional, Integer, 0] the millisecond duration to add to the
'''                                <see cref="ReadAfterWriteDelay"/>. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [String] The received message. </returns>
Public Function QueryLine(ByVal a_message As String, _
    Optional ByVal a_delay As Integer = 0, _
    Optional ByVal a_trimEnd As Boolean = True, _
    Optional ByVal a_termination As String = vbLf) As String

    Const p_procedureName As String = "QueryLine"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_buffer As String: p_buffer = VBA.vbNullString
    Dim p_length As Long
    p_length = Me.TryQueryLine(a_message, p_buffer, p_details, a_delay, a_trimEnd, a_termination)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
    QueryLine = p_buffer
    
End Function


''' <summary>   Sends an unterminated query message and reads back from the
''' server until reaching the specified
''' termination. </summary>
''' <param name="a_message">       [String] The unterminated query message. </param>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <param name="a_delay">         [Optional, Integer, 0] the millisecond duration to add to the
'''                                <see cref="ReadAfterWriteDelay"/>. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [Long] If no error, the number of bytes that were received.
''' Otherwise, RECEIVE_ERROR (-1).
''' </returns>
Public Function TryQueryLine(ByVal a_message As String, _
                             ByRef a_buffer As String, _
                             ByRef a_details As String, _
                             Optional ByVal a_delay As Integer = 0, _
                             Optional ByVal a_trimEnd As Boolean = True, _
                             Optional ByVal a_termination As String = VBA.vbLf) As Long
                         
    Const p_procedureName As String = "TryQueryLine"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long: p_length = 0
    
    p_length = Me.TryWriteLine(a_message, p_details, a_delay)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If Not p_success Then
        a_buffer = VBA.vbNullString
    End If
    
    If p_success Then
        p_length = Me.TryRead(a_buffer, p_details, a_trimEnd, a_termination)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryQueryLine = p_length
                         
End Function

''' <summary>   Sends an operation complete query command and waits for a reply. </summary>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if error.
''' </returns>
Public Function TrySynchronize(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TrySynchronize"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    Dim p_message As String: p_message = This.SynchronizationQueryMessage
    Dim p_buffer As String: p_buffer = VBA.vbNullString
    Dim p_length As Long
    
    p_length = Me.TryQueryLine(p_message, p_buffer, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If p_success Then
        p_success = This.SynchronizationQueryReply = p_buffer
        If Not p_success Then _
            p_details = "Failed synchronizing; '" & _
                This.SynchronizationQueryMessage & "' returned '" & p_buffer & _
                "' instead of '" & This.SynchronizationQueryReply & "'."
    Else
        p_details = "Failed synchronizing; " & p_details
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TrySynchronize = p_success

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  GPIB Lan Controller Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Long]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = This.GpibLanControllerPort
End Property

''' <summary>   Sets the port number of the GPIB Lan controller. </summary>
''' <param name="a_value">   [Long] the GPIB Lan controller port. </param>
Public Property Let GpibLanControllerPort(ByVal a_value As Long)
    This.GpibLanControllerPort = a_value
End Property

''' <summary>   True if a GPIB Lan controller is attached. </summary>
''' <value>   [Boolean]. </value>
Public Property Get GpibLanControllerAttached() As Boolean
    If This.Socket Is Nothing Then
        GpibLanControllerAttached = False
    Else
        GpibLanControllerAttached = This.Socket.Port = This.GpibLanControllerPort
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  GPIB Lan Controller Methods
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Tries to get and validate a GPIB-Lan controller settings returning the actual
''' value and true if the correct value was set or false if not. </summary>
''' <remarks>   This method assumes that the bare command is the query command.  </remarks>
''' <param name="a_command">   [String] the bare command, e.g.,<c>++eos</c>. </param>
''' <param name="a_value">     [String] the expected value to get. </param>
''' <param name="a_result">    [Out, String] the actual value that was set. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <value>   [Boolean]. </value>
Public Function TryGetValue(ByVal a_command As String, ByVal a_value As String, _
    ByRef a_result As String, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryGetValue"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    If Me.GpibLanControllerAttached Then
        p_success = True
    Else
        p_success = False
        p_details = "The GPIB-Lan controller must be attached for setting the '" & _
            a_command & "' value to '" & a_value & "'."
    End If

    ' send the query command.
    Dim p_length As Integer
    
    If p_success Then
    
        p_length = This.Socket.TrySendMessage(a_command & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    End If
    
    If p_success Then
    
        p_length = This.Socket.TryReceive(a_result, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
        If Not p_success Then
            p_details = "Command '" & a_command & "' failed; " & p_details
        End If
    
    End If
    
    If p_success Then
    
        p_success = 0 = VBA.StrComp(a_value, a_result, VBA.VbCompareMethod.vbTextCompare)
        
        If Not p_success Then
            p_details = "Command '" & a_command & "' failed; Result '" & _
                a_result & "' is not equal to the expected value '" & a_value & "'."
        End If
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryGetValue = p_success
 
End Function

''' <summary>   Tries to set a GPIB-Lan controller settings. </summary>
''' <param name="a_command">   [String] The command. </param>
''' <param name="a_value">     [String] the value to set. </param>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TrySetValue(ByVal a_command As String, ByVal a_value As String, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TrySetValue"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
        p_success = True
    Else
        p_success = False
        p_details = "The GPIB-Lan controller must be attached for setting the '" & _
            a_command & "' value to '" & a_value & "'."
    End If
    
    If p_success Then
    
        ' send the command
        p_length = Me.Socket.TrySendMessage(a_command & " " & a_value & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    End If
    
    If p_success Then
    
        Dim p_result As String

        ' validate the command
        p_success = Me.TryGetValue(a_command, a_value, p_result, p_details)
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TrySetValue = p_success

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  GPIB Lan Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   This command performs a power-on reset of the controller. </summary>
''' <remarks>   The process takes about 5 seconds. All input received over the network during
''' this time are ignored.
''' Syntax: <c>++rst</c>. </remarks>
Public Property Get PowerOnResetCommand() As String
    PowerOnResetCommand = "++rst"
End Property

''' <summary>   Preforms a power on reset of the controller and wait until its completion. </summary>
''' <remarks>   While the manual indicates a delay of 5 seconds, a delay of 1 seconds still
''' permitted reconnecting after power on reset although with some wait time.  With a delay of
''' 3 seconds, reconnecting was virtually immediated. </remarks>
''' <param name="a_socketAddress">   [String] The socket address to reset. </param>
''' <param name="a_details">         [Out, String] Details why the method failed. </param>
''' <param name="a_delay">           [Optional, Double, 3] the power reset delay in seconds. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryPowerOnReset(ByVal a_socketAddress As String, _
    ByRef a_details As String, Optional a_delay As Double = 3) As Boolean

    Const p_procedureName = "TryPowerOnReset"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_socket As cc_isr_Winsock.IPv4StreamSocket
    Set p_socket = cc_isr_Winsock.Factory.NewIPv4StreamSocket().Initialize()
    
    p_success = p_socket.TryOpenConnection(a_socketAddress, 3000, p_details)
    
    If p_success And p_socket.Port = This.GpibLanControllerPort Then
    
        p_length = p_socket.TrySendMessage(Me.PowerOnResetCommand & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
        If p_success Then
            ' now wait the power on reset delay. Must be over 5 seconds.
            Dim p_endTime As Double: p_endTime = VBA.CDbl(VBA.Now) + a_delay / (60# * 60 * 24)
            While VBA.CDbl(VBA.Now) < p_endTime
                DoEvents
            Wend
        End If
    
    End If
    
    If p_success Then _
        p_success = p_socket.TryCloseConnection(p_details)
   
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details

    TryPowerOnReset = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB METHODS
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   This command enables front panel operation of the currently addressed instrument. </summary>
''' <remarks>   Syntax: <c>++loc</c>. </remarks>
Public Property Get GoToLocalCommand() As String
    GoToLocalCommand = "++loc"
End Property

''' <summary>   Issues a Go To Local (GTL) (++loc). </summary>
''' <param name="a_delay">   [Optional, Long, 100] milliseconds duration to wait after
'''                          executing the command. </param>
Public Sub GoToLocal(Optional ByVal a_delay As Long = 100)
    
    Const p_procedureName = "GoToLocal"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
        p_length = Me.Socket.TrySendMessage(Me.GoToLocalCommand & This.Termination, p_details, a_delay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    End If
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Sub

''' <summary>   This command disables front panel operation of the currently addressed instrument. </summary>
''' <remarks>   Syntax: <c>++llo</c>. </remarks>
Public Property Get LocalLockoutCommand() As String
    LocalLockoutCommand = "++llo"
End Property

''' <summary>   Issues a local lockout (++llo). </summary>
''' <param name="a_delay">   [Optional, Long, 100] milliseconds duration to wait after
'''                          executing the command. </param>
Public Sub LocalLockout(Optional ByVal a_delay As Long = 100)
    
    Const p_procedureName = "LocalLockout"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
        p_length = Me.Socket.TrySendMessage(Me.LocalLockoutCommand & This.Termination, p_details, a_delay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    End If
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Sub

''' <summary>   This command issues a selective device clear on the currently addressed instrument. </summary>
''' <remarks>   Syntax: <c>++clr</c>. </remarks>
Public Property Get SelectiveDeviceClearCommand() As String
    SelectiveDeviceClearCommand = "++clr"
End Property

''' <summary>   Issues an SDC (selective device clear). </summary>
''' <remarks>   On the 2700 synchronization failed if the instrument was in
''' external triggering mode.
''' This can be avoided by issueing the <c>*RST</c> command first. </remakrs>
''' <param name="a_details">       [Out, String] Details why the method failed. </param>
''' <param name="a_synchronize">   [Optional, Boolean, False] True to query operation completion
'''                                after executing the command. </param>
''' <param name="a_delay">         [Optional, Long, 10] milliseconds duration to wait after
'''                                executing the command. </param>
Public Function TrySelectiveDeviceClear(ByRef a_details As String, _
    Optional ByVal a_synchronize As Boolean = False, _
    Optional ByVal a_delay As Long = 10)
    
    Const p_procedureName As String = "TrySelectiveDeviceClear"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
        
        p_length = Me.Socket.TrySendMessage(Me.SelectiveDeviceClearCommand & This.Termination, _
            p_details, a_delay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
        
        If Not p_success Then _
            p_details = "Failed sending the '" & Me.SelectiveDeviceClearCommand & "'; " & p_details
        
        If p_success And a_synchronize Then _
            p_success = Me.TrySynchronize(p_details)
        
        If Not p_success Then _
            p_details = "Failed trying to synchronize following a selective device clear; " & p_details
            
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TrySelectiveDeviceClear = p_success
    
End Function

''' <summary>   Gets the GPIB primary address. </summary>
''' <value>   [Integer]. </value>
Public Property Get PrimaryGpibAddress() As Integer
    PrimaryGpibAddress = This.PrimaryGpibAddress
End Property

''' <summary>   Gets the GPIB secondary address. </summary>
''' <value>   [Integer]. </value>
Public Property Get SecondaryGpibAddress() As Integer
    SecondaryGpibAddress = This.SecondaryGpibAddress
End Property

''' <summary>   Configures or queries the GPIB address of the GPIB Lan Controller.
''' The meaning of the GPIB address depends on the operating mode of the controller.
'''
''' In Controller mode, it refers to the GPIB address of the instrument being controlled.
''' In DEVICE mode, it is the address of the GPIB peripheral that the controller is emulating. </summary>
''' <remarks>
''' An optional secondary address may also be specified.
'''
''' Internally, the secondary address, which is offset by 96, must be separated from the primary
''' address by a space character. Specifying secondary address has no effect in DEVICE mode.
'''
''' Syntax: ++addr [<PAD> [<SAD>]]
''' PAD (Primary Address) is a decimal value between 0 and 30.
''' SAD (Secondary Address) is a decimal value between 96 and 126. SAD is optional.
''' </remarks>
Public Property Get GpibAddressCommand() As String
    GpibAddressCommand = "++addr"
End Property

''' <summary>   Gets the GPIB address reading. </summary>
''' <value>   [String]. </value>
Public Property Get GpibAddressReading() As String
    GpibAddressReading = This.GpibAddressReading
End Property

''' <summary>   Query the GPIB address. </summary>
''' <value>   An string including the primary and secondary address separated by a space. </value>
Public Function GpibAddressGetter() As String
    
    Const p_procedureName As String = "GpibAddressGetter"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    Dim p_length As Long
    
    Dim p_result As String: p_result = "-1"
    
    If Me.GpibLanControllerAttached Then
        
        p_length = Me.Socket.TrySendMessage(Me.GpibAddressCommand & This.Termination, _
            p_details, Me.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))

        If p_success Then
        
            Dim p_receivedMessage As String
            p_length = Me.Socket.TryReceive(p_receivedMessage, p_details)
            p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
            
        End If
        
        If p_success Then
        
            p_result = p_receivedMessage
            
            Dim p_delimiter As String: p_delimiter = " "
            Dim p_value As String
            
            p_value = cc_isr_Core.StringExtensions.Pop(p_receivedMessage, p_delimiter)
            If VBA.vbNullString = p_value Then
                This.PrimaryGpibAddress = -1
            Else
                If Not cc_isr_Core.StringExtensions.TryParseInteger(p_value, _
                    This.PrimaryGpibAddress, p_details) Then
                    This.PrimaryGpibAddress = -1
                End If
            End If
        
            p_value = cc_isr_Core.StringExtensions.Pop(p_receivedMessage, p_delimiter)
            If VBA.vbNullString = p_value Then
                This.SecondaryGpibAddress = -1
            Else
                If Not cc_isr_Core.StringExtensions.TryParseInteger(p_value, _
                    This.SecondaryGpibAddress, p_details) Then
                    This.SecondaryGpibAddress = -1
                End If
            End If
        
        End If
    
    End If
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
    This.GpibAddressReading = p_result
    GpibAddressGetter = p_result
    
End Function

''' <summary>   Configure the GPIB address of the GPIB Lan Controller.
''' The meaning of the GPIB address depends on the operating mode of the controller.
'''
''' In Controller mode, it refers to the GPIB address of the instrument being controlled.
''' In DEVICE mode, it is the address of the GPIB peripheral that the controller is emulating. </summary>
''' <remarks>
''' An optional secondary address may also be specified.
'''
''' Internally, the secondary address, which is offset by 96, must be separated from the primary
''' address by a space character. Specifying secondary address has no effect in DEVICE mode.
''' </remarks>
''' <para name="a_primary">     [Integer] Specifies the primary GPIB address
'''                             between 0 and 30. </param>
''' <para name="a_secondary">   [Optional, Integer] Specifies the second GPIB Address
'''                             between 0 and 30. </param>
''' <param name="a_delay">      [Optional, Long, 100] milliseconds duration to wait after
'''                             executing the command. </param>
''' <returns>   The status byte. </returns>
Public Sub GpibAddressSetter(ByVal a_primary As Integer, _
                             Optional ByVal a_secondary As Integer = -1, _
                             Optional ByVal a_delay As Long = 100)
    
    Const p_procedureName As String = "GpibAddressGetter"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
    
        If a_primary >= 0 And a_secondary < 0 Then
        
            a_primary = cc_isr_Core_IO.CoreExtensions.ClampInteger(a_primary, 0, 30)
            
            p_length = Me.Socket.TrySendMessage("++addr " & CStr(a_primary) & This.Termination, _
                p_details, a_delay)
                
            p_success = p_length > 0
            
        ElseIf a_primary >= 0 And a_secondary >= 0 Then
            
            a_primary = cc_isr_Core_IO.CoreExtensions.ClampInteger(a_primary, 0, 30)
            a_secondary = cc_isr_Core_IO.CoreExtensions.ClampInteger(a_secondary, 0, 30)
            
            p_length = Me.Socket.TrySendMessage("++addr " & CStr(a_primary) & _
                " " & CStr(a_secondary + 96) & This.Termination, _
                p_details, a_delay)
        
            p_success = p_length > 0
        
        End If
    
        If Not p_success Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
    End If
    
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  GPIB Lan Settings and query commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++auto
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <sumamry>   Automatically address instruments to talk after sending a command. </summary>
''' <remarks>
''' The GPIB-Lan controller can be configured to automatically address
''' instruments to talk after sending them a command in order to read their response. The
''' feature called, Read-After-Write, saves the user from having to issue read commands
''' repeatedly. This command enabled or disabled the Read-After-Write feature.
''' In addition, auto command also addresses the instrument at the currently specified
''' address to TALK or LISTEN. <c>++auto 0</c> addresses the instrument to LISTEN and
''' <c>++auto 1</c> addresses the instrument to TALK.
''' If the command is issued without any arguments it returns the current state of the
''' readafter-write feature.
''' Syntax: <c>++auto [0|1]</c>
''' Some instruments generate “Query Unterminated” or “-420” error if they are addressed
''' to talk after sending a command that does not generate a response (often called nonquery commands).
''' In effect the instrument is saying, I have been asked to talk but I have
''' nothing to say. The error is often benign and may be ignored. Otherwise, use the
''' <c>++read</c> command to read the instrument response.
''' </remarks>
''' <value>   [String]. </value>
Public Property Get AutoAssertTalkCommand() As String
    AutoAssertTalkCommand = "++auto"
End Property

''' <summary>   GPIB Lan Controller Auto Assert Talk. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoAssertTalk() As Boolean
    AutoAssertTalk = This.AutoAssertTalk
End Property

''' <summary>   Gets the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <returns>   [Boolean]. </returns>
Public Function AutoAssertTalkGetter() As Boolean

    Const p_procedureName As String = "AutoAssertTalkGetter"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long

    If Me.GpibLanControllerAttached Then
    
        p_length = Me.Socket.TrySendMessage(Me.AutoAssertTalkCommand & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        
        p_success = p_length > 0
        
        If p_success Then
        
            Dim p_result As String
            p_length = This.Socket.TryReceive(p_result, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
            
        End If
    
        If p_success Then
            p_success = Me.TryParseAutoAssertTalk(p_result, p_details)
        End If
        
    End If
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
    AutoAssertTalkGetter = This.AutoAssertTalk

End Function

''' <summary>   Tries to parse the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryParseAutoAssertTalk(ByVal a_value As String, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseAutoAssertTalk"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    If "1" = a_value Then
        This.AutoAssertTalk = True
        p_success = True
    ElseIf "0" = a_value Then
        This.AutoAssertTalk = False
        p_success = True
    Else
        p_success = False
        p_details = "The return value '" & a_value & "' is invalid for the '" & _
            Me.AutoAssertTalkCommand & "' command."
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryParseAutoAssertTalk = p_success
    
End Function

''' <summary>   Sets the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
Public Sub AutoAssertTalkSetter(ByVal a_value As Boolean)

    Const p_procedureName As String = "AutoAssertTalkSetter"
    
    If Me.GpibLanControllerAttached Then
    
        Dim p_details As String
        If Not Me.TryAutoAssertTalkSetter(a_value, p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
    End If

End Sub

''' <summary>   Tries to set the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryAutoAssertTalkSetter(ByVal a_value As Boolean, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryAutoAssertTalkSetter"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    If Me.GpibLanControllerAttached Then
    
        Dim p_command As String
        Dim p_value As String
        Dim p_result As String
        Dim p_sentCount As Long
        
        p_command = Me.AutoAssertTalkCommand
        p_value = IIf(a_value, "1", "0")
        
        ' send the command
        p_sentCount = Me.Socket.TrySendMessage(p_command & " " & p_value & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
                
        p_success = (p_sentCount > 0) And (0 = VBA.Len(p_details))
        
        ' validate the command
        If p_success Then _
            p_success = Me.TryGetValue(p_command, p_value, p_result, p_details)
        
        If p_success Then _
            p_success = Me.TryParseAutoAssertTalk(p_result, p_details)
        
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryAutoAssertTalkSetter = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++eos
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   The Append Termination command or <c>++eos</c>, specified the GPIB Lan Controller
''' option for appending a termination string to the sent message. </summary>
''' <remarks>  This command specifies GPIB termination characters. When data from host is received
''' over the network, all non-escaped LF, CR and ESC characters are removed and GPIB
''' terminators, as specified by this command, are appended before sending the data to
''' instruments.
'''
''' Although the manual indicates that "This command does not affect data from instruments
''' received over GPIB port.", we found that messages received from the instrument were
''' terminated with CR LF when EOS was set at 0 and LF when EOS was set at 3.
''' Syntax: <c>++eos [0|1|2|3]
''' </remarks>
''' <value>   [String]. </value>
Public Property Get AppendTerminationCommand() As String
    AppendTerminationCommand = "++eos"
End Property

''' <summary>   GPIB Lan Controller Termination options. </summary>
''' <value>   [<see cref="AppendTerminationOption"/>]. </value>
Public Property Get AppendTermination() As AppendTerminationOption
    AppendTermination = This.AppendTermination
End Property

''' <summary>   Gets the GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <returns>   [<see cref="AppendTerminationOption"/>]. </returns>
Public Function AppendTerminationGetter() As AppendTerminationOption

    Const p_procedureName As String = "AppendTerminationGetter"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long

    If Me.GpibLanControllerAttached Then
    
        p_length = Me.Socket.TrySendMessage(Me.AppendTerminationCommand & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        Dim p_result As String
        If p_success Then
            p_length = This.Socket.TryReceive(p_result, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        End If
        
        If p_success Then _
            p_success = Me.TryParseAppendTermination(p_result, p_details)
    
    End If
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".AppendTerminationGetter", p_details
    
    AppendTerminationGetter = This.AppendTermination

End Function

''' <summary>   Sets the GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_value">   [<see cref="AppendTerminationOption"/>] the option to set. </param>
Public Sub AppendTerminationSetter(ByVal a_value As AppendTerminationOption)

    Const p_procedureName As String = "AppendTerminationSetter"
    
    If Me.GpibLanControllerAttached Then
    
        Dim p_details As String
        If Not Me.TryAppendTerminationSetter(a_value, p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".AppendTerminationSetter", p_details
    
    End If
    
End Sub

''' <summary>   Tries to parse GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryParseAppendTermination(ByVal a_value As String, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseAppendTermination"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_value As Integer
    p_success = cc_isr_Core.StringExtensions.TryParseInteger(a_value, p_value, p_details)
    If p_success Then
        This.AppendTermination = p_value
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryParseAppendTermination = p_success
    
End Function

''' <summary>   Tries to set the GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_value">     [<see cref="AppendTerminationOption"/>] The value to set. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryAppendTerminationSetter(ByVal a_value As AppendTerminationOption, _
    ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryAppendTerminationSetter"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
    
        Dim p_command As String
        Dim p_value As String
        Dim p_result As String
        
        p_command = Me.AppendTerminationCommand
        p_value = VBA.CStr(a_value)
        
        ' send the command
        p_length = Me.Socket.TrySendMessage(p_command & " " & p_value & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        ' validate the command
        If p_success Then _
            p_success = Me.TryGetValue(p_command, p_value, p_result, p_details)
            
        If p_success Then _
            p_success = Me.TryParseAppendTermination(p_result, p_details)
    
    End If
   
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryAppendTerminationSetter = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++read_tmo_ms
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   The read timeout command or <c>read_tmo_ms</c> sets the timeout
''' of the <c>++read</c> and <c>++spoll</c> commands. </summary>
''' <remarks>
''' This command specifies the timeout value, in milliseconds, to be used in the read
''' command and spoll command. Timeout may be set to any value between 1 and 3000
''' milliseconds.
''' Syntax: <c>++read_tmo_ms {time}</c> where {time} is decimal value between 1 and 3000.
''' </remarks>
''' <value>   [String]. </value>
Public Property Get ReadTimeoutCommand() As String
    ReadTimeoutCommand = "++read_tmo_ms"
End Property

''' <summary>   GPIB Lan Controller Read Timeout. </summary>
''' <value>   [<see cref="Integer"/>]. </value>
Public Property Get ReadTimeout() As Integer
    ReadTimeout = This.ReadTimeout
End Property

''' <summary>   Gets the GPIB Lan Controller read timeout. </summary>
''' <returns>   [Integer]. </returns>
Public Function ReadTimeoutGetter() As Integer

    Const p_procedureName As String = "ReadTimeoutGetter"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long

    If Me.GpibLanControllerAttached Then
    
        Dim p_sentCount As Integer
        p_length = Me.Socket.TrySendMessage(Me.ReadTimeoutCommand & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        Dim p_result As String
        If p_success Then
            p_length = This.Socket.TryReceive(p_result, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        End If
        
        If p_success Then _
            p_success = Me.TryParseReadTimeout(p_result, p_details)
    
    End If
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".ReadTimeoutGetter", p_details
            
    ReadTimeoutGetter = This.ReadTimeout

End Function

''' <summary>   Sets the GPIB Lan Controller read timeout in milliseconds. </summary>
''' <param name="a_value">   [<see cref="Integer"/>] the read timeout in milliseconds to set. </param>
Public Sub ReadTimeoutSetter(ByVal a_value As Integer)

    Const p_procedureName As String = "ReadTimeoutSetter"

    If Me.GpibLanControllerAttached Then
    
        Dim p_details As String
        If Not Me.TryReadTimeoutSetter(a_value, p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".ReadTimeoutSetter", p_details
    
    End If
    
End Sub

''' <summary>   Tries to parse GPIB Lan Controller <see cref="Integer"/>. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryParseReadTimeout(ByVal a_value As String, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseReadTimeout"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_value As Integer
    p_success = cc_isr_Core.StringExtensions.TryParseInteger(a_value, p_value, p_details)
    If p_success Then _
        This.ReadTimeout = p_value
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryParseReadTimeout = p_success

End Function

''' <summary>   Tries to set the GPIB Lan Controller read timeout. </summary>
''' <param name="a_value">     [Integer] The value to set. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryReadTimeoutSetter(ByVal a_value As Integer, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryReadTimeoutSetter"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
    
        Dim p_command As String
        Dim p_value As String
        Dim p_result As String
        
        If a_value < This.MinimumTimeout Then
        
            p_details = "Read timeout '" & VBA.CStr(a_value) & _
                "' must be larger than '" & VBA.CStr(This.MinimumTimeout) & "'."
            p_success = False
        
        ElseIf a_value > This.ReadTimeoutMaximum Then
        
            p_details = "Read timeout '" & VBA.CStr(a_value) & _
                "' must be smaller than '" & VBA.CStr(This.ReadTimeoutMaximum) & "'."
            p_success = False
        
        Else
        
            p_command = Me.ReadTimeoutCommand
            p_value = VBA.CStr(a_value)
            
            ' send the command
            p_length = Me.Socket.TrySendMessage(p_command & " " & p_value & This.Termination, _
                    p_details, This.ReadAfterWriteDelay)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
            
            ' validate the command
            If p_success Then _
                p_success = Me.TryGetValue(p_command, p_value, p_result, p_details)
                
            If p_success Then _
                p_success = Me.TryParseReadTimeout(p_result, p_details)
                
        End If
        
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryReadTimeoutSetter = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++eoi
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <sumamry>   This command enables or disables the assertion of the EOI signal with the
''' last character of any command sent over GPIB port. </summary>
''' <remarks>   Some instruments require EOI signal to be asserted in order to properly
''' detect the end of a command. Syntax: <c>++eoi [0|1]</c>. </remarks>
Public Property Get EndOrIdentifyCommand() As String
    EndOrIdentifyCommand = "++eoi"
End Property

''' <summary>   GPIB Lan Controller End Or Identify (EOI). </summary>
''' <value>   [Boolean]. </value>
Public Property Get EndOrIdentify() As Boolean
    EndOrIdentify = This.EndOrIdentify
End Property

''' <summary>   Gets the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <returns>   [Boolean]. </returns>
Public Function EndOrIdentifyGetter() As Boolean

    Const p_procedureName As String = "EndOrIdentifyGetter"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
    
        p_length = Me.Socket.TrySendMessage(Me.EndOrIdentifyCommand & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
         
        Dim p_result As String
        If p_success Then
            p_length = This.Socket.TryReceive(p_result, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        End If
        
        If p_success Then _
            p_success = Me.TryParseEndOrIdentify(p_result, p_details)
    
    End If
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".EndOrIdentifyGetter", p_details
    
    EndOrIdentifyGetter = This.EndOrIdentify

End Function

''' <summary>   Tries to parse the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryParseEndOrIdentify(ByVal a_value As String, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseEndOrIdentify"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    If "1" = a_value Then
        This.EndOrIdentify = True
        p_success = True
    ElseIf "0" = a_value Then
        This.EndOrIdentify = False
        p_success = True
    Else
        p_success = False
        p_details = "The return value '" & a_value & "' is invalid for the '" & _
            Me.EndOrIdentifyCommand & "' command."
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryParseEndOrIdentify = p_success
        
End Function

''' <summary>   Sets the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
Public Sub EndOrIdentifySetter(ByVal a_value As Boolean)

    Const p_procedureName As String = "EndOrIdentifySetter"
    
    If Me.GpibLanControllerAttached Then
    
        Dim p_details As String
        If Not Me.TryEndOrIdentifySetter(a_value, p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".EndOrIdentifySetter", p_details
    
    End If
    
End Sub

''' <summary>   Tries to set the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryEndOrIdentifySetter(ByVal a_value As Boolean, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryEndOrIdentifySetter"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
    
        Dim p_command As String
        Dim p_value As String
        Dim p_result As String
        
        p_command = Me.EndOrIdentifyCommand
        p_value = IIf(a_value, "1", "0")
        
        ' send the command
        p_length = Me.Socket.TrySendMessage(p_command & " " & p_value & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        ' validate the command
        If p_success Then _
            p_success = Me.TryGetValue(p_command, p_value, p_result, p_details)
            
        If p_success Then _
            p_success = Me.TryParseEndOrIdentify(p_result, p_details)
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryEndOrIdentifySetter = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++spoll
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Queries the GPIB SRQ signal status. </summary>
''' <remarks>   This command returns the current state of the GPIB SRQ signal. The controller returns
''' <c>1</c> if the SRQ signal is asserted (low) and <c>0</c> if the signal is not asserted (high). </remarks>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Integer] <c>1</c> if the SRQ signal is asserted (low),
''' <c>0</c> if the signal is not asserted (high), or -1 if the query failed. </returns>
Public Function TryQueryServiceRequested(ByRef a_details As String) As Integer
    
    Const p_procedureName As String = "ServiceRequested"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    Dim p_result As Integer
    Dim p_affirmativeReply As String: p_affirmativeReply = "1"
    Dim p_negativeReply As String: p_negativeReply = "0"
    
    If Me.GpibLanControllerAttached Then

        Dim p_reply As String
        p_length = This.Socket.TryQuery("++srq" & This.Termination, p_reply, p_details)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        If p_success Then
            If 0 = VBA.StrComp(p_reply, p_affirmativeReply, VBA.VbCompareMethod.vbTextCompare) Then
                p_result = 1
            ElseIf 0 = VBA.StrComp(p_reply, p_negativeReply, VBA.VbCompareMethod.vbTextCompare) Then
                p_result = 0
            Else
                p_result = -1
                p_details = "The query reply '" & p_reply & _
                    "' matched neither the affirmative '" & p_affirmativeReply & _
                    "' nor the negative '" & p_negativeReply & "' values."
            End If
        Else
            p_result = -1
        End If
        
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
End Function

''' <summary>   Performs a serial poll of the instrument at the specified address. </summary>
''' <remarks>
''' This command performs a serial poll of the instrument at the specified address. If no
''' address is specified then this command serial polls the currently addressed instrument (as
''' set by a previous <c>++addr</c> command). This command uses the time-out value specified
''' by the read_tmo_ms command.
''' Syntax: <c>++spoll [{primary address} {secondary address}}</c> where
''' Primary Address is a decimal value between 0 and 30.
''' Secondary Address is a decimal value between 96 and 126.
''' </remarks>
''' <value>   [String]. </value>
Public Property Get SerialPollCommand() As String
    SerialPollCommand = "++spoll"
End Property

''' <summary>   Reads the status byte of the current GPIB instrument. </summary>
''' <exception cref="cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError"> Thrown
''' when a socket send or receive error condition occurs. </exception>
''' <returns>   [Integer] The status byte or RECEIVE_ERROR (-1) if failed receiving a reading or
''' failed parsing the reading to an integer.
''' </returns>
Public Function SerialPoll() As Integer
    
    Const p_procedureName As String = "SerialPoll"
    
    Dim p_statusByte As Integer: p_statusByte = 0
    
    If Me.GpibLanControllerAttached Then
    
        Dim p_details As String
        p_statusByte = Me.TrySerialPoll(p_details)
        If p_statusByte < 0 Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
                ThisWorkbook.VBProject & "." & TypeName(Me) & ".SerialPoll", _
                " Failed reading '" & Me.SerialPollCommand & "'; " & p_details
    End If
    
    SerialPoll = p_statusByte

End Function

''' <summary>   Tries to read the status byte of the current GPIB instrument. </summary>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Integer] The status byte or RECEIVE_ERROR (-1) if failed receiving a reading or
''' failed parsing the reading to an integer.
''' </returns>
Public Function TrySerialPoll(ByRef a_details As String) As Integer
    
    Const p_procedureName As String = "TrySerialPoll"
    Dim p_statusByte As Integer: p_statusByte = 0
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached Then
    
        p_length = Me.Socket.TrySendMessage(Me.SerialPollCommand & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        Dim p_reply As String
        If p_success Then
        
            p_length = This.Socket.TryReceive(p_reply, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))

            If p_success Then _
                p_success = cc_isr_Core.StringExtensions.TryParseInteger(p_reply, p_statusByte, p_details)
        
        End If
    
    End If
    
    If Not p_success Then
        p_statusByte = -1
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    End If
    
    TrySerialPoll = p_statusByte

End Function

''' <summary>   Waits for the expected status bits or timeout. </summary>
''' <param name="a_bitsStatus"/>   [Integer] The expected status of the specified status bits. </param>
''' <param name="a_statusBits"/>   [Integer] The expected status bits. </param>
''' <param name="a_timeout"/>      [Long] The timeout in milliseconds. </param>
''' <param name="a_statusByte"/>   [Out, Integer] The last received status byte.
'''                                Negative if serial poll failed. </param>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if the status byte has the expected bits value. </returns>
Public Function AwaitStatusBits(ByVal a_bitsStatus As Integer, ByVal a_statusBits As Integer, _
    ByVal a_timeout As Long, ByRef a_statusByte As Integer, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "AwaitStatusBits"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_gotIt As Boolean

    If Me.GpibLanControllerAttached Then
    
        Dim p_stopper As cc_isr_Core_IO.Stopwatch
        Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
        p_stopper.Restart
        
        Do
            DoEvents
            ' read the status bit.
            a_statusByte = Me.TrySerialPoll(p_details)
            p_gotIt = (a_statusByte >= 0) And (a_bitsStatus = (a_statusBits And a_statusByte))
        Loop Until p_gotIt Or (a_statusByte < 0) Or (p_stopper.ElapsedMilliseconds > a_timeout)
        p_success = a_statusByte >= 0
        
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    AwaitStatusBits = p_gotIt
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Restoring state
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns true and a message if the instrument is nothing or not conneced. </summary>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            should connect. </param>
''' <returns>   [Boolean]. </returns>
Public Function ShouldConnect(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "ShouldConnect"

    If This.Socket Is Nothing Then
        a_details = "The Tcp Session socket instance is nothing."
        ShouldConnect = True
    ElseIf Not Me.Connected Then
        If This.Socket.Connected Then
            a_details = "A connection with the instrument at '" & This.SocketAddress & _
                "' was not established."
        Else
            a_details = "The session socket is not connected to '" & This.SocketAddress & "'."
        End If
        ShouldConnect = True
    Else
        ShouldConnect = False
    End If
    
End Function

''' <summary>   Checks if the GPIB-Lan device needs to be restored to its Initial State, i.e.,
''' or if <see cref="AutoAssertTalk"/> is enabeld. </summary>
''' <remarks>   Note that the GPIB-Lan Auto assert to talk mnust be disabled to allow
''' communications with the 2700 instrument or monitoring of the status
''' byte using serial poll. </remarks>.
''' <param name="a_details">    [Out, String] Details why the instrument state
'''                             should be restored. </param>
''' <returns>   [Boolean]. </returns>
Public Function ShouldRestoreInitialState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "ShouldRestoreInitialState"

    Dim p_details As String: p_details = VBA.vbNullString
    If Me.ShouldConnect(p_details) Then
        
        ShouldRestoreInitialState = True
    
    Else
    
        Dim p_autoAssertTalk As Boolean
        p_autoAssertTalk = Me.AutoAssertTalkGetter()
        DoEvents
        If p_autoAssertTalk Then
            a_details = "GPIB Lan Read After Write is enabled. It should be disabled."
            ShouldRestoreInitialState = True
        Else
            Dim p_timeout As Integer
            p_timeout = Me.ReadTimeoutGetter()
            If p_timeout <> This.SocketTimeout Then
                a_details = "GPIB Lan read timeout '" & VBA.CStr(p_timeout) & _
                    "' is not equal to the socket timeout '& VBA.CStr(This.SocketTimeout) & " '."
                ShouldRestoreInitialState = True
            Else
                ShouldRestoreInitialState = False
            End If
        End If
    
    End If

End Function

''' <summary>   Clears the instrument output and then clears the instrument execution state. </summary>
''' <remarks>   This method catches and hides any execution errors. </remarks>
''' <param name="a_command">   [Optional, String, <c>*CLS;*WAI;*OPC?</c>] the command. </param>
''' <param name="a_timeout">   [Optional, Integer, 100] the timeout. </param>
''' <retursn>   [String. </returns>
Public Function QueryClearExecutionState(Optional ByVal a_command As String = "*CLS;*WAI;*OPC?", _
    Optional ByVal a_timeout As Integer = 100) As String
        
    Const p_procedureName As String = "QueryClearExecutionState"
        
    On Error Resume Next
    
    Dim p_details As String
    Dim p_reply As String
    Dim p_success As Boolean
    Dim p_length As Long
    
    ' and then clear the execution state.
    p_length = Me.TryQueryLine(a_command, p_reply, p_details)
    DoEvents
    p_length = Me.TryQueryLine(a_command, p_reply, p_details)
    DoEvents
    p_length = Me.TryQueryLine(a_command, p_reply, p_details)
    DoEvents
    
    On Error GoTo 0

    QueryClearExecutionState = p_reply

End Function

''' <summary>   Configures the GPIB-Lan controller initial state. </summary>
''' <remarks>   Call this method upon connecting and restoring Initial State. </remarks>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
Public Function TryConfigureInitialState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryConfigureInitialState"

    Dim p_success As Boolean:     p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If p_success Then
    
        Dim p_timeout As Integer
        p_timeout = VBA.CInt(cc_isr_Core_IO.CoreExtensions.Clamp(This.SocketTimeout, _
            This.ReadTimeoutMinimum, This.ReadTimeoutMaximum))
        p_success = Me.TrySetValue(Me.ReadTimeoutCommand, _
            VBA.CStr(p_timeout), _
            p_details)
        If p_success Then This.ReadTimeout = p_timeout
    
    End If
    
    If p_success Then
    
        ' set the GPIB termination characters to none - do not append termination characters.
        p_success = Me.TrySetValue(Me.AppendTerminationCommand, _
            VBA.CStr(AppendTerminationOption.AppendNothing), _
            p_details)
        If p_success Then This.AppendTermination = AppendTerminationOption.AppendNothing
    
    End If
    
    If p_success Then
    
        ' uturn on senting end or identity at the end of each message to the instrument.
        p_success = Me.TrySetValue(Me.EndOrIdentifyCommand, "1", p_details)
        If p_success Then This.EndOrIdentify = True
    
    End If
    
    If p_success Then
    
        ' turn off auto assert talk
        p_success = Me.TrySetValue(Me.AutoAssertTalkCommand, "0", p_details)
        If p_success Then This.AutoAssertTalk = False
    
    End If
    
    If p_success Then
    
        ' disable front panel operation of the currently addressed instrument.
        p_length = Me.Socket.TrySendMessage(Me.LocalLockoutCommand & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
   
    End If
    
    If p_success Then
    
        p_success = Me.TryEstablishConnection(False, p_details)
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryConfigureInitialState = p_success


End Function

''' <summary>   Tries to establish a connection with the instrument. </summary>
''' <remarks>   Call this method upon connecting and restoring Initial State.
''' This method might be called from parent classes if the connection failed to establish
''' where instrument reset is required. </remarks>
''' <param name="a_established">   [Boolean] Used when called from a parent class to indicate that
'''                                the patent class was able to establish a connection. </param>
''' <param name="a_details">       [Out, String] Details why the method failed. </param>
Public Function TryEstablishConnection(ByVal a_established As Boolean, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryEstablishConnection"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If a_established Then
    
        This.Established = a_established
        
    Else
    
        ' establish a connection to the actual instrument because a connection
        ' to the GPIB-Lan controller is insuffcient to ascertain that the instrument
        ' is responsive.
        Dim p_reply As String
        p_length = Me.TryQueryLine(This.ConnectionQueryMessage, p_reply, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
        
        If p_success Then
            p_success = 0 = VBA.StrComp(p_reply, This.ConnectionQueryReply, vbTextCompare)
            This.Established = p_success
            If Not p_success Then _
                p_details = "Failed validating connection query reply; received '" & p_reply & _
                    "' instead of '" & This.ConnectionQueryReply & _
                    "' in reply to '" & This.ConnectionQueryMessage & "'."
        Else
            p_details = "Failed validating connection; & p_details"
        End If
                
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryEstablishConnection = p_success

End Function

''' <summary>   Configures the GPIB-Lan controller final state. </summary>
''' <remarks>   Call this method upon disconnecting. </remarks>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
Public Function TryConfigureFinalState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryConfigureFinalState"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    If p_success Then
    
        ' turn off auto assert talk
        p_success = Me.TrySetValue(Me.AutoAssertTalkCommand, "0", p_details)
        If p_success Then This.AutoAssertTalk = False
    
    End If
    
    If p_success Then
    
        ' enable front panel operation of the currently addressed instrument.
        p_length = Me.Socket.TrySendMessage(Me.GoToLocalCommand & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
   
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
   
    
    TryConfigureFinalState = p_success

End Function

''' <summary>   Restores the  GPIB-Lan controller to its Initial State. </summary>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean] true if restored. </returns>
Public Function TryRestoreInitialState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryRestoreInitialState"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    If p_success And VBA.vbNullString = This.SocketAddress Then
        p_details = "Unable to connect because the socket address is empty."
        p_success = False
    End If
    
    If p_success And This.MinimumTimeout > This.SocketTimeout Then
        p_details = "Unable to connect because the socket timeout '" & VBA.CStr(This.SocketTimeout) & _
        " ms is below the minimum of " & VBA.CStr(This.MinimumTimeout) & " ms."
        p_success = False
    End If
    
    If p_success And Not (Me.Connected Or This.Socket.CanConnect) Then
        
        ' if not connected and not 'can connect' it means that the
        ' socket was disposed and needs to be rebuilt.
        Me.Initialize cc_isr_Winsock.Factory.NewIPv4StreamSocket
            
        p_success = This.Socket.CanConnect
        
        If Not p_success Then
            p_details = "Failed initializing the socket; the socket id is invalid ('" & _
                VBA.CStr(cc_isr_Winsock.ws32_INVALID_SOCKET) & "')."
        End If
    
    End If
    
    If p_success Then
    
        If Not Me.Connected Then
        
            ' the socket address and timeout are saved from the previous socket connection
            p_success = This.Socket.TryOpenConnection(This.SocketAddress, This.SocketTimeout, p_details)
        
        Else
        
            ' configure the initial state
            p_success = Me.TryConfigureInitialState(p_details)
        
        End If
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    
    TryRestoreInitialState = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Socket Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Winsock.IPv4SreamSocket.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_socket_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Const p_procedureName As String = "m_Socket_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_eventArgs.Connected Then
    
        ' save the socket connection information
        This.SocketAddress = This.Socket.Address
        This.SocketTimeout = This.Socket.ReceiveTimeout
         
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

''' <summary>   Handles the <see cref="cc_isr_Winsock.IPv4SreamSocket.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_socket_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

    Const p_procedureName As String = "m_Socket_ConnectionChanging"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_eventArgs.Connected Then
    
        Dim p_details As String
        
        ' configure the final state
        If Not Me.TryConfigureFinalState(p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.EnqueueError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Failed configuring the final state; " & p_details
   
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



