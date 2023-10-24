VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GpibLanController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Gpib Lan Controller extensions. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    
    ' socket settings
    SocketTimeout As Long
    
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

''' <summary>   GPIB Lan Controller Termination options. </summary>
Public Enum AppendTerminationOption

    ''' <summary>   Enum option for appending carriage return and line feed. </summary>
    CarriageReturnLineFeed = 0
    ''' <summary>   Enum option for appending carriage return. </summary>
    CarriageReturn = 1
    ''' <summary>   Enum option for appending line feed. </summary>
    LineFeed = 2
    ''' <summary>   Enum option for appending nothing. </summary>
    AppendNothing = 3

End Enum

''' <summary>   Handles the Class Initialize event. </summary>
Private Sub Class_Initialize()

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

''' <summary>   Handles the Class terminate event. </summary>
Private Sub Class_Terminate()
End Sub

''' <summary>   Returns this instance. No actual initialization is done
''' at this time. </summary>
''' <returns>   [<see cref="TcpSession"/>]. </returns>
Public Function Initialize() As GpibLanController
    
    Set Initialize = Me
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Socket Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

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
'  Instrument send and receive methods.
'  use socket send and receive to communicate with the controller.
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sends an unterminated message to the instrument. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_message">   [String] The unterminated message. </param>
''' <returns>   [Long] The number of bytes that were sent to the server. </returns>
Public Function WriteLine(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_message As String) As Long
        
    Const p_procedureName As String = "WriteLine"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Integer
    
    p_length = Me.TryWriteLine(a_socket, a_message, p_details)
    p_success = (p_length > 0) And (0 = VBA.Len(p_details))
    
    If Not p_success Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    End If
    
    WriteLine = p_length
    
End Function

''' <summary>   Sends an unterminated message to the instrument. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_message">   [String] The unterminated message. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Long]. 0 if nothing to send, negative (-1) if error,
''' positive if the entire message was sent. </returns>
Public Function TryWriteLine(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_message As String, ByRef a_details As String) As Long

    Const p_procedureName As String = "TryWriteLine"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    p_length = a_socket.TrySendMessage(a_message & This.Termination, _
        p_details, This.ReadAfterWriteDelay)
    
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
''' <param name="a_socket">        [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [String] The received message. </returns>
Public Function Read(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    Optional ByVal a_trimEnd As Boolean = True, _
    Optional ByVal a_termination As String = vbLf) As String

    Const p_procedureName As String = "Read"

    Dim p_buffer As String
    Dim p_details As String
    
    Dim p_length As Long
    p_length = Me.TryRead(a_socket, p_buffer, p_details, a_trimEnd, a_termination)
    
    If 0 > p_length Or 0 < VBA.Len(p_details) Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Function

''' <summary>   Tries to receive a message from the instrument until reaching the specified termination. </summary>
''' <remarks>   The <see cref="ReceiveTimeout"/> must be set long enough the exceed any instrument
''' or controller delays. </remarks>
''' <param name="a_socket">        [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [Long] If no error, the number of bytes that were received.
''' Otherwise, RECEIVE_ERROR (-1).
''' </returns>
Public Function TryRead(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_buffer As String, _
    ByRef a_details As String, _
    Optional ByVal a_trimEnd As Boolean = True, _
    Optional ByVal a_termination As String = vbLf) As Long
                           
    Const p_procedureName As String = "TryRead"
                           
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
                           
    If a_socket.Port = This.GpibLanControllerPort Then
        p_length = a_socket.TrySendMessage(Me.ReadEOICommand & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    End If
    
    If p_success Then
        p_length = a_socket.TryReceive(a_buffer, p_details, a_trimEnd, a_termination)
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
''' <param name="a_socket">        [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_message">       [String] The unterminated query message. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [String] The received message. </returns>
Public Function QueryLine(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_message As String, _
    Optional ByVal a_trimEnd As Boolean = True, _
    Optional ByVal a_termination As String = vbLf) As String

    Const p_procedureName As String = "QueryLine"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_buffer As String: p_buffer = VBA.vbNullString
    Dim p_length As Long
    p_length = Me.TryQueryLine(a_socket, a_message, p_buffer, p_details, a_trimEnd, a_termination)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
    QueryLine = p_buffer
    
End Function

''' <summary>   Sends an unterminated query message and reads back from the
''' server until reaching the specified
''' termination. </summary>
''' <param name="a_socket">        [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_message">       [String] The unterminated query message. </param>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [Long] If no error, the number of bytes that were received.
''' Otherwise, RECEIVE_ERROR (-1).
''' </returns>
Public Function TryQueryLine(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_message As String, _
    ByRef a_buffer As String, _
    ByRef a_details As String, _
    Optional ByVal a_trimEnd As Boolean = True, _
    Optional ByVal a_termination As String = VBA.vbLf) As Long
                         
    Const p_procedureName As String = "TryQueryLine"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long: p_length = 0
    
    p_length = Me.TryWriteLine(a_socket, a_message, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If Not p_success Then
        a_buffer = VBA.vbNullString
    End If
    
    If p_success Then
        p_length = Me.TryRead(a_socket, a_buffer, p_details, a_trimEnd, a_termination)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryQueryLine = p_length
                         
End Function

''' <summary>   Sends an operation complete query command and waits for a reply. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TrySynchronize(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TrySynchronize"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    Dim p_message As String: p_message = This.SynchronizationQueryMessage
    Dim p_buffer As String: p_buffer = VBA.vbNullString
    Dim p_length As Long
    
    p_length = Me.TryQueryLine(a_socket, p_message, p_buffer, p_details)
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
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <value>   [Boolean] true if a GPIB Lan is attached to this LAN connection. </value>
Public Property Get GpibLanControllerAttached(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket) As Boolean
    If a_socket Is Nothing Then
        GpibLanControllerAttached = False
    Else
        GpibLanControllerAttached = a_socket.Port = This.GpibLanControllerPort
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  GPIB Lan Controller Methods
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++rst - power on reset
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
''' <returns>   [Boolean] True if success. </returns>
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
'  ++loc: Go to local
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   This command enables front panel operation of the currently addressed instrument. </summary>
''' <remarks>   Syntax: <c>++loc</c>. </remarks>
Public Property Get GoToLocalCommand() As String
    GoToLocalCommand = "++loc"
End Property

''' <summary>   Issues a Go To Local (GTL) (++loc). </summary>
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_delay">    [Optional, Long, 100] milliseconds duration to wait after
'''                           executing the command. </param>
Public Sub GoToLocal(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    Optional ByVal a_delay As Long = 100)
    
    Const p_procedureName = "GoToLocal"
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    p_success = Me.TryGoToLocal(a_socket, p_details, a_delay)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Sub

''' <summary>   Issues a Go To Local (GTL) (++loc). </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <param name="a_delay">     [Optional, Long, 100] milliseconds duration to wait after
'''                            executing the command. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryGoToLocal(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String, Optional ByVal a_delay As Long = 100) As Boolean
    
    Const p_procedureName = "TryGoToLocal"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached(a_socket) Then
        p_length = a_socket.TrySendMessage(Me.GoToLocalCommand & This.Termination, p_details, a_delay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
        
    TryGoToLocal = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++llo - local lockout
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   This command disables front panel operation of the currently addressed instrument. </summary>
''' <remarks>   Syntax: <c>++llo</c>. </remarks>
Public Property Get LocalLockoutCommand() As String
    LocalLockoutCommand = "++llo"
End Property

''' <summary>   Issues a local lockout (++llo). </summary>
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <param name="a_delay">    [Optional, Long, 100] milliseconds duration to wait after
'''                           executing the command. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Sub LocalLockout(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    Optional ByVal a_delay As Long = 100)
    
    Const p_procedureName = "LocalLockout"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    p_success = Me.TryLocalLockout(a_socket, p_details, a_delay)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Sub

''' <summary>   Issues a local lockout (++llo). </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <param name="a_delay">     [Optional, Long, 100] milliseconds duration to wait after
'''                            executing the command. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryLocalLockout(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String, Optional ByVal a_delay As Long = 100) As Boolean
    
    Const p_procedureName = "TryLocalLockout"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached(a_socket) Then
        p_length = a_socket.TrySendMessage(Me.LocalLockoutCommand & This.Termination, p_details, a_delay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryLocalLockout = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++rst - selective device clear
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   This command issues a selective device clear on the currently addressed instrument. </summary>
''' <remarks>   Syntax: <c>++clr</c>. </remarks>
Public Property Get SelectiveDeviceClearCommand() As String
    SelectiveDeviceClearCommand = "++clr"
End Property

''' <summary>   Issues an SDC (selective device clear). </summary>
''' <remarks>   On the 2700 synchronization failed if the instrument was in
''' external triggering mode.
''' This can be avoided by issueing the <c>*RST</c> command first. </remakrs>
''' <param name="a_socket">        [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_synchronize">   [Optional, Boolean, False] True to query operation completion
'''                                after executing the command. </param>
''' <param name="a_delay">         [Optional, Long, 10] milliseconds duration to wait after
'''                                executing the command. </param>
Public Sub SelectiveDeviceClear(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    Optional ByVal a_synchronize As Boolean = False, _
    Optional ByVal a_delay As Long = 100)
    
    Const p_procedureName = "SelectiveDeviceClear"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    p_success = Me.TrySelectiveDeviceClear(a_socket, p_details, a_synchronize, a_delay)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Sub


''' <summary>   Issues an SDC (selective device clear). </summary>
''' <remarks>   On the 2700 synchronization failed if the instrument was in
''' external triggering mode.
''' This can be avoided by issueing the <c>*RST</c> command first. </remakrs>
''' <param name="a_socket">        [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">       [Out, String] Details why the method failed. </param>
''' <param name="a_synchronize">   [Optional, Boolean, False] True to query operation completion
'''                                after executing the command. </param>
''' <param name="a_delay">         [Optional, Long, 10] milliseconds duration to wait after
'''                                executing the command. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TrySelectiveDeviceClear(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String, _
    Optional ByVal a_synchronize As Boolean = False, _
    Optional ByVal a_delay As Long = 10)
    
    Const p_procedureName As String = "TrySelectiveDeviceClear"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached(a_socket) Then
        
        p_length = a_socket.TrySendMessage(Me.SelectiveDeviceClearCommand & This.Termination, _
            p_details, a_delay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
        
        If Not p_success Then _
            p_details = "Failed sending the '" & Me.SelectiveDeviceClearCommand & "'; " & p_details
        
        If p_success And a_synchronize Then _
            p_success = Me.TrySynchronize(a_socket, p_details)
        
        If Not p_success Then _
            p_details = "Failed trying to synchronize following a selective device clear; " & p_details
            
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TrySelectiveDeviceClear = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++addr - Gpib address
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

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

''' <summary>   Query the GPIB address. </summary>
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <value>   [String] including the primary and secondary address separated by a space. </value>
Public Function GpibAddressGetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_primary As Integer, ByRef a_secondary As Integer) As String
    
    Const p_procedureName As String = "GpibAddressGetter"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    Dim p_reading As String
    p_success = Me.TryGetGpibAddress(a_socket, p_reading, a_primary, a_secondary, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    GpibAddressGetter = p_reading
    
End Function

''' <summary>   Query the GPIB address. </summary>
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_reading">   [Out, String] the reading including the primary and secondary
''' addresses separated by a space. </param>
''' <param name="a_primary">   [Out, integer] the primary address. </param>
''' <param name="a_primary">   [Out, integer] the secondart address. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryGetGpibAddress(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_reading As String, ByRef a_primary As Integer, ByRef a_secondary As Integer, _
    ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryGetGpibAddress"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    Dim p_length As Long
    
    Dim p_reading As String: p_reading = "-1"
    
    If Me.GpibLanControllerAttached(a_socket) Then
        
        p_length = a_socket.TrySendMessage(Me.GpibAddressCommand & This.Termination, _
            p_details, Me.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))

        If p_success Then
        
            p_length = a_socket.TryReceive(p_reading, p_details)
            p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
            
        End If
        
        If p_success Then
        
            a_reading = p_reading
        
            Dim p_delimiter As String: p_delimiter = " "
            Dim p_value As String
            
            p_value = cc_isr_Core.StringExtensions.Pop(p_reading, p_delimiter)
            If VBA.vbNullString = p_value Then
                a_primary = -1
            Else
                If cc_isr_Core.StringExtensions.TryParseInteger(p_value, _
                    a_primary, p_details) Then
                Else
                    a_primary = -1
                End If
            End If
        
            p_value = cc_isr_Core.StringExtensions.Pop(p_reading, p_delimiter)
            If VBA.vbNullString = p_value Then
                a_secondary = -1
            Else
                If Not cc_isr_Core.StringExtensions.TryParseInteger(p_value, _
                    a_secondary, p_details) Then
                    a_secondary = -1
                End If
            End If
        
            p_success = a_primary >= 0
            
        End If
    
        If p_success Then
            This.PrimaryGpibAddress = a_primary
            This.GpibAddressReading = a_reading
            This.SecondaryGpibAddress = a_secondary
        Else
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                "; " & p_details
        End If
    
    End If
    
    TryGetGpibAddress = p_success
    
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
''' <param name="a_socket">     [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <para name="a_primary">     [Integer] Specifies the primary GPIB address
'''                             between 0 and 30. </param>
''' <para name="a_secondary">   [Optional, Integer] Specifies the second GPIB Address
'''                             between 0 and 30. </param>
''' <param name="a_delay">      [Optional, Long, 100] milliseconds duration to wait after
'''                             executing the command. </param>
Public Sub GpibAddressSetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_primary As Integer, _
    Optional ByVal a_secondary As Integer = -1, _
    Optional ByVal a_delay As Long = 100)
    
    Const p_procedureName As String = "GpibAddressGetter"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    Dim p_length As Long
    
    p_success = Me.TrySetGpibAddress(a_socket, a_primary, p_details, a_secondary, a_delay)
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Sub

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
''' <param name="a_socket">     [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <para name="a_primary">     [Integer] Specifies the primary GPIB address
'''                             between 0 and 30. </param>
''' <param name="a_details">    [Out, String] details the failure. </param>
''' <para name="a_secondary">   [Optional, Integer] Specifies the second GPIB Address
'''                             between 0 and 30. </param>
''' <param name="a_delay">      [Optional, Long, 100] milliseconds duration to wait after
'''                             executing the command. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TrySetGpibAddress(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_primary As Integer, _
    ByRef a_details As String, _
    Optional ByVal a_secondary As Integer = -1, _
    Optional ByVal a_delay As Long = 100) As Boolean
    
    Const p_procedureName As String = "TrySetGpibAddress"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached(a_socket) Then
    
        If a_primary >= 0 And a_secondary < 0 Then
        
            a_primary = cc_isr_Core_IO.CoreExtensions.ClampInteger(a_primary, 0, 30)
            
            p_length = a_socket.TrySendMessage("++addr " & CStr(a_primary) & This.Termination, _
                p_details, a_delay)
                
            p_success = p_length > 0
            
        ElseIf a_primary >= 0 And a_secondary >= 0 Then
            
            a_primary = cc_isr_Core_IO.CoreExtensions.ClampInteger(a_primary, 0, 30)
            a_secondary = cc_isr_Core_IO.CoreExtensions.ClampInteger(a_secondary, 0, 30)
            
            p_length = a_socket.TrySendMessage("++addr " & CStr(a_primary) & _
                " " & CStr(a_secondary + 96) & This.Termination, _
                p_details, a_delay)
        
            p_success = p_length > 0
        
        End If
    
        If p_success Then
            This.PrimaryGpibAddress = a_primary
            This.SecondaryGpibAddress = a_secondary
            
        Else
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                "; " & p_details
        End If
    
    End If
    
    TrySetGpibAddress = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  GPIB Lan configuration commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++auto - auto address to talk
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
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <returns>   [Boolean] True if auto assert talk is enabled. </returns>
Public Function AutoAssertTalkGetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket) As Boolean

    Const p_procedureName As String = "AutoAssertTalkGetter"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_autoAssertTalk As Boolean
    p_success = Me.TryGetAutoAssertTalk(a_socket, p_autoAssertTalk, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
    AutoAssertTalkGetter = p_autoAssertTalk

End Function

''' <summary>   Gets the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">     [Out, Boolean] True to Auto Assert Talk. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryGetAutoAssertTalk(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_value As Boolean, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryGetAutoAssertTalk"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    Dim p_autoAssertTalk As Boolean

    If Me.GpibLanControllerAttached(a_socket) Then
    
        p_length = a_socket.TrySendMessage(Me.AutoAssertTalkCommand & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        
        p_success = p_length > 0
        
        If p_success Then
        
            Dim p_reading As String
            p_length = a_socket.TryReceive(p_reading, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
            
        End If
    
        If p_success Then _
            p_success = Me.TryParseAutoAssertTalk(p_reading, p_autoAssertTalk, p_details)
        
        If p_success Then _
            a_value = p_autoAssertTalk
        
        If p_success Then _
            This.AutoAssertTalk = p_autoAssertTalk
        
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryGetAutoAssertTalk = p_success

End Function

''' <summary>   Tries to parse the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryParseAutoAssertTalk(ByVal a_value As String, _
    ByRef a_autoAssertTalk As Boolean, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseAutoAssertTalk"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    If "1" = a_value Then
        a_autoAssertTalk = True
        p_success = True
    ElseIf "0" = a_value Then
        a_autoAssertTalk = False
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
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
Public Sub AutoAssertTalkSetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_value As Boolean)

    Const p_procedureName As String = "AutoAssertTalkSetter"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    p_success = Me.TryAutoAssertTalkSetter(a_socket, a_value, p_details)
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details

End Sub

''' <summary>   Tries to set the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryAutoAssertTalkSetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_value As Boolean, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryAutoAssertTalkSetter"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    If Me.GpibLanControllerAttached(a_socket) Then
    
        Dim p_command As String
        Dim p_value As String
        Dim p_reading As String
        Dim p_sentCount As Long
        
        p_command = Me.AutoAssertTalkCommand
        p_value = IIf(a_value, "1", "0")
        
        ' send the command
        p_sentCount = a_socket.TrySendMessage(p_command & " " & p_value & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
                
        p_success = (p_sentCount > 0) And (0 = VBA.Len(p_details))
        
        ' validate the command
        If p_success Then _
            p_success = Me.TryGetValue(a_socket, p_command, p_value, p_reading, p_details)
        
        Dim p_autoAssertTalk As Boolean
        If p_success Then _
            p_success = Me.TryParseAutoAssertTalk(p_reading, p_autoAssertTalk, p_details)
        
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryAutoAssertTalkSetter = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++eos - Append termination
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
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
Public Function AppendTerminationGetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket) As AppendTerminationOption

    Const p_procedureName As String = "AppendTerminationGetter"
   
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_appendTermination As cc_isr_Ieee488.AppendTerminationOption
    p_appendTermination = cc_isr_Ieee488.AppendTerminationOption.LineFeed
    p_success = Me.TryGetAppendTermination(a_socket, p_appendTermination, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
            
    AppendTerminationGetter = p_appendTermination
    
End Function

''' <summary>   Gets the GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">     [Out, <see cref="AppendTerminationOption"/>] the value. </value>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [<see cref="AppendTerminationOption"/>]. </returns>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryGetAppendTermination(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_value As AppendTerminationOption, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryGetAppendTermination"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    Dim p_appendTermination As cc_isr_Ieee488.AppendTerminationOption
    p_appendTermination = cc_isr_Ieee488.AppendTerminationOption.LineFeed

    If Me.GpibLanControllerAttached(a_socket) Then
    
        p_length = a_socket.TrySendMessage(Me.AppendTerminationCommand & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        Dim p_reading As String
        If p_success Then
            p_length = a_socket.TryReceive(p_reading, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        End If
        
        If p_success Then _
            p_success = Me.TryParseAppendTermination(p_reading, p_appendTermination, p_details)
            
        If p_success Then _
            This.AppendTermination = p_appendTermination
            
        If p_success Then _
            a_value = p_appendTermination
            
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
        "; " & p_details
    
    TryGetAppendTermination = p_success

End Function

''' <summary>   Tries to parse GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryParseAppendTermination(ByVal a_value As String, _
    ByRef a_appendTermination As AppendTerminationOption, _
    ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseAppendTermination"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_value As Integer
    p_success = cc_isr_Core.StringExtensions.TryParseInteger(a_value, p_value, p_details)
    If p_success Then _
        a_appendTermination = p_value
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryParseAppendTermination = p_success
    
End Function


''' <summary>   Sets the GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">    [<see cref="AppendTerminationOption"/>] the option to set. </param>
Public Sub AppendTerminationSetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_value As AppendTerminationOption)

    Const p_procedureName As String = "AppendTerminationSetter"
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_appendTermination As cc_isr_Ieee488.AppendTerminationOption
    p_appendTermination = cc_isr_Ieee488.AppendTerminationOption.LineFeed
    p_success = Me.TrySetAppendTermination(a_socket, a_value, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Sub

''' <summary>   Sets the GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">     [<see cref="AppendTerminationOption"/>] The value to set. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TrySetAppendTermination(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_value As AppendTerminationOption, _
    ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TrySetAppendTermination"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    Dim p_appendTermination As cc_isr_Ieee488.AppendTerminationOption
    p_appendTermination = cc_isr_Ieee488.AppendTerminationOption.LineFeed
    
    If Me.GpibLanControllerAttached(a_socket) Then
    
        Dim p_command As String
        Dim p_value As String
        Dim p_reading As String
        
        p_command = Me.AppendTerminationCommand
        p_value = VBA.CStr(a_value)
        
        ' send the command
        p_length = a_socket.TrySendMessage(p_command & " " & p_value & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        ' validate the command
        If p_success Then _
            p_success = Me.TryGetValue(a_socket, p_command, p_value, p_reading, p_details)
            
        If p_success Then _
            p_success = Me.TryParseAppendTermination(p_reading, p_appendTermination, p_details)
    
        If p_success Then _
            This.AppendTermination = p_appendTermination
    
    End If
   
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TrySetAppendTermination = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++read_tmo_ms - read timeout
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
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <returns>   [Long]. </returns>
Public Function ReadTimeoutGetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket) As Long

    Const p_procedureName As String = "ReadTimeoutGetter"
   
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_readTimeout As Long
    p_success = Me.TryGetReadTimeout(a_socket, p_readTimeout, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
            
    ReadTimeoutGetter = p_readTimeout
    
End Function

''' <summary>   Gets the GPIB Lan Controller read timeout. </summary>
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">    [Out, Long] The parsed value. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryGetReadTimeout(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_value As Long, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryGetReadTimeout"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    Dim p_timeout As Long

    If Me.GpibLanControllerAttached(a_socket) Then
    
        Dim p_sentCount As Integer
        p_length = a_socket.TrySendMessage(Me.ReadTimeoutCommand & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        Dim p_reading As String
        If p_success Then
            p_length = a_socket.TryReceive(p_reading, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        End If
        
        If p_success Then _
            p_success = Me.TryParseReadTimeout(p_reading, p_timeout, p_details)
    
    End If
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
            
    a_value = p_timeout
    TryGetReadTimeout = p_success

End Function

''' <summary>   Tries to parse GPIB Lan Controller <see cref="Integer"/>. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_timeout">   [Out, Long] The parsed value. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryParseReadTimeout(ByVal a_value As String, _
    ByRef a_timeout As Long, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseReadTimeout"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_value As Long
    p_success = cc_isr_Core.StringExtensions.TryParseLong(a_value, p_value, p_details)
    If p_success Then _
        a_timeout = p_value
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryParseReadTimeout = p_success

End Function

''' <summary>   Sets the GPIB Lan Controller read timeout in milliseconds. </summary>
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">    [<see cref="Integer"/>] the read timeout in milliseconds to set. </param>
Public Sub ReadTimeoutSetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_value As Integer)

    Const p_procedureName As String = "ReadTimeoutSetter"

    If Me.GpibLanControllerAttached(a_socket) Then
    
        Dim p_details As String
        If Not Me.TryReadTimeoutSetter(a_socket, a_value, p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
    End If
    
End Sub

''' <summary>   Tries to set the GPIB Lan Controller read timeout. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">     [Integer] The value to set. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryReadTimeoutSetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_value As Integer, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryReadTimeoutSetter"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached(a_socket) Then
    
        Dim p_command As String
        Dim p_value As String
        Dim p_reading As String
        
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
            p_length = a_socket.TrySendMessage(p_command & " " & p_value & This.Termination, _
                    p_details, This.ReadAfterWriteDelay)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
            
            ' validate the command
            If p_success Then _
                p_success = Me.TryGetValue(a_socket, p_command, p_value, p_reading, p_details)
                
            Dim p_readTimeout As Long
            If p_success Then _
                p_success = Me.TryParseReadTimeout(p_reading, p_readTimeout, p_details)
                
            If p_success Then _
                This.ReadTimeout = p_readTimeout
                
        End If
        
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryReadTimeoutSetter = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++eoi - End or identify
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
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function EndOrIdentifyGetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket) As Boolean

    Const p_procedureName As String = "EndOrIdentifyGetter"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_result As Boolean
    
    p_success = Me.TryGetEndOrIdentify(a_socket, p_result, p_details)
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
    EndOrIdentifyGetter = p_result

End Function

''' <summary>   Gets the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">     [Out, Boolean] enables (true) or disables (false) the assertion of the
'''                            EOI signal with the last character of any command sent over GPIB port. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryGetEndOrIdentify(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_value As Boolean, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryGetEndOrIdentify"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached(a_socket) Then
    
        p_length = a_socket.TrySendMessage(Me.EndOrIdentifyCommand & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
         
        Dim p_reading As String
        If p_success Then
            p_length = a_socket.TryReceive(p_reading, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        End If
        
        If p_success Then _
            p_success = Me.TryParseEndOrIdentify(p_reading, a_value, p_details)
            
        If p_success Then _
            This.EndOrIdentify = a_value
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryGetEndOrIdentify = p_success

End Function


''' <summary>   Tries to parse the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryParseEndOrIdentify(ByVal a_value As String, _
    ByRef a_endOrIdentify As Boolean, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseEndOrIdentify"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    If "1" = a_value Then
        a_endOrIdentify = True
        p_success = True
    ElseIf "0" = a_value Then
        a_endOrIdentify = False
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
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">    [Boolean] enables (true) or disables (false) the assertion of the EOI signal with the
'''                           last character of any command sent over GPIB port. </param>
Public Sub EndOrIdentifySetter(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_value As Boolean)

    Const p_procedureName As String = "EndOrIdentifySetter"
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    p_success = Me.TrySetEndOrIdentify(a_socket, a_value, p_details)
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
End Sub

''' <summary>   Tries to set the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_value">     [Boolean] enables (true) or disables (false) the assertion of the EOI
'''                            signal with the last character of any command sent over GPIB port. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TrySetEndOrIdentify(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_value As Boolean, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TrySetEndOrIdentify"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached(a_socket) Then
    
        Dim p_command As String
        Dim p_value As String
        Dim p_reading As String
        
        p_command = Me.EndOrIdentifyCommand
        p_value = IIf(a_value, "1", "0")
        
        ' send the command
        p_length = a_socket.TrySendMessage(p_command & " " & p_value & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        ' validate the command
        If p_success Then _
            p_success = Me.TryGetValue(a_socket, p_command, p_value, p_reading, p_details)
            
        Dim p_endOrIdentify As Boolean
        If p_success Then _
            p_success = Me.TryParseEndOrIdentify(p_reading, p_endOrIdentify, p_details)
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TrySetEndOrIdentify = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++spoll - serial poll
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Queries the GPIB SRQ signal status. </summary>
''' <remarks>   This command returns the current state of the GPIB SRQ signal. The controller returns
''' <c>1</c> if the SRQ signal is asserted (low) and <c>0</c> if the signal is not asserted (high). </remarks>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Integer] <c>1</c> if the SRQ signal is asserted (low),
''' <c>0</c> if the signal is not asserted (high), or -1 if the query failed. </returns>
Public Function TryQueryServiceRequested(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String) As Integer
    
    Const p_procedureName As String = "ServiceRequested"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    Dim p_result As Integer
    Dim p_affirmativeReply As String: p_affirmativeReply = "1"
    Dim p_negativeReply As String: p_negativeReply = "0"
    
    If Me.GpibLanControllerAttached(a_socket) Then

        Dim p_reading As String
        p_length = a_socket.TryQuery("++srq", p_reading, p_details)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        If p_success Then
            If 0 = VBA.StrComp(p_reading, p_affirmativeReply, VBA.VbCompareMethod.vbTextCompare) Then
                p_result = 1
            ElseIf 0 = VBA.StrComp(p_reading, p_negativeReply, VBA.VbCompareMethod.vbTextCompare) Then
                p_result = 0
            Else
                p_result = -1
                p_details = "The query reply '" & p_reading & _
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
''' <param name="a_socket">   [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <returns>   [Integer] The status byte or RECEIVE_ERROR (-1) if failed receiving a reading or
''' failed parsing the reading to an integer.
''' </returns>
Public Function SerialPoll(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket) As Integer
    
    Const p_procedureName As String = "SerialPoll"
    
    Dim p_statusByte As Integer: p_statusByte = 0
    Dim p_details As String: p_details = VBA.vbNullString
    
    p_statusByte = Me.TrySerialPoll(a_socket, p_details)
    If p_statusByte < 0 Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
            ThisWorkbook.VBProject & "." & TypeName(Me) & ".SerialPoll", _
            " Failed reading '" & Me.SerialPollCommand & "'; " & p_details
    
    SerialPoll = p_statusByte

End Function

''' <summary>   Tries to read the status byte of the current GPIB instrument. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Integer] The status byte or RECEIVE_ERROR (-1) if failed receiving a reading or
''' failed parsing the reading to an integer.
''' </returns>
Public Function TrySerialPoll(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String) As Integer
    
    Const p_procedureName As String = "TrySerialPoll"
    Dim p_statusByte As Integer: p_statusByte = 0
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If Me.GpibLanControllerAttached(a_socket) Then
    
        p_length = a_socket.TrySendMessage(Me.SerialPollCommand & This.Termination, _
                p_details, This.ReadAfterWriteDelay)
        p_success = (p_length > 0) And (0 = VBA.Len(p_details))
        
        Dim p_reading As String
        If p_success Then
        
            p_length = a_socket.TryReceive(p_reading, p_details)
            p_success = (p_length > 0) And (0 = VBA.Len(p_details))

            If p_success Then _
                p_success = cc_isr_Core.StringExtensions.TryParseInteger(p_reading, p_statusByte, p_details)
        
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
''' <param name="a_socket">        [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_bitsStatus"/>   [Integer] The expected status of the specified status bits. </param>
''' <param name="a_statusBits"/>   [Integer] The expected status bits. </param>
''' <param name="a_timeout"/>      [Long] The timeout in milliseconds. </param>
''' <param name="a_statusByte"/>   [Out, Integer] The last received status byte.
'''                                Negative if serial poll failed. </param>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if the status byte has the expected bits value. </returns>
Public Function AwaitStatusBits(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_bitsStatus As Integer, ByVal a_statusBits As Integer, _
    ByVal a_timeout As Long, ByRef a_statusByte As Integer, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "AwaitStatusBits"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_gotIt As Boolean

    If Me.GpibLanControllerAttached(a_socket) Then
    
        Dim p_stopper As cc_isr_Core_IO.Stopwatch
        Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
        p_stopper.Restart
        
        Do
            DoEvents
            ' read the status bit.
            a_statusByte = Me.TrySerialPoll(a_socket, p_details)
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
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            should connect. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function ShouldConnect(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "ShouldConnect"

    If a_socket Is Nothing Then
        a_details = "The Tcp Session socket instance is nothing."
        ShouldConnect = True
    ElseIf Not a_socket.Connected Then
        a_details = "The session socket is not connected."
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
''' <param name="a_socket">     [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">    [Out, String] Details why the instrument state
'''                             should be restored. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function ShouldRestoreInitialState(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "ShouldRestoreInitialState"
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    If Me.ShouldConnect(a_socket, p_details) Then
        
        ShouldRestoreInitialState = True
    
    Else
    
        Dim p_autoAssertTalk As Boolean
        p_autoAssertTalk = Me.AutoAssertTalkGetter(a_socket)
        DoEvents
        If p_autoAssertTalk Then
            a_details = "GPIB Lan Read After Write is enabled. It should be disabled."
            ShouldRestoreInitialState = True
        Else
            Dim p_timeout As Long
            p_timeout = Me.TryGetReadTimeout(a_socket, p_timeout, p_details)
            If p_timeout <> This.SocketTimeout Then
                a_details = "GPIB Lan read timeout '" & VBA.CStr(p_timeout) & _
                    "' is not equal to the socket timeout '& VBA.CStr(this.SocketTimeout) & " '."
                ShouldRestoreInitialState = True
            Else
                ShouldRestoreInitialState = False
            End If
        End If
    
    End If

End Function

''' <summary>   Clears the instrument output and then clears the instrument execution state. </summary>
''' <remarks>   This method catches and hides any execution errors. </remarks>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_command">   [Optional, String, <c>*CLS;*WAI;*OPC?</c>] the command. </param>
''' <param name="a_timeout">   [Optional, Integer, 100] the timeout. </param>
''' <returns>   [String. </returns>
Public Function QueryClearExecutionState(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    Optional ByVal a_command As String = "*CLS;*WAI;*OPC?", _
    Optional ByVal a_timeout As Integer = 100) As String
        
    Const p_procedureName As String = "QueryClearExecutionState"
        
    On Error Resume Next
    
    Dim p_details As String
    Dim p_reading As String
    Dim p_success As Boolean
    Dim p_length As Long
    
    ' and then clear the execution state.
    p_length = Me.TryQueryLine(a_socket, a_command, p_reading, p_details)
    DoEvents
    p_length = Me.TryQueryLine(a_socket, a_command, p_reading, p_details)
    DoEvents
    p_length = Me.TryQueryLine(a_socket, a_command, p_reading, p_details)
    DoEvents
    
    On Error GoTo 0

    QueryClearExecutionState = p_reading

End Function

''' <summary>   Configures the GPIB-Lan controller initial state. </summary>
''' <remarks>   Call this method upon connecting and restoring Initial State. </remarks>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryConfigureInitialState(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryConfigureInitialState"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If p_success And Me.GpibLanControllerAttached(a_socket) Then
    
        Dim p_timeout As Integer
        p_timeout = VBA.CInt(cc_isr_Core_IO.CoreExtensions.Clamp(This.ReadTimeout, _
            This.MinimumTimeout, This.ReadTimeoutMaximum))
            
        p_success = Me.TrySetValue(a_socket, Me.ReadTimeoutCommand, _
            VBA.CStr(p_timeout), _
            p_details)
        If p_success Then This.SocketTimeout = p_timeout
    
    End If
    
    If p_success And Me.GpibLanControllerAttached(a_socket) Then
    
        ' set the GPIB termination characters to none - do not append termination characters.
        p_success = Me.TrySetValue(a_socket, Me.AppendTerminationCommand, _
            VBA.CStr(AppendTerminationOption.AppendNothing), _
            p_details)
        If p_success Then This.AppendTermination = AppendTerminationOption.AppendNothing
    
    End If
    
    If p_success And Me.GpibLanControllerAttached(a_socket) Then
    
        ' uturn on senting end or identity at the end of each message to the instrument.
        p_success = Me.TrySetValue(a_socket, Me.EndOrIdentifyCommand, "1", p_details)
        If p_success Then This.EndOrIdentify = True
    
    End If
    
    If p_success And Me.GpibLanControllerAttached(a_socket) Then
    
        ' turn off auto assert talk
        p_success = Me.TrySetValue(a_socket, Me.AutoAssertTalkCommand, "0", p_details)
        If p_success Then This.AutoAssertTalk = False
    
    End If
    
    If p_success And Me.GpibLanControllerAttached(a_socket) Then
    
        ' disable front panel operation of the currently addressed instrument.
        p_length = a_socket.TrySendMessage(Me.LocalLockoutCommand & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
   
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryConfigureInitialState = p_success

End Function

''' <summary>   Configures the GPIB-Lan controller final state. </summary>
''' <remarks>   Call this method upon disconnecting. </remarks>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryConfigureFinalState(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryConfigureFinalState"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    If p_success Then
    
        ' turn off auto assert talk
        p_success = Me.TrySetValue(a_socket, Me.AutoAssertTalkCommand, "0", p_details)
        If p_success Then This.AutoAssertTalk = False
    
    End If
    
    If p_success Then
    
        ' enable front panel operation of the currently addressed instrument.
        p_length = a_socket.TrySendMessage(Me.GoToLocalCommand & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
   
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
   
    
    TryConfigureFinalState = p_success

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' controller setter and getter
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Tries to get and validate a GPIB-Lan controller settings returning the actual
''' value and true if the correct value was set or false if not. </summary>
''' <remarks>   This method assumes that the bare command is the query command.  </remarks>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_command">   [String] the bare command, e.g.,<c>++eos</c>. </param>
''' <param name="a_value">     [String] the expected value to get. </param>
''' <param name="a_reading">   [Out, String] the actual value that was get. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <value>   [Boolean]. </value>
Public Function TryGetValue(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_command As String, ByVal a_value As String, _
    ByRef a_reading As String, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryGetValue"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    If a_socket.Port = This.GpibLanControllerPort Then
        p_success = True
    Else
        p_success = False
        p_details = "The GPIB-Lan controller must be attached for getting the '" & _
            a_command & "' value expected as '" & a_value & "'."
    End If

    ' send the query command.
    Dim p_length As Integer
    
    If p_success Then
    
        p_length = a_socket.TrySendMessage(a_command & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    End If
    
    If p_success Then
    
        p_length = a_socket.TryReceive(a_reading, p_details)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
        If Not p_success Then
            p_details = "Command '" & a_command & "' failed; " & p_details
        End If
    
    End If
    
    If p_success Then
    
        p_success = 0 = VBA.StrComp(a_value, a_reading, VBA.VbCompareMethod.vbTextCompare)
        
        If Not p_success Then
            p_details = "Command '" & a_command & "' failed; Result '" & _
                a_reading & "' is not equal to the expected value '" & a_value & "'."
        End If
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryGetValue = p_success
 
End Function

''' <summary>   Tries to set a GPIB-Lan controller settings. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_command">   [String] The command. </param>
''' <param name="a_value">     [String] the value to set. </param>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TrySetValue(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_command As String, ByVal a_value As String, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TrySetValue"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If a_socket.Port = This.GpibLanControllerPort Then
        p_success = True
    Else
        p_success = False
        p_details = "The GPIB-Lan controller must be attached for setting the '" & _
            a_command & "' value to '" & a_value & "'."
    End If
    
    If p_success Then
    
        ' send the command
        p_length = a_socket.TrySendMessage(a_command & " " & a_value & This.Termination, _
            p_details, This.ReadAfterWriteDelay)
        p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    End If
    
    If p_success Then
    
        Dim p_reading As String

        ' validate the command
        p_success = Me.TryGetValue(a_socket, a_command, a_value, p_reading, p_details)
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TrySetValue = p_success

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Ping controller and device
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Tries to open and initialize a socket connection to the controller. </summary>
''' <param name="a_address">   [String] The socket dotted address including the port. </param>
''' <param name="a_timeout">   [Long] The socket and controller read timeouts. </param>
''' <param name="a_socket">    [Out, <see cref="cc_isr_Winsock.IPv4StreamSocket"/>] The
'''                            connected socket. </param>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryOpenConnection(ByVal a_address As String, _
    ByVal a_timeout As Long, _
    ByRef a_socket As cc_isr_Winsock.IPv4StreamSocket, ByRef a_details As String) As Boolean

    Const p_procedureName = "TryOpenConnection"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' set the read timeout for initial configuration
    This.ReadTimeout = a_timeout
    
    Set a_socket = cc_isr_Winsock.Factory.NewIPv4StreamSocket().Initialize()
    
    p_success = a_socket.TryOpenConnection(a_address, a_timeout, p_details)
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details

    TryOpenConnection = p_success
    
End Function

''' <summary>   Tries to ping the Interface and the Instrument. </summary>
''' <param name="a_address">   [String] The socket dotted address including the port. </param>
''' <param name="a_timeout">   [Long] The socket and controller read timeouts. </param>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryPing(ByVal a_address As String, _
    ByVal a_timeout As Long, _
    ByRef a_details As String) As Boolean

    Const p_procedureName = "TryPing"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_socket As cc_isr_Winsock.IPv4StreamSocket
    p_success = Me.TryOpenConnection(a_address, a_timeout, p_socket, p_details)
    
    If p_success Then
    
         p_success = Me.TryConfigureInitialState(p_socket, p_details)
         
         ' if failed, try to power on reset
         If Not p_success Then
         
             p_success = Me.TryPowerOnReset(p_socket, p_details)
             
             If p_success Then _
                 p_success = Me.TryConfigureInitialState(p_socket, p_details)
             
        End If
    
    End If
    
    If p_success Then _
         p_success = Me.TryPingDevice(p_socket, p_details)
    
    If p_success And p_socket.Connected Then _
        p_success = p_socket.TryCloseConnection(p_details)
   
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details

    TryPing = p_success
    
End Function

''' <summary>   Tries to ping the device with the <see cref="ConnectionQueryMessage"/>. </summary>
''' <param name="a_socket">    [<see cref="cc_isr_Winsock.IPv4StreamSocket"/>] a connected socket. </param>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean] True if success. </returns>
Public Function TryPingDevice(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryPingDevice"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    ' establish a connection to the actual instrument because a connection
    ' to the GPIB-Lan controller is insuffcient to ascertain that the instrument
    ' is responsive.
    Dim p_reading As String
    p_length = Me.TryQueryLine(a_socket, This.ConnectionQueryMessage, p_reading, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If p_success Then
        p_success = 0 = VBA.StrComp(p_reading, This.ConnectionQueryReply, vbTextCompare)
        If Not p_success Then _
            p_details = "Failed validating connection query reply; received '" & p_reading & _
                "' instead of '" & This.ConnectionQueryReply & _
                "' in reply to '" & This.ConnectionQueryMessage & "'."
    Else
        p_details = "Failed validating connection; & p_details"
    End If
                
    If Not p_success Then _
        a_details = ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryPingDevice = p_success

End Function


