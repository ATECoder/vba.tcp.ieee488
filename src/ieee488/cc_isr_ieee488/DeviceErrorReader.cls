VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeviceErrorReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Implements the Device Error Reader interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Implements IDeviceErrorReader

Private Type this_
    Session As cc_isr_Ieee488.TcpSession
    ClearCommand As String
    QueryCommand As String
    LastError As String
    LastErrors As String
End Type

Private This As this_

''' <summary>   Returns a reference to the <see cref="IDeviceErrorReader"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.TcpSession"/>] A connected session. </param>
''' <value>   [<see cref="IDeviceErrorReader"/>]. </returns>
Private Function IDeviceErrorReader_Initialize(ByVal a_session As cc_isr_Ieee488.TcpSession) As IDeviceErrorReader
    Set This.Session = a_session
    Set IDeviceErrorReader_Initialize = Me
End Function

''' <summary>   Gets the device error clear command. </summary>
''' <value>   [String] </value>
Public Property Get IDeviceErrorReader_ClearCommand() As String
    IDeviceErrorReader_ClearCommand = This.ClearCommand
End Property

''' <summary>   Sets the device error clear command. </summary>
''' <param name="a_value">    [String] the clear command, e.g., ':SYST:CLE'. </param>
Public Property Let IDeviceErrorReader_ClearCommand(ByVal a_value As String)
    This.ClearCommand = a_value
End Property

''' <summary>   Gets the device error Query command. </summary>
''' <value>   [String] </value>
Public Property Get IDeviceErrorReader_QueryCommand() As String
    IDeviceErrorReader_QueryCommand = This.QueryCommand
End Property

''' <summary>   Sets the device error Query command. </summary>
''' <param name="a_value">    [String] the query command, e.g., ':SYST:ERR?'. </param>
Public Property Let IDeviceErrorReader_QueryCommand(ByVal a_value As String)
    This.QueryCommand = a_value
End Property

''' <summary>   Returns a reference to the <see cref="IDeviceErrorReader"/>. </summary>

''' <summary>   The last error that was dequeued from the device. </summary>
''' <value>   [String] </value>
Private Property Get IDeviceErrorReader_LastError() As String
    IDeviceErrorReader_LastError = This.LastError
End Property

''' <summary>   The last errors that were dequeued from the device. </summary>
''' <value>   [String] </value>
Private Property Get IDeviceErrorReader_LastErrors() As String
    IDeviceErrorReader_LastErrors = This.LastErrors
End Property

''' <summary>   Dequeues the top device error from the device error queue. </summary>
''' <param name="a_details">    [Out, String] Details why this method failed. </param>
''' <value>   [String] the top error on the error queue. </value>
Private Function IDeviceErrorReader_TryDequeueError(ByRef a_details As String) As String
    
    Const p_procedureName As String = "IDeviceErrorReader_TryDequeueError"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_error As String: p_error = VBA.vbNullString
    p_success = Me.TryDequeueDeviceError(p_error, p_details)
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    This.LastError = p_error
    IDeviceErrorReader_TryDequeueError = This.LastError

End Function

''' <summary>   Dequeues a device error. </summary>
''' <param name="a_error">      [Out, String] The device error. </param>
''' <param name="a_details">    [Out, String] Details why this method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryDequeueDeviceError(ByRef a_error As String, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryDequeueDeviceError"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_command As String: p_command = This.QueryCommand
    
    If 0 = VBA.Len(p_command) Then
        TryDequeueDeviceError = p_success
        Exit Function
    End If
    
    Dim p_reply As String: p_reply = VBA.vbNullString: a_error = VBA.vbNullString
    
    p_length = This.Session.TryQueryLine(p_command, p_reply, p_details)
    p_success = (p_length >= 0) And (0 = VBA.Len(p_details))
    
    If p_success Then
        a_error = p_reply
    Else
        p_details = "Failed querying the '" & p_command & "' query; " & p_details
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryDequeueDeviceError = p_success

End Function

''' <summary>   Clears the device error queue. </summary>
''' <param name="a_details">    [Out, String] Details why this method failed. </param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <c>*CLS</c> command by querying <c>*CLS;*WAI;*OPC?</c></param>
''' <returns>   [Boolean]. </returns>
Public Function TryClearDeviceErrorQueue(ByRef a_details As String, Optional ByVal a_awaitOpc As Boolean = True) As Boolean

    Const p_procedureName As String = "TryClearDeviceErrorQueue"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Dim p_command As String: p_command = This.ClearCommand
    
    If 0 = VBA.Len(p_command) Then
        TryClearDeviceErrorQueue = p_success
        Exit Function
    End If
    
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
            p_details = "Failed clearing the device errors; The instrument returned '" & _
                p_reply & "' instead of '" & "" & _
                "' in response to the '" & p_command & "' query."
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryClearDeviceErrorQueue = p_success

End Function

''' <summary>   Clears the device error queue. </summary>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <see cref=""/>. </param>
''' <returns>   [String] The operation completion reply; defaults to "1". </returns>
Private Function IDeviceErrorReader_TryClearErrors(ByRef a_details As String, _
    Optional ByVal a_awaitOpc As Boolean = True) As Boolean

    IDeviceErrorReader_TryClearErrors = Me.TryClearDeviceErrorQueue(a_details, a_awaitOpc)
        
End Function

''' <summary>   Queries and tries to parse a device error into a number and a message. </summary>
''' <param name="a_number">        [Out, String] the error number; set to the <see cref="Err.Number"/>
'''                                if exception occurred parsing. </param>
''' <param name="a_description">   [Out, String] the error description; set to the
'''                                <see cref="Err.Description"/> if an exception occurred parsing. </param>
''' <param name="a_decorator">     [String] the decorator around the message part of the
'''                                device error message. </param>
''' <value>   [Boolean] true if queires and parsed. </value>
Private Function IDeviceErrorReader_TryDequeueParseError(ByRef a_number As String, _
    ByRef a_description As String, Optional ByVal a_decorator As String = """") As Boolean
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_error As String: p_error = VBA.vbNullString
    a_number = VBA.vbNullString
    a_description = VBA.vbNullString
    p_success = Me.TryDequeueDeviceError(p_error, p_details)
    If p_success Then _
        p_success = Me.TryParseDeviceError(p_error, a_number, a_description, a_decorator)
    
    IDeviceErrorReader_TryDequeueParseError = p_success
    
End Function

''' <summary>   Queries and concatenates device errors into a delimited replay. </summary>
''' <param name="a_errors">      [Out, String] a delimited string with errors or the no error message
'''                              if none. </param>
''' <param name="a_delimiter">   [Optional, String, VBA.vbCrLf] the delimiter to use when building
'''                              the reply. Note that a comma separates the device error number and error
'''                              description. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Integer] the number of dequeued errors. </value>
Private Function IDeviceErrorReader_TryDequeueErrors(ByRef a_errors As String, _
    Optional ByVal a_decorator As String = """", _
    Optional ByVal a_delimiter As String = VBA.vbCrLf) As Integer

    Dim p_delimiter As String
    Dim p_errorsCount As Integer
    On Error Resume Next
    
    Dim p_builder As String
    p_builder = VBA.vbNullString
    
    Dim p_errorNumber As String
    Dim a_errorDescription As String
    Dim p_done As Boolean
    p_done = False
    
    Do
    
        DoEvents
        
        p_delimiter = IIf(VBA.vbNullString = p_builder, VBA.vbNullString, a_delimiter)
        
        If IDeviceErrorReader_TryDequeueParseError(p_errorNumber, a_errorDescription, a_decorator) Then
            p_builder = p_builder & p_delimiter & p_errorNumber & "," & a_errorDescription
            p_done = "0" = p_errorNumber
        Else
            p_builder = p_builder & p_delimiter & p_errorNumber & "," & a_errorDescription
            p_done = True
            p_errorsCount = p_errorsCount + 1
        End If
        
        p_delimiter = IIf(VBA.vbNullString = p_builder, VBA.vbNullString, a_delimiter)
        
        ' enqueue a VBA Error as a device error.
        If Err.Number <> 0 Then
            p_builder = p_builder & p_delimiter & VBA.CStr(Err.Number) & ",Application error: " & Err.Description
            p_done = True
            p_errorsCount = p_errorsCount + 1
        End If
    
    Loop Until p_done
    
    On Error GoTo 0
    
    a_errors = p_builder
    IDeviceErrorReader_TryDequeueErrors = p_errorsCount

End Function

''' <summary>   Tries to parse a device error into a number and a message. </summary>
''' <param name="a_value">         [String] the device error. </param>
''' <param name="a_number">        [Out, String] the error number; set to the <see cref="Err.Number"/>
'''                                if exception occurred parsing. </param>
''' <param name="a_description">   [Out, String] the error description; set to the
'''                                <see cref="Err.Description"/> if an exception occurred parsing. </param>
''' <param name="a_decorator">     [String] the decorator around the message part of the
'''                                device error message. </param>
''' <value>   [Boolean] true if parsed. </value>
Public Function TryParseDeviceError(ByVal a_value As String, _
    ByRef a_number As String, ByRef a_description As String, _
    Optional ByVal a_decorator As String = """") As Boolean
    
    On Error Resume Next
    
    Dim p_success As Boolean
    
    a_number = cc_isr_Core.StringExtensions.Pop(a_value, ",")
    
    If VBA.vbNullString = a_number Then
        a_number = "0"
        a_description = "Parsing '" & a_value & "' failed getting the error number."
        p_success = False
    Else
        a_description = cc_isr_Core.StringExtensions.Pop(a_value, ",")
        If VBA.vbNullString = a_description Then
            a_description = "Parsing '" & a_value & "' failed getting the error message."
            p_success = False
        Else
            a_description = cc_isr_Core.StringExtensions.TrimLeft(a_description, a_decorator)
            a_description = cc_isr_Core.StringExtensions.TrimRight(a_description, a_decorator)
            p_success = True
        End If
        
    End If
    
    ' if an exception, set the error number and message to the error info.
    If Err.Number <> 0 Then
        a_number = VBA.CStr(Err.Number)
        a_description = "Application error: " & Err.Description
        p_success = False
    End If
    
    On Error GoTo 0

    TryParseDeviceError = p_success

End Function


''' <summary>   Tries to parse a device error into a number and a message. </summary>
''' <param name="a_value">         [String] the device error. </param>
''' <param name="a_number">        [Out, String] the error number; set to the <see cref="Err.Number"/>
'''                                if exception occurred parsing. </param>
''' <param name="a_description">   [Out, String] the error description; set to the
'''                                <see cref="Err.Description"/> if an exception occurred parsing. </param>
''' <param name="a_decorator">     [String] the decorator around the message part of the
'''                                device error message. </param>
''' <value>   [Boolean] true if parsed. </value>
Private Function IDeviceErrorReader_TryParseError(ByVal a_value As String, _
    ByRef a_number As String, ByRef a_description As String, _
    Optional ByVal a_decorator As String = """") As Boolean
    
    IDeviceErrorReader_TryParseError = Me.TryParseDeviceError(a_value, a_number, a_description, a_decorator)

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Class events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Initializes the class properties. </summary>
Private Sub Class_Initialize()
    This.ClearCommand = ":SYST:CLE"
    This.QueryCommand = ":SYST:ERR?"
    This.LastError = VBA.vbNullString
    This.LastErrors = VBA.vbNullString
End Sub
