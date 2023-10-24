VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDeviceErrorReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object that can read device errors. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Gets the device error clear command. </summary>
''' <value>   [String] </value>
Public Property Get ClearCommand() As String
End Property

''' <summary>   Sets the device error clear command. </summary>
''' <param name="a_value">    [String] the clear command, e.g., ':SYST:CLE'. </param>
Public Property Let ClearCommand(ByVal a_value As String)
End Property

''' <summary>   Gets the device error Query command. </summary>
''' <value>   [String] </value>
Public Property Get QueryCommand() As String
End Property

''' <summary>   Sets the device error Query command. </summary>
''' <param name="a_value">    [String] the query command, e.g., ':SYST:ERR?'. </param>
Public Property Let QueryCommand(ByVal a_value As String)
End Property

''' <summary>   Returns a reference to the <see cref="IDeviceErrorReader"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.TcpSession"/>] A connected session. </param>
''' <value>   [<see cref="IDeviceErrorReader"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.TcpSession) As IDeviceErrorReader
End Function

''' <summary>   The last error that was dequeued from the device. </summary>
''' <value>   [String] </value>
Public Property Get LastError() As String
End Property

''' <summary>   The last errors that were dequeued from the device. </summary>
''' <value>   [String] </value>
Public Property Get LastErrors() As String
End Property

''' <summary>   Dequeues the top device error from the device error queue. </summary>
''' <value>   [String] the top error on the error queue. </value>
Public Function TryDequeueError(ByRef a_details As String) As String
End Function

''' <summary>   Clears the device error queue. </summary>
''' <param name="a_details">    [Out, String] Details why this method failed. </param>
''' <param name="a_awaitOpc">   [Optional, Boolean, True] True to wait for operation completion
'''                             after issuing the <see cref=""/>. </param>
''' <returns>   [String] The operation completion reply; defaults to "1". </returns>
Public Function TryClearErrors(ByRef a_details As String, Optional ByVal a_awaitOpc As Boolean = True) As Boolean
End Function

''' <summary>   Queries and tries to parse a device error into a number and a message. </summary>
''' <param name="a_number">      [Out, String] the error number; set to the <see cref="Err.Number"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_message">     [Out, String] the error message; set to the <see cref="Err.Description"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Boolean] true if parsed. </value>
Public Function TryDequeueParseError(ByRef a_number As String, _
    ByRef a_message As String, Optional ByVal a_decorator As String = """") As Boolean
End Function

''' <summary>   Queries and concatenates device errors into a delimited replay. </summary>
''' <param name="a_errors">      [Out, String] a delimited string with errors or the no error message
'''                              if none. </param>
''' <param name="a_delimiter">   [Optional, String, VBA.vbCrLf] the delimiter to use when building
'''                              the reply. Note that a comma separates the device error number and error
'''                              description. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Integer] the number of dequeued errors. </value>
Public Function TryDequeueErrors(ByRef a_errors As String, _
    Optional ByVal a_decorator As String = """", _
    Optional ByVal a_delimiter As String = VBA.vbCrLf) As Integer
End Function

''' <summary>   Tries to parse a device error into a number and a message. </summary>
''' <param name="a_value">       [String] the device error. </param>
''' <param name="a_number">      [Out, String] the error number; set to the <see cref="Err.Number"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_message">     [Out, String] the error message; set to the <see cref="Err.Description"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Boolean] true if parsed. </value>
Public Function TryParseError(ByVal a_value As String, _
    ByRef a_number As String, ByRef a_message As String, _
    Optional ByVal a_decorator As String = """") As Boolean
End Function

