VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ieee488Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_socketReadTimeoutCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_socketAddressCell As String = "C2"
Private Const m_socketIdCell As String = "E2"
Private Const m_connectedCell As String = "F2"
Private Const m_sentMessageCell As String = "G2"
Private Const m_receivedMessageLengthCell As String = "K2"
Private Const m_receivedMessageCell As String = "L2"
Private Const m_elapsedTimeCell As String = "M2"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_lastMessageCell As String = "B4"
Private Const m_lastErrorMessageCell As String = "B22"

Private Const m_serialPollCell As String = "F12"
Private Const m_statusByteCell As String = "G12"
Private Const m_StandardByteCell As String = "I12"
Private Const m_srqCell As String = "C10"
Private Const m_gpibAddressCell As String = "C12"
Private Const m_gpibLanReadTimeoutCell  As String = "C15"
Private Const m_readAfterWriteEnabledCell As String = "C18"

Private WithEvents m_socket As cc_isr_Winsock.IPv4StreamSocket
Attribute m_socket.VB_VarHelpID = -1

Private Type this_
    Initialized As Boolean
    StopWatch As StopWatch
    Device As cc_isr_Ieee488.Device
End Type

Private This As this_

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_disconnected As Boolean
    
    If This.Device Is Nothing Then
        p_disconnected = True
    Else
        p_disconnected = Not This.Device.Connected
    End If
    
    If Not p_disconnected Then p_disconnected = This.Device.Connectable.CloseConnection()
    
    If p_disconnected Then
        
        If Not This.Device Is Nothing Then _
            This.Device.Dispose
    End If

    Set m_socket = Nothing
    Set This.Device = Nothing
    Set This.StopWatch = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Toggle the enable state of the controls. </summary>
''' <param name="a_connected">   [Boolean] True if connected; otherwise, false. </param>
Friend Sub EnableControls(ByVal a_connected As Boolean)

    Const p_procedureName = "EnableControls"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
   
    Me.ConnectToggleButton.Value = a_connected
    
    ' report the connection state
    Range(m_connectedCell).Value = a_connected
        
    Me.RSTButton.Enabled = a_connected
    Me.CLSButton.Enabled = a_connected
    Me.QueryButton.Enabled = a_connected
    Me.ReadButton.Enabled = a_connected
    Me.ReadStatusByteButton.Enabled = a_connected
    Me.ReadStandardEventStatusButton.Enabled = a_connected
    Me.WriteButton.Enabled = a_connected
    
    If a_connected Then
        EnableGpibLanControls This.Device.Session.GpibLanControllerAttached
    Else
        EnableGpibLanControls False
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Populates the Commands combo box. </summary>
Friend Sub PopulateCommandsComboBox()
    
    Const p_procedureName = "PopulateCommandsComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.CommandsComboBox.Clear
    Me.CommandsComboBox.AddItem ClearExecutionStateCommand
    Me.CommandsComboBox.AddItem ClearExecutionStateCommand & ";" & _
                                OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem IdentityQueryCommand
    Me.CommandsComboBox.AddItem OperationCompleteCommand
    Me.CommandsComboBox.AddItem OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem OptionsQueryCommand
    Me.CommandsComboBox.AddItem ResetKnownStateCommand
    Me.CommandsComboBox.AddItem ResetKnownStateCommand & ";" & _
                                OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
                                StandardEventEnableCommand, &H7F) & ";" & _
                                OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem StandardEventEnableQueryCommand
    Me.CommandsComboBox.AddItem StandardEventStatusQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
                                ServiceRequestEnableCommand, &H7F) & ";" & _
                                OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
                                StandardServiceEnableCommand, &H7F, &H7F) & ";" & _
                                OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
                                StandardServiceEnableCommand, &H7F, &H7F) & ";" & _
                                OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem ServiceRequestEnableQueryCommand
    Me.CommandsComboBox.AddItem ServiceRequestQueryCommand
    Me.CommandsComboBox.AddItem WaitCommand
    Me.CommandsComboBox.AddItem WaitCommand & ";" & _
                                OperationCompletedQueryCommand

    Me.CommandsComboBox.ListIndex = 1

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Restarts the stop watch and clears the elapsed time cell. </summary>
Friend Sub RestartStopwatch()
    
    Range(m_elapsedTimeCell).Value = vbNullString
    This.StopWatch.Restart

End Sub

''' <summary>   Reads the stop watch and updates the elapsed time cell. </summary>
Friend Sub ReadStopwatch()
    
    Range(m_elapsedTimeCell).Value = Format(This.StopWatch.ElapsedMilliseconds, m_elapsedTimeFormat)

End Sub

''' <summary>   Write line handling an error messages. </summary>
''' <param name="a_message">   [String] to send to the instrument. </param>
Friend Sub WriteLine(ByVal a_message As String)

    Const p_procedureName = "WriteLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.Device.Session.WriteLine a_message
    
    Me.ReadStopwatch
    
    If This.Device.Session.GpibLanControllerAttached Then
        Me.ReadAfterWriteEnabled = This.Device.Session.AutoAssertTalk
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Query a message from the instrument handling any errors. </summary>
''' <param name="a_message">   [String] to send to the instrument. </param>
Friend Function QueryLine(ByVal a_message As String) As String

    Const p_procedureName = "QueryLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    QueryLine = This.Device.Session.QueryLine(a_message, False)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Read from the instrument handling any errors. </summary>
Friend Function Read() As String

    Const p_procedureName = "Read"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Read = This.Device.Session.Read()
    
    Me.ReadStopwatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function


Public Function IsConnected() As Boolean

    If This.Device Is Nothing Then
        IsConnected = False
    Else
        IsConnected = This.Device.Connected
    End If
    
End Function

''' <summary>   Handles the Toggle connection button click event. </summary>
Friend Sub ConnectToggleButton_Click()

    Const p_procedureName = "ConnectToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_host As String: p_host = Me.Host
    Dim p_port As Long: p_port = Me.Port
    Dim p_sessionReadTimeout As Long: p_sessionReadTimeout = Me.SessionReadTimeout
    
    If Not ConnectToggleButton.Value And Me.IsConnected Then
        
        This.Device.Connectable.CloseConnection
        
    ElseIf ConnectToggleButton.Value And Not Me.IsConnected Then
       
        Set This.Device = cc_isr_Ieee488.Factory.NewDevice.Initialize()
        Set m_socket = This.Device.Session.Socket
        Set This.StopWatch = cc_isr_Core_IO.Factory.NewStopwatch()
        Me.EnableControls False
       
        Me.RestartStopwatch

        ' populate the combo box commands
        PopulateCommandsComboBox
       
        This.Device.Connectable.OpenConnection p_host & ":" & VBA.CStr(p_port), _
            p_sessionReadTimeout
    
    End If
    
    Me.ReadStopwatch
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Clear Know State button click event. </summary>
Private Sub CLSButton_Click()
    
    Const p_procedureName = "CLSButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.Device.ClearExecutionState
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()
    
    Const p_procedureName = "RSTButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    p_success = This.Device.TryResetKnownState(p_details)
    If Not p_success Then _
        Me.LastErrorMessage = p_details
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    
    Const p_procedureName = "SDCButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    p_success = This.Device.Session.TrySelectiveDeviceClear(p_details)
    
    If Not p_success Then _
        Me.LastErrorMessage = p_details
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Private Sub ReadStandardEventStatusButton_Click()

    Const p_procedureName = "ReadStandardEventStatusButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Me.StandardByte = This.Device.QueryStandardEventsStatus
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub


''' <summary>   Handles the Read Status Byte button click event. </summary>
Friend Sub ReadStatusByteButton_Click()
    
    Const p_procedureName = "ReadStatusByteButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Me.StatusByte = This.Device.QueryServiceRequestStatus
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()

    Const p_procedureName = "WriteButton_Click"
    
   
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.SentMessage = vbNullString
    Dim p_messageToSend As String: p_messageToSend = Me.CommandsComboBox.Value
    Me.WriteLine p_messageToSend
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_value As Integer: p_value = -1
    Me.SentMessage = p_messageToSend
    If This.Device.Session.GpibLanControllerAttached And Me.AutoStatusReadCheckBox.Value Then
        p_value = This.Device.Session.TrySerialPoll(p_details)
        Range(m_serialPollCell).Value = p_value
        If p_value < 0 Then _
            Me.LastErrorMessage = p_details
    End If
    
    Dim p_isQuery As Boolean: p_isQuery = cc_isr_Core.StringExtensions.EndsWith(p_messageToSend, "?")
    If Not p_isQuery And Me.AutoStatusReadCheckBox.Value Then
        Me.StatusByte = This.Device.QueryServiceRequestStatus
        Me.StandardByte = This.Device.QueryStandardEventsStatus
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()

    Const p_procedureName = "ReadButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReceivedMessage = vbNullString
    Dim p_messageReceived As String
    
    Me.RestartStopwatch
    
    p_messageReceived = Me.Read
    
    Me.ReadStopwatch
    
    Me.ReceivedMessage = p_messageReceived
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    
    Const p_procedureName = "QueryButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReceivedMessage = vbNullString
    Me.SentMessage = vbNullString
    Dim p_messageToSend As String: p_messageToSend = Me.CommandsComboBox.Value
    Dim p_messageReceived As String
    
    p_messageReceived = Me.QueryLine(p_messageToSend)
    
    Me.ReceivedMessage = p_messageReceived
    Me.SentMessage = p_messageToSend
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Gets the host dotted IP Address. </summary>
Friend Property Get Host() As String
    Host = Range(m_inputHostCell).Value
End Property

''' <summary>   Gets the port number. </summary>
Friend Property Get Port() As Integer
    Port = CInt(Range(m_inputPortCell).Value)
End Property

''' <summary>   Get the session read timeout. </summary>
Friend Property Get SessionReadTimeout() As Integer
    SessionReadTimeout = CInt(Range(m_socketReadTimeoutCell).Value)
End Property

''' <summary>   Get the GPIB-Lan Read timeout. </summary>
Friend Property Get GpibLanReadTimeout() As Integer
    GpibLanReadTimeout = CInt(Range(m_gpibLanReadTimeoutCell).Value)
End Property

''' <summary>   Set the GPIB-Lan Read timeout. </summary>
Friend Property Let GpibLanReadTimeout(ByVal a_value As Integer)
    Range(m_gpibLanReadTimeoutCell).Value = a_value
End Property


''' <summary>   Gets the received message. </summary>
Friend Property Get ReceivedMessage() As String
    ReceivedMessage = Range(m_receivedMessageCell).Value
End Property

''' <summary>   Sets the received message. </summary>
Friend Property Let ReceivedMessage(ByVal a_value As String)
    Range(m_receivedMessageCell).Value = a_value
    Range(m_receivedMessageLengthCell).Value = Len(a_value)
End Property

''' <summary>   Gets the sent message. </summary>
Friend Property Get SentMessage() As String
    SentMessage = Range(m_sentMessageCell).Value
End Property

''' <summary>   Sets the sent message. </summary>
Friend Property Let SentMessage(ByVal a_value As String)
    Range(m_sentMessageCell).Value = a_value
End Property

''' <summary>   Gets the Last Error Message. </summary>
Friend Property Get LastErrorMessage() As String
    LastErrorMessage = Range(m_lastErrorMessageCell).Value
End Property

''' <summary>   Sets the Last Error Message. </summary>
Friend Property Let LastErrorMessage(ByVal a_value As String)
    Range(m_lastErrorMessageCell).Value = a_value
End Property

''' <summary>   Gets the Last Message. </summary>
Friend Property Get LastMessage() As String
    LastMessage = Range(m_lastMessageCell).Value
End Property

''' <summary>   Sets the Last Message. </summary>
Friend Property Let LastMessage(ByVal a_value As String)
    Range(m_lastMessageCell).Value = a_value
End Property

''' <summary>   Gets the status byte. </summary>
Friend Property Get StatusByte() As Integer
    StatusByte = Range(m_statusByteCell).Value
End Property

''' <summary>   Sets the status byte. </summary>
Friend Property Let StatusByte(ByVal a_value As Integer)
    Range(m_statusByteCell).Value = a_value
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
Friend Property Get ReadAfterWriteEnabled() As Boolean
    ReadAfterWriteEnabled = CBool(Range(m_readAfterWriteEnabledCell).Value)
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
Friend Property Let ReadAfterWriteEnabled(ByVal a_value As Boolean)
    Range(m_readAfterWriteEnabledCell).Value = a_value
End Property

''' <summary>   Gets the Standard byte. </summary>
Friend Property Get StandardByte() As Integer
    StandardByte = Range(m_StandardByteCell).Value
End Property

''' <summary>   Sets the Standard byte. </summary>
Friend Property Let StandardByte(ByVal a_value As Integer)
    Range(m_StandardByteCell).Value = a_value
End Property

' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +
'
' GPIB LAN
'
' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +

''' <summary>   Toggle the enabled state of the GPIB Lan controls. </summary>
''' <param name="a_usingGpibLanController">   [Boolean] True if using the GPIB-Lan controller. </param>
Friend Sub EnableGpibLanControls(ByVal a_usingGpibLanController As Boolean)
    Me.SDCButton.Enabled = a_usingGpibLanController
    Me.GoToLocalButton.Enabled = a_usingGpibLanController
    Me.ReadAfterWriteGetButton.Enabled = a_usingGpibLanController
    Me.ReadAfterWriteSetButton.Enabled = a_usingGpibLanController
    Me.LocalLockoutButton.Enabled = a_usingGpibLanController
    Me.SerialPollButton.Enabled = a_usingGpibLanController
    Me.SRQButton.Enabled = a_usingGpibLanController
    Me.GpibAddressGetButton.Enabled = a_usingGpibLanController
    Me.GpibAddressSetButton.Enabled = a_usingGpibLanController
    Me.ReadTimeoutGetButton.Enabled = a_usingGpibLanController
    Me.ReadTimeoutSetButton.Enabled = a_usingGpibLanController
End Sub

Private Sub GoToLocalButton_Click()

    Const p_procedureName = "GoToLocalButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.Device.Session.GoToLocal
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub LocalLockoutButton_Click()

    Const p_procedureName = "LocalLockoutButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.Device.Session.LocalLockout
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub ReadAfterWriteSetButton_Click()

    Const p_procedureName = "ReadAfterWriteSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.Device.Session.AutoAssertTalkSetter Me.ReadAfterWriteEnabled
    
    Me.ReadStopwatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub ReadAfterWriteGetButton_Click()

    Const p_procedureName = "ReadAfterWriteGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Me.ReadAfterWriteEnabled = This.Device.Session.AutoAssertTalkGetter
    
    Me.ReadStopwatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler


End Sub

Private Sub SerialPollButton_Click()

    Const p_procedureName = "SerialPollButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    Range(m_serialPollCell).Value = This.Device.Session.TrySerialPoll(p_details)
    
    Me.LastErrorMessage = p_details
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub SRQButton_Click()

    Const p_procedureName = "SRQButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim a_value As Integer
    a_value = This.Device.Session.TryQueryServiceRequested(p_details)
    Select Case a_value
        Case 1
            Range(m_srqCell).Value = "True"
        Case 0
            Range(m_srqCell).Value = "False"
        Case -1
            Range(m_srqCell).Value = "Failed"
    End Select
    Me.LastErrorMessage = p_details
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub GpibAddressSetButton_Click()

    Const p_procedureName = "GpibAddressSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.Device.Session.GpibAddressSetter VBA.CInt(Range(m_gpibAddressCell).Value)

    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub GpibAddressGetButton_Click()

    Const p_procedureName = "GpibAddressGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Range(m_gpibAddressCell).Value = This.Device.Session.GpibAddressGetter

    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub ReadTimeoutGetButton_Click()

    Const p_procedureName = "ReadTimeoutGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Range(m_gpibLanReadTimeoutCell).Value = This.Device.Session.ReadTimeoutGetter

    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Private Sub ReadTimeoutSetButton_Click()

    Const p_procedureName = "ReadTimeoutSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.Device.Session.ReadTimeoutSetter VBA.CInt(Range(m_gpibLanReadTimeoutCell).Value)
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0

    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +
'
' IEEE 488 Device Events
'
' +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +  +

''' <summary>   Handles the <see cref="cc_isr_Ieee488.Device.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_socket_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
        
    Const p_procedureName = "m_socket_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.EnableControls a_eventArgs.Connected

    If a_eventArgs.Connected Then
    
        ' report the socket number
        Range(m_socketAddressCell).Value = This.Device.Session.SocketAddress
        Range(m_socketIdCell).Value = This.Device.Session.Socket.SocketId
    
    Else
    
        Range(m_socketAddressCell).Value = vbNullString
        Range(m_socketIdCell).Value = vbNullString
        
    End If
    
    If This.Device.Session.GpibLanControllerAttached Then
    
        Me.ReadAfterWriteEnabled = This.Device.Session.AutoAssertTalk
    
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
   
End Sub

''' <summary>   Handles the <see cref="cc_isr_Ieee488.Device.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_socket_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Worksheet events
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Initialize this worksheet if not initialized. </summary>
Private Sub Worksheet_Activate()
    
    Const p_procedureName = "Initialize"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Not This.Initialized Then
    
        'Set This.Device = cc_isr_Ieee488.Factory.NewDevice.Initialize()
        'Set m_socket = This.Device.Session.Socket
        'Set This.StopWatch = cc_isr_Core_IO.Factory.NewStopwatch()
        'Me.EnableControls False
        
        ' populate the combo box commands
        'PopulateCommandsComboBox
    
    End If
    This.Initialized = True
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

