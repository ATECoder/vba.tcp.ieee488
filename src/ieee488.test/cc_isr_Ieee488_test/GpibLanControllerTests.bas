Attribute VB_Name = "GpibLanControllerTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Gpib Lan Controller Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   This class properties. </summary>
Private Type this_
    Name As String
    TestNumber As Integer
    PreviousTestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    SocketAddress As String
    Controller As cc_isr_Ieee488.GpibLanController
    Socket As cc_isr_Winsock.IPv4StreamSocket
    Session As cc_isr_Ieee488.TcpSession
    SessionTimeout As Long
    DelayStopper As cc_isr_Core_IO.Stopwatch
    TestStopper As cc_isr_Core_IO.Stopwatch
    ErrTracer As IErrTracer
    TestCount As Integer
    RunCount As Integer
    PassedCount As Integer
    FailedCount As Integer
    InconclusiveCount As Integer
    
    ' expected values
    IdentityCompany As String
    GpibAddress As String
    PrimaryGpibAddress As Integer
    SecondaryGpibAddress As Integer
End Type

Private This As this_

''' <summary>   Runs the specified test. </summary>
Public Function RunTest(ByVal a_testNumber As Integer) As cc_isr_Test_Fx.Assert
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.TestNumber = a_testNumber
    BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestShouldConnect
        Case 2
            Set p_outcome = TestShouldQueryIdentity
        Case 3
            Set p_outcome = TestShouldSyncrhonize
        Case 4
            Set p_outcome = TestShouldRecoverFromSyntaxError
        Case 5
            Set p_outcome = TestShouldRestoreInitialState
        Case 6
            Set p_outcome = TestShouldRecoverFromAutoAssertTalk
        Case 7
            Set p_outcome = TestShouldRestoreFromClosedConnection
        Case 8
            Set p_outcome = TestShouldReadGpibAddress
        Case 9
            Set p_outcome = TestShouldReadStatusByte
        Case 10
            Set p_outcome = TestShouldSerialPoll
        Case 10
            Set p_outcome = TestShouldToggleAutoAssertTalk
        Case 11
            Set p_outcome = TestShouldToggleAppendTermination
        Case 12
            Set p_outcome = TestShouldToggleReadTimeout
        Case 13
            Set p_outcome = TestShouldToggleEndOrIdentify
        Case 14
            Set p_outcome = TestShouldGoToLocal
        Case 15
            Set p_outcome = TestShouldLocalLockout
        Case 16
            Set p_outcome = TestShouldClearSelectiveDevice
        Case 17
            Set p_outcome = TestShouldPowerOnReset
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 12
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
''' <remarks>
''' Test 01 TestShouldConnect passed. Elapsed time: 0.1 ms.
''' Test 02 TestShouldQueryIdentity passed. Elapsed time: 24.0 ms.
''' Test 03 TestShouldSyncrhonize passed. Elapsed time: 25.0 ms.
''' Test 04 TestShouldRecoverFromSyntaxError passed. Elapsed time: 116.2 ms.
''' Test 05 TestShouldRestoreInitialState passed. Elapsed time: 786.8 ms.
''' Test 06 TestShouldRecoverFromAutoAssertTalk passed. Elapsed time: 622.1 ms.
''' Test 07 TestShouldRestoreFromClosedConnection passed. Elapsed time: 17.2 ms.
''' Test 08 TestShouldReadGpibAddress passed. Elapsed time: 3.3 ms.
''' Test 09 TestShouldReadStatusByte passed. Elapsed time: 43.9 ms.
''' Test 10 TestShouldSerialPoll passed. Elapsed time: 29.7 ms.
''' Test 11 TestShouldToggleAppendTermination passed. Elapsed time: 209.1 ms.
''' Test 12 TestShouldToggleEndOrIdentify passed. Elapsed time: 209.2 ms.
''' Test 13 TestShouldGoToLocal passed. Elapsed time: 100.1 ms.
''' Test 14 TestShouldLocalLockout passed. Elapsed time: 100.1 ms.
''' Test 15 TestShouldClearSelectiveDevice passed. Elapsed time: 10.2 ms.
''' 13:20:36 Power on reset starting. This could take 3 seconds. Please wait...
''' 13:20:39 done power on reset.
''' Test 16 TestShouldPowerOnReset passed. Elapsed time: 2598.0 ms.
''' Ran 16 out of 16 tests.
''' Passed: 16; Failed: 0; Inconclusive: 0.
''' </remakrs>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 17
    Dim p_testNumber As Integer
    For p_testNumber = 1 To This.TestCount
        Set p_outcome = RunTest(p_testNumber)
        If Not p_outcome Is Nothing Then
            This.RunCount = This.RunCount + 1
            If p_outcome.AssertInconclusive Then
                This.InconclusiveCount = This.InconclusiveCount + 1
            ElseIf p_outcome.AssertSuccessful Then
                This.PassedCount = This.PassedCount + 1
            Else
                This.FailedCount = This.FailedCount + 1
            End If
        End If
        DoEvents
    Next p_testNumber
    AfterAll
    Debug.Print "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
                "; Inconclusive: " & VBA.CStr(This.InconclusiveCount) & "."
End Sub

''' <summary>   Prepares all tests. </summary>
Public Sub BeforeAll()

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed to run all tests.")
    
    This.Name = "TcpSessionTests"
    
    Set This.Controller = cc_isr_Ieee488.GpibLanController.Initialize()
    
    This.SocketAddress = "192.168.0.252:1234"
    This.SessionTimeout = 3000
    
    ' expected settings
    This.GpibAddress = "16"
    This.PrimaryGpibAddress = 16
    This.SecondaryGpibAddress = -1
    This.IdentityCompany = "KEITHLEY INSTRUMENTS INC."
    
    Set This.ErrTracer = New ErrTracer
    
    This.TestNumber = 0
    This.PreviousTestNumber = 0
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' prime all tests
    
    Set This.DelayStopper = cc_isr_Core_IO.Factory.NewStopwatch
    Set This.TestStopper = cc_isr_Core_IO.Factory.NewStopwatch
        
    ' open a connection to a new session.
    Set p_outcome = AssertPingDevice()
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Primed to run all tests; the device can be pingged.")
    Else
        Set p_outcome = Assert.Inconclusive( _
            "Failed priming all tests; The device could not be pingged; " & p_outcome.AssertMessage)
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed to run all tests.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Failed priming all tests;" & _
                VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeAllAssert = p_outcome
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Prepares each test before it is run. </summary>
Public Sub BeforeEach()

    Const p_procedureName As String = "BeforeEach"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.TestNumber = This.PreviousTestNumber Then _
        This.TestNumber = This.PreviousTestNumber + 1

    Dim p_outcome As cc_isr_Test_Fx.Assert

    If Not This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If

    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    ' Prepare the next test
    If This.BeforeAllAssert.AssertSuccessful Then
    
        Set This.Socket = cc_isr_Winsock.Factory.NewIPv4StreamSocket().Initialize()
        Set p_outcome = AssertShouldConnect(This.Socket)
    
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
             Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber))
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                ";" & VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeEachAssert = p_outcome

    On Error GoTo 0
    
    This.TestStopper.Restart
    
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases test elements after each tests is run. </summary>
Public Sub AfterEach()
    
    Const p_procedureName As String = "AfterEach"
    
    ' Trap errors to the error handler.
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")

    ' check if we can proceed with cleanup.
    
    If Not This.BeforeEachAssert.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to cleanup test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeEachAssert.AssertMessage)

    ' cleanup after each test.
        
    ' Release the socket
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(This.Socket, "Socket should not be nothing.")
        
    If p_outcome.AssertSuccessful Then
        ' socket might get disconnected in some tests
        If This.Socket.Connected Then _
            Set p_outcome = AssertShouldDisconnect(This.Socket)
    End If
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' record the previous test number
    This.PreviousTestNumber = This.TestNumber
    
    ' release the 'Before Each' cc_isr_Test_Fx.Assert.
    Set This.BeforeEachAssert = Nothing

    If p_outcome.AssertSuccessful Then
    
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Errors reported cleaning up test #" & VBA.CStr(This.TestNumber) & _
                ";" & VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    
    End If

    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases the test class after all tests run. </summary>
Public Sub AfterAll()
    
    Const p_procedureName As String = "AfterAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = cc_isr_Test_Fx.Assert.Pass("All tests cleaned up.")
    
    ' cleanup after all tests.

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before All' cc_isr_Test_Fx.Assert.
    Set This.BeforeAllAssert = Nothing

    ' report any leftover errors.
    Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Errors reported cleaning up all tests;" & _
            VBA.vbCrLf & p_outcome.AssertMessage)
    End If
    
    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Assert that the device can be pinged at the provided address. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertPingDevice() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "AssertOpenNewSession"

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Ready to ping the device.")
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Controller.TryPing(This.SocketAddress, This.SessionTimeout, p_details), _
            p_details)
     
    Set AssertPingDevice = p_outcome

End Function

Private Function AssertShouldValidateQuery(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket, _
    ByVal a_command As String, _
    ByVal a_value As String) As cc_isr_Test_Fx.Assert
    
    Dim p_elapsed As Double
    Dim p_stopper As cc_isr_Core_IO.Stopwatch
    Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_result As String
    Dim p_details As String
    Dim p_setCommand As String
    p_setCommand = a_command & " " & a_value
    p_stopper.Restart
    
    If This.Controller.TrySetValue(a_socket, a_command, a_value, p_details) Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass()
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    End If
    
    p_elapsed = p_stopper.ElapsedMilliseconds
    Set AssertShouldValidateQuery = p_outcome
    Debug.Print "    '" & p_setCommand & "' value set to " & p_result & _
        " in " & Format(p_elapsed, "0.0") & "ms."
End Function

Private Function AssertShouldConnect(ByRef a_socket As cc_isr_Winsock.IPv4StreamSocket) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    p_success = This.Controller.TryOpenConnection(This.SocketAddress, This.SessionTimeout, a_socket, p_details)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, p_details)
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(a_socket.Connected, _
            "Socket should be connected after opening a connection.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Controller.TryConfigureInitialState(a_socket, p_details), _
            p_details)
    
    Set AssertShouldConnect = p_outcome
    
End Function

Private Function AssertShouldDisconnect(ByVal a_socket As cc_isr_Winsock.IPv4StreamSocket) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(a_socket.Connected, _
        "Socket should be connected before disconnection.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(a_socket.TryCloseConnection(p_details), p_details)
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(a_socket.Connected, _
            "Socket should be disconnected after disconnection.")
    
    Set AssertShouldDisconnect = p_outcome
    
End Function


''' <summary>   Unit test. Asserts that the Controller should connect by checking a query. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' Test 01 TestShouldConnect passed. Elapsed time: 0.5 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldConnect() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldConnect"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Socket.Connected, "Socket would be connected after connecting")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertShouldDisconnect(This.Socket)
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Socket.Connected, "Socket would be disconnected after disconnecting.")
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldConnect = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts that the Controller should query a device identity. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' Test 02 TestShouldQueryIdentity passed. Elapsed time: 22.1 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldQueryIdentity() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldQueryIdentity"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String: p_command = "*IDN?"
    Dim p_sentCount As Integer
    Dim p_identity As String
    Dim p_readCount As Integer
    Dim p_reading As String
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then
            
        ' send the command
        If 0 >= This.Controller.TryWriteLine(This.Socket, p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    
    End If

    If p_outcome.AssertSuccessful Then
        
        If 0 > This.Controller.TryRead(This.Socket, p_identity, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
            1 = VBA.InStr(1, p_identity, This.IdentityCompany, VBA.VbCompareMethod.vbTextCompare), _
            "Identity '" & p_identity & " should start with '" & This.IdentityCompany & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldQueryIdentity = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts that the controller should synchronize (await operation completion). </summary>
''' <remarks>
''' <code>
''' With 1ms read sfter write delay.
''' Test 03 TestShouldSyncrhonize passed. Elapsed time: 37.7 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldSyncrhonize() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldSyncrhonize"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    
    Dim p_socket As cc_isr_Winsock.IPv4StreamSocket
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertShouldConnect(p_socket)
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Controller.TrySynchronize(p_socket, p_details), p_details)
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertShouldDisconnect(p_socket)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldSyncrhonize = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts that the stream socket should serial poll. </summary>
''' <remarks>
''' <code>
''' With 1ms read sfter write delay.
''' Test 10 TestShouldSerialPoll passed. Elapsed time: 30.1 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldSerialPoll() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldSerialPoll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    
    Dim p_command As String
    Dim p_sentCount As Integer
    
    If p_outcome.AssertSuccessful Then
            
        ' clear execution state, enable OPC Standard Event and Service Request on the standard event bit.
        p_command = "*CLS;*ESE 1;*SRE 32"
        If 0 >= This.Controller.TryWriteLine(This.Socket, p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
        If p_outcome.AssertSuccessful Then
        
            ' syncrhronize.
            p_command = "*OPC?"
            If 0 >= This.Controller.TryWriteLine(This.Socket, p_command, p_details) Then
                Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
            End If
        
        End If
    End If
    
    Dim p_reading As String
    If p_outcome.AssertSuccessful Then
        If 0 > This.Controller.TryRead(This.Socket, p_reading, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reading, _
            "Operation completion query should return the correct reply.")
        
    If p_outcome.AssertSuccessful Then
        
        p_command = "*OPC"
        If 0 >= This.Controller.TryWriteLine(This.Socket, p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
        
    ' wait for the operation completion bit.
    Dim p_stadnardEventBit As Integer
    p_stadnardEventBit = 32
    
    Dim p_requestingServiceBit As Integer
    p_requestingServiceBit = 64
    
    Dim p_timeout As Long: p_timeout = 500
    
    Dim p_statusByte As Integer
    If p_outcome.AssertSuccessful Then
        p_success = This.Controller.AwaitStatusBits(This.Socket, p_requestingServiceBit, p_requestingServiceBit, _
            p_timeout, p_statusByte, p_details)
        
        If Not p_success Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail( _
                "Status byte '" & VBA.CStr(p_statusByte) & _
                "' requesting service bit '" & VBA.CStr(p_requestingServiceBit) & "' should be set; " & p_details)
        
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldSerialPoll = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts that the controller should read the status byte. </summary>
''' <remarks>
''' <code>
''' With 1ms read sfter write delay.
''' Test 09 TestShouldReadStatusByte passed. Elapsed time: 34.0 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldReadStatusByte() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldReadStatusByte"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then
            
        ' clear execution state, enable OPC Standard Event and Service Request on the standard event bit.
        p_command = "*CLS;*ESE 1;*SRE 32"
        If 0 >= This.Controller.TryWriteLine(This.Socket, p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
        If p_outcome.AssertSuccessful Then
        
            ' syncrhronize.
            p_command = "*OPC?"
            If 0 >= This.Controller.TryWriteLine(This.Socket, p_command, p_details) Then
                Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
            End If
        
        End If
    End If
    
    Dim p_reading As String
    If p_outcome.AssertSuccessful Then
        If 0 > This.Controller.TryRead(This.Socket, p_reading, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reading, _
            "Operation completion query should return the correct reply.")
        
    If p_outcome.AssertSuccessful Then
        
        p_command = "*OPC"
        If 0 >= This.Controller.TryWriteLine(This.Socket, p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
        
    If p_outcome.AssertSuccessful Then
        
        p_command = "*STB?"
        If 0 >= This.Controller.TryWriteLine(This.Socket, p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
        
    If p_outcome.AssertSuccessful Then
        If 0 > This.Controller.TryRead(This.Socket, p_reading, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    Dim p_statusByte As Integer
    If p_outcome.AssertSuccessful Then
        If Not cc_isr_core.StringExtensions.TryParseInteger(p_reading, p_statusByte, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    ' wait for the operation completion bit.
    Dim p_stadnardEventBit As Integer
    p_stadnardEventBit = 32
    
    Dim p_requestingServiceBit As Integer
    p_requestingServiceBit = 64
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_requestingServiceBit, _
            p_requestingServiceBit And p_statusByte, _
            "Status byte '" & VBA.CStr(p_statusByte) & _
            "' requesting service bit 6 '" & VBA.CStr(p_requestingServiceBit) & "' should be set.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldReadStatusByte = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should recover from a Syntax error. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' 04 TestShouldRecoverFromSyntaxError passed. Elapsed time: 116.1 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRecoverFromSyntaxError() As Assert

    Const p_procedureName As String = "TestShouldRecoverFromSyntaxError"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_actualReply As String
    Dim p_expectedReply As String
    
    If p_outcome.AssertSuccessful Then
        
        ' issue a bad command
        On Error Resume Next
        This.Controller.WriteLine This.Socket, "**OPC"
        On Error GoTo 0
        
        ' clear the error state
        cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
        
        DoEvents
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 100
        
        p_expectedReply = "1"
        p_actualReply = This.Controller.QueryLine(This.Socket, "*CLS;*WAI;*OPC?")
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "Controller should query operation completion.")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldRecoverFromSyntaxError = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should restore initial state. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' Test 05 TestShouldRestoreInitialState passed. Elapsed time: 6835.3 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRestoreInitialState() As Assert

    Const p_procedureName As String = "TestShouldRestoreInitialState"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    If p_outcome.AssertSuccessful Then
        ' turn on auto assert TALK settings.
        This.Controller.AutoAssertTalkSetter This.Socket, True
        Set p_outcome = Assert.IsTrue(This.Controller.AutoAssertTalkGetter(This.Socket), _
            "GPIB-Lan controller Auto Assert Talk should be true.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_success = This.Controller.TryConfigureInitialState(This.Socket, p_details)
        Set p_outcome = Assert.IsTrue(p_success, p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        ' alter the read timeout
        This.Controller.ReadTimeoutSetter This.Socket, 2999
        p_success = This.Controller.ShouldRestoreInitialState(This.Socket, p_details)
        Set p_outcome = Assert.IsTrue(p_success, _
            "GPIB-Lan controller should need to restore state; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        p_success = This.Controller.TryConfigureInitialState(This.Socket, p_details)
        Set p_outcome = Assert.IsTrue(p_success, _
            "GPIB-Lan controller should configure initial state; " & p_details)
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldRestoreInitialState = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should recover from auto assert TALK condition. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' Test 06 TestShouldRecoverFromAutoAssertTalk passed. Elapsed time: 620.8 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRecoverFromAutoAssertTalk() As Assert

    Const p_procedureName As String = "TestShouldRecoverFromAutoAssertTalk"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_expectedAutoAssertTalk As Boolean
    Dim p_actualAutoAssertTalk As Boolean
    p_expectedAutoAssertTalk = True
    If p_outcome.AssertSuccessful Then
        ' turn on auto assert TALK condition.
        This.Controller.AutoAssertTalkSetter This.Socket, p_expectedAutoAssertTalk
        p_success = This.Controller.TryGetAutoAssertTalk(This.Socket, p_actualAutoAssertTalk, p_details)
        Set p_outcome = Assert.IsTrue(p_success, _
            "GPIB-Lan controller Auto Assert Talk should be " & VBA.CStr(p_expectedAutoAssertTalk) & ".")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Socket.CloseConnection
        Set p_outcome = Assert.IsFalse(This.Socket.Connected, _
            "GPIB-Lan controller socket should be disconnected.")
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertShouldConnect(This.Socket)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldRecoverFromAutoAssertTalk = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should Restore after closed connection. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' Test 07 TestShouldRestoreFromClosedConnection passed. Elapsed time: 14.9 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRestoreFromClosedConnection() As Assert

    Const p_procedureName As String = "TestShouldRestoreFromClosedConnection"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Socket.TryCloseConnection(p_details), _
            "Controller socket should disconnect; " & p_details)
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsFalse(This.Socket.Connected, "Controller should be disconnected.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Controller.ShouldConnect(This.Socket, p_details), _
            "Controller should report that the socket should connect after connection closed; " & p_details)
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Controller.ShouldRestoreInitialState(This.Socket, p_details), _
            "Controller should report that the socket should restore initial state; " & p_details)
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = AssertShouldConnect(This.Socket)
        
        If Not p_outcome.AssertSuccessful Then _
            Set p_outcome = Assert.Fail("Controller should be reconnected to restore initial state; " & _
                p_outcome.AssertMessage)
            
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldRestoreFromClosedConnection = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should read GPIB address. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' Test 08 TestShouldReadGpibAddress passed. Elapsed time: 4.6 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldReadGpibAddress() As Assert

    Const p_procedureName As String = "TestShouldReadGpibAddress"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_primary As Integer
    Dim p_secondary As Integer
    Dim p_reading As String
    If p_outcome.AssertSuccessful Then
        
        p_reading = This.Controller.GpibAddressGetter(This.Socket, p_primary, p_secondary)
        Set p_outcome = Assert.AreEqual(This.GpibAddress, p_reading, _
            "Controller should read GPIB address.")
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.PrimaryGpibAddress, p_primary, _
            "Controller method should return the expected primary GPIB address.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.PrimaryGpibAddress, This.Controller.PrimaryGpibAddress, _
            "Controller should parse the primary GPIB address.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.SecondaryGpibAddress, p_secondary, _
            "Controller method should return the expected secondary GPIB address.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.SecondaryGpibAddress, This.Controller.SecondaryGpibAddress, _
            "Controller should parse the secondary GPIB address.")
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldReadGpibAddress = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts that the Gpib Lan Controller should pwer on reset. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' 12:54:38 Power on reset starting. This could take 3 seconds. Please wait...
''' 12:54:41 done power on reset.
''' Test 16 TestShouldPowerOnReset passed. Elapsed time: 2841.0 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldPowerOnReset() As Assert

    Const p_procedureName As String = "TestShouldPowerOnReset"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    Dim p_details As String
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_delay As Double: p_delay = 3
        
        Debug.Print VBA.Format$(Now, "h:mm:ss"); " Power on reset starting. This could take "; _
            VBA.CStr(p_delay); " seconds. Please wait..."
        
        Dim p_success As Boolean
        p_success = This.Controller.TryPowerOnReset(This.SocketAddress, p_details, p_delay)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, p_details)
        
        Debug.Print VBA.Format$(Now, "h:mm:ss"); " done power on reset."
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldPowerOnReset = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function


''' <summary>   Unit test. Asserts Controller should toggle auto assert TALK condition. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldToggleAutoAssertTalk() As Assert

    Const p_procedureName As String = "TestShouldToggleAutoAssertTalk"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsFalse(This.Controller.AutoAssertTalkGetter(This.Socket), _
            "GPIB-Lan controller Auto Assert Talk should be False initially.")
    
    If p_outcome.AssertSuccessful Then
        This.Controller.AutoAssertTalkSetter This.Socket, True
        Set p_outcome = Assert.IsTrue(This.Controller.AutoAssertTalkGetter(This.Socket), _
            "GPIB-Lan controller Auto Assert Talk should be changed to True.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Controller.AutoAssertTalkSetter This.Socket, False
        Set p_outcome = Assert.IsFalse(This.Controller.AutoAssertTalkGetter(This.Socket), _
            "GPIB-Lan controller Auto Assert Talk should be be restored to False.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldToggleAutoAssertTalk = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should toggle auto assert TALK condition. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' Test 11 TestShouldToggleAppendTermination passed. Elapsed time: 210.4 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldToggleAppendTermination() As Assert

    Const p_procedureName As String = "TestShouldToggleAppendTermination"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_expectedValue As cc_isr_Ieee488.AppendTerminationOption
    p_expectedValue = cc_isr_Ieee488.AppendTerminationOption.AppendNothing
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(p_expectedValue, _
            This.Controller.AppendTerminationGetter(This.Socket), _
            "GPIB-Lan controller Append Termination Option should be 'Append Nothing' (3) initially.")
    
    p_expectedValue = cc_isr_Ieee488.AppendTerminationOption.LineFeed
    If p_outcome.AssertSuccessful Then
        This.Controller.AppendTerminationSetter This.Socket, p_expectedValue
        Set p_outcome = Assert.AreEqual(p_expectedValue, _
            This.Controller.AppendTerminationGetter(This.Socket), _
            "GPIB-Lan controller Auto Assert Talk should be changed to 'Line Feed' (2).")
    End If
    
    p_expectedValue = cc_isr_Ieee488.AppendTerminationOption.AppendNothing
    
    If p_outcome.AssertSuccessful Then
        This.Controller.AppendTerminationSetter This.Socket, p_expectedValue
        Set p_outcome = Assert.AreEqual(p_expectedValue, _
            This.Controller.AppendTerminationGetter(This.Socket), _
            "GPIB-Lan controller Auto Assert Talk should be restored to 'Append Nothing' (3) initially.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldToggleAppendTermination = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should toggle auto assert TALK condition. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' Test 12 TestShouldToggleEndOrIdentify passed. Elapsed time: 210.7 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldToggleEndOrIdentify() As Assert

    Const p_procedureName As String = "TestShouldToggleEndOrIdentify"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Controller.EndOrIdentifyGetter(This.Socket), _
            "GPIB-Lan controller End or Identify (EOI) should be true initially.")
    
    If p_outcome.AssertSuccessful Then
        This.Controller.EndOrIdentifySetter This.Socket, False
        Set p_outcome = Assert.IsFalse(This.Controller.EndOrIdentifyGetter(This.Socket), _
            "GPIB-Lan controller End or Identify (EOI) should be changed to False.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Controller.EndOrIdentifySetter This.Socket, True
        Set p_outcome = Assert.IsTrue(This.Controller.EndOrIdentifyGetter(This.Socket), _
            "GPIB-Lan controller End or Identify (EOI) should be be restored to True.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldToggleEndOrIdentify = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function


''' <summary>   Unit test. Asserts Controller should toggle auto assert TALK condition. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldToggleReadTimeout() As Assert

    Const p_procedureName As String = "TestShouldToggleReadTimeout"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.SessionTimeout, _
            This.Controller.ReadTimeoutGetter(This.Socket), _
            "GPIB-Lan controller read timeout should equal initial value.")
    
    Dim p_expectedTimeout As Long: p_expectedTimeout = 2999
    If p_outcome.AssertSuccessful Then
        This.Controller.ReadTimeoutSetter This.Socket, p_expectedTimeout
        Set p_outcome = Assert.AreEqual(p_expectedTimeout, _
            This.Controller.ReadTimeoutGetter(This.Socket), _
            "GPIB-Lan controller read timeout should equal modified value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Controller.ReadTimeoutSetter This.Socket, This.SessionTimeout
        Set p_outcome = Assert.AreEqual(This.SessionTimeout, _
            This.Controller.ReadTimeoutGetter(This.Socket), _
            "GPIB-Lan controller read timeout should equal restored initial value.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldToggleReadTimeout = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should go to local. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' Test 13 TestShouldGoToLocal passed. Elapsed time: 100.5 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldGoToLocal() As Assert

    Const p_procedureName As String = "TestShouldGoToLocal"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Controller.TryGoToLocal(This.Socket, p_details), _
            "GPIB-Lan controller should go to local; " & p_details)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldGoToLocal = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should local lockout. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' Test 14 TestShouldLocalLockout passed. Elapsed time: 100.5 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldLocalLockout() As Assert

    Const p_procedureName As String = "TestShouldLocalLockout"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Controller.TryLocalLockout(This.Socket, p_details), _
            "GPIB-Lan controller should go to local; " & p_details)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldLocalLockout = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts Controller should clear selective device. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' Test 15 TestShouldClearSelectiveDevice passed. Elapsed time: 11.3 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldClearSelectiveDevice() As Assert

    Const p_procedureName As String = "TestShouldClearSelectiveDevice"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Controller.TrySelectiveDeviceClear(This.Socket, p_details), _
            "GPIB-Lan controller should clear selective device; " & p_details)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldClearSelectiveDevice = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function



