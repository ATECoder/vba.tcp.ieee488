Attribute VB_Name = "TcpSessionTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Tcp Session Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   This class properties. </summary>
Private Type this_
    Name As String
    TestNumber As Integer
    PreviousTestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    Address As String
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
            Set p_outcome = TestShouldAwaitOperationCompletion
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
            Set p_outcome = TestGpibLanControllerShouldPowerOnReset
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
''' <remarks>
''' Test 01 TestShouldConnect passed. Elapsed time: 14.1 ms.
''' Test 02 TestShouldQueryIdentity passed. Elapsed time: 21.0 ms.
''' Test 03 TestShouldAwaitOperationCompletion passed. Elapsed time: 33.0 ms.
''' Test 04 TestShouldRecoverFromSyntaxError passed. Elapsed time: 117.5 ms.
''' Test 05 TestShouldRestoreInitialState passed. Elapsed time: 925.1 ms.
''' Test 06 TestShouldRecoverFromAutoAssertTalk passed. Elapsed time: 402.8 ms.
''' Test 07 TestShouldRestoreFromClosedConnection passed. Elapsed time: 56.0 ms.
''' Test 08 TestShouldReadGpibAddress passed. Elapsed time: 3.9 ms.
''' 9:12:09 Power on reset starting. This could take 6 seconds. Please wait...
''' 9:12:15 done power on reset.
''' Test 09 TestGpibLanControllerShouldPowerOnReset passed. Elapsed time: 5385.3 ms.
''' Ran 9 out of 9 tests.
''' Passed: 9; Failed: 0; Inconclusive: 0.
''' 9:20:59 Power on reset starting. This could take 3 seconds. Please wait...
''' 9:21:02 done power on reset.
''' Test 09 TestGpibLanControllerShouldPowerOnReset passed. Elapsed time: 2927.2 ms.
''' </remakrs>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 9
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
    
    This.Address = "192.168.0.252:1234"
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
    Set p_outcome = AssertOpenNewSession()
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Primed to run all tests; TCP Session is connected.")
    Else
        Set p_outcome = Assert.Inconclusive( _
            "Failed priming all tests; TCP Session should be connected; " & p_outcome.AssertMessage)
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

    If This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = IIf(This.Session.Connected, _
            cc_isr_Test_Fx.Assert.Pass("Ready to prime pre-test #" & VBA.CStr(This.TestNumber) & _
                "; IPV4 Stream Client is connected."), _
            cc_isr_Test_Fx.Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
                ";" & " IPV4 Stream Client should be connected"))
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If

    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    ' Prepare the next test
    
    Dim p_details As String: p_details = VBA.vbNullString
   
    If p_outcome.AssertSuccessful Then
        
        ' clear execution state before each test.
        ' clear errors if any so as to leave the instrument without errors.
        ' here we add *OPC? to prevent the query unterminated error.
        
        Dim p_command As String
        p_command = "*CLS;*WAI;*OPC?"
        If 0 >= This.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
    
    Dim p_reply As String
    If p_outcome.AssertSuccessful Then
        If 0 > This.Session.TryRead(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reply, _
            "Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            "; Operation completion query should return the correct reply.")
    
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
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_command As String
        Dim p_reply As String
        Dim p_details As String: p_details = VBA.vbNullString
    
        ' clear errors if any so as to leave the instrument without errors.
        p_command = "*CLS;*WAI;*OPC?"
        If 0 >= This.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
        If 0 > This.Session.TryRead(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
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
    
    ' disconnect if connected
    Dim p_details As String: p_details = VBA.vbNullString
    If Not This.Session Is Nothing Then
        If Not This.Session.Socket Is Nothing Then
            If Not This.Session.Socket.TryCloseConnection(p_details) Then
                Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
            End If
        End If
    End If
    Set This.Session = Nothing

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

''' <summary>   Construct a new session and open a connection. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertOpenNewSession() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "AssertOpenNewSession"

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Ready to construct a new session.")
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    
    Set This.Session = cc_isr_Ieee488.Factory.NewTcpSession()
    This.Session.Initialize cc_isr_Winsock.Factory.NewIPv4StreamSocket()
    This.Session.GpibLanControllerPort = 1234
    This.Session.Termination = VBA.vbLf
    This.Session.ReadAfterWriteDelay = 1
   
   
    If Not This.Session.Connectable.TryOpenConnection(This.Address, This.SessionTimeout, p_details) Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Session.Connected, _
            "The newly constructed session should be connected.")
    
    Dim p_expectedReply As String: p_expectedReply = "1"
    Dim p_actualReply As String
    If p_outcome.AssertSuccessful Then
        p_length = This.Session.TryQueryLine("*OPC?", p_actualReply, p_details)
        p_success = (0 < p_length) And (0 = VBA.Len(p_details))
        Set p_outcome = Assert.IsTrue(p_success, _
            "Operation complete query should succeed after making a new connection; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "Operation complete query reply should equal expected value after making a new connection.")
    End If
    
    Set AssertOpenNewSession = p_outcome

End Function

Private Function AssertShouldValidateQuery(ByVal a_command As String, _
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
    
    If This.Session.TrySetValue(a_command, a_value, p_details) Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass()
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    End If
    
    p_elapsed = p_stopper.ElapsedMilliseconds
    Set AssertShouldValidateQuery = p_outcome
    Debug.Print "    '" & p_setCommand & "' value set to " & p_result & _
        " in " & Format(p_elapsed, "0.0") & "ms."
End Function

''' <summary>   Unit test. Asserts that the session should connect by checking a query. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' TestShouldConnect passed. in 13.8 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldConnect() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldConnect"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_reply As String
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then
            
        ' check if connected and clear errors.
        p_command = "*CLS;*WAI;*OPC?"
        If 0 >= This.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        If 0 > This.Session.TryRead(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    
    End If
    
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

''' <summary>   Unit test. Asserts that the session should query a device identity. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' TestShouldQueryIdentity passed. in 21.5 ms.
''' TestShouldQueryIdentity passed. in 20.4 ms.
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
    Dim p_reply As String
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then
            
        ' send the command
        If 0 >= This.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    
    End If

    If p_outcome.AssertSuccessful Then
        
        If 0 > This.Session.TryRead(p_identity, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
            1 = VBA.InStr(1, p_identity, This.IdentityCompany, VBA.VbCompareMethod.vbTextCompare), _
            "Identity '" & p_identity & " should start with '" & This.IdentityCompany & "'.")

    End If

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

''' <summary>   Unit test. Asserts that the stream socket should await operation completion. </summary>
''' <remarks>
''' <code>
''' With 1ms read sfter write delay.
''' TestShouldAwaitOperationCompletion passed. in 33.2 ms.
''' TestShouldAwaitOperationCompletion passed. in 34.1 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldAwaitOperationCompletion() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldAwaitOperationCompletion"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then
            
        ' clear execution state, enable OPC Standard Event and Service Request on the standard event bit.
        p_command = "*CLS;*ESE 1;*SRE 32"
        If 0 >= This.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
        If p_outcome.AssertSuccessful Then
        
            ' syncrhronize.
            p_command = "*OPC?"
            If 0 >= This.Session.TryWriteLine(p_command, p_details) Then
                Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
            End If
        
        End If
    End If
    
    Dim p_reply As String
    If p_outcome.AssertSuccessful Then
        If 0 > This.Session.TryRead(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reply, _
            "Operation completion query should return the correct reply.")
        
    If p_outcome.AssertSuccessful Then
        
        p_command = "*OPC"
        If 0 >= This.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
        
    If p_outcome.AssertSuccessful Then
        
        p_command = "*STB?"
        If 0 >= This.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
        
    If p_outcome.AssertSuccessful Then
        If 0 > This.Session.TryRead(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    Dim p_statusByte As Integer
    If p_outcome.AssertSuccessful Then
        If Not cc_isr_core.StringExtensions.TryParseInteger(p_reply, p_statusByte, p_details) Then
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
    
    Set TestShouldAwaitOperationCompletion = p_outcome
    
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

''' <summary>   Unit test. Asserts session should recover from a Syntax error. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' TestShouldRecoverFromSyntaxError passed. in 116.9 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRecoverFromSyntaxError() As Assert

    Const p_procedureName As String = "TestShouldRecoverFromSyntaxError"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_actualReply As String
    Dim p_expectedReply As String
    
    If p_outcome.AssertSuccessful Then
        
        ' issue a bad command
        On Error Resume Next
        This.Session.WriteLine "**OPC"
        On Error GoTo 0
        
        ' clear the error state
        cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
        
        DoEvents
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 100
        
        p_expectedReply = "1"
        p_actualReply = This.Session.QueryLine("*CLS;*WAI;*OPC?")
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "Session should query operation completion.")
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

''' <summary>   Unit test. Asserts session should restore initial state. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' TestShouldRestoreInitialState passed. in 563.6 ms.
''' After adding timeout test we get:
''' TestShouldRestoreInitialState passed. in 887.2 ms.
''' Note that these values might explain the instability issues when the
''' read timeout was at 500 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRestoreInitialState() As Assert

    Const p_procedureName As String = "TestShouldRestoreInitialState"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    If p_outcome.AssertSuccessful Then
        ' turn on auto assert TALK settings.
        This.Session.AutoAssertTalkSetter True
        Set p_outcome = Assert.IsTrue(This.Session.AutoAssertTalkGetter(), _
            "GPIB-Lan controller Auto Assert Talk should be true.")
    End If
    
    Dim p_details As String
    If p_outcome.AssertSuccessful Then
        If Not This.Session.TryRestoreInitialState(p_details) Then _
            Set p_outcome = Assert.Fail(p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        ' alter the read timeout
        This.Session.ReadTimeoutSetter 2999
        Set p_outcome = Assert.IsTrue(This.Session.ShouldRestoreInitialState(p_details), _
            "GPIB-Lan controller should need to restore state; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        If Not This.Session.TryRestoreInitialState(p_details) Then _
            Set p_outcome = Assert.Fail(p_details)
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

''' <summary>   Unit test. Asserts session should recover from auto assert TALK condition. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' TestShouldRecoverFromAutoAssertTalk passed. in 380.6 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRecoverFromAutoAssertTalk() As Assert

    Const p_procedureName As String = "TestShouldRecoverFromAutoAssertTalk"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_actualReply As String
    Dim p_expectedReply As String
    
    If p_outcome.AssertSuccessful Then
        ' turn on auto assert TALK condition.
        This.Session.AutoAssertTalkSetter True
        Set p_outcome = Assert.IsTrue(This.Session.AutoAssertTalkGetter(), _
            "GPIB-Lan controller Auto Assert Talk should be true.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Session.Socket.CloseConnection
        Set p_outcome = Assert.IsFalse(This.Session.Connected, _
            "GPIB-Lan controller should be disconnected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Session.Initialize cc_isr_Winsock.Factory.NewIPv4StreamSocket
        This.Session.Socket.OpenConnection This.Address, This.SessionTimeout
        Set p_outcome = Assert.IsTrue(This.Session.Connected, _
            "GPIB-Lan controller should be connected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.Session.QueryLine("*OPC?")
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "Session should query operation completion.")
    End If

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

''' <summary>   Unit test. Asserts session should Restore after closed connection. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' TestShouldRestoreFromClosedConnection passed. in 83.5 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRestoreFromClosedConnection() As Assert

    Const p_procedureName As String = "TestShouldRestoreFromClosedConnection"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_actualReply As String
    Dim p_expectedReply As String
    Dim p_details As String
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(This.Session.Socket.TryCloseConnection(p_details), _
            "Session should disconnect; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsFalse(This.Session.Connected, _
            "Session should be disconnected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        ' This.Session.Initialize cc_isr_Winsock.Factory.NewIPv4StreamSocket
        Set p_outcome = Assert.IsTrue(This.Session.TryRestoreInitialState(p_details), _
            "Session should restore initial state after connection closed; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(This.Session.Connected, _
            "Session should be connected after restoring initial state.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.Session.QueryLine("*OPC?")
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "Session should query operation completion.")
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

''' <summary>   Unit test. Asserts session should read GPIB address. </summary>
''' <remarks>
''' With 1ms read sfter write delay.
''' TestShouldReadGpibAddress passed. in 1.4 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldReadGpibAddress() As Assert

    Const p_procedureName As String = "TestShouldReadGpibAddress"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_reading As String
    If p_outcome.AssertSuccessful Then
        p_reading = This.Session.GpibAddressGetter()
        Set p_outcome = Assert.AreEqual(This.GpibAddress, p_reading, _
            "Session should read GPIB address.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(This.PrimaryGpibAddress, This.Session.PrimaryGpibAddress, _
            "Session should parse the primary GPIB address.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(This.SecondaryGpibAddress, This.Session.SecondaryGpibAddress, _
            "Session should parse the secondary GPIB address.")
    End If
    
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
''' 9:20:59 Power on reset starting. This could take 3 seconds. Please wait...
''' 9:21:02 done power on reset.
''' Test 09 TestGpibLanControllerShouldPowerOnReset passed. Elapsed time: 2927.2 ms.
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestGpibLanControllerShouldPowerOnReset() As Assert

    Const p_procedureName As String = "TestGpibLanControllerShouldPowerOnReset"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    Dim p_details As String
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(This.Session.Socket.TryCloseConnection(p_details), _
            "Session should disconnect; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_delay As Double: p_delay = 3
        
        Debug.Print VBA.Format$(Now, "h:mm:ss"); " Power on reset starting. This could take "; _
            VBA.CStr(p_delay); " seconds. Please wait..."
        
        Dim p_success As Boolean
        p_success = This.Session.TryPowerOnReset(This.Session.SocketAddress, p_details, p_delay)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, p_details)
        
        Debug.Print VBA.Format$(Now, "h:mm:ss"); " done power on reset."
    End If
   
    '  open a connection to a new session
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertOpenNewSession()
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestGpibLanControllerShouldPowerOnReset = p_outcome
    
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



