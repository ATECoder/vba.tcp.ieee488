Attribute VB_Name = "DeviceTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Device Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    PreviousTestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
    Device As cc_isr_Ieee488.Device
    Session As cc_isr_Ieee488.TcpSession
    Address As String
    SessionTimeout As Integer
    TestStopper As cc_isr_Core_IO.Stopwatch
    ErrTracer As IErrTracer
    TestCount As Integer
    RunCount As Integer
    PassedCount As Integer
    FailedCount As Integer
    InconclusiveCount As Integer
End Type

Private This As this_

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Test runners
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Runs the specified test. </summary>
Public Function RunTest(ByVal a_testNumber As Integer) As cc_isr_Test_Fx.Assert
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.TestNumber = a_testNumber
    BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestShouldConnect
        Case 2
            Set p_outcome = TestShouldRecoverFromSyntaxError
        Case 3
            Set p_outcome = TestShouldRecoverFromAutoAssertTalk
        Case 4
            Set p_outcome = TestShouldRestoreInitialState
        Case 5
            Set p_outcome = TestShouldRestoreFromClosedConnection
        Case 6
            Set p_outcome = TestQueryUnterminatedErrorShouldRecover
        Case 7
            Set p_outcome = TestQueryInterruptedErrorShouldRecover
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 6
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
''' <remarks>
''' <code>
''' Test 01 TestShouldConnect passed. Elapsed time: 8.4 ms.
''' Test 02 TestShouldRecoverFromSyntaxError passed. Elapsed time: 109.2 ms.
''' Test 03 TestShouldRecoverFromAutoAssertTalk passed. Elapsed time: 1832.4 ms.
''' Test 04 TestShouldRestoreInitialState passed. Elapsed time: 2380.4 ms.
''' Test 05 TestShouldRestoreFromClosedConnection passed. Elapsed time: 2929.0 ms.
''' Reset clear try count =  1
''' Test 06 TestQueryUnterminatedErrorShouldRecover passed. Elapsed time: 4435.8 ms.
''' Reset clear try count =  1
''' Test 07 TestQueryInterruptedErrorShouldRecover passed. Elapsed time: 1439.9 ms.
''' Ran 7 out of 7 tests.
''' Passed: 7; Failed: 0; Inconclusive: 0.
''' </code>
''' </remarks>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 7
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Tests initialize and cleanup.
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Prepares all tests. </summary>
''' <remarks>   This method sets up the 'Before All' <see cref="cc_isr_Test_Fx.Assert"/>
''' which serves to set the 'Before Each' <see cref="cc_isr_Test_Fx.Assert"/>.
''' The error object and user defined errors state are left clear after this method. </remarks>
Public Sub BeforeAll()

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = Assert.Pass("Primed to run all tests.")

    This.Name = "DeviceTests"
    
    This.Address = "192.168.0.252:1234"
    This.SessionTimeout = 3000
    
    Set This.TestStopper = cc_isr_Core_IO.Factory.NewStopwatch
    Set This.ErrTracer = New ErrTracer
    
    This.TestNumber = 0
    This.PreviousTestNumber = 0
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState

    ' Prime all tests
    
    ' open a connection to a new session.
    Set p_outcome = AssertOpenNewSession()
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Primed to run all tests; IEEE488 is connected.")
    Else
        Set p_outcome = Assert.Inconclusive( _
            "Failed priming all tests; IEEE488 Device should be connected; " & p_outcome.AssertMessage)
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
            Set p_outcome = Assert.Pass("Primed to run all tests.")
        Else
            Set p_outcome = Assert.Inconclusive("Failed priming all tests;" & _
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
''' <remarks>   This method sets up the 'Before Each' <see cref="cc_isr_Test_Fx.Assert"/>
''' which serves to initialize the <see cref="cc_isr_Test_Fx.Assert"/> of each test.
''' The error object and user defined errors state are left clear after this method. </remarks>
Public Sub BeforeEach()

    Const p_procedureName As String = "BeforeEach"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.TestNumber = This.PreviousTestNumber Then _
        This.TestNumber = This.PreviousTestNumber + 1
    
    Dim p_outcome As cc_isr_Test_Fx.Assert

    If This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = IIf(This.Device.Connected, _
            Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber) & "; IEEE488 Device is Connected."), _
            Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                "; IEEE488 Device should be connected."))
    Else
        Set p_outcome = Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
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
             Set p_outcome = Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber))
        Else
            Set p_outcome = Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
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
''' <remarks>   This method uses the <see cref="ErrTracer"/> to report any leftover errors
''' in the user defined errors queue and stack. The error object and user defined errors
''' state are left clear after this method. </remarks>
Public Sub AfterEach()
    
    Const p_procedureName As String = "AfterEach"
    
    ' Trap errors to the error handler.
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")

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
''' <remarks>   This method uses the <see cref="ErrTracer"/> to report any leftover errors
''' in the user defined errors queue and stack. The error object and user defined errors
''' state are left clear after this method. </remarks>
Public Sub AfterAll()
    
    Const p_procedureName As String = "AfterAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = Assert.Pass("All tests cleaned up.")
    
    ' cleanup after all tests.
    If This.BeforeAllAssert.AssertSuccessful Then
    End If
    
    ' disconnect if connected
    If Not This.Device Is Nothing Then _
        This.Device.Dispose
    Set This.Session = Nothing
    Set This.Device = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before All' assert.
    Set This.BeforeAllAssert = Nothing

    ' report any leftover errors.
    Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = Assert.Inconclusive("Errors reported cleaning up all tests;" & _
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Tests
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

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
    
    Set This.Device = cc_isr_Ieee488.Factory.NewDevice
    This.Device.GpibLanControllerPort = 1234
    This.Device.ReadAfterWriteDelay = 1
    This.Device.Termination = VBA.vbLf
    
    ' this also initializes the session.
    This.Device.Initialize
    Set This.Session = This.Device.Session
   
    If Not This.Session.Connectable.TryOpenConnection(This.Address, This.SessionTimeout, p_details) Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Session.Connected, _
            "The newly constructed IEEE488 Device should be connected.")
    
    If p_outcome.AssertSuccessful Then
        Dim p_reply As String
        p_success = This.Device.TryQueryOperationCompleted(p_reply, p_details)
        Set p_outcome = Assert.IsTrue(p_success, _
            "The newly constructed IEEE488 Device should query operation completion; " & p_details)
    End If
    
    Set AssertOpenNewSession = p_outcome

End Function

''' <summary>   Unit test. Asserts that the session should connect. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldConnect() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldConnect"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_actualReply As String
    Dim p_expectedReply As String
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "IEEE488 Device should query operation completion.")

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

''' <summary>   Unit test. Asserts recovery from Syntax error. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' TestShouldRecoverFromSyntaxError passed. in 132.0 ms.
''' TestShouldRecoverFromSyntaxError passed. in 124.0 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRecoverFromSyntaxError() As Assert

    Const p_procedureName As String = "TestShouldRecoverFromSyntaxError"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_reply As String: p_reply = VBA.vbNullString
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        ' issue a bad command
        On Error Resume Next
        This.Session.WriteLine "**OPC"
        On Error GoTo 0
        
        ' clear the error state
        cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
        
        DoEvents
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 100
        
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = Assert.IsTrue(This.Device.TryQueryOperationCompleted(p_reply, p_details), _
                "IEEE488 device should query operation completion." & p_details)
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

''' <summary>   Unit test. Asserts device should restore initial state after Auto Assert
''' TALK condition. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' TestShouldRestoreInitialState passed. in 967.2 ms.
''' TestShouldRestoreInitialState passed. in 979.1 ms.
''' </code>
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
        ' turn on auto assert TALK condition.
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
        Set p_outcome = Assert.IsTrue(This.Device.ShouldRestoreInitialState(p_details), _
            "IEEE488 device should restore state; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(This.Device.ShouldRestoreInitialState(p_details), _
            "IEEE488 Device should need to restore initial state.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(This.Device.TryRestoreInitialState(p_details), _
            "IEEE488 Device should restore initial state; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Dim p_actualReply As String
        Dim p_expectedReply As String
        p_expectedReply = "1"
        p_actualReply = This.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "IEEE488 Device should query operation completion.")
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

''' <summary>   Unit test. Asserts the device should recover from auto assert TALK condition. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' TestShouldRecoverFromAutoAssertTalk passed. in 410.9 ms.
''' TestShouldRecoverFromAutoAssertTalk passed. in 397.4 ms.
''' TestShouldRecoverFromAutoAssertTalk passed. in 397.9 ms.
''' </code>
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
    
    If p_outcome.AssertSuccessful And This.Session.GpibLanControllerAttached Then
        ' turn on auto assert TALK condition.
        This.Session.AutoAssertTalkSetter True
        Set p_outcome = Assert.IsTrue(This.Session.AutoAssertTalkGetter, _
            "GPIB-Lan controller Auto Assert Talk should be true.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Session.Socket.CloseConnection
        Set p_outcome = Assert.IsFalse(This.Session.Connected, _
            "IEEE488 Device should be disconnected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Device.Initialize
        This.Session.Socket.OpenConnection This.Address, This.SessionTimeout
        Set p_outcome = Assert.IsTrue(This.Device.Connected, _
            "IEEE488 Device should be connected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "IEEE488 Device should query operation completion.")
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

''' <summary>   Unit test. Asserts the device should restore initial state from a closed connection. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' TestShouldRestoreFromClosedConnection passed. in 109.0 ms.
''' </code>
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
            "IEEE488 Device should disconnect; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsFalse(This.Session.Connected, _
            "IEEE488 Device should be disconnected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        'This.Device.Initialize
        Set p_outcome = Assert.IsTrue(This.Device.TryRestoreInitialState(p_details), _
            "IEEE488 Device should restore its initial state; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(This.Session.Connected, _
            "IEEE488 Device should be connected after restoring initial state.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "IEEE488 Device should query operation completion.")
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

''' <summary>   Unit test. Asserts the device should recover from query unterminated error. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' Reset clear try count =  1
''' Test 06 TestQueryUnterminatedErrorShouldRecover passed. Elapsed time: 4446.1 ms.
''' Reset clear try count =  1
''' Test 06 TestQueryUnterminatedErrorShouldRecover passed. Elapsed time: 4464.8 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestQueryUnterminatedErrorShouldRecover() As Assert

    Const p_procedureName As String = "TestQueryUnterminatedErrorShouldRecover"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_actualReply As String
    Dim p_details As String
    Dim p_success As Boolean
    Dim p_length As Long
    
    If p_outcome.AssertSuccessful Then
        ' create an unterminated error
        Dim p_command As String: p_command = "*OPC"
        p_length = This.Device.Session.TryQueryLine(p_command, p_actualReply, p_details)
        p_success = (0 < p_length) And (0 = VBA.Len(p_details))
        Set p_outcome = Assert.IsFalse(p_success, _
            "IEEE488 Device query should fail on unterminated error; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Dim p_tries As Integer: p_tries = 0
        p_success = This.Device.TryResetClears(p_details, p_tries, 3)
        If p_success Then _
            Debug.Print "Reset clear try count = "; p_tries
        Set p_outcome = Assert.IsTrue(This.Session.Connected, _
            "IEEE488 Device should reset and clear after query unterminated error; " & p_details)
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestQueryUnterminatedErrorShouldRecover = p_outcome
    
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

''' <summary>   Unit test. Asserts the device should recover from query Interrupted error. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' Reset clear try count =  1
''' Test 07 TestQueryInterruptedErrorShouldRecover passed. Elapsed time: 1438.2 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestQueryInterruptedErrorShouldRecover() As Assert

    Const p_procedureName As String = "TestQueryInterruptedErrorShouldRecover"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_details As String
    Dim p_success As Boolean
    Dim p_length As Long
    Dim p_command As String
    
    If p_outcome.AssertSuccessful Then
        
        ' send a query command
        p_command = "*OPC?"
        p_length = This.Device.Session.TryWriteLine(p_command, p_details)
        p_success = (0 < p_length) And (0 = VBA.Len(p_details))
        Set p_outcome = Assert.IsTrue(p_success, _
            "IEEE488 Device should receive the command #1; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        
        ' send the command again without fetching the reply
        p_length = This.Device.Session.TryWriteLine(p_command, p_details)
        p_success = (0 < p_length) And (0 = VBA.Len(p_details))
        Set p_outcome = Assert.IsTrue(p_success, _
            "IEEE488 Device should receive the command #2; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Dim p_tries As Integer: p_tries = 0
        p_success = This.Device.TryResetClears(p_details, p_tries, 3)
        If p_success Then _
            Debug.Print "Reset clear try count = "; p_tries
        Set p_outcome = Assert.IsTrue(This.Session.Connected, _
            "IEEE488 Device should reset and clear after query interrupted error; " & p_details)
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestQueryInterruptedErrorShouldRecover = p_outcome
    
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





