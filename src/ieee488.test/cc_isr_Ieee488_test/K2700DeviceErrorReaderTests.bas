Attribute VB_Name = "K2700DeviceErrorReaderTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   IEEE488 Device Error Reader Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   This class properties. </summary>
Private Type this_
    Name As String
    TestNumber As Integer
    PreviousTestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    Device As cc_isr_Ieee488.Device
    Session As cc_isr_Ieee488.TcpSession
    Address As String
    SessionTimeout As Integer
    TestStopper As cc_isr_Core_IO.Stopwatch
    ErrTracer As cc_isr_Test_Fx.IErrTracer
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
            Set p_outcome = TestNoErrorShouldParse
        Case 2
            Set p_outcome = TestUndefinedHeaderErrorShouldParse
        Case Else
    End Select
    AfterEach
    Set RunTest = p_outcome
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 2
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
''' <remarks>
''' TestNoErrorShouldParse passed. in 35.4 ms.
''' TestNoErrorShouldParse passed. in 22.9 ms.
''' TestUndefinedHeaderErrorShouldParse passed. in 109.5 ms.
''' Ran 2 out of 2 tests.
''' Passed: 2; Failed: 0; Inconclusive: 0.
''' </remarks>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 2
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

    This.Name = "K2700DeviceErrorReaderTests"
    
    This.Address = "192.168.0.252:1234"
    This.SessionTimeout = 3000
    
    Set This.TestStopper = cc_isr_Core_IO.Factory.NewStopwatch
    Set This.ErrTracer = New ErrTracer
    
    This.TestNumber = 0
    This.PreviousTestNumber = 0
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' Prime all tests
    
    Set This.Device = cc_isr_Ieee488.Factory.NewDevice
    This.Device.GpibLanControllerPort = 1234
    This.Device.ReadAfterWriteDelay = 1
    This.Device.Termination = VBA.vbLf
    
    ' this also initializes the session.
    This.Device.Initialize
    Set This.Session = This.Device.Session
    
    Dim p_details As String
    If Not This.Session.Connectable.TryOpenConnection(This.Address, This.SessionTimeout, p_details) Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    End If
    
    If This.Device.Connected Then
        Set p_outcome = Assert.Pass("Primed to run all tests; IEEE488 Device is connected.")
    Else
        Set p_outcome = Assert.Inconclusive( _
            "Failed priming all tests; IEEE488 Device should be connected.")
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
    
    Set This.ErrTracer = Nothing
    
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

''' <summary>   Unit test. Asserts the device <c>No Error</c> should. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' TestNoErrorShouldParse passed. in 30.2 ms.
''' TestNoErrorShouldParse passed. in 22.8 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestNoErrorShouldParse() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestNoErrorShouldParse"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_success As Boolean
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsNotNothing(This.Device.DeviceErrorReader, _
            "Device Error Reader should be instantiated.")
    
    If p_outcome.AssertSuccessful Then
        
        p_success = This.Device.DeviceErrorReader.TryDequeueParseError(p_errorNumber, p_errorMessage)
        Set p_outcome = Assert.IsTrue(p_success, _
            "Device Error Reader should dequeue and parse the last device error.")

    End If

    Dim p_expectedErrorNumber As String: p_expectedErrorNumber = "0"
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorNumber, p_errorNumber, _
            "Device Error Reader should dequeue the 'No Error' error number.")

    End If

    Dim p_expectedErrorMessage As String: p_expectedErrorMessage = "No error"
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessage, p_errorMessage, _
            "Device Error Reader should dequeue the 'No Error' error message.")

    End If

    Dim p_actualErrorMessages As String
    Dim p_expectedErrorMessages As String: p_expectedErrorMessages = "0,No error"
    If p_outcome.AssertSuccessful Then
        
        p_success = Not This.Device.DeviceErrorReader.TryDequeueErrors(p_actualErrorMessages)
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessages, p_actualErrorMessages, _
            "Device Error Reader should dequeue the '0,No Error' error messages.")

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestNoErrorShouldParse = p_outcome
    
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


''' <summary>   Unit test. Asserts parsing device <c>UndefinedHeader</c> Error. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' TestNoErrorShouldParse passed. in 22.8 ms.
''' TestUndefinedHeaderErrorShouldParse passed. in 109.8 ms.
''' TestNoErrorShouldParse passed. in 22.8 ms.
''' TestUndefinedHeaderErrorShouldParse passed. in 109.6 ms.
''' TestNoErrorShouldParse passed. in 22.7 ms.
''' TestUndefinedHeaderErrorShouldParse passed. in 108.8 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestUndefinedHeaderErrorShouldParse() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestUndefinedHeaderErrorShouldParse"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_success As Boolean
    
    ' validate the existence of no errors.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = TestNoErrorShouldParse()
    
    If p_outcome.AssertSuccessful Then
    
        ' create and fetch the error.
        This.Session.WriteLine "**CLS"
        cc_isr_Core_IO.Factory.NewStopwatch.Wait 50
        p_success = This.Device.DeviceErrorReader.TryDequeueParseError(p_errorNumber, p_errorMessage)
        Set p_outcome = Assert.IsTrue(p_success, _
            "Device Error Reader should dequeue and parse the last device error.")

    End If

    Dim p_expectedErrorMessage As String: p_expectedErrorMessage = "Undefined header"
    Dim p_expectedErrorNumber As String: p_expectedErrorNumber = "-113"
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorNumber, p_errorNumber, _
            "Device Error Reader should dequeue the '" & _
            p_expectedErrorMessage & "' error number.")

    End If

    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessage, p_errorMessage, _
            "Device Error Reader should dequeue the expected error message.")

    End If

    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedReply As String: p_expectedReply = "1"
        Dim p_actualReply As String: p_actualReply = This.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "The Device shoudl query operation completion.")

    End If

    Dim p_actualErrorMessages As String
    Dim p_expectedErrorMessages As String: p_expectedErrorMessages = "0,No error"
    If p_outcome.AssertSuccessful Then
        
        p_success = Not This.Device.DeviceErrorReader.TryDequeueErrors(p_actualErrorMessages)
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessages, p_actualErrorMessages, _
            "Device Error Reader should dequeue the '0,No Error' error messages.")

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestUndefinedHeaderErrorShouldParse = p_outcome
   
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

