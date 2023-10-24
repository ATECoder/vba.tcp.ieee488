Attribute VB_Name = "Syntax"
Option Explicit

''' <summary>   Gets the SCPI reading for infinity. </summary>
Public Const InfinityReading As String = "9.90000E+37"

''' <summary>   Gets the SCPI value for infinity. </summary>
Public Const Infinity As Double = 9.9E+37

''' <summary>   Gets the SCPI reading for negative infinity. </summary>
Public Const NegativeInfinityReading As String = "-9.91000E+37"

''' <summary>   Gets the SCPI value for negative infinity. </summary>
Public Const NegativeInfinity As Double = -9.91E+37

''' <summary>   Gets the SCPI reading for 'not-a-number' (NAN). </summary>
Public Const NotANumberReading As String = "9.91000E+37"

''' <summary>   Gets the SCPI value for 'not-a-number' (NAN). </summary>
Public Const NotANumber As Double = 9.91E+37

''' <summary>   Gets the Clear Status (*CLS) command. </summary>
''' <remarks>
''' <see href="https://rfmw.em.keysight.com/spdhelpfiles/33500/webhelp/US/Content/__I_SCPI/IEEE-488_Subsystem.htm"/>
''' Clears the event registers in all register groups. Also clears the error queue.
''' </remarks>
Public Const ClearExecutionStateCommand As String = "*CLS"

''' <summary>   Gets the Identity query (*IDN?) command. </summary>
''' <remarks>
''' Identification string contains four comma separated fields:
''' Manufacturer name, Model Number, Serial Number, Revision Code.
''' </remarks>
Public Const IdentityQueryCommand As String = "*IDN?"

''' <summary>   Gets the operation complete (*OPC) command. </summary>
''' <remarks>   Sets "Operation Complete" (bit 0) in the Standard Event register at the completion
''' of the current operation.
''' The purpose of this command is to synchronize your application with the instrument.
''' Used in triggered sweep, triggered burst, list, or arbitrary waveform sequence modes to provide
''' a way to poll or interrupt the computer when the *TRG or INITiate[:IMMediate] is complete.
''' Other commands may be executed before Operation Complete bit is set.
''' The difference between *OPC and *OPC? is that *OPC? returns "1" to the output buffer when the
''' current operation completes. This means that no further commands can be sent after an *OPC? until
''' it has responded. In this way an explicit polling loop can be avoided. That is, the IO driver will
''' wait for the response.
''' </remarks>
Public Const OperationCompleteCommand As String = "*OPC"

''' <summary>   Gets the operation complete query (*OPC?) command. </summary>
''' <remarks>
''' Returns 1 to the output buffer after all pending commands complete.
''' The purpose of this command is to synchronize your application with the instrument.
''' Other commands cannot be executed until this command completes.
''' The difference between *OPC and *OPC? is that *OPC? returns "1" to the output buffer when the
''' current operation completes. This means that no further commands can be sent after an *OPC?
''' until it has responded. In this way an explicit polling loop can be avoided. That is, the IO
''' driver will wait for the response.
''' </remarks>
Public Const OperationCompletedQueryCommand As String = "*OPC?"

''' <summary>   Gets the options query (*OPT?) command. </summary>
''' <remarks> Returns a quoted string identifying any installed options. </remarks>
Public Const OptionsQueryCommand As String = "*OPT?"

''' <summary>   Power-On Status Clear (*PSC {0}). </summary>
''' <remarks>
''' Enables (1) or disables (0) clearing of two specific registers at power on:
''' Standard Event enable register (*ESE).
''' Status Byte condition register (*SRE).
''' Questionable Data Register
''' Standard Operation Register
''' This setting is non-volatile through a power-cycle. If it therefore useful for GPIB connection as follows:
''' <code>
''' *PSC 0 to disable enable clearing
''' *ESE 128 to enable power-on event
''' *SRE 32 to enable a SRQ on std event
''' </code>
''' This short program now provides a GPIB SRQ signal when the unit is turned on.
''' </remarks>
Public Const PowerOnStatusClearCommand As String = "*PSC {0}"

''' <summary>   Power-On Status Clear query command (*PSC?). </summary>
''' <remarks>
''' </remarks>
Public Const PowerOnStatusClearQueryCommand As String = "*PSC?"

''' <summary>   Gets the Standard Event Enable (*ESE {0}) command. </summary>
''' <remarks>
''' Event Status Enable Command and Query. Enables bits in the enable register for the Standard Event Register group. The selected bits are then reported to bit 5 of the Status Byte Register.
''' Use *PSC to control whether the Standard Event enable register is cleared at power on.
''' For example, *PSC 0 preserves the enable register contents through power cycles.
''' *CLS does not clear enable register, does clear event register.
''' </remarks>
Public Const StandardEventEnableCommand = "*ESE"

Public Const StandardEventEnableCommandFormat = "*ESE {0}"

''' <summary>   Gets the Standard Event Enable query (*ESE?) command. </summary>
''' <remarks> </remarks>
Public Const StandardEventEnableQueryCommand As String = "*ESE?"

''' <summary>   Gets the Standard Event Status (*ESR?) command. </summary>
''' <remarks>
''' Standard Event Status Register Query. Queries the event register for the Standard Event Register group.
''' Register is read-only; bits not cleared when read.
''' Any or all conditions can be reported to the Standard Event summary bit through the enable register.
''' To set the enable register mask, write a decimal value to the register using *ESE.
''' Once a bit is set, it remains set until cleared by this query or *CLS.
''' </remarks>
Public Const StandardEventStatusQueryCommand As String = "*ESR?"

''' <summary>   Gets the Service Request Enable (*SRE {0}) command. </summary>
''' <remarks>
''' Service Request Enable. This command enables bits in the enable register for the Status Byte Register group.
''' To enable specific bits, specify the decimal value corresponding to the binary-weighted sum of the bits in
''' the register. The selected bits are summarized in the "Master Summary" bit (bit 6) of the Status Byte Register.
''' If any of the selected bits change from 0 to 1, the instrument generates a Service Request signal.
''' *CLS clears the event register, but not the enable register.
''' *PSC (power-on status clear) determines whether Status Byte enable register is cleared at power on.
''' For example, *PSC 0 preserves the contents of the enable register through power cycles.
''' Status Byte enable register is not cleared by *RST.
''' </remarks>
Public Const ServiceRequestEnableCommand = "*SRE"

Public Const ServiceRequestEnableCommandFormat = "*SRE {0}"

''' <summary>   Gets the Standard Event and Service Request Enable '*CLS; *ESE {0}; *SRE {1}' command format. </summary>
Public Const StandardServiceEnableCommand = "*CLS; *ESE {0}; *SRE {1}"

''' <summary>   Gets the Service Request Enable query (*SRE?) command. </summary>
''' <remarks> </remarks>
Public Const ServiceRequestEnableQueryCommand As String = "*SRE?"

''' <summary>   Gets the Service Request Status query (*STB?) command. </summary>
''' <remarks>
''' Similar to a Serial Poll, but processed like any other instrument command.
''' Register is read-only; bits not cleared when read.
''' Returns same result as a Serial Poll, but "Master Summary" bit (bit 6) is not cleared by *STB?.
''' Power cycle or *RST clears all bits in condition register.
''' Returns a decimal value that corresponds to the binary-weighted sum of all bits set in the register.
''' For example, with bit 3 (decimal 8) and bit 5 (decimal 32) set (and corresponding bits enabled),
''' the query returns +40.
''' </remarks>
Public Const ServiceRequestQueryCommand As String = "*STB?"

''' <summary>   Gets the reset to known state (*RST) command. </summary>
''' <remarks>
''' Resets instrument to factory default state, independent of MEMory:STATe:RECall:AUTO setting.
''' Does not affect stored instrument states, stored arbitrary waveforms, or I/O settings;
''' these are stored in non-volatile memory.
''' Aborts a sweep or burst in progress.
''' </remarks>
Public Const ResetKnownStateCommand As String = "*RST"

''' <summary>   Trigger command (*TRG). </summary>
''' <remarks>
''' Trigger Command. Triggers a sweep, burst, arbitrary waveform advance, or LIST advance from the
''' remote interface if the bus (software) trigger source is currently selected (TRIGger[1|2]:SOURce BUS).
''' <remarks>
Public Const TriggerCommand As String = "*TRG"

''' <summary>   Self test query command (*TST?). </summary>
''' <remarks>
''' Self-Test Query. Performs a complete instrument self-test. If test fails, one or more error messages
''' will provide additional information. Use SYSTem:ERRor? to read error queue.
''' A power-on self-test occurs when you turn on the instrument. This limited test assures you that the
''' instrument is operational.
''' A complete self-test (*TST?) takes approximately 15 seconds. If all tests pass, you have high confidence
''' that the instrument is fully operational.
''' Passing *TST displays "Self-Test Passed" on the front panel. Otherwise, it displays "Self-Test Failed"
''' and an error number. See Service and Repair - Introduction for instructions on contacting support or
''' returning the instrument for service.
''' </remarks>
Public Const SelfTestQueryQueryCommand As String = "*TST?"

''' <summary>   Gets the Wait (*WAI) command. </summary>
''' <remarks> Configures the instrument to wait for all pending operations to complete before executing any
''' additional commands over the interface.
''' For example, you can use this with the *TRG command to ensure that the instrument is ready for a trigger:
''' <code>
''' *TRG;*WAI;*TRG
''' </code>
''' </remarks>
Public Const WaitCommand As String = "*WAI"

''' <summary>   Enumerates the status byte flags of the standard event register. </summary>
''' <remarks>
''' Enumerates the Standard Event Status Register Bits. Read this information using ESR? or
''' status.standard.event. Use *ESE or status.standard.enable or event status enable to enable
''' this register.''' These values are used when reading or writing to the standard event
''' registers. Reading a status register returns a value. The binary equivalent of the returned
''' value indicates which register bits are set. The least significant bit of the binary number
''' is bit 0, and the most significant bit is bit 15. For example, assume value 9 is returned for
''' the enable register. The binary equivalent is
''' 0000000000001001. This value indicates that bit 0 (OPC) and bit 3 (DDE)
''' are set.
''' </remarks>
Public Enum StandardEventFlags

    ''' <summary>   The None option. </summary>
    None = 0

    ''' <summary>
    ''' Bit B0, Operation Complete (OPC). Set bit indicates that all
    ''' pending selected device operations are completed and the unit is ready to
    ''' accept new commands. The bit is set in response to an *OPC command.
    ''' The ICL function OPC() can be used in place of the *OPC command.
    ''' </summary>
    OperationComplete = 1

    ''' <summary>
    ''' Bit B1, Request Control (RQC). Set bit indicates that....
    ''' </summary>
    RequestControl = &H2

    ''' <summary>
    ''' Bit B2, Query Error (QYE). Set bit indicates that you attempted
    ''' to read data from an empty Output Queue.
    ''' </summary>
    QueryError = &H4

    ''' <summary>
    ''' Bit B3, Device-Dependent Error (DDE). Set bit indicates that a
    ''' device operation did not execute properly due to some internal
    ''' condition.
    ''' </summary>
    DeviceDependentError = &H8

    ''' <summary>
    ''' Bit B4 (16), Execution Error (EXE). Set bit indicates that the unit
    ''' detected an error while trying to execute a command.
    ''' This is used by QUATECH to report No Contact.
    ''' </summary>
    ExecutionError = &H10

    ''' <summary>
    ''' Bit B5 (32), Command Error (CME). Set bit indicates that a
    ''' command error has occurred. Command errors include:<p>
    ''' IEEE-488.2 syntax error — unit received a message that does not follow
    ''' the defined syntax of the IEEE-488.2 standard.  </p><p>
    ''' Semantic error — unit received a command that was misspelled or received
    ''' an optional IEEE-488.2 command that is not implemented.  </p><p>
    ''' The device received a Group Execute Trigger (GET) inside a program
    ''' message.  </p>
    ''' </summary>
    CommandError = &H20

    ''' <summary>
    ''' Bit B6 (64), User Request (URQ). Set bit indicates that the LOCAL
    ''' key on the SourceMeter front panel was pressed.
    ''' </summary>
    UserRequest = &H40

    ''' <summary>
    ''' Bit B7 (128), Power ON (PON). Set bit indicates that the device
    ''' has been turned off and turned back on since the last time this register
    ''' has been read.
    ''' </summary>
    PowerToggled = &H80

    ''' <summary>
    ''' Unknown value due to, for example, error trying to get value from the device.
    ''' </summary>
    Unknown = &H100

    ''' <summary>Includes all bits. </summary>
    all = &HFF ' 255

End Enum

''' <summary>   Gets or sets the status byte bits of the service request register. </summary>
''' <remarks>
''' Enumerates the Status Byte Register Bits. Use *STB? or status.request_event to read this
''' register. Use *SRE or status.request_enable to enable these services. This attribute is used
''' to read the status byte, which is returned as a numeric value. The binary equivalent of the
''' returned value indicates which register bits are set. <para>
''' (c) 2005 Integrated Scientific Resources, Inc. All rights reserved. </para><para>
''' Licensed under The MIT License. </para>
''' </remarks>
Public Enum ServiceRequestFlags

    ''' <summary>   The None option. </summary>
    None = 0

    ''' <summary>
    ''' Bit B0, Measurement Summary Bit (MSB). Set summary bit indicates
    ''' that an enabled measurement event has occurred.
    ''' </summary>
    MeasurementEvent = &H1

    ''' <summary>
    ''' Bit B1, System Summary Bit (SSB). Set summary bit indicates
    ''' that an enabled system event has occurred.
    ''' </summary>
    SystemEvent = &H2

    ''' <summary>
    ''' Bit B2, Error Available (EAV). Set summary bit indicates that
    ''' an error or status message is present in the Error Queue.
    ''' </summary>
    ErrorAvailable = &H4

    ''' <summary>
    ''' Bit B3, Questionable Summary Bit (QSB). Set summary bit indicates
    ''' that an enabled questionable event has occurred.
    ''' </summary>
    QuestionableEvent = &H8

    ''' <summary>
    ''' Bit B4 (16), Message Available (MAV). Set summary bit indicates that
    ''' a response message is present in the Output Queue.
    ''' </summary>
    MessageAvailable = &H10

    ''' <summary>Bit B5, Event Summary Bit (ESB). Set summary bit indicates
    ''' that an enabled standard event has occurred.
    ''' </summary>
    standardEvent = &H20 ' (32) ESB

    ''' <summary>
    ''' Bit B6 (64), Request Service (RQS)/Master Summary Status (MSS).
    ''' Set bit indicates that an enabled summary bit of the Status Byte Register
    ''' is set. Depending on how it is used, Bit B6 of the Status Byte Register
    ''' is either the Request for Service (RQS) bit or the Master Summary Status
    ''' (MSS) bit: When using the GPIB serial poll sequence of the unit to obtain
    ''' the status byte (serial poll byte), B6 is the RQS bit. When using
    ''' status.condition or the *STB? common command to read the status byte,
    ''' B6 is the MSS bit.
    ''' </summary>
    RequestingService = &H40

    ''' <summary>
    ''' Bit B7 (128), Operation Summary (OSB). Set summary bit indicates that
    ''' an enabled operation event has occurred.
    ''' </summary>
    OperationEvent = &H80

    ''' <summary>
    ''' Includes all bits.
    ''' </summary>
    all = &HFF ' 255

    ''' <summary>
    ''' Unknown value due to, for example, error trying to get value from the device.
    ''' </summary>
    Unknown = &H100

End Enum

