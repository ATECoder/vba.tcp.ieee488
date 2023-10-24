# About

[cc.isr.tcp.Ieee488] is an Excel workbook for control and communication with an instrument that supports the IEEE 488.2 standard with commands such as `*IDN?` and `*CLS`.

## Note
An attempt was made to rename the project to cc_isr_Tcp_ieee488. This failed causing Excel to display the infamous [User-Defined Type Not Defined] error.

## Workbook references

* [cc.isr.Winsock] - Implements TCP Client and Server classes with Windows Winsock API.
* [cc.isr.Core] - Core work book.
* [cc.isr.core.io] - Core I/O workbook.

## Object Libraries references

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]
* [Microsoft VBScript Regular Expression 5.5]

## Key Features

* Provides commands and queries for communicating with IEEE488.2 instrument.
* Uses Windows Winsock32 calls to construct sockets for communicating with the instrument by way of a GPIB-Lan controller such as the [Prologix GPIB-Lan controller].
* Provides GPIB-Lan commands and queries for communicating with the GPIB-Lan controller.

## Main Types

The main types provided by this library are:

* _GpibLanController_ -- Communicates with the instrument by way of a GPIB-Lan controller.
* _ViSession_ -- Uses a _TcpCllient_ to communicate with the instrument by sending and receiving messages by way of the GPIB-Lan controller.
* _IEEE488Session_ -- Implements the core methods for communicating with an IEEE488.2 Instrument.

## Unit Testing

See [cc.isr.tcp.ieee488.test]

## Integration Testing

See [cc.isr.tcp.ieee488.demo]

## Feedback

[cc.isr.tcp.Ieee488] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.tcp.Ieee488] repository.

[cc.isr.tcp.ieee488]: https://github.com/ATECoder/vba.tcp.ieee488
[cc.isr.tcp.ieee488.test]: https://github.com/ATECoder/vba.tcp.ieee488/src/test
[cc.isr.tcp.ieee488.demo]: https://github.com/ATECoder/vba.tcp.ieee488/src/demo

[cc.isr.winsock]: https://github.com/ATECoder/vba.winsock/src/

[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.core.io]: https://github.com/ATECoder/vba.core/src/io
[cc.isr.test.fx]: https://github.com/ATECoder/vba.core/src/testfx

[unit test]: ./unit.test.lnk
[deploy]: ./deploy.ps1
[localize]: ./localize.ps1

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
[Microsoft VBScript Regular Expression 5.5]: <c:/windows/system32/vbscript.dll/3>

[User-Defined Type Not Defined]: https://stackoverflow.com/questions/19680402/compile-throws-a-user-defined-type-not-defined-error-but-does-not-go-to-the-of#:~:text=So%20the%20solution%20is%20to%20declare%20every%20referenced,objXML%20As%20Variant%20Set%20objXML%20%3D%20CreateObject%20%28%22MSXML2.DOMDocument%22%29
