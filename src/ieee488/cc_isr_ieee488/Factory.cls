VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' <summary>   Instantiates the<see cref="Device"/> class to
''' share with external workbooks. </summary>
''' <returns>   [<see cref="Device"/>]. </returns>
Public Function NewDevice() As Device
    Set NewDevice = New Device
End Function

''' <summary>   Instantiates the<see cref="TcpSession"/> class to
''' share with external workbooks. </summary>
''' <returns>   [<see cref="TcpSession"/>]. </returns>
Public Function NewTcpSession() As TcpSession
    Set NewTcpSession = New TcpSession
End Function

