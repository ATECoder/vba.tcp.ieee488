VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scpi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   SCPI extensions. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Checks if the specified command is query command. </summary>
''' <param name="a_command">   [String] the command. </param>
''' <returns>   [Boolean]. </returns>
Public Function IsQuery(ByVal a_command As String) As Boolean
    IsQuery = cc_isr_Core.StringExtensions.EndsWith(a_command, "?")
End Function

''' <summary>   Checks if the specified value is not a number per the SCPI standard. </summary>
''' <param name="a_value">   [String] the value. </param>
''' <returns>   [Boolean]. </returns>
Public Function IsNaN(ByVal a_value As String) As Boolean
    IsNaN = 1 = InStr(1, a_value, Syntax.NotANumberReading, vbTextCompare)
End Function

''' <summary>   Checks if the specified value is infinity per the SCPI standard. </summary>
''' <param name="a_value">   [String] the value. </param>
''' <returns>   [Boolean]. </returns>
Public Function IsInfinity(ByVal a_value As String) As Boolean
    IsInfinity = 1 = InStr(1, a_value, Syntax.InfinityReading, vbTextCompare)
End Function

''' <summary>   Checks if the specified value is negative infinity per the SCPI standard. </summary>
''' <param name="a_value">   [String] the value. </param>
''' <returns>   [Boolean]. </returns>
Public Function IsNegativeInfinity(ByVal a_value As String) As Boolean
    IsNegativeInfinity = 1 = InStr(1, a_value, Syntax.NegativeInfinityReading, vbTextCompare)
End Function

