Attribute VB_Name = "MNew"
Option Explicit

Public Function Version(ByVal aMajor As Long, ByVal aMinor As Long, Optional ByVal aBuild As Long = -1, Optional ByVal aRevision As Long = -1) As Version
    Set Version = New Version: Version.New_ aMajor, aMinor, aBuild, aRevision
End Function
Public Function VersionS(ByVal VersionString As String) As Version
    Set VersionS = New Version: VersionS.NewS VersionString
End Function
Public Function VersionA() As Version
    Set VersionA = New Version: VersionA.NewA
End Function
Public Function VersionD(Optional DateValue) As Version
    Set VersionD = New Version: VersionD.NewD DateValue
End Function

Public Function FileVersionInfo(aPathFileName As String) As FileVersionInfo
    Set FileVersionInfo = New FileVersionInfo: FileVersionInfo.New_ aPathFileName
End Function

