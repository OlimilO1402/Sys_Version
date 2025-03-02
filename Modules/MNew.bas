Attribute VB_Name = "MNew"
Option Explicit

Public Function version(ByVal aMajor As Long, ByVal aMinor As Long, Optional ByVal aBuild As Long = -1, Optional ByVal aRevision As Long = -1) As version
    Set version = New version: version.New_ aMajor, aMinor, aBuild, aRevision
End Function

Public Function VersionS(ByVal VersionString As String) As version
    Set VersionS = New version: VersionS.NewS VersionString
End Function

Public Function VersionA() As version
    Set VersionA = New version: VersionA.NewA
End Function

Public Function FileVersionInfo(aPathFileName As String) As FileVersionInfo
    Set FileVersionInfo = New FileVersionInfo: FileVersionInfo.New_ aPathFileName
End Function

