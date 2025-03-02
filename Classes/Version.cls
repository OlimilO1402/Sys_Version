VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Version"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TInt16HiLo
    Lo As Integer
    Hi As Integer
End Type

Private Type TLong
    Value As Long
End Type

Private Type TVersion
    Major    As Long
    Minor    As Long
    Build    As Long
    Revision As TLong
End Type

Private m_Ver As TVersion

Private Sub Class_Initialize()
    m_Ver.Build = -1: m_Ver.Revision.Value = -1
End Sub

Friend Sub New_(aMajor As Long, aMinor As Long, Optional aBuild As Long = -1, Optional aRevision As Long = -1)
    m_Ver.Major = aMajor: m_Ver.Minor = aMinor: m_Ver.Build = IIf(aBuild < 0, -1, aBuild): m_Ver.Revision.Value = IIf(aRevision < 0, -1, aRevision)
End Sub
Friend Sub NewS(VersionString As String)
    Dim sa() As String: sa = Split(VersionString, ".")
    Dim u As Long: u = UBound(sa)
    If u >= 0 Then m_Ver.Major = CLng(sa(0))
    If u >= 1 Then m_Ver.Minor = CLng(sa(1))
    If u >= 2 Then m_Ver.Build = CLng(sa(2))
    If u >= 3 Then m_Ver.Revision.Value = CLng(sa(3))
End Sub
Friend Sub NewA()
    New_ App.Major, App.Minor, 0, App.Revision
End Sub

'Copy to MNew:
'Public Function Version(ByVal aMajor As Long, ByVal aMinor As Long, Optional ByVal aRevision As Long = -1, Optional ByVal aBuild As Long = -1) As Version
'    Set Version = New Version: Version.New_ aMajor, aMinor, aRevision, aBuild
'End Function
'Public Function VersionS(ByVal VersionString As String) As Version
'    Set VersionS = New Version: VersionS.NewS VersionString
'End Function
'Public Function VersionA() As Version
'    Set VersionA = New Version: VersionA.NewA
'End Function

Friend Sub NewC(Other As Version)
    With Other: m_Ver.Major = .Major: m_Ver.Minor = .Minor: m_Ver.Build = .Build: m_Ver.Revision.Value = .Revision: End With
End Sub

Public Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_Ver)
End Property

Public Function Clone() As Version
    Set Clone = New Version: Clone.NewC Me
End Function

Public Function Equals(Other As Version) As Boolean
    If Me.Major <> Other.Major Then Exit Function
    If Me.Minor <> Other.Minor Then Exit Function
    If Me.Build <> Other.Build Then Exit Function
    If Me.Revision <> Other.Revision Then Exit Function
    Equals = True
End Function

Private Function IsGreaterOrEqual(Version As Version, Other As Version) As Boolean
    If Version.Major < Other.Major Then Exit Function
    If Version.Minor < Other.Minor Then Exit Function
    If Version.Build < Other.Build Then Exit Function
    If Version.Revision < Other.Revision Then Exit Function
    IsGreaterOrEqual = True
End Function

Public Function IsLessThen(Other As Version) As Boolean
    IsLessThen = IsGreater(Other, Me)
End Function

Public Function IsLessThenOrEqual(Other As Version) As Boolean
    IsLessThenOrEqual = IsGreaterOrEqual(Other, Me)
End Function

Public Function IsGreaterThen(Other As Version) As Boolean
    IsGreaterThen = IsGreater(Me, Other)
End Function

Public Function IsGreaterThenOrEqual(Other As Version) As Boolean
    IsGreaterThenOrEqual = IsGreaterOrEqual(Me, Other)
End Function

Private Function IsGreater(Version As Version, Other As Version) As Boolean
    If Not IsGreaterOrEqual(Version, Other) Then Exit Function
    IsGreater = Not Version.Equals(Other)
End Function

Public Function CompareTo(Other As Version) As Long
    If Me.Equals(Other) Then Exit Function
    If Me.IsLessThen(Other) Then CompareTo = -1: Exit Function
    If Me.IsGreaterThen(Other) Then CompareTo = 1: Exit Function
End Function

Friend Function ToStr() As String
    ToStr = m_Ver.Major & "." & m_Ver.Minor & IIf(m_Ver.Build = -1, "", "." & m_Ver.Build) & IIf(m_Ver.Revision.Value = -1, "", "." & m_Ver.Revision.Value)
End Function

Public Property Get Major() As Long
    Major = m_Ver.Major
End Property
Friend Property Let Major(ByVal Value As Long)
    m_Ver.Major = Value
End Property

Public Property Get Minor() As Long
    Minor = m_Ver.Minor
End Property
Friend Property Let Minor(ByVal Value As Long)
    m_Ver.Minor = Value
End Property

Public Property Get Build() As Long
    Build = m_Ver.Build
End Property
Friend Property Let Build(ByVal Value As Long)
    m_Ver.Build = Value
End Property

Public Property Get Revision() As Long
    Revision = m_Ver.Revision.Value
End Property
Friend Property Let Revision(ByVal Value As Long)
    m_Ver.Revision.Value = Value
End Property

Public Property Get MajorRevision() As Integer
    Dim thilo As TInt16HiLo: LSet thilo = m_Ver.Revision
    MajorRevision = thilo.Hi
End Property

Public Property Get MinorRevision() As Integer
    Dim thilo As TInt16HiLo: LSet thilo = m_Ver.Revision
    MinorRevision = thilo.Lo
End Property

