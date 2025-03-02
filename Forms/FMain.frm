VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   10215
   End
   Begin VB.CommandButton BtnTestVersion 
      Caption         =   "Test Class Version"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_IndentStack As Byte

Private Sub Form_Resize()
    Dim L As Single, t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Text1.Move L, t, W, H
End Sub

Private Sub IndentStack_Push()
    m_IndentStack = m_IndentStack + 1
End Sub

Private Sub IndentStack_Pop()
    m_IndentStack = m_IndentStack - 1
End Sub

Private Sub BtnTestVersion_Click()
    DebugPrint "Test Class Version"
    DebugPrint "=================="
    DebugPrint ""
    
    IndentStack_Push
    TestCtors
    
    TestComparing
    IndentStack_Pop
    
End Sub

Sub TestCtors()
    
    DebugPrint "Test Constructors"
    DebugPrint "-----------------"
    IndentStack_Push
    
    Dim ver As Version
    
    Set ver = New Version
    DebugPrint ver.ToStr '0.0.-1.-1
    DebugPrint ver.Major & "." & ver.Minor & "." & ver.Build & "." & ver.Revision & "." & ver.MajorRevision & "." & ver.MinorRevision   '0.0.-1.-1.-1.-1
    
    Set ver = MNew.Version(1, 2, 3, 4)
    DebugPrint ver.ToStr '1.2.3.4
    DebugPrint ver.Major & "." & ver.Minor & "." & ver.Build & "." & ver.Revision & "." & ver.MajorRevision & "." & ver.MinorRevision   '1.2.3.4.0.4
    
    Set ver = MNew.VersionS("1.2.3.4")
    DebugPrint ver.ToStr '1.2.3.4
    DebugPrint ver.Major & "." & ver.Minor & "." & ver.Build & "." & ver.Revision & "." & ver.MajorRevision & "." & ver.MinorRevision '1.2.3.4.0.4
    
    Set ver = MNew.Version(1, 2, &H1234, &H43215678)
    DebugPrint ver.ToStr '1.2.4660.1126258296.17185.22136
    DebugPrint ver.Major & "." & ver.Minor & "." & ver.Build & "." & ver.Revision & "." & ver.MajorRevision & "." & ver.MinorRevision '1.2.4660.1126258296.17185.22136
    
    Set ver = MNew.VersionA
    DebugPrint ver.ToStr '2025.3.1
    DebugPrint ver.Major & "." & ver.Minor & "." & ver.Build & "." & ver.Revision & "." & ver.MajorRevision & "." & ver.MinorRevision '2025.3.1.0.1
    
    DebugPrint ""
    
    IndentStack_Pop
    
End Sub

Sub TestComparing()
    
    DebugPrint "Test Comparing"
    DebugPrint "--------------"
    IndentStack_Push

    Dim v1 As Version, v2 As Version
    
    Set v1 = MNew.Version(2025, 3, 1, 1): Set v2 = v1.Clone
    DoAllComparings v1, v2
    
    v2.Revision = v2.Revision + 1
    DoAllComparings v1, v2
    
    v1.Revision = v2.Revision + 1
    DoAllComparings v1, v2
        
    IndentStack_Pop
End Sub

Sub DoAllComparings(v1 As Version, v2 As Version)
    If v1.Equals(v2) Then DebugPrint "v1(" & v1.ToStr & ") = v2(" & v2.ToStr & ")"
    If Not v1.Equals(v2) Then DebugPrint "v1(" & v1.ToStr & ") <> v2(" & v2.ToStr & ")"
    If v1.IsLessThen(v2) Then DebugPrint "v1(" & v1.ToStr & ") < v2(" & v2.ToStr & ")"
    If v1.IsGreaterThen(v2) Then DebugPrint "v1(" & v1.ToStr & ") > v2(" & v2.ToStr & ")"
    If v1.IsLessThenOrEqual(v2) Then DebugPrint "v1(" & v1.ToStr & ") <= v2(" & v2.ToStr & ")"
    If v1.IsGreaterThenOrEqual(v2) Then DebugPrint "v1(" & v1.ToStr & ") >= v2(" & v2.ToStr & ")"
    Dim c As Long: c = v1.CompareTo(v2): DebugPrint "v1.CompareTo(v2) = " & c
    DebugPrint ""
End Sub

Sub DebugPrint(ByVal s As String)
    s = Space(m_IndentStack * 2) & s
    Text1.Text = Text1.Text & s & vbCrLf
End Sub

'Easy Comparing
'==============
'
'for comparing e.g. two numbers in any serious programming language, there are the typical operator characters of course everybody knows.
'like the following:
' VB    C#    meaning
'----------------------------------
'* =    ==    Equality
'* <>   !=    Not Equal
'* <=   <=    Less then or equal
'* >=   >=    Greater then or equal
'* <    <     Less then
'* >    >     Greater then
'** special: int CompareTo(other)
'
'for comparing two objects (of a class), in other languages, there is something called operator overloading.
'What just means you can write a function that has an operator-character as the "function name".
'In VBA/VBC we do not have operator overloading, but we do not bother, we even do not need this.
'It is just "syntactic sugar" to make code more readable, but imho it does not fulfill it's purpose in every
'situation. In fact writing named member-functions is readable enough for comparing two objects.
'
'So have a look at the list above. Do we need a function for every operator, for every possible comparison?
'Yes, we maybe actually need all the above functions, but did you know that we actually need only 2 functions,
'and all the other operations are just a combination of that two functions?
'
'In VBA/VBC we dim a Boolean and per se the boolean has the value "False". VB does this for us, so there is no
'need for an extra initialization of a Boolean variable or also even a Boolean function.
'
'the 2 functions we need are
'* a public member function "Equals" where we just hand over the "other" object and
'* a private function "IsGreaterOrEqual" where we give two objects;
'  this functions could also be static/shared in a standard module ...
'
'...and all the other operator-functions are just combinations of this two functions.
'
'To give this something what actually makes sense we could imagine a class "Version" with the
'member properties Major, Minor, Build And Revision
'(compare: Version class https://learn.microsoft.com/en-us/dotnet/api/system.version?view=net-8.0 )
'Maybe we have a situation where we have different versions of a file or a program, and in our program
'we want to react on it
'
'here are the 2 main full-size functions we need:
'
'Public Function Equals(Other As Version) As Boolean
'    If Me.Major <> Other.Major Then Exit Function
'    If Me.Minor <> Other.Minor Then Exit Function
'    If Me.Build <> Other.Build Then Exit Function
'    If Me.Revision <> Other.Revision Then Exit Function
'    Equals = True
'End Function
'
'Private Function IsGreaterOrEqual(Version As Version, Other As Version) As Boolean
'    If Version.Major < Other.Major Then Exit Function
'    If Version.Minor < Other.Minor Then Exit Function
'    If Version.Build < Other.Build Then Exit Function
'    If Version.Revision < Other.Revision Then Exit Function
'    IsGreaterOrEqual = True
'End Function
'
'and here are the very slim functions for all other comparisons:
'
'Public Function IsLessThen(Other As Version) As Boolean
'    IsLessThen = IsGreater(Other, Me)
'End Function
'
'Public Function IsLessThenOrEqual(Other As Version) As Boolean
'    IsLessThenOrEqual = IsGreaterOrEqual(Other, Me)
'End Function
'
'Public Function IsGreaterThen(Other As Version) As Boolean
'    IsGreaterThen = IsGreater(Me, Other)
'End Function
'
'Public Function IsGreaterThenOrEqual(Other As Version) As Boolean
'    IsGreaterThenOrEqual = IsGreaterOrEqual(Me, Other)
'End Function
'
'Private Function IsGreater(Version As Version, Other As Version) As Boolean
'    If Not IsGreaterOrEqual(Version, Other) Then Exit Function
'    IsGreater = Not Version.Equals(Other)
'End Function
'
'Public Function CompareTo(Other As Version) As Long
'    If Me.Equals(Other) Then Exit Function
'    If Me.IsLessThen(Other) Then CompareTo = -1: Exit Function
'    If Me.IsGreaterThen(Other) Then CompareTo = 1: Exit Function
'End Function
'
'and pay attention on what comparing-operator charcters we actually needed:
'just "<> Not Equal" and "< Less then"
'
'there acually is no need for the other operator-characters
'
