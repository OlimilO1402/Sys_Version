VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Version"
   ClientHeight    =   11295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9375
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
   ScaleHeight     =   11295
   ScaleWidth      =   9375
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
      Height          =   10695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   9375
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
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_IndentStack As Byte

Private Sub Form_Load()
    Me.Caption = "Version Class v" & MNew.VersionA.ToStr
    BtnTestVersion.Value = True
End Sub

Private Sub Form_Resize()
    Dim l As Single, t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Text1.Move l, t, W, H
End Sub

Private Sub IndentStack_Push()
    m_IndentStack = m_IndentStack + 1
End Sub

Private Sub IndentStack_Pop()
    m_IndentStack = m_IndentStack - 1
End Sub

Private Sub BtnTestVersion_Click()
    
    Text1.Text = vbNullString
    
    DebugPrint "Test Class Version"
    DebugPrint "=================="
    DebugPrint ""
    
    IndentStack_Push
    
    TestCtors
    TestComparing
    TestFileVersionInfo
    TestRandom
    TestTodayNYesterday
    
    IndentStack_Pop
    
End Sub

Sub TestCtors()
    
    DebugPrint "Test Constructors"
    DebugPrint "-----------------"
    IndentStack_Push
    
    Dim Ver As Version
    
    Set Ver = New Version
    DebugPrint Ver.ToStr '0.0.-1.-1
    DebugPrint Ver.Major & "." & Ver.Minor & "." & Ver.Build & "." & Ver.Revision & "." & Ver.MajorRevision & "." & Ver.MinorRevision   '0.0.-1.-1.-1.-1
    
    Set Ver = MNew.Version(1, 2, 3, 4)
    DebugPrint Ver.ToStr '1.2.3.4
    DebugPrint Ver.Major & "." & Ver.Minor & "." & Ver.Build & "." & Ver.Revision & "." & Ver.MajorRevision & "." & Ver.MinorRevision   '1.2.3.4.0.4
    
    Set Ver = MNew.VersionS("1.2.3.4")
    DebugPrint Ver.ToStr '1.2.3.4
    DebugPrint Ver.Major & "." & Ver.Minor & "." & Ver.Build & "." & Ver.Revision & "." & Ver.MajorRevision & "." & Ver.MinorRevision '1.2.3.4.0.4
    
    Set Ver = MNew.Version(1, 2, &H1234, &H43215678)
    DebugPrint Ver.ToStr '1.2.4660.1126258296.17185.22136
    DebugPrint Ver.Major & "." & Ver.Minor & "." & Ver.Build & "." & Ver.Revision & "." & Ver.MajorRevision & "." & Ver.MinorRevision '1.2.4660.1126258296.17185.22136
    
    Set Ver = MNew.VersionA
    DebugPrint Ver.ToStr '2025.3.1
    DebugPrint Ver.Major & "." & Ver.Minor & "." & Ver.Build & "." & Ver.Revision & "." & Ver.MajorRevision & "." & Ver.MinorRevision '2025.3.1.0.1
    
    DebugPrint ""
    
    IndentStack_Pop
    
End Sub

Sub TestComparing()
    
    DebugPrint "Test Comparisons"
    DebugPrint "----------------"
    IndentStack_Push
    
    Dim v1 As Version, v2 As Version
    
    Set v1 = MNew.Version(2025, 3, 1, 1): Set v2 = v1.Clone
    DoAllComparisons v1, v2
    
    v2.Revision = v2.Revision + 1
    DoAllComparisons v1, v2
    
    v1.Revision = v2.Revision + 1
    DoAllComparisons v1, v2
    
    IndentStack_Pop
    
End Sub

Sub TestRandom()
    
    DebugPrint "Test Random Version"
    DebugPrint "-------------------"
    IndentStack_Push
    Dim v1 As Version, v2 As Version
    
    Set v1 = MNew.Version(MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8)
    Set v2 = v1.Clone 'MNew.Version(MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8)
    DoAllComparisons v1, v2
    
    v1.Major = v1.Major + 1
    DoAllComparisons v1, v2
    
    v2.Major = v1.Major + 1
    DoAllComparisons v1, v2
    
    Set v1 = MNew.Version(MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8)
    Set v2 = v1.Clone 'MNew.Version(MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8)
    DoAllComparisons v1, v2
    
    v1.Minor = v1.Minor + 1
    DoAllComparisons v1, v2
    
    v2.Minor = v1.Minor + 1
    DoAllComparisons v1, v2
    
    Set v1 = MNew.Version(MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8)
    Set v2 = MNew.Version(MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8, MPtr.RndUInt8)
    DoAllComparisons v1, v2
    
    IndentStack_Pop
    
End Sub

Private Sub TestFileVersionInfo()
    
    DebugPrint "Test FileVersionInfo"
    DebugPrint "--------------------"
    IndentStack_Push
    
    Dim FVI1 As FileVersionInfo: Set FVI1 = MNew.FileVersionInfo(App.Path & "\" & App.ProductName & "1.exe")
    Dim FVI2 As FileVersionInfo: Set FVI2 = MNew.FileVersionInfo(App.Path & "\" & App.ProductName & "2.exe")
    
    Dim Ver1 As Version:         Set Ver1 = MNew.VersionS(FVI1.ProductVersion)
    Dim Ver2 As Version:         Set Ver2 = MNew.VersionS(FVI2.ProductVersion)
    
    DoAllComparisons Ver1, Ver2
    
    IndentStack_Pop
    
End Sub

Private Sub TestTodayNYesterday()
    
    DebugPrint "Test Today and Yesterday"
    DebugPrint "------------------------"
    IndentStack_Push

    Dim tod As Version: Set tod = MNew.VersionD
    Dim yed As Version: Set yed = MNew.VersionD(tod.ToDate - 1)
    
    DoAllComparisons tod, yed
    
    IndentStack_Pop
    
End Sub

Sub DoAllComparisons(v1 As Version, v2 As Version)
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

'Easy Comparison
'===============
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
'(compare: Version class https://learn.microsoft.com/en-us/dotnet/api/system.Version?view=net-8.0 )
'Maybe we have a situation where we have different Versions of a file or a program, and in our program
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
