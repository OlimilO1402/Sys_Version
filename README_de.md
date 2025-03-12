# Sys_Version  
## Easy Comparison  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/Sys_Version?style=plastic)](https://github.com/OlimilO1402/Sys_Version/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Sys_Version?style=plastic)](https://github.com/OlimilO1402/Sys_Version/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Sys_Version/total.svg)](https://github.com/OlimilO1402/Sys_Version/releases/download/v2025.3.2/Version_v2025.3.2.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)

Leichteres Vergleichen  
======================  
  
Um zwei Zahlen miteinandeer zu vergleichen, gibt es in jeder seriösen Programmiersprache die typischen Operatorzeichen, die natürlich jeder kennt.  
Wie die folgenden:  
  
 |  VB   |  C#   |  Bedeutung   
 |:-----:|:-----:|:----------------  
 |  =    |  ==   |  Gleichheit  
 |  \<>  |  !=   |  Nicht Gleich  
 |  \<=  |  \<=  |  Kleiner oder Gleich  
 |  \>=  |  \>=  |  Größer oder Gleich  
 |  \<   |  \<   |  Kleiner als  
 |  \>   |  \>   |  Größer als  
 |       |       |  int CompareTo(other)  

Um zwei Objekte (einer Klasse) miteinander zu vergleichen, gibt es in anderen Sprachen etwas, das als Operatorüberladung bezeichnet wird.  
Das heißt, dass man eine Funktion schreiben kann, die ein Operator-Zeichen als „Funktionsname“ hat.  
In VBA/VBC haben wir keine Operatorüberladung, aber das stört uns nicht, man braucht das nicht wirklich.  
Es ist nur „syntaktischer Zucker“, um den Code lesbarer zu machen, und imho erfüllt es seinen Zweck nicht in jeder Situation.  
In der Tat ist das Schreiben von benannten Member-Funktionen lesbar genug, um zwei Objekte zu vergleichen.  
  
Schauen wir uns also die obige Liste an. Brauchen wir eine Funktion für jeden Operator, für jeden möglichen Vergleich?  
Ja, wir brauchen vielleicht tatsächlich alle oben genannten Funktionen, aber wussten Sie, dass wir eigentlich nur 2 Funktionen brauchen?
und alle anderen Operationen nur eine Kombination aus diesen beiden Funktionen sind?  
  
In VBA/VBC definieren wir eine Boolesche Variable und diese hat per se den Wert „False“. VB macht das für uns, also ist die Initialisierung 
einer booleschen Variable oder auch eine boolesche Funktion nicht erforderlich.    

Die 2 Funktionen die wir brauchen sind:  
* eine Public Member Funktion "Equals" und  
* eine Private Member Funktion "CheckGreater";  
  
an diese übergeben wir einfach das „andere“ Objekt, und alle anderen Operator-Funktionen sind nur Kombinationen dieser beiden Funktionen.  
  
Um dem Ganzen etwas Sinnvolles zu geben, könnten wir uns eine Klasse „Version“ mit den Mitgliedseigenschaften Major, Minor, Build und Revision ausdenken.

([Vergleiche: Version class](https://learn.microsoft.com/en-us/dotnet/api/system.version?view=net-8.0))  
Vielleicht haben wir eine Situation, in der wir verschiedene Versionen einer Datei oder eines Programms haben, und in unserem Programm
wollen wir darauf reagieren.  
  
Hier sind die 2 wichtigsten Funktionen in voller Größe, die wir brauchen:  

```vba  
Public Function Equals(Other As Version) As Boolean
    If Other Is Nothing Then Exit Function
    If Me.Major <> Other.Major Then Exit Function
    If Me.Minor <> Other.Minor Then Exit Function
    If Me.Build <> Other.Build Then Exit Function
    If Me.Revision <> Other.Revision Then Exit Function
    Equals = True
End Function

Private Function CheckGreater(Other As Version) As Boolean
    If Other Is Nothing Then CheckGreater = True: Exit Function
    If Me.Major < Other.Major Then Exit Function
    If Me.Minor < Other.Minor Then Exit Function
    If Me.Build < Other.Build Then Exit Function
    If Me.Revision < Other.Revision Then Exit Function
    CheckGreater = True
End Function
```
  
Und hier sind die sehr schlanken Funktionen für alle anderen Vergleiche:  
  
```vba  
Public Function IsLessThen(Other As Version) As Boolean
    If Me.Equals(Other) Then Exit Function
    IsLessThen = Not CheckGreater(Other)
End Function

Public Function IsLessThenOrEqual(Other As Version) As Boolean
    If Me.Equals(Other) Then IsLessThenOrEqual = True: Exit Function
    IsLessThenOrEqual = Not CheckGreater(Other)
End Function

Public Function IsGreaterThen(Other As Version) As Boolean
    If Me.Equals(Other) Then Exit Function
    IsGreaterThen = CheckGreater(Other)
End Function

Public Function IsGreaterThenOrEqual(Other As Version) As Boolean
    If Me.Equals(Other) Then IsGreaterThenOrEqual = True: Exit Function
    IsGreaterThenOrEqual = CheckGreater(Other)
End Function

Public Function CompareTo(Other As Version) As Long
    If Me.Equals(Other) Then Exit Function
    If CheckGreater(Other) Then CompareTo = 1 Else CompareTo = -1
End Function
```  
  
Achten Sie darauf, welche Vergleichsoperatorzeichen tatsächlich benötigt werden:  
„<> Nicht gleich“ und „< Kleiner als“.  
  
![Version Image](Resources/Version.png "Version Image")  
  