# Sys_Version  
## Easy Comparison  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/Sys_Version?style=plastic)](https://github.com/OlimilO1402/Sys_Version/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Sys_Version?style=plastic)](https://github.com/OlimilO1402/Sys_Version/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Sys_Version/total.svg)](https://github.com/OlimilO1402/Sys_Version/releases/download/v2025.3.2/Version_v2025.3.2.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)
   
([zur Version des Textes auf deutsch](https://github.com/OlimilO1402/Sys_Version/README_de.md))  
  
Easy Comparison  
===============  
  
For comparing two numbers, in any serious programming language, there are the typical operator characters of course everybody knows.  
like the following:  
  
 |  VB   |  C#   |  meaning   
 |:-----:|:-----:|:----------------  
 |  =    |  ==   |  Equality  
 |  \<>  |  !=   |  Not Equal  
 |  \<=  |  \<=  |  Less then or equal  
 |  \>=  |  \>=  |  Greater then or equal  
 |  \<   |  \<   |  Less then  
 |  \>   |  \>   |  Greater then  
 |       |       |  int CompareTo(other)  

For comparing two objects (of a class), in other languages, there is something called operator overloading.  
What means something like you can write a function that has an operator-character as the "function name".  
In VBA/VBC we do not have operator overloading, but we do not bother, we even do not need this.  
It is just "syntactic sugar" to make code more readable, and imho it does not fulfill his purpose in every
situation.  
In fact writing named member-functions is readable enough for comparing two objects.  
  
So have a look at the list above. Do we need a function for every operator, for every possible comparison?  
Yes, we maybe actually need all the above functions, but did you know that we actually need only 2 functions,
and all the other operations are just a combination of that two functions?  
  
In VBA/VBC we dim a Boolean and per se the boolean has the value "False". VB does this for us, so there is no
need for an extra initialization of a Boolean variable or also even a Boolean function.  
  
The 2 functions we need are:
* a public member function "Equals" and  
* a private member function "CheckGreater";  
 
where we just hand over the "other" object, and all the other operator-functions are just combinations of this two functions.  
  
To give this something what actually makes sense we could imagine a class "Version" with the  
member properties Major, Minor, Build And Revision ([compare: Version class](https://learn.microsoft.com/en-us/dotnet/api/system.version?view=net-8.0))  
Maybe we have a situation where we have different versions of a file or a program, and in our program
we want to react on it.  
Here are the 2 main full-size functions we need:  

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
  
And here are the very slim functions for all other comparisons:  
  
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
  
Pay attention on what comparing operator characters was actually needed:  
just "<> Not Equal" and "< Less then", no need for the other operator-characters.  
I mean you could, but you don't have to, and it could make the code not better but even less readable.  
  
![Version Image](Resources/Version.png "Version Image")  
  