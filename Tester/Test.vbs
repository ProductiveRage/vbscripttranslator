Option Explicit

WScript.Echo 1.1
WScript.Echo a.Name

WScript.Echo 1+1
WScript.Echo 1-1
WScript.Echo 1+-1
WScript.Echo 1+++1
WScript.Echo 1+++-1
WScript.Echo 1+++--1

WScript.Echo "Test1"

WScript.Echo _
  "Test2"

WScript.Echo "Test 2.1": WScript.Echo "Test 2.2"

Dim objTest: Set objTest = new CTest
objTest.F1 1, "Two", 3
objTest.F1(1, "Two", 3)
Call objTest.F1(1, "Two", 3)
Call(objTest.F1(_
	1, _
	"Two", _
	3 _
))

WScript.Echo T1(1, 2)
WScript.Echo T1(1, 3) + T1(1, 4)

Function T1(a1, a2)
	T1 = a1 + a2
End Function

Class CTest
	Public m0
	Private m1
	Dim m2
	Public Function F1(ByRef a1, ByVal a2, a3)
		On Error Resume Next
		WScript.Echo "At F1"
		Err.Clear
		WScript.Echo "Still at F1"
		On Error Goto 0
	End Function
End Class

WScript.Echo "Test3"