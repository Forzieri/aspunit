<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	Dim objLifecycle
	Set objLifecycle = ASPUnit.CreateLifeCycle("Setup", "Teardown")

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester Ok Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterOkPassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterOkPassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester Equal Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterEqualPassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterEqualPassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester NotEqual Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterNotEqualPassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterNotEqualPassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester Same Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterSamePassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterSamePassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester NotSame Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterNotSamePassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterNotSamePassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester assert true & false Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterAssertTruePassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertTruePassedFalsey"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertFalsePassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertFalsePassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester assert instance of Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterAssertInstanceOfWithRightType"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertInstanceOfWithWrongType") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester assert is null | is empty", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterAssertIsNullWithNull"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertIsNullWithObject"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertIsNullWithPrimitive"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertIsEmptyWithEmpty"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertIsEmptyWithObject"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertIsEmptyWithPrimitive") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester assert date", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterAssertTomorrowGreaterThanToday"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertYesterdayGreaterThanToday"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertYesterdayLessThanToday"), _
				ASPUnit.CreateTest("ASPUnitTesterAssertTomorrowLessThanToday") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.Run()

	' Create a global instance of ASPUnitTester for testing

	Sub Setup()
		Call ExecuteGlobal("Dim objService")
		Set objService = New ASPUnitTester
	End Sub

	Sub Teardown()
		Set objService = Nothing
	End Sub

	' Ok Assertion Method Tests

	Sub ASPUnitTesterOkPassedTruthy()
		Call ASPUnit.Equal(objService.Ok(True, ""), True, "Ok method should return truthy")
	End Sub

	Sub ASPUnitTesterOkPassedFalsey()
		Call ASPUnit.Equal(objService.Ok(False, ""), False, "Ok method should return falsey")
	End Sub

	' Equal Assertion Method Tests

	Sub ASPUnitTesterEqualPassedTruthy()
		Call ASPUnit.Equal(objService.Equal(True, True, ""), True, "Equal method should return truthy with equal values")
	End Sub

	Sub ASPUnitTesterEqualPassedFalsey()
		Call ASPUnit.Equal(objService.Equal(True, False, ""), False, "Equal method should return falsey with unequal values")
	End Sub

	' NotEqual Assertion Method Tests

	Sub ASPUnitTesterNotEqualPassedTruthy()
		Call ASPUnit.Equal(objService.NotEqual(True, False, ""), True, "NotEqual method should return truthy with unequal values")
	End Sub

	Sub ASPUnitTesterNotEqualPassedFalsey()
		Call ASPUnit.Equal(objService.NotEqual(True, True, ""), False, "NotEqual method should return falsey with equal values")
	End Sub

	' Same Assertion Method Tests

	Sub ASPUnitTesterSamePassedTruthy()
		Dim objA, _
			objB

		Set objA = New RegExp
		Set objB = objA

		Call ASPUnit.Equal(objService.Same(objA, objB, ""), True, "Same method should return truthy with same references")

		Set objB = Nothing
		Set objA = Nothing
	End Sub

	' Assert truthy or falsey Method Tests

	Sub ASPUnitTesterAssertTruePassedTruthy()
		Call ASPUnit.Equal(objService.assertTrue(True, ""), True, "Assert True of true should be true")
	End Sub

	Sub ASPUnitTesterAssertTruePassedFalsey()
		Call ASPUnit.Equal(objService.assertTrue(False, ""), False, "Assert True of false should be false")
	End Sub

	Sub ASPUnitTesterAssertFalsePassedTruthy()
		Call ASPUnit.Equal(objService.assertFalse(True, ""), False, "Assert false of True should be false")
	End Sub

	Sub ASPUnitTesterAssertFalsePassedFalsey()
		Call ASPUnit.Equal(objService.assertFalse(False, ""), True, "Assert false of false should be true")
	End Sub

	Sub ASPUnitTesterAssertInstanceOfWithRightType()
		Dim regex
		Set regex = New RegExp
		Call ASPUnit.Equal(objService.assertInstanceOf(regex, "IRegExp2", ""), True, "Regexp instance should be of IRegExp2 type")
	End Sub

	Sub ASPUnitTesterAssertInstanceOfWithWrongType()
		Dim notRegexp
		notRegexp = "Not a string"
		Call ASPUnit.Equal(objService.assertInstanceOf(notRegexp, "IRegExp2", ""), False, "String instance should not be of IRegExp2 type")
	End Sub

	Sub ASPUnitTesterSamePassedFalsey()
		Dim objA, _
			objB

		Set objA = New RegExp
		Set objB = New RegExp

		Call ASPUnit.Equal(objService.Same(objA, objB, ""), False, "Same method should return falsey with different references")

		Set objB = Nothing
		Set objA = Nothing
	End Sub

	' NotSame Assertion Method Tests

	Sub ASPUnitTesterNotSamePassedTruthy()
		Dim objA, _
			objB

		Set objA = New RegExp
		Set objB = New RegExp

		Call ASPUnit.Equal(objService.NotSame(objA, objB, ""), True, "NotSame method should return truthy with different references")

		Set objB = Nothing
		Set objA = Nothing
	End Sub

	Sub ASPUnitTesterNotSamePassedFalsey()
		Dim objA, _
			objB

		Set objA = New RegExp
		Set objB = objA

		Call ASPUnit.Equal(objService.NotSame(objA, objB, ""), False, "NotSame method should return falsey with same references")

		Set objB = Nothing
		Set objA = Nothing
	End Sub

	' Is Null assertions
	Sub ASPUnitTesterAssertIsNullWithNull()
		Dim A
		A = Null
		Call ASPUnit.Equal(objService.assertIsNull(A,""), True, "Assert is null on NULL should return true")
	End Sub

	Sub ASPUnitTesterAssertIsNullWithObject()
		Dim A
		Set A = new RegExp
		Call ASPUnit.Equal(objService.assertIsNull(A,""), False, "Assert is null on RegExp should return false")
		Set A = Nothing
	End Sub

	Sub ASPUnitTesterAssertIsNullWithPrimitive()
		Dim A
		A = 21
		Call ASPUnit.Equal(objService.assertIsNull(A,""), False, "Assert is null on integer should return false")
	End Sub

	Sub ASPUnitTesterAssertIsEmptyWithEmpty()
		Dim A
		A = Empty
		Call ASPUnit.Equal(objService.assertIsEmpty(A,""), True, "Assert is empty on EMPTY should return true")
	End Sub

	Sub ASPUnitTesterAssertIsEmptyWithObject()
		Dim A
		Set A = new RegExp
		Call ASPUnit.Equal(objService.assertIsEmpty(A,""), False, "Assert is empty on RegExp should return false")
		Set A = Nothing
	End Sub

	Sub ASPUnitTesterAssertIsEmptyWithPrimitive()
		Dim A
		A = 21
		Call ASPUnit.Equal(objService.assertIsEmpty(A,""), False, "Assert is empty on integer should return false")
	End Sub

	Sub ASPUnitTesterAssertTomorrowGreaterThanToday()
		Dim today : today = now()
		Dim tomorrow : tomorrow = DateAdd("d", 1, today)
		Call ASPUnit.Equal(objService.AssertDateGreaterThan(tomorrow, today, ""), True, "Tomorrow is greater then today")
	End Sub

	Sub ASPUnitTesterAssertYesterdayLessThanToday()
		Dim today : today = now()
		Dim yesterday : yesterday = DateAdd("d", -1, today)
		Call ASPUnit.Equal(objService.AssertDateLessThan(yesterday, today, ""), True, "Yesterday is less then today")
	End Sub

	Sub ASPUnitTesterAssertYesterdayGreaterThanToday()
		Dim today : today = now()
		Dim yesterday : yesterday = DateAdd("d", -1, today)
		Call ASPUnit.Equal(objService.AssertDateGreaterThan(yesterday, today, ""), False, "Yesterday is not greater then today")
	End Sub

	Sub ASPUnitTesterAssertTomorrowLessThanToday()
		Dim today : today = now()
		Dim tomorrow : tomorrow = DateAdd("d", 1, today)
		Call ASPUnit.Equal(objService.AssertDateLessThan(tomorrow, today, ""), False, "Tomorrow is not less then today")
	End Sub

%>