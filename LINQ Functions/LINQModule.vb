Module LINQModule
    'Transcribed by Andrew Williamson from lecture given by jp2code regarding LINQ functions

    'Language Integrated Query or LINQ, pronounced link, is a Microsoft .Net component that adds native data querying capabilities
    'to .Net languages (VB.Net)

    'YOUTUBE LECTURE VIDEO LINK: https://www.youtube.com/watch?v=eVe_NOQELaI&ab_channel=jp2code
    'YOUTUBE CHANNEL: https://www.youtube.com/@jp2code

    'TYPES: all the same syntax using IQueryable to create a data provider
    '   - LINQ to Objects
    '   - LINQ to SQL - ties in with ADO and some ORM's
    '   - LINQ to XML
    '   - LINQ to Datasets

    'This example uses LINQ to objects to demonstrate examples, but all examples can be adapted to the data providers.

    Public Class cEmployee
        Public Property ID As Integer
        Public Property FirstName As String
        Public Property LastName As String
        Public Property YearsWorked As Integer
        Public Property PhoneNumber As String
        Public Property Address As New cAddress
        Public Property AccessPoints As New List(Of cAccessPoint)

        Public ReadOnly Property FullName As String
            Get
                Return FirstName & " " & LastName
            End Get
        End Property
    End Class

    Public Class cAddress
        Public Property Address1 As String
        Public Property Address2 As String
        Public Property City As String
        Public Property State As String
        Public Property Zip As String
    End Class

    Public Class cAccessPoint
        Public Property ID As Integer
        Public Property Name As String
    End Class

    Public Class cFullTimer
        Inherits cEmployee
    End Class

    Public Class cPartTimer
        Inherits cEmployee
    End Class

    Public Class cAccessJoin
        Public Property EmployeeID As Integer
        Public Property AccessID As Integer
    End Class

    Sub Main()
        'Shorthand declarations of Named Types (Strongly Typed)
        Dim oAllAccess As New cAccessPoint() With {
            .ID = 0,
            .Name = "All Access"
        }
        Dim oMainMenuAccess As New cAccessPoint() With {
            .ID = 1,
            .Name = "Main Menu"
        }
        Dim oPayRollAccess As New cAccessPoint() With {
            .ID = 2,
            .Name = "Payroll"
        }
        Dim oPhoneNumbersAccess As New cAccessPoint() With {
            .ID = 3,
            .Name = "Phone Numbers"
        }
        Dim oTimeSheetAccess As New cAccessPoint() With {
            .ID = 4,
            .Name = "Time Sheets"
        }

        Dim lstEmployees As New List(Of cEmployee)({
            New cEmployee() With {.ID = 1,
                                  .FirstName = "Daniel",
                                  .LastName = "Bewley",
                                  .YearsWorked = 5,
                                  .PhoneNumber = "9403683921",
                                  .Address = New cAddress() With {.Address1 = "901 Reen Dr.",
                                                                  .City = "Lufkin",
                                                                  .State = "TX",
                                                                  .Zip = "75904"},
                                  .AccessPoints = New List(Of cAccessPoint)({oAllAccess})
                                 },
            New cEmployee() With {.ID = 2,
                                  .FirstName = "Jared",
                                  .LastName = "Coleson",
                                  .YearsWorked = 13,
                                  .PhoneNumber = "9361112222",
                                  .Address = New cAddress() With {.Address1 = "1001 Test Ave.",
                                                                  .City = "Nacogdoches",
                                                                  .State = "TX",
                                                                  .Zip = "75964"},
                                  .AccessPoints = New List(Of cAccessPoint)({oAllAccess})
                                 },
            New cEmployee() With {.ID = 3,
                                  .FirstName = "James",
                                  .LastName = "Sanders",
                                  .YearsWorked = 2,
                                  .PhoneNumber = "9363334444",
                                  .Address = New cAddress() With {.Address1 = "753 Testing Rd.",
                                                                  .City = "Nacogdoches",
                                                                  .State = "TX",
                                                                  .Zip = "75964"},
                                  .AccessPoints = New List(Of cAccessPoint)({oMainMenuAccess, oPhoneNumbersAccess})
                                 },
            New cEmployee() With {.ID = 4,
                                  .FirstName = "Matt",
                                  .LastName = "Golden",
                                  .YearsWorked = 0,
                                  .PhoneNumber = "9365556666",
                                  .Address = New cAddress() With {.Address1 = "123 Main St.",
                                                                  .City = "Nacogdoches",
                                                                  .State = "TX",
                                                                  .Zip = "75964"},
                                  .AccessPoints = New List(Of cAccessPoint)({oMainMenuAccess, oTimeSheetAccess})
                                 },
            New cEmployee() With {.ID = 5,
                                  .FirstName = "Jeff",
                                  .LastName = "Moreau",
                                  .YearsWorked = 1,
                                  .PhoneNumber = "9367778888",
                                  .Address = New cAddress() With {.Address1 = "123 Main St.",
                                                                  .City = "Nacogdoches",
                                                                  .State = "TX",
                                                                  .Zip = "75964"},
                                  .AccessPoints = New List(Of cAccessPoint)({oMainMenuAccess, oPayRollAccess})
                                 }
        })

        'Class members being initialised cannot be shared members, read-only members, constants, or method calls.
        'They cannot be indexed or qualified either. 
        'The following examples raise compile errors:

        'FAIL 
        'Dim oInvalidEmployee As New cEmployee() With {.AccessPoints(0) = oAllAccess}

        'FAIL
        'Dim oInvalidEmployee As New cEmployee() With {.Address.City = "Test City"}

        'Basic operations in SQL format

        Dim lstSelectedEmployees As IEnumerable(Of cEmployee) =
            From oEmployee As cEmployee In lstEmployees
            Where oEmployee.YearsWorked >= 2
            Select oEmployee

        Console.WriteLine("Basic operations in SQL format")
        Console.WriteLine()
        Console.WriteLine("The following calculates the employees who have given 2 or more years of service to the company")
        Console.WriteLine()
        Console.WriteLine("        Dim lstSelectedEmployees As IEnumerable(Of cEmployee) =")
        Console.WriteLine("            From oEmployee As cEmployee In lstEmployees")
        Console.WriteLine("            Where oEmployee.YearsWorked >= 2")
        Console.WriteLine("            Select oEmployee")
        Console.WriteLine()
        Console.WriteLine("E.G.")
        Console.WriteLine()

        For Each oEmp As cEmployee In lstSelectedEmployees
            Console.WriteLine("Employee: " & oEmp.FullName)
        Next oEmp

        'Prints 
        'Employee: Daniel Bewley
        'Employee: Jared Coleson
        'Employee: James Sanders

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        'Ordered by Last Name
        lstSelectedEmployees =
            From oEmployee As cEmployee In lstEmployees
            Order By oEmployee.LastName
            Select oEmployee

        Console.WriteLine()
        Console.WriteLine("The following Orders the last names of the employees at the company, in alphabetical order")
        Console.WriteLine()
        Console.WriteLine("        lstSelectedEmployees =")
        Console.WriteLine("            From oEmployee As cEmployee In lstEmployees")
        Console.WriteLine("            Order By oEmployee.LastName")
        Console.WriteLine("            Select oEmployee")
        Console.WriteLine()
        Console.WriteLine("E.G.")
        Console.WriteLine()

        For Each oEmp As cEmployee In lstSelectedEmployees
            Console.WriteLine(oEmp.LastName)
        Next oEmp

        'Prints 
        'Bewley
        'Coleson
        'Golden
        'Moreau
        'Sanders

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()



        Console.WriteLine()
        Console.WriteLine("Combining Where/Order By")
        Console.WriteLine("Multiple Order By")

        'Combining Where/Order By
        'Multiple Order By

        lstSelectedEmployees =
            From oEmployee As cEmployee In lstEmployees
            Where oEmployee.Address.City = "Nacogdoches"
            Order By oEmployee.FirstName Descending, oEmployee.LastName Descending
            Select oEmployee

        Console.WriteLine()
        Console.WriteLine("The following calculates the employees at the company, who come from Nacogdoches,")
        Console.WriteLine("with their names ordered in descending alphabetical order,")
        Console.WriteLine("first by first name, then by last name")
        Console.WriteLine()
        Console.WriteLine("        lstSelectedEmployees =")
        Console.WriteLine("            From oEmployee As cEmployee In lstEmployees")
        Console.WriteLine("            Where oEmployee.Address.City = " & ControlChars.Quote & "Nacogdoches" & ControlChars.Quote)
        Console.WriteLine("            Order By oEmployee.FirstName Descending, oEmployee.LastName Descending")
        Console.WriteLine("            Select oEmployee")
        Console.WriteLine()
        Console.WriteLine("E.G.")
        Console.WriteLine()

        For Each oEmp As cEmployee In lstSelectedEmployees
            Console.WriteLine(oEmp.FirstName + " " + oEmp.LastName)
        Next oEmp

        'Prints 
        'Matt Golden
        'Jeff Moreau
        'Jared Coleson
        'James Sanders

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()



        Console.WriteLine()
        Console.WriteLine("Distinct results")
        Console.WriteLine("Selecting type other than list")

        'Distinct results
        'Selecting type other than list

        Dim lstDistinctCities As IEnumerable(Of String) =
            From oEmployee As cEmployee In lstEmployees
            Select oEmployee.Address.City Distinct

        Console.WriteLine()
        Console.WriteLine("The following lists the distinct (different) cities that the employees come from.")
        Console.WriteLine()
        Console.WriteLine("        Dim lstDistinctCities As IEnumerable(Of String) =")
        Console.WriteLine("            From oEmployee As cEmployee In lstEmployees")
        Console.WriteLine("            Select oEmployee.Address.City Distinct")
        Console.WriteLine()
        Console.WriteLine("E.G.")
        Console.WriteLine()

        For Each strCity As String In lstDistinctCities
            Console.WriteLine("City: " & strCity)
        Next strCity

        'Prints 
        'City:   Lufkin
        'City:   Nacogdoches

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()



        'SUM / AVERAGE / MIN / MAX

        Dim lstAllYearsWorked As IEnumerable(Of Integer) =
            From oEmployee As cEmployee In lstEmployees
            Select oEmployee.YearsWorked

        Dim intSum As Integer = lstAllYearsWorked.Sum()
        Dim dblAverage As Double = lstAllYearsWorked.Average()
        Dim intMin As Integer = lstAllYearsWorked.Min()

        Console.Clear()
        Console.WriteLine("When lists of iEnumerable are of type Integer, Double or Long")
        Console.WriteLine("The functions Sum(), Average(), Min() and Max() can be applied to them.")
        Console.WriteLine()
        Console.WriteLine("        Dim lstAllYearsWorked As IEnumerable(Of Integer) =")
        Console.WriteLine("            From oEmployee As cEmployee In lstEmployees")
        Console.WriteLine("            Select oEmployee.YearsWorked()")
        Console.WriteLine()
        Console.WriteLine("        Dim intSum As Integer = lstAllYearsWorked.Sum()")
        Console.WriteLine("        Dim dblAverage As Double = lstAllYearsWorked.Average()")
        Console.WriteLine("        Dim intMin As Integer = lstAllYearsWorked.Min()")
        Console.WriteLine()
        Console.WriteLine("E.G.")
        Console.WriteLine()
        Console.WriteLine("Sum() of years worked by employees = " & intSum)
        Console.WriteLine("Average() of years worked by employees = " & dblAverage)
        Console.WriteLine("Min() years worked by an employee = " & intMin)

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()

        Console.WriteLine()
        Console.WriteLine("Extension method of IEnumberable(Of Integer/Double/Float/Long)")
        Console.WriteLine()

        'Extension method of IEnumberable(Of Integer/Double/Float/Long)

        Dim intMax As Integer =
            (From oEmployee As cEmployee In lstEmployees
             Select oEmployee.YearsWorked).Max()

        Console.WriteLine()
        Console.WriteLine("Max() can be determined by the extension method. More on that later!")
        Console.WriteLine()
        Console.WriteLine("        Dim intMax As Integer =")
        Console.WriteLine("            (From oEmployee As cEmployee In lstEmployees")
        Console.WriteLine("             Select oEmployee.YearsWorked).Max()")
        Console.WriteLine()
        Console.WriteLine("E.G.")
        Console.WriteLine()
        Console.WriteLine("Max() years worked by an employee = " & intMax)
        Console.WriteLine()
        Console.WriteLine("Coming back to the extension method later")

        'Coming back to the extension method later

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()



        'Anonymous types (Used in following examples)
        'Useful in UI building
        'Be careful with these (.Net concession to loosely typed language advocates)
        '   NOTE: Technically anonymous types violate Elliots Standards, please still
        '         declare types with queries as much as you can.

        Console.Clear()
        Console.WriteLine("Anonymous Types")
        Console.WriteLine()
        Console.WriteLine("Anonymous types are useful in UI building")
        Console.WriteLine("Be careful with these (.Net concession to loosely typed language advocates)")
        Console.WriteLine("   NOTE: Technically anonymous types violate Elliots Standards, please still")
        Console.WriteLine("         declare types with queries as much as you can.")

        Dim oAnonyousEmployee = New With {.ID = -1,
                                          .FirstName = "Anonymous",
                                          .LastName = "Employee",
                                          .PhoneNUmber = "1112223333",
                                          .AccessPoints = New List(Of cAccessPoint)()
                                         }

        Console.WriteLine()
        Console.WriteLine("        Dim oAnonyousEmployee = New With {.ID = -1,")
        Console.WriteLine("                                          .FirstName = " & ControlChars.Quote & "Anonymous" & ControlChars.Quote & ",")
        Console.WriteLine("                                          .LastName = " & ControlChars.Quote & "Employee" & ControlChars.Quote & ",")
        Console.WriteLine("                                          .PhoneNUmber = " & ControlChars.Quote & "1112223333" & ControlChars.Quote & ",")
        Console.WriteLine("                                          .AccessPoints = New List(Of cAccessPoint)()")
        Console.WriteLine("                                         }")
        Console.WriteLine()
        Console.WriteLine("E.G.")
        Console.WriteLine()
        Console.WriteLine("Anonymous Employee Full Name: " & oAnonyousEmployee.FirstName & " " & oAnonyousEmployee.LastName)

        'Prints 
        '"Anonynous Employee Full Name: Anonynous Employee"

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("FAILS - Anonymous types do not have functions")
        Console.WriteLine("oAnonyousEmployee.Fullname could not be possible using the Fullname function in the cEmployee class")
        Console.WriteLine()
        Console.WriteLine("LINQ statement taking advantage of anonymous types - won't know type until runtime")
        Console.WriteLine("Equivalent to 'IN' clause in SQL")
        Console.WriteLine("Using function from original cEmployee definition in anonymous type creation")
        Console.WriteLine()

        'FAILS - Anonymous types do not have functions
        'Console.WriteLine("Anonymous Employee Full Name: " & oAnonyousEmployee.Fullname)

        'LINQ statement taking advantage of anonymous types - won't know type until runtime
        'Equivalent to 'IN' clause in SQL
        'Using function from original cEmployee definition in anonymous type creation

        Dim lstSelectedIDs As New List(Of Integer)({2, 3, 4})

        Dim lstOtherEmployees =
            From oEmployee In lstEmployees
                Where lstSelectedIDs.Contains(oEmployee.ID)
                Select New With {.FirstName = oEmployee.FirstName,
                                 .LastName = oEmployee.LastName,
                                 .FullName = oEmployee.FullName()}

        'Lazy loading - query not executed until used

        Console.WriteLine("Dim lstSelectedIDs As New List(Of Integer)({2, 3, 4})")
        Console.WriteLine()
        Console.WriteLine("Dim lstOtherEmployees =")
        Console.WriteLine("    From oEmployee In lstEmployees")
        Console.WriteLine("    Where lstSelectedIDs.Contains(oEmployee.ID)")
        Console.WriteLine("    Select New With {.FirstName = oEmployee.FirstName,")
        Console.WriteLine("                     .LastName = oEmployee.LastName,")
        Console.WriteLine("                     .FullName = oEmployee.FullName()}")
        Console.WriteLine()
        Console.WriteLine("Lazy loading - query not executed until used")
        Console.WriteLine()
        Console.WriteLine("E.G.")
        Console.WriteLine()

        For Each oEmployee In lstOtherEmployees
            Console.WriteLine(oEmployee.FirstName)
        Next oEmployee

        'Prints
        'Jared
        'James
        'Matt

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()

        'Grouping: Getting counts

        Console.WriteLine()
        Console.WriteLine("Grouping: Getting counts")
        Console.WriteLine()

        Dim oGroupedEmployeesByCity =
            From oEmployee In lstEmployees
            Group oEmployee By oEmployee.Address.City
            Into Group, Count()

        Console.WriteLine("Dim oGroupedEmployeesByCity =")
        Console.WriteLine("    From oEmployee In lstEmployees")
        Console.WriteLine("    Group oEmployee By oEmployee.Address.City")
        Console.WriteLine("    Into Group, Count()")


        For Each oCityGroup In oGroupedEmployeesByCity
            Console.WriteLine("City: " & oCityGroup.City & "; Count: " & oCityGroup.Count)
        Next oCityGroup
        'Prints
        '   City: Lufkin; Count 1
        '   City: Nagogdoches; Count 4

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()

        'Grouping: Getting selected objects from group

        Console.WriteLine()
        Console.WriteLine("Grouping: Getting selected objects from group")
        Console.WriteLine()
        Console.WriteLine("Dim oGroupedEmployeesByCityWithEmployeeInfo =")
        Console.WriteLine("    From oEmployee In lstEmployees")
        Console.WriteLine("    Group oEmployee By oEmployee.Address.City")
        Console.WriteLine("    Into CurrentGroup = Group, Count()")
        Console.WriteLine()
        Console.WriteLine("Uses anonymous types behind the scenes to define 'CurrentGroup'")
        Console.WriteLine()
        Console.WriteLine("For Each oCityGroup In oGroupedEmployeesByCityWithEmployeeInfo")
        Console.WriteLine()
        Console.WriteLine("    Console.WriteLine(" & ControlChars.Quote & "City: " & ControlChars.Quote & " & oCityGroup.City & " & ControlChars.Quote & "; Count: " & ControlChars.Quote & " & oCityGroup.Count)")
        Console.WriteLine()
        Console.WriteLine("    'Can loop through group of items")
        Console.WriteLine("     For Each strEmpFullName As String In")
        Console.WriteLine("         From oCurEmp As cEmployee In oCityGroup.CurrentGroup")
        Console.WriteLine("         Order By oCurEmp.FirstName")
        Console.WriteLine("         Select oCurEmp.FullName")
        Console.WriteLine()
        Console.WriteLine("         Console.WriteLine(" & ControlChars.Quote & " . " & ControlChars.Quote & " & strEmpFullName)")
        Console.WriteLine("     Next strEmpFullName")
        Console.WriteLine("Next oCityGroup")
        Console.WriteLine()


        Dim oGroupedEmployeesByCityWithEmployeeInfo =
            From oEmployee In lstEmployees
            Group oEmployee By oEmployee.Address.City
            Into CurrentGroup = Group, Count()

        'Uses anonymous types behind the scenes to define 'CurrentGroup'

        For Each oCityGroup In oGroupedEmployeesByCityWithEmployeeInfo
            Console.WriteLine("City: " & oCityGroup.City & "; Count: " & oCityGroup.Count)

            'Can loop through group of items
            For Each strEmpFullName As String In
                From oCurEmp As cEmployee In oCityGroup.CurrentGroup
                Order By oCurEmp.FirstName
                Select oCurEmp.FullName

                Console.WriteLine(" . " & strEmpFullName)
            Next strEmpFullName
        Next oCityGroup

        'Prints
        '   City: Lufkin; Count 1
        '       - Daniel Bewley
        '   City: Nagogdoches; Count 4
        '       - James Sanders
        '       - Jared Coleson
        '       - Jeff Moreau
        '       - Matt Golden

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        'Join: Get all access points' names a certain employee has access to (using many to many relationship)

        'List representing table of all access points
        Dim lstAllAccessPoints As IEnumerable(Of cAccessPoint) = New List(Of cAccessPoint)({
            oAllAccess, oMainMenuAccess, oPayRollAccess, oPhoneNumbersAccess, oTimeSheetAccess})

        'List representing many to many table specifying what access a user has
        Dim lstAccessJoin As IEnumerable(Of cAccessJoin) = New List(Of cAccessJoin)({
            New cAccessJoin() With {.EmployeeID = 1,
                                    .AccessID = oAllAccess.ID},
            New cAccessJoin() With {.EmployeeID = 2,
                                    .AccessID = oAllAccess.ID},
            New cAccessJoin() With {.EmployeeID = 3,
                                    .AccessID = oMainMenuAccess.ID},
            New cAccessJoin() With {.EmployeeID = 1,
                                    .AccessID = oPayRollAccess.ID}
        })

        Dim lstUserAccessNames =
            From oEmployee As cEmployee In lstEmployees
            Join oAccessJoin As cAccessJoin In lstAccessJoin On oEmployee.ID Equals oAccessJoin.EmployeeID
            Join oAccessPoint As cAccessPoint In lstAllAccessPoints On oAccessJoin.AccessID Equals oAccessPoint.ID
            Where oEmployee.ID = 3
            Select oAccessPoint.Name


        Console.WriteLine()
        Console.WriteLine("Join: Get all access points' names a certain employee has access to (using many to many relationship)")
        Console.WriteLine()
        Console.WriteLine("List representing table of all access points")
        Console.WriteLine("Dim lstAllAccessPoints As IEnumerable(Of cAccessPoint) = New List(Of cAccessPoint)({")
        Console.WriteLine("    oAllAccess, oMainMenuAccess, oPayRollAccess, oPhoneNumbersAccess, oTimeSheetAccess})")
        Console.WriteLine()
        Console.WriteLine("List representing many to many table specifying what access a user has")
        Console.WriteLine("Dim lstAccessJoin As IEnumerable(Of cAccessJoin) = New List(Of cAccessJoin)({")
        Console.WriteLine("    New cAccessJoin() With {.EmployeeID = 1,")
        Console.WriteLine("                            .AccessID = oAllAccess.ID},")
        Console.WriteLine("    New cAccessJoin() With {.EmployeeID = 2,")
        Console.WriteLine("                            .AccessID = oAllAccess.ID},")
        Console.WriteLine("    New cAccessJoin() With {.EmployeeID = 3,")
        Console.WriteLine("                            .AccessID = oMainMenuAccess.ID},")
        Console.WriteLine("    New cAccessJoin() With {.EmployeeID = 1,")
        Console.WriteLine("                            .AccessID = oPayRollAccess.ID},")
        Console.WriteLine("})")
        Console.WriteLine()
        Console.WriteLine("Dim lstUserAccessNames =")
        Console.WriteLine("    From oEmployee As cEmployee In lstEmployees")
        Console.WriteLine("    Join oAccessJoin As cAccessJoin In lstAccessJoin On oEmployee.ID Equals oAccessJoin.EmployeeID")
        Console.WriteLine("    Join oAccessPoint As cAccessPoint In lstAllAccessPoints On oAccessJoin.AccessID Equals oAccessPoint.ID")
        Console.WriteLine("    Where oEmployee.ID = 3")
        Console.WriteLine("    Select oAccessPoint.Name")
        Console.WriteLine()

        For Each strAccessName As String In lstUserAccessNames
            Console.WriteLine(strAccessName)
        Next
        'Prints
        '   Main Menu
        '   Payroll

        'NOTE: Joins are expensive operations, just like in databases. Use with extreme cuation.
        'caution.Or() don't use at all. Complex joins should only be used in DB2
        'Exception: Cross Database queries.

        Console.WriteLine("NOTE: Joins are expensive operations, just like in databases. Use with extreme cuation.")
        Console.WriteLine("caution.Or() don't use at all. Complex joins should only be used in DB2")
        Console.WriteLine("Exception: Cross Database queries.")

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.Clear()

        'Union: Unordered concatenation of collections
        'Useful scenario: Unioning 2 lists that have a shared inherited type 
        'Sub Select example - Can perform as many selects on anything implementing IEnumerator
        Dim lstFullTimers As New List(Of cFullTimer)
        Dim lstPartTimers As New List(Of cPartTimer)

        lstSelectedEmployees =
            From oEmployee As cEmployee In (
                (From oFullTimer As cEmployee In lstFullTimers
                 Select oFullTimer).
                Union(
                From oPartTime As cEmployee In lstPartTimers
                Select oPartTime)
            )
            Order By oEmployee.LastName
            Select oEmployee


        Console.WriteLine("Dim lstFullTimers As New List(Of cFullTimer)")
        Console.WriteLine("Dim lstPartTimers As New List(Of cPartTimer)")
        Console.WriteLine()
        Console.WriteLine("lstSelectedEmployees =")
        Console.WriteLine("    From oEmployee As cEmployee In (")
        Console.WriteLine("         (From oFullTimer As cEmployee In lstFullTimers")
        Console.WriteLine("          Select oFullTimer).")
        Console.WriteLine("         Union(")
        Console.WriteLine("         From oPartTime As cEmployee In lstPartTimers")
        Console.WriteLine("         Select oPartTime)")
        Console.WriteLine("    )")
        Console.WriteLine("    Order By oEmployee.LastName")
        Console.WriteLine("    Select oEmployee")
        Console.WriteLine()

        'Returns all records from lstFullTimers and lstPartTimers as cEmployee objects
        'ordered by last name.




        'Lambda Expressions / Anonymous Methods

        'A lambda expression is a function or subroutine with a name that can be used
        'wherever a delegate is valid. Lambda expressions can be functions or subroutines
        'and can be single-line or multi-line. You can pass values from the current scope to 
        'a lambda expression. (MSDN)

        Console.Clear()
        Console.WriteLine("Lambda Expressions / Anonymous Methods")
        Console.WriteLine()
        Console.WriteLine("A lambda expression is a function or subroutine with a name that can be used")
        Console.WriteLine("wherever a delegate is valid. Lambda expressions can be functions or subroutines")
        Console.WriteLine("and can be single-line or multi-line. You can pass values from the current scope to ")
        Console.WriteLine("a lambda expression. (MSDN)")

        Console.WriteLine()
        Console.WriteLine("Dim subPrintLine = Sub(pOutput As String)")
        Console.WriteLine("                       Console.WriteLine(pOutput)")
        Console.WriteLine("                   End Sub")
        Console.WriteLine()
        Console.WriteLine("Dim funcIncrement1 = Function(x As Integer) x + 1")
        Console.WriteLine()
        Console.WriteLine("The body of a single-line function must be an expression that returns a value, not")
        Console.WriteLine("a statement. There is no return statement for single-line functions. The value returned")
        Console.WriteLine("by the single-line function is the value of the expression in the body of the function.")
        Console.WriteLine()
        Console.WriteLine("Return types inferred or can be explicitly stated")
        Console.WriteLine()
        Console.WriteLine("Dim funcIncrement2 = Function(x As Integer) As Integer")
        Console.WriteLine("                         Return x + 2")
        Console.WriteLine("                     End Function")
        Console.WriteLine()
        Console.WriteLine("subPrintLine(" & ControlChars.Quote & "Testing 123" & ControlChars.Quote & ")")
        Console.WriteLine("Console.WriteLine(funcIncrement1(1).ToString)")
        Console.WriteLine("Console.WriteLine(funcIncrement2(1).ToString)")
        Console.WriteLine()

        Dim subPrintLine = Sub(pOutput As String)
                               Console.WriteLine(pOutput)
                           End Sub


        Dim funcIncrement1 = Function(x As Integer) x + 1

        'The body of a single-line function must be an expression that returns a value, not
        'a statement. There is no return statement for single-line functions. The value returned
        'by the single-line function is the value of the expression in the body of the function.

        'Return types inferred or can be explicitly stated
        Dim funcIncrement2 = Function(x As Integer) As Integer
                                 Return x + 2
                             End Function

        subPrintLine("Testing 123")
        Console.WriteLine(funcIncrement1(1).ToString)
        Console.WriteLine(funcIncrement2(1).ToString)
        Console.WriteLine()

        'PRINTS
        'Testing 123
        '2
        '3

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("Action and Func are the implicit type of lambda expressions / anonymous methods")
        Console.WriteLine("Important to understand: Can be used as parameter types.")
        Console.WriteLine()
        Console.WriteLine("Dim subPrintDeclaredType As Action(Of String) =")
        Console.WriteLine("    Sub(pOutput As String) Console.Write(pOutput)")
        Console.WriteLine()
        Console.WriteLine("Dim subPrintLineDeclaredType As Action(Of String) =")
        Console.WriteLine("    Sub(pOutput As String)")
        Console.WriteLine("        Console.WriteLine(pOutput)")
        Console.WriteLine("    End Sub")
        Console.WriteLine()
        Console.WriteLine("Dim funcInrement1DeclaredType As Func(Of Integer, String) =")
        Console.WriteLine("    Function(x As Integer) (x + 1).ToString()")
        Console.WriteLine()
        Console.WriteLine("Dim funcInrement2DeclaredType As Func(Of Integer, String) =")
        Console.WriteLine("    Function(x As Integer) As Integer")
        Console.WriteLine("        Return (x + 2).ToString()")
        Console.WriteLine("    End Function")
        Console.WriteLine()

        'Action and Func are the implicit type of lambda expressions / anonymous methods
        'Important to understand: Can be used as parameter types.

        Dim subPrintDeclaredType As Action(Of String) =
            Sub(pOutput As String) Console.Write(pOutput)

        Dim subPrintLineDeclaredType As Action(Of String) =
            Sub(pOutput As String)
                Console.WriteLine(pOutput)
            End Sub

        Dim funcInrement1DeclaredType As Func(Of Integer, String) =
            Function(x As Integer) (x + 1).ToString()

        Dim funcInrement2DeclaredType As Func(Of Integer, String) =
            Function(x As Integer) As Integer
                Return (x + 2).ToString()
            End Function

        'NOTES: 
        '----------------------
        '
        'Lambda expressions cannot have modifiers, such as Overloads and Overrides.

        'Optional and Paramarray paramaters are not permitted.

        'Generic paramaters are not permitted.
        '(MSDN)


        Console.WriteLine("NOTES:")
        Console.WriteLine("----------------------")
        Console.WriteLine()
        Console.WriteLine("Lambda expressions cannot have modifiers, such as Overloads and Overrides.")
        Console.WriteLine()
        Console.WriteLine("Optional and Paramarray paramaters are not permitted.")
        Console.WriteLine()
        Console.WriteLine("Generic paramaters are not permitted.")
        Console.WriteLine("(MSDN)")
        Console.WriteLine()

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        'SCOPE
        'A lambda expression shares it's context with the scope within which it is defined.
        'It has the same access rights as any code written in the containing scope. 

        Console.Clear()
        Console.WriteLine("SCOPE")
        Console.WriteLine("A lambda expression shares it's context with the scope within which it is defined.")
        Console.WriteLine("It has the same access rights as any code written in the containing scope. ")
        Console.WriteLine()
        Console.WriteLine("Dim intAge As Integer = 35")
        Console.WriteLine()
        Console.WriteLine("Dim funcGetAge As Func(Of Integer) =")
        Console.WriteLine("    Function()")
        Console.WriteLine("        Return intAge")
        Console.WriteLine("    End Function")
        Console.WriteLine()
        Console.WriteLine("Some previous examples using LINQ extension methods.")
        Console.WriteLine()
        Console.WriteLine("lstSelectedEmployees = lstEmployees.Where(Function(oEmp) oEmp.YearsWorked >= 2)")
        Console.WriteLine()
        Console.WriteLine("Used in foreach")
        Console.WriteLine()
        Console.WriteLine("For Each oEmployee As cEmployee In lstEmployees.OrderBy(Function(oEmp) oEmp.LastName)")
        Console.WriteLine("Next oEmployee")
        Console.WriteLine()
        Console.WriteLine("Large chaining should be readable by putting next statement on different lines")
        Console.WriteLine("New Method: 'ThenBy/ThenByDescending': Used in multiple column ordering")
        Console.WriteLine()
        Console.WriteLine("lstSelectedEmployees = lstEmployees.Where(Function(oEmp) oEmp.Address.City = " & ControlChars.Quote & "Nagogdoches" & ControlChars.Quote & ").")
        Console.WriteLine("    OrderByDescending(Function(oEmp) oEmp.FirstName).")
        Console.WriteLine("    ThenByDescending(Function(oEmp) oEmp.LastName)")
        Console.WriteLine()
        Console.WriteLine("Small chaining can be on one line if readable and direct")
        Console.WriteLine("lstDistinctCities = lstEmployees.Select(Function(oEmp) oEmp.Address.City).Distinct()")
        Console.WriteLine()
        Console.WriteLine("Combine both forms with ease and readability")
        Console.WriteLine("intSum = (From oEmployee As cEmployee In lstEmployees")
        Console.WriteLine("            Select oEmployee.YearsWorked).Sum()")
        Console.WriteLine()
        Console.WriteLine("Select Many")
        Console.WriteLine("Powerful extension method that a flattens a collection of lists into a single list")
        Console.WriteLine("Also demonstrates use of custom Distinct Comparer")
        Console.WriteLine("Dim lstDistinctAccessPoints = lstEmployees.SelectMany(Function(oEmp) oEmp.AccessPoints).Distinct(New cAccessPointEqualityComparer())")
        Console.WriteLine()
        Console.WriteLine("Aggregate: combines values into single result")
        Console.WriteLine("Dim strAllFirstNames As String = lstEmployees.Select(Function(oEmp) oEmp.FirstName).")
        Console.WriteLine("                                    Aggregate(Function(pPrev, pNext) pPrev & ", " & pNext)")
        Console.WriteLine()
        Console.WriteLine("NOTE: Uses string concatenation behind the scenes (Scenes are still immutable)")
        Console.WriteLine(">>>Please still use StringBuilder on large concatenation loops<<<")
        Console.WriteLine()
        Console.WriteLine("Reverse")
        Console.WriteLine("Dim lstEmployeesReversed = lstEmployees.Reverse()")
        Console.WriteLine()
        Console.WriteLine("SequenceEqual")
        Console.WriteLine()
        Console.WriteLine("Dim lstFirstNames1 = lstEmployees.Select(Function(pEmp) pEmp.FirstName)")
        Console.WriteLine("Dim lstFirstNames2 = lstEmployees.Select(Function(pEmp) pEmp.FirstName)")
        Console.WriteLine()
        Console.WriteLine("Comparing primitive types")
        Console.WriteLine()
        Console.WriteLine("If (lstFirstNames1.SequenceEqual(lstFirstNames2)) Then")
        Console.WriteLine("    Console.WriteLine(" & ControlChars.Quote & "Both are equal" & ControlChars.Quote & ")")
        Console.WriteLine("End If")
        Console.WriteLine()
        Console.WriteLine("Dim lstEmployeesCopy As New List(Of cEmployee)(lstEmployees)")
        Console.WriteLine()
        Console.WriteLine("Comparing complex types requires custom Comparer object, otherwise will just compare references.")
        Console.WriteLine()
        Console.WriteLine("If (lstEmployees.SequenceEqual(lstEmployeesCopy, New cEmployeeEqualityComparer())) Then")
        Console.WriteLine("Console.WriteLine(" & ControlChars.Quote & "Both are equal" & ControlChars.Quote & ")")
        Console.WriteLine("End If")
        Console.WriteLine()

        Dim intAge As Integer = 35

        Dim funcGetAge As Func(Of Integer) =
            Function()
                Return intAge
            End Function

        'Some previous examples using LINQ extension methods.

        lstSelectedEmployees = lstEmployees.Where(Function(oEmp) oEmp.YearsWorked >= 2)

        'Used in foreach
        For Each oEmployee As cEmployee In lstEmployees.OrderBy(Function(oEmp) oEmp.LastName)
        Next oEmployee

        'Large chaining should be readable by putting next statement on different lines
        'New Method: 'ThenBy/ThenByDescending': Used in multiple column ordering
        lstSelectedEmployees = lstEmployees.Where(Function(oEmp) oEmp.Address.City = "Nagogdoches").
            OrderByDescending(Function(oEmp) oEmp.FirstName).
            ThenByDescending(Function(oEmp) oEmp.LastName)

        'Small chaining can be on one line if readable and direct
        lstDistinctCities = lstEmployees.Select(Function(oEmp) oEmp.Address.City).Distinct()

        'Combine both forms with ease and readability
        intSum = (From oEmployee As cEmployee In lstEmployees
                    Select oEmployee.YearsWorked).Sum()

        'Select Many
        'Powerful extension method that a flattens a collection of lists into a single list
        'Also demonstrates use of custom Distinct Comparer
        Dim lstDistinctAccessPoints = lstEmployees.SelectMany(Function(oEmp) oEmp.AccessPoints).Distinct(New cAccessPointEqualityComparer())


        'Aggregate: combines values into single result
        Dim strAllFirstNames As String = lstEmployees.Select(Function(oEmp) oEmp.FirstName).
                                            Aggregate(Function(pPrev, pNext) pPrev & ", " & pNext)

        'NOTE: Uses string concatenation behind the scenes (Scenes are still immutable)
        '>>>Please still use StringBuilder on large concatenation loops<<<


        'Reverse
        'Dim lstEmployeesReversed = lstEmployees.Reverse()


        'SequenceEqual
        Dim lstFirstNames1 = lstEmployees.Select(Function(pEmp) pEmp.FirstName)
        Dim lstFirstNames2 = lstEmployees.Select(Function(pEmp) pEmp.FirstName)

        'Comparing primitive types
        If (lstFirstNames1.SequenceEqual(lstFirstNames2)) Then
            Console.WriteLine("Both are equal")
        End If

        Dim lstEmployeesCopy As New List(Of cEmployee)(lstEmployees)

        'Comparing complex types requires custom Comparer object, otherwise will just compare references.
        If (lstEmployees.SequenceEqual(lstEmployeesCopy, New cEmployeeEqualityComparer())) Then
            Console.WriteLine("Both are equal")
        End If

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        'Single / First / Last
        'Quick queries to get a single item from a collection

        'Note the type, this method *does not return a query/collection* but immediately
        'returns a single instance of the item type
        '***WILL THROW EXCEPTION if returning empty list or more than one record
        '***Do not use unless programmatically wanting to throw exception on failure of condition

        Console.Clear()
        Console.WriteLine("Single / First / Last")
        Console.WriteLine("Quick queries to get a single item from a collection")
        Console.WriteLine()
        Console.WriteLine("Note the type, this method *does not return a query/collection* but immediately")
        Console.WriteLine("returns a single instance of the item type")
        Console.WriteLine("***WILL THROW EXCEPTION if returning empty list or more than one record")
        Console.WriteLine("***Do not use unless programmatically wanting to throw exception on failure of condition")
        Console.WriteLine()
        Console.WriteLine("Dim oSingleEmployee As cEmployee = lstEmployees.Where(Function(pEmp) pEmp.ID = 1).Single()")
        Console.WriteLine()
        Console.WriteLine("NOTE: Many extension methods have optional 'where' clause as lambda method (slightly shorter code)")
        Console.WriteLine()
        Console.WriteLine("oSingleEmployee = lstEmployees.First(Function(pEmp) pEmp.FirstName = " & ControlChars.Quote & "Daniel" & ControlChars.Quote & ")")
        Console.WriteLine("oSingleEmployee = lstEmployees.Last()")
        Console.WriteLine()


        Dim oSingleEmployee As cEmployee = lstEmployees.Where(Function(pEmp) pEmp.ID = 1).Single()

        'NOTE: Many extension methods have optional 'where' clause as lambda method (slightly shorter code)
        oSingleEmployee = lstEmployees.First(Function(pEmp) pEmp.FirstName = "Daniel")
        oSingleEmployee = lstEmployees.Last()


        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("SingleOrDefault / FirstOrDefault / LastOrDefault")
        Console.WriteLine()
        Console.WriteLine("SAFE Alternative to 'Single / First / Last' that returns NULL on query finding no items")
        Console.WriteLine()
        Console.WriteLine("Returns Nothing")
        Console.WriteLine()
        Console.WriteLine("oSingleEmployee = lstEmployees.SingleOrDefault(Function(pEmp) pEmp.ID = -999)")
        Console.WriteLine()
        Console.WriteLine("NOTE: Still throws exception if *more than one* record return")
        Console.WriteLine("oSingleEmployee = lstEmployees.FirstOrDefault()")
        Console.WriteLine("oSingleEmployee = lstEmployees.LastOrDefault()")
        Console.WriteLine()

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("SingleOrDefault / FirstOrDefault / LastOrDefault")
        Console.WriteLine("SAFE Alternative to 'Single / First / Last' that returns NULL on query finding no items")
        Console.WriteLine() 


        Console.WriteLine("Returns nothing")
        Console.WriteLine()
        Console.WriteLine("oSingleEmployee = lstEmployees.SingleOrDefault(Function(pEmp) pEmp.ID = -999)")
        Console.WriteLine()
        Console.WriteLine("NOTE: Still throws exception if *more than one* record return")
        Console.WriteLine()
        Console.WriteLine("oSingleEmployee = lstEmployees.FirstOrDefault()")
        Console.WriteLine("oSingleEmployee = lstEmployees.LastOrDefault()")
        Console.WriteLine()

        'SingleOrDefault / FirstOrDefault / LastOrDefault
        'SAFE Alternative to 'Single / First / Last' that returns NULL on query finding no items

        'Returns nothing 

        oSingleEmployee = lstEmployees.SingleOrDefault(Function(pEmp) pEmp.ID = -999)

        'NOTE: Still throws exception if *more than one* record return

        oSingleEmployee = lstEmployees.FirstOrDefault()
        oSingleEmployee = lstEmployees.LastOrDefault()


        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("ElementAt / ElementAtOrDefault")
        Console.WriteLine("Index operator function")
        Console.WriteLine()
        Console.WriteLine("oSingleEmployee = lstEmployees.ElementAt(0) 'Will throw exception if out of range")
        Console.WriteLine("oSingleEmployee = lstEmployees.ElementAtOrDefault(0) 'Returns nothing if out of range")
        Console.WriteLine()

        'ElementAt / ElementAtOrDefault
        'Index operator function

        oSingleEmployee = lstEmployees.ElementAt(0) 'Will throw exception if out of range
        oSingleEmployee = lstEmployees.ElementAtOrDefault(0) 'Returns nothing if out of range

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()



        Console.WriteLine()
        Console.WriteLine("Any / All")
        Console.WriteLine("Methods that return booleans based on whether clause is met or not")
        Console.WriteLine()
        Console.WriteLine("Called on existing list in method chain")
        Console.WriteLine()
        Console.WriteLine("Dim blnEmployeesExistInLufkin As Boolean =")
        Console.WriteLine("    lstEmployees.Where(Function(pEmp) pEmp.Address.City = " & ControlChars.Quote & "Lufkin" & ControlChars.Quote & ").Any()")
        Console.WriteLine()
        Console.WriteLine("Dim blnAllEmployeesFromNacogdoches As Boolean =")
        Console.WriteLine("    lstEmployees.All(Function(pEmp) pEmp.Address.City = " & ControlChars.Quote & "Nacogdoches" & ControlChars.Quote & ")")
        Console.WriteLine()

        'Any / All
        'Methods that return booleans based on whether clause is met or not
        'Both accept where clauses or can be called on existing list

        'Called on existing list in method chain
        Dim blnEmployeesExistInLufkin As Boolean =
            lstEmployees.Where(Function(pEmp) pEmp.Address.City = "Lufkin").Any()
        'Returns True

        Dim blnAllEmployeesFromNacogdoches As Boolean =
            lstEmployees.All(Function(pEmp) pEmp.Address.City = "Nacogdoches")
        'Returns False

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("Take / TakeWhile / Skip / SkipWhile")
        Console.WriteLine()
        Console.WriteLine("Dim lstSubSelectedEmployees As IEnumerable(Of cEmployee)")
        Console.WriteLine()
        Console.WriteLine("lstSubSelectedEmployees = lstEmployees.Take(2)")
        Console.WriteLine("lstSubSelectedEmployees = lstEmployees.TakeWhile(Function(pEmp) pEmp.ID < 3)")
        Console.WriteLine("lstSubSelectedEmployees = lstEmployees.Skip(3)")
        Console.WriteLine("lstSubSelectedEmployees = lstEmployees.SkipWhile(Function(pEmp) pEmp.ID < 3)")
        Console.WriteLine()
        Console.WriteLine("Use chaining for easy list manipulation!")
        Console.WriteLine()
        Console.WriteLine("lstSubSelectedEmployees = lstEmployees.OrderBy(Function(pEmp) pEmp.YearsWorked).")
        Console.WriteLine("                                        SkipWhile(Function(pEmp) pEmp.YearsWorked < 2).")
        Console.WriteLine("                                        TakeWhile(Function(pEmp) pEmp.YearsWorked < 10)")
        Console.WriteLine()

        'Take / TakeWhile / Skip / SkipWhile

        Dim lstSubSelectedEmployees As IEnumerable(Of cEmployee)

        lstSubSelectedEmployees = lstEmployees.Take(2)
        lstSubSelectedEmployees = lstEmployees.TakeWhile(Function(pEmp) pEmp.ID < 3)
        lstSubSelectedEmployees = lstEmployees.Skip(3)
        lstSubSelectedEmployees = lstEmployees.SkipWhile(Function(pEmp) pEmp.ID < 3)

        'Use chaining for easy list manipulation!
        lstSubSelectedEmployees = lstEmployees.OrderBy(Function(pEmp) pEmp.YearsWorked).
                                                SkipWhile(Function(pEmp) pEmp.YearsWorked < 2).
                                                TakeWhile(Function(pEmp) pEmp.YearsWorked < 10)


        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("To{Type}' Functions")
        Console.WriteLine("Primitive types: Makes copies of values")
        Console.WriteLine("Complex types: Makes copies of references")
        Console.WriteLine()
        Console.WriteLine("Very useful! Use these!")
        Console.WriteLine()
        Console.WriteLine("Dim oQueryEmployees As IEnumerable(Of cEmployee) = lstEmployees.AsEnumerable()")
        Console.WriteLine()
        Console.WriteLine("Dim lstEmployeesList As List(Of cEmployee) = oQueryEmployees.ToList()")
        Console.WriteLine("Dim aryEmployees As cEmployee() = oQueryEmployees.ToArray()")
        Console.WriteLine("Dim dictEmployeesDictionary As Dictionary(Of Integer, cEmployee) = oQueryEmployees.ToDictionary(Function(pEmp) pEmp.ID)")
        Console.WriteLine()

        'To{Type}' Functions
        'Primitive types: Makes copies of values
        'Complex types: Makes copies of references

        'Very useful! Use these!

        Dim oQueryEmployees As IEnumerable(Of cEmployee) = lstEmployees.AsEnumerable()

        Dim lstEmployeesList As List(Of cEmployee) = oQueryEmployees.ToList()
        Dim aryEmployees As cEmployee() = oQueryEmployees.ToArray()
        Dim dictEmployeesDictionary As Dictionary(Of Integer, cEmployee) = oQueryEmployees.ToDictionary(Function(pEmp) pEmp.ID)


        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("Custom collections")
        Console.WriteLine("   Only need to implement IEnumerable(Of T) for simple collections")
        Console.WriteLine()
        Console.WriteLine("Dim dictEmployees As New cEmployeeDictionary()")
        Console.WriteLine("lstEmployees.ToList().ForEach(Sub(oEmp) dictEmployees.Add(oEmp))")
        Console.WriteLine()
        Console.WriteLine("lstSelectedEmployees = dictEmployees.Where(Function(w) w.YearsWorked <= 3)")
        Console.WriteLine()

        'Custom collections
        '   Only need to implement IEnumerable(Of T) for simple collections
        Dim dictEmployees As New cEmployeeDictionary()
        lstEmployees.ToList().ForEach(Sub(oEmp) dictEmployees.Add(oEmp))

        lstSelectedEmployees = dictEmployees.Where(Function(w) w.YearsWorked <= 3)


        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("LINQ to Datasets")
        Console.WriteLine("Make sure to reference 'System.Data.DataSetExtension' is included in projects")
        Console.WriteLine()
        Console.WriteLine("Dim oOriginalTblEmployees As New DataTable()")
        Console.WriteLine("oOriginalTblEmployees.Columns.Add(" & ControlChars.Quote & "ID" & ControlChars.Quote & ", GetType(Integer))")
        Console.WriteLine("oOriginalTblEmployees.Columns.Add(" & ControlChars.Quote & "FirstName" & ControlChars.Quote & ", GetType(String))")
        Console.WriteLine("oOriginalTblEmployees.Columns.Add(" & ControlChars.Quote & "LastName" & ControlChars.Quote & ", GetType(String))")
        Console.WriteLine("oOriginalTblEmployees.Columns.Add(" & ControlChars.Quote & "YearsWorked" & ControlChars.Quote & ", GetType(Integer))")
        Console.WriteLine("oOriginalTblEmployees.Columns.Add(" & ControlChars.Quote & "PhoneNumber" & ControlChars.Quote & ", GetType(String))")
        Console.WriteLine()
        Console.WriteLine("For Each oEmployee As cEmployee In lstEmployees")
        Console.WriteLine("    Dim oNewRow As DataRow = oOriginalTblEmployees.NewRow()")
        Console.WriteLine("    oNewRow(" & ControlChars.Quote & "ID" & ControlChars.Quote & ") = oEmployee.ID")
        Console.WriteLine("    oNewRow(" & ControlChars.Quote & "FirstName" & ControlChars.Quote & ") = oEmployee.FirstName")
        Console.WriteLine("    oNewRow(" & ControlChars.Quote & "LastName" & ControlChars.Quote & ") = oEmployee.LastName")
        Console.WriteLine("    oNewRow(" & ControlChars.Quote & "YearsWorked" & ControlChars.Quote & ") = oEmployee.YearsWorked")
        Console.WriteLine("    oNewRow(" & ControlChars.Quote & "PhoneNumber" & ControlChars.Quote & ") = oEmployee.PhoneNumber")
        Console.WriteLine("    oOriginalTblEmployees.Rows.Add(oNewRow)")
        Console.WriteLine("Next oEmployee")
        Console.WriteLine()
        Console.WriteLine("Dim oTableQuery =")
        Console.WriteLine("    From oEmployeeRow In oOriginalTblEmployees")
        Console.WriteLine("    Order By oEmployeeRow(" & ControlChars.Quote & "FirstName" & ControlChars.Quote & ")")
        Console.WriteLine("    Select oEmployeeRow")
        Console.WriteLine()
        Console.WriteLine("A query that returns and IEnumerable(Of DataRow) has 'CopyToDataTable'")
        Console.WriteLine("method - allows you to easily copy into new Data Table ")
        Console.WriteLine()
        Console.WriteLine("Dim oNewTblEmployees As DataTable = oTableQuery.CopyToDataTable()")
        Console.WriteLine()
        Console.WriteLine("For Each oEmployeeRow As DataRow In oNewTblEmployees.Rows")
        Console.WriteLine("    Console.WriteLine(oEmployeeRow(" & ControlChars.Quote & "FirstName" & ControlChars.Quote & ").ToString)")
        Console.WriteLine("Next oEmployeeRow")
        Console.WriteLine()

        'LINQ to Datasets
        'Make sure to reference 'System.Data.DataSetExtension' is included in projects
        Dim oOriginalTblEmployees As New DataTable()
        oOriginalTblEmployees.Columns.Add("ID", GetType(Integer))
        oOriginalTblEmployees.Columns.Add("FirstName", GetType(String))
        oOriginalTblEmployees.Columns.Add("LastName", GetType(String))
        oOriginalTblEmployees.Columns.Add("YearsWorked", GetType(Integer))
        oOriginalTblEmployees.Columns.Add("PhoneNumber", GetType(String))

        For Each oEmployee As cEmployee In lstEmployees
            Dim oNewRow As DataRow = oOriginalTblEmployees.NewRow()
            oNewRow("ID") = oEmployee.ID
            oNewRow("FirstName") = oEmployee.FirstName
            oNewRow("LastName") = oEmployee.LastName
            oNewRow("YearsWorked") = oEmployee.YearsWorked
            oNewRow("PhoneNumber") = oEmployee.PhoneNumber
            oOriginalTblEmployees.Rows.Add(oNewRow)
        Next oEmployee

        Dim oTableQuery =
            From oEmployeeRow In oOriginalTblEmployees
            Order By oEmployeeRow("FirstName")
            Select oEmployeeRow

        'A query that returns and IEnumerable(Of DataRow) has 'CopyToDataTable'
        'method - allows you to easily copy into new Data Table 
        Dim oNewTblEmployees As DataTable = oTableQuery.CopyToDataTable()

        For Each oEmployeeRow As DataRow In oNewTblEmployees.Rows
            Console.WriteLine(oEmployeeRow("FirstName").ToString)
        Next oEmployeeRow
        'Prints
        '   Daniel
        '   James
        '   Jared
        '   Jeff
        '   Matt

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.WriteLine()
        Console.WriteLine("Use extension methods as well w/ DataTables")
        Console.WriteLine()
        Console.WriteLine("For Each oEmployeeRow As DataRow In oNewTblEmployees.AsEnumerable().")
        Console.WriteLine("                                                        Where(Function(oEmp) CInt(oEmp(" & ControlChars.Quote & "YearsWorked" & ControlChars.Quote & ")))")
        Console.WriteLine("    Console.WriteLine(oEmployeeRow(" & ControlChars.Quote & "FirstName" & ControlChars.Quote & ").ToString)")
        Console.WriteLine("Next oEmployeeRow")
        Console.WriteLine()

        'Use extension methods as well w/ DataTables
        For Each oEmployeeRow As DataRow In oNewTblEmployees.AsEnumerable().
                                                                Where(Function(oEmp) CInt(oEmp("YearsWorked")))
            Console.WriteLine(oEmployeeRow("FirstName").ToString)
        Next oEmployeeRow
        'Prints
        '   Daniel
        '   Jared

        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()

        Console.WriteLine()
        Console.WriteLine("LILNQ to XML")
        Console.WriteLine("VERY POWERFUL Compared to old .Net XmlDocument and XmlElement classes/API")
        Console.WriteLine("Elliott hardly ever uses XML parsing so not very useful to us.")
        Console.WriteLine()
        Console.WriteLine("Dim oXmlContacts As XElement =")
        Console.WriteLine("    <Contacts>")
        Console.WriteLine("        <Contact>")
        Console.WriteLine("            <Name>Patrick Hines</Name>")
        Console.WriteLine("            <Phone Type=" & ControlChars.Quote & "Home" & ControlChars.Quote & ">206-555-0144</Phone>")
        Console.WriteLine("            <Phone Type=" & ControlChars.Quote & "Work" & ControlChars.Quote & ">425-555-0145</Phone>")
        Console.WriteLine("            <Address>")
        Console.WriteLine("                <Street1>123 Main Street</Street1>")
        Console.WriteLine("                 <City>Mercer Island</City>")
        Console.WriteLine("                 <State>WA</State>")
        Console.WriteLine("                 <Postal>68042</Postal>")
        Console.WriteLine("             </Address>")
        Console.WriteLine("         </Contact>")
        Console.WriteLine("     </Contacts>")
        Console.WriteLine()

        'LILNQ to XML
        'VERY POWERFUL Compared to old .Net XmlDocument and XmlElement classes/API
        'Elliott hardly ever uses XML parsing so not very useful to us.

        Dim oXmlContacts As XElement =
            <Contacts>
                <Contact>
                    <Name>Patrick Hines</Name>
                    <Phone Type="Home">206-555-0144</Phone>
                    <Phone Type="Work">425-555-0145</Phone>
                    <Address>
                        <Street1>123 Main Street</Street1>
                        <City>Mercer Island</City>
                        <State>WA</State>
                        <Postal>68042</Postal>
                    </Address>
                </Contact>
            </Contacts>

        'Gets all <Contact> elements w/ a Name of 'Patrick Hines'
        Dim oContact = oXmlContacts.Elements("Contact").
            Where(Function(oCurElement)
                      Return oCurElement.Element("Name").Value = "Patrick Hines"
                  End Function)


        Console.WriteLine()
        Console.WriteLine("Press and key to continue ...")
        Console.ReadKey()


        Console.Clear()
        Console.WriteLine("Final Notes")
        Console.WriteLine()
        Console.WriteLine("Performance:")
        Console.WriteLine(" - LINQ: in general more 'expensive' operation, not that bad though")
        Console.WriteLine(" - However: extremely readable and adds to maintainability")
        Console.WriteLine(" - I personally would not use on collections greater than 10,000 in memeory")
        Console.WriteLine(" - Collections greater than 1,000: try to do ordering, grouping and other")
        Console.WriteLine("   operations that have to enumnerate every item.")
        Console.WriteLine(" - Note: Ordering in LINQ uses QuickSort")
        Console.WriteLine()
        Console.WriteLine("Moral of the story: ALWAYS lean on DB2 for complex querying operations")
        Console.WriteLine()
        Console.WriteLine("For true performance improvements: AsParallel")
        Console.WriteLine("PLINQ (Parallel LINQ): Uses .Net built in ability to take advantage of multiple cores.")

        'Final Notes

        'Performance:
        ' - LINQ: in general more 'expensive' operation, not that bad though
        ' - However: extremely readable and adds to maintainability
        ' - I personally would not use on collections greater than 10,000 in memeory
        ' - Collections greater than 1,000: try to do ordering, grouping and other
        '   operations that have to enumnerate every item.
        ' - Note: Ordering in LINQ uses QuickSort

        'Moral of the story: ALWAYS lean on DB2 for complex querying operations

        'For true performance improvements: AsParallel
        'PLINQ (Parallel LINQ): Uses .Net built in ability to take advantage of multiple cores.

        Console.WriteLine()
        Console.WriteLine("Press Any Key to EXIT")
        Console.ReadKey()

    End Sub

    Public Class cAccessPointEqualityComparer
        Implements IEqualityComparer(Of cAccessPoint)

        Public Function Equals1(x As cAccessPoint, y As cAccessPoint) As Boolean Implements IEqualityComparer(Of cAccessPoint).Equals
            Return x.ID.Equals(y.ID)
        End Function

        Public Function GetHashCode1(obj As cAccessPoint) As Integer Implements IEqualityComparer(Of cAccessPoint).GetHashCode
            Return obj.GetHashCode()
        End Function
    End Class


    Public Class cEmployeeEqualityComparer
        Implements IEqualityComparer(Of cEmployee)

        Public Function Equals1(x As cEmployee, y As cEmployee) As Boolean Implements IEqualityComparer(Of cEmployee).Equals
            Return x.ID.Equals(y.ID)
        End Function

        Public Function GetHashCode1(obj As cEmployee) As Integer Implements IEqualityComparer(Of cEmployee).GetHashCode
            Return obj.GetHashCode()
        End Function
    End Class

    Public Class cEmployeeDictionary
        Implements IEnumerable(Of cEmployee)

        'NOTE: IDictionary too complex for this example

        Public Sub Add(oEmployee As cEmployee)
            _Employees.Add(oEmployee)
        End Sub

        Public Function GetEnumerator() As IEnumerator(Of cEmployee) Implements IEnumerable(Of cEmployee).GetEnumerator
            Return _Employees.GetEnumerator()
        End Function

        Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _Employees.GetEnumerator()
        End Function

        Private _Employees As New List(Of cEmployee)

        Default Public ReadOnly Property Item(pID As Integer) As cEmployee
            Get
                Return _Employees.SingleOrDefault(Function(oEmp) oEmp.ID = pID)
            End Get
        End Property
    End Class

End Module
