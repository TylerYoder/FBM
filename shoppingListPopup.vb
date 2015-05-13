Imports Microsoft.Office.Interop
Imports System.IO

Public Class shoppingListPopup

    Private Sub shoppingListPopup_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim startDay = Date.Today
        For x = 0 To 6
            If startDay.AddDays(x).DayOfWeek.ToString = main.startPrintDay Then
                StartTimePicker.Value = startDay.AddDays(x)
                Exit For
            End If
        Next
        EndTimePicker.Value() = StartTimePicker.Value.AddDays(6)
    End Sub

    Private Sub printShoppingListButton_Click(sender As Object, e As EventArgs) Handles printShoppingListButton.Click
        If (StartTimePicker.Value() > EndTimePicker.Value()) Then
            MsgBox("Start Time cannot be after End Time. Please try again.")
        Else
            Me.Close()
            printshoppingList(StartTimePicker.Value, EndTimePicker.Value)
        End If
    End Sub

    Public Function printshoppingList(startDate As Date, endDate As Date)
        If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\menu.xml") = False Then
            Return Nothing
        End If
        Dim menuFile As XDocument = XDocument.Load((Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\menu.xml"), LoadOptions.PreserveWhitespace)
        If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\recipes.xml") = False Then
            Return Nothing
        End If
        Dim recipeFile As XDocument = XDocument.Load((Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\recipes.xml"), LoadOptions.PreserveWhitespace)

        'fix corner case of year change
        Dim numOfDays As Integer
        If (startDate.Year = endDate.Year) Then
            numOfDays = endDate.DayOfYear - startDate.DayOfYear
        Else
            numOfDays = (365 - startDate.DayOfYear) + endDate.DayOfYear
        End If

        Dim word As New Word.Application
        Dim shoppingFile As Word.Document = word.Documents.Add
        Dim counter As Integer = 0
        Dim para As Word.Paragraph = shoppingFile.Paragraphs.Add
        Dim ingredientsList(300, 2) As String '300 is arbitrary for now

        shoppingFile.Range.Font.Size = 8
        shoppingFile.Range.ParagraphFormat.SpaceAfter = 0
        shoppingFile.Sections(1).Headers(1).Range.Text = "Shopping List for: " + DateAndTime.MonthName(startDate.Month) + " " + startDate.Day.ToString + " to " + DateAndTime.MonthName(endDate.Month) + " " + endDate.Day.ToString


        For x = 0 To numOfDays
            Dim day = startDate.AddDays(x).Day
            Dim curYear = startDate.AddDays(x).Year
            Dim curMonth = startDate.AddDays(x).Month
            Dim lunchEntree As String
            Dim lunchSide1 As String
            Dim lunchSide2 As String
            Dim dinnerEntree As String
            Dim dinnerSide1 As String
            Dim dinnerSide2 As String

            Dim meals = _
                From nm In menuFile.<Calendar>.<Day>.Elements("Date") _
                Where CStr(nm.Element("Year")) = curYear.ToString And CStr(nm.Element("Day")) = day.ToString And CStr(nm.Element("Month")) = curMonth.ToString
                Select {nm.Parent().Element("Lunch"), nm.Parent().Element("Dinner")}


            If meals.Count <> 0 Then
                If meals(0).ElementAt(0) IsNot Nothing Then
                    Dim lunchItems = meals(0).ElementAt(0).Value.Split(vbLf)
                    lunchEntree = lunchItems(1).Substring(6)
                    lunchSide1 = lunchItems(2).Substring(6)
                    lunchSide2 = lunchItems(3).Substring(6)

                    'first check reg recipes, since that's where the majority of recipes will come from
                    Dim lunchRecipe = _
                        From nm In recipeFile.<Recipes>.<RegEntrees>.Elements("Recipe")
                        Where CStr(nm.Element("EntreeName")) = lunchEntree
                        Select {nm.Element("Ingredients")}
                    'if that doesn't work, then check temp recipes
                    If lunchRecipe.Count = 0 Then
                        lunchRecipe = _
                            From nm In recipeFile.<Recipes>.<TempEntrees>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = lunchEntree
                            Select {nm.Element("Ingredients")}
                    End If
                    If lunchRecipe.Count = 0 Then
                        Continue For
                    End If
                    For Each element In lunchRecipe(0).Elements
                        Dim ingredientItems = element.Descendants()
                        'check if ingredient name is already in list. If so, check if ingredient qty type is same. If so, add qtys. Otherwise create new item.
                        Dim flag As Boolean = False
                        For y = 0 To counter
                            If ingredientsList(y, 2) = ingredientItems(2).Value Then
                                If ingredientsList(y, 1) = ingredientItems(1).Value Then
                                    Dim old As Double = CDbl(ingredientsList(y, 0))
                                    Dim nu As Double = CDbl(ingredientItems(0).Value)
                                    ingredientsList(y, 0) = (old + nu).ToString
                                    flag = True
                                    Exit For
                                End If
                            End If
                        Next
                        If flag = True Then
                            Continue For
                        End If
                        ingredientsList(counter, 0) = ingredientItems(0).Value
                        ingredientsList(counter, 1) = ingredientItems(1).Value
                        ingredientsList(counter, 2) = ingredientItems(2).Value
                        counter += 1
                    Next

                    If lunchSide1 <> "" Then
                        'first check reg recipes, since that's where the majority of recipes will come from
                        Dim lunchSide1Recipe = _
                            From nm In recipeFile.<Recipes>.<RegSides>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = lunchSide1
                            Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                        'if that doesn't work, then check temp recipes
                        If lunchSide1Recipe.Count = 0 Then
                            lunchSide1Recipe = _
                                From nm In recipeFile.<Recipes>.<TempSides>.Elements("Recipe")
                                Where CStr(nm.Element("EntreeName")) = lunchSide1
                                Select {nm.Element("Ingredients")}
                        End If
                        If lunchSide1Recipe.Count = 0 Then
                            Continue For
                        End If
                        For Each element In lunchSide1Recipe(0).Elements
                            Dim ingredientItems = element.Descendants()
                            'check if ingredient name is already in list. If so, check if ingredient qty type is same. If so, add qtys. Otherwise create new item.
                            Dim flag As Boolean = False
                            For y = 0 To counter
                                If ingredientsList(y, 2) = ingredientItems(2).Value Then
                                    If ingredientsList(y, 1) = ingredientItems(1).Value Then
                                        Dim old As Double = CDbl(ingredientsList(y, 0))
                                        Dim nu As Double = CDbl(ingredientItems(0).Value)
                                        ingredientsList(y, 0) = (old + nu).ToString
                                        flag = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If flag = True Then
                                Continue For
                            End If
                            ingredientsList(counter, 0) = ingredientItems(0).Value
                            ingredientsList(counter, 1) = ingredientItems(1).Value
                            ingredientsList(counter, 2) = ingredientItems(2).Value
                            counter += 1
                        Next
                    End If

                    If lunchSide2 <> "" Then
                        'first check reg recipes, since that's where the majority of recipes will come from
                        Dim lunchSide2Recipe = _
                            From nm In recipeFile.<Recipes>.<RegSides>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = lunchSide2
                            Select {nm.Element("Ingredients")}
                        'if that doesn't work, then check temp recipes
                        If lunchSide2Recipe.Count = 0 Then
                            lunchSide2Recipe = _
                                From nm In recipeFile.<Recipes>.<TempSides>.Elements("Recipe")
                                Where CStr(nm.Element("EntreeName")) = lunchSide2
                                Select {nm.Element("Ingredients")}
                        End If
                        If lunchSide2Recipe.Count = 0 Then
                            Continue For
                        End If
                        For Each element In lunchSide2Recipe(0).Elements
                            Dim ingredientItems = element.Descendants()
                            'check if ingredient name is already in list. If so, check if ingredient qty type is same. If so, add qtys. Otherwise create new item.
                            Dim flag As Boolean = False
                            For y = 0 To counter
                                If ingredientsList(y, 2) = ingredientItems(2).Value Then
                                    If ingredientsList(y, 1) = ingredientItems(1).Value Then
                                        Dim old As Double = CDbl(ingredientsList(y, 0))
                                        Dim nu As Double = CDbl(ingredientItems(0).Value)
                                        ingredientsList(y, 0) = (old + nu).ToString
                                        flag = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If flag = True Then
                                Continue For
                            End If
                            ingredientsList(counter, 0) = ingredientItems(0).Value
                            ingredientsList(counter, 1) = ingredientItems(1).Value
                            ingredientsList(counter, 2) = ingredientItems(2).Value
                            counter += 1
                        Next
                    End If
                End If


                If meals(0).ElementAt(1) IsNot Nothing Then
                    Dim dinnerItems = meals(0).ElementAt(1).Value.Split(vbLf)
                    dinnerEntree = dinnerItems(1).Substring(6)
                    dinnerSide1 = dinnerItems(2).Substring(6)
                    dinnerSide2 = dinnerItems(3).Substring(6)

                    'first check reg recipes, since that's where the majority of recipes will come from
                    Dim dinnerRecipe = _
                        From nm In recipeFile.<Recipes>.<RegEntrees>.Elements("Recipe")
                        Where CStr(nm.Element("EntreeName")) = dinnerEntree
                        Select {nm.Element("Ingredients")}
                    'if that doesn't work, then check temp recipes
                    If dinnerRecipe.Count = 0 Then
                        dinnerRecipe = _
                            From nm In recipeFile.<Recipes>.<TempEntrees>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = dinnerEntree
                            Select {nm.Element("Ingredients")}
                    End If
                    If dinnerRecipe.Count = 0 Then
                        Continue For
                    End If
                    For Each element In dinnerRecipe(0).Elements
                        Dim ingredientItems = element.Descendants()
                        'check if ingredient name is already in list. If so, check if ingredient qty type is same. If so, add qtys. Otherwise create new item.
                        Dim flag As Boolean = False
                        For y = 0 To counter
                            If ingredientsList(y, 2) = ingredientItems(2).Value Then
                                If ingredientsList(y, 1) = ingredientItems(1).Value Then
                                    Dim old As Double = CDbl(ingredientsList(y, 0))
                                    Dim nu As Double = CDbl(ingredientItems(0).Value)
                                    ingredientsList(y, 0) = (old + nu).ToString
                                    flag = True
                                    Exit For
                                End If
                            End If
                        Next
                        If flag = True Then
                            Continue For
                        End If
                        ingredientsList(counter, 0) = ingredientItems(0).Value
                        ingredientsList(counter, 1) = ingredientItems(1).Value
                        ingredientsList(counter, 2) = ingredientItems(2).Value
                        counter += 1
                    Next

                    If dinnerSide1 <> "" Then
                        'first check reg recipes, since that's where the majority of recipes will come from
                        Dim dinnerSide1Recipe = _
                            From nm In recipeFile.<Recipes>.<RegSides>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = dinnerSide1
                            Select {nm.Element("Ingredients")}
                        'if that doesn't work, then check temp recipes
                        If dinnerSide1Recipe.Count = 0 Then
                            dinnerSide1Recipe = _
                                From nm In recipeFile.<Recipes>.<TempSides>.Elements("Recipe")
                                Where CStr(nm.Element("EntreeName")) = dinnerSide1
                                Select {nm.Element("Ingredients")}
                        End If
                        If dinnerSide1Recipe.Count = 0 Then
                            Continue For
                        End If
                        For Each element In dinnerSide1Recipe(0).Elements
                            Dim ingredientItems = element.Descendants()
                            'check if ingredient name is already in list. If so, check if ingredient qty type is same. If so, add qtys. Otherwise create new item.
                            Dim flag As Boolean = False
                            For y = 0 To counter
                                If ingredientsList(y, 2) = ingredientItems(2).Value Then
                                    If ingredientsList(y, 1) = ingredientItems(1).Value Then
                                        Dim old As Double = CDbl(ingredientsList(y, 0))
                                        Dim nu As Double = CDbl(ingredientItems(0).Value)
                                        ingredientsList(y, 0) = (old + nu).ToString
                                        flag = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If flag = True Then
                                Continue For
                            End If
                            ingredientsList(counter, 0) = ingredientItems(0).Value
                            ingredientsList(counter, 1) = ingredientItems(1).Value
                            ingredientsList(counter, 2) = ingredientItems(2).Value
                            counter += 1
                        Next
                    End If

                    If dinnerSide2 <> "" Then
                        'first check reg recipes, since that's where the majority of recipes will come from
                        Dim dinnerSide2Recipe = _
                            From nm In recipeFile.<Recipes>.<RegSides>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = dinnerSide2
                            Select {nm.Element("Ingredients")}
                        'if that doesn't work, then check temp recipes
                        If dinnerSide2Recipe.Count = 0 Then
                            dinnerSide2Recipe = _
                                From nm In recipeFile.<Recipes>.<TempSides>.Elements("Recipe")
                                Where CStr(nm.Element("EntreeName")) = dinnerSide2
                                Select {nm.Element("Ingredients")}
                        End If
                        If dinnerSide2Recipe.Count = 0 Then
                            Continue For
                        End If
                        For Each element In dinnerSide2Recipe(0).Elements
                            Dim ingredientItems = element.Descendants()
                            'check if ingredient name is already in list. If so, check if ingredient qty type is same. If so, add qtys. Otherwise create new item.
                            Dim flag As Boolean = False
                            For y = 0 To counter
                                If ingredientsList(y, 2) = ingredientItems(2).Value Then
                                    If ingredientsList(y, 1) = ingredientItems(1).Value Then
                                        Dim old As Double = CDbl(ingredientsList(y, 0))
                                        Dim nu As Double = CDbl(ingredientItems(0).Value)
                                        ingredientsList(y, 0) = (old + nu).ToString
                                        flag = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If flag = True Then
                                Continue For
                            End If
                            ingredientsList(counter, 0) = ingredientItems(0).Value
                            ingredientsList(counter, 1) = ingredientItems(1).Value
                            ingredientsList(counter, 2) = ingredientItems(2).Value
                            counter += 1
                        Next
                    End If
                End If
            End If

        Next
        For y = 0 To counter
            para.Range.Text = ingredientsList(y, 0) + " " + ingredientsList(y, 1) + " " + ingredientsList(y, 2)
            para.Range.InsertParagraphAfter()
        Next
        para.Range.InsertParagraphAfter()
        word.Visible = True

        Return Nothing
    End Function




End Class