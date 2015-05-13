Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Xml

Public Class printWeeklyPopup

    Private Sub printWeeklyPopup_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim startDay = Date.Today
        For x = 0 To 6
            If startDay.AddDays(x).DayOfWeek.ToString = main.startPrintDay Then
                StartTimePicker.Value = startDay.AddDays(x)
                Exit For
            End If
        Next
        EndTimePicker.Value() = StartTimePicker.Value.AddDays(6)
    End Sub

    Private Sub printWeeklyPopupButton_Click(sender As Object, e As EventArgs) Handles printWeeklyPopupButton.Click
        If (StartTimePicker.Value() > EndTimePicker.Value()) Then
            MsgBox("Start Time cannot be after End Time. Please try again.")
        Else
            Me.Close()
            printWeekly(StartTimePicker.Value, EndTimePicker.Value, printShoppingListCheckBox.Checked)
        End If
    End Sub

    Private Function printWeekly(startDate As Date, endDate As Date, printShoppingListFlag As Boolean)
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
        Dim weeklyFile As Word.Document = word.Documents.Add
        weeklyFile.Range.Font.Size = 8
        weeklyFile.Range.ParagraphFormat.SpaceAfter = 0
        weeklyFile.Sections(1).Headers(1).Range.Text = "Meals List for: " + DateAndTime.MonthName(startDate.Month) + " " + startDate.Day.ToString + " to " + DateAndTime.MonthName(endDate.Month) + " " + endDate.Day.ToString

        For x = 0 To numOfDays
            Dim day = startDate.AddDays(x).Day
            Dim curYear = startDate.AddDays(x).Year
            Dim curMonth = startDate.AddDays(x).Month
            Dim lunchEntree As String
            Dim lunchSide1 As String
            Dim lunchSide2 As String
            Dim lunchIngredients As New List(Of String)
            Dim lunchInstructions As New List(Of String)
            Dim dinnerEntree As String
            Dim dinnerSide1 As String
            Dim dinnerSide2 As String
            Dim dinnerIngredients As New List(Of String)
            Dim dinnerInstructions As New List(Of String)

            Dim para As Word.Paragraph = weeklyFile.Paragraphs.Add

            para.Range.Font.Size = 10
            para.Range.Font.Bold = True
            para.Range.Text = startDate.AddDays(x).DayOfWeek.ToString + ", " + DateAndTime.MonthName(startDate.AddDays(x).Month) + " " + startDate.AddDays(x).Day.ToString
            para.Range.InsertParagraphAfter()

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

                    para.Range.Font.Bold = True
                    If lunchSide1 <> "" Then
                        If lunchSide2 <> "" Then
                            para.Range.Text = "Lunch: " + lunchEntree + " with " + lunchSide1 + " and " + lunchSide2
                        Else
                            para.Range.Text = "Lunch: " + lunchEntree + " with " + lunchSide1
                        End If
                    Else
                        para.Range.Text = "Lunch: " + lunchEntree
                    End If
                    para.Range.InsertParagraphAfter()
                    para.Range.Font.Bold = False

                    'first check reg recipes, since that's where the majority of recipes will come from
                    Dim lunchRecipe = _
                        From nm In recipeFile.<Recipes>.<RegEntrees>.Elements("Recipe")
                        Where CStr(nm.Element("EntreeName")) = lunchEntree
                        Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                    'if that doesn't work, then check temp recipes
                    If lunchRecipe.Count = 0 Then
                        lunchRecipe = _
                            From nm In recipeFile.<Recipes>.<TempEntrees>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = lunchEntree
                            Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                    End If
                    lunchIngredients.Add("For " + lunchEntree + ":")
                    lunchInstructions.Add("For " + lunchEntree + ":")
                    For Each element In lunchRecipe(0).Elements
                        If element.Name = "Ingredient" Then
                            Dim ingredientString As String = ""
                            Dim ingredientItems = element.Descendants()
                            For y = 0 To 2
                                If y <> 0 Then
                                    ingredientString += " "
                                End If
                                ingredientString += ingredientItems(y).Value
                            Next
                            lunchIngredients.Add(ingredientString)
                        Else
                            lunchInstructions.Add(element.Value)
                        End If
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
                                Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                        End If
                        lunchIngredients.Add("For " + lunchSide1 + ":")
                        lunchInstructions.Add("For " + lunchSide1 + ":")
                        For Each element In lunchSide1Recipe(0).Elements
                            If element.Name = "Ingredient" Then
                                Dim ingredientString As String = ""
                                Dim ingredientItems = element.Descendants()
                                For y = 0 To 2
                                    If y <> 0 Then
                                        ingredientString += " "
                                    End If
                                    ingredientString += ingredientItems(y).Value
                                Next
                                lunchIngredients.Add(ingredientString)
                            Else
                                lunchInstructions.Add(element.Value)
                            End If
                        Next
                    End If

                    If lunchSide2 <> "" Then
                        'first check reg recipes, since that's where the majority of recipes will come from
                        Dim lunchSide2Recipe = _
                            From nm In recipeFile.<Recipes>.<RegSides>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = lunchSide2
                            Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                        'if that doesn't work, then check temp recipes
                        If lunchSide2Recipe.Count = 0 Then
                            lunchSide2Recipe = _
                                From nm In recipeFile.<Recipes>.<TempSides>.Elements("Recipe")
                                Where CStr(nm.Element("EntreeName")) = lunchSide2
                                Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                        End If
                        lunchIngredients.Add("For " + lunchSide2 + ":")
                        lunchInstructions.Add("For " + lunchSide2 + ":")
                        For Each element In lunchSide2Recipe(0).Elements
                            If element.Name = "Ingredient" Then
                                Dim ingredientString As String = ""
                                Dim ingredientItems = element.Descendants()
                                For y = 0 To 2
                                    If y <> 0 Then
                                        ingredientString += " "
                                    End If
                                    ingredientString += ingredientItems(y).Value
                                Next
                                lunchIngredients.Add(ingredientString)
                            Else
                                lunchInstructions.Add(element.Value)
                            End If
                        Next
                    End If
                    Dim ingredientStringToList As String = ""
                    For Each item In lunchIngredients
                        ingredientStringToList += (item + ", ")
                    Next
                    ingredientStringToList = ingredientStringToList.Substring(0, ingredientStringToList.Length - 2)
                    para.Range.Text = ingredientStringToList
                    For Each item In lunchInstructions
                        para.Range.Text += item
                    Next
                    para.Range.InsertParagraphAfter()
                End If


                If meals(0).ElementAt(1) IsNot Nothing Then
                    Dim dinnerItems = meals(0).ElementAt(1).Value.Split(vbLf)
                    dinnerEntree = dinnerItems(1).Substring(6)
                    dinnerSide1 = dinnerItems(2).Substring(6)
                    dinnerSide2 = dinnerItems(3).Substring(6)

                    para.Range.Font.Bold = True
                    If dinnerSide1 <> "" Then
                        If dinnerSide2 <> "" Then
                            para.Range.Text = "Dinner: " + dinnerEntree + " with " + dinnerSide1 + " and " + dinnerSide2
                        Else
                            para.Range.Text = "Dinner: " + dinnerEntree + " with " + dinnerSide1
                        End If
                    Else
                        para.Range.Text = "Dinner: " + dinnerEntree
                    End If
                    para.Range.InsertParagraphAfter()
                    para.Range.Font.Bold = False

                    'first check reg recipes, since that's where the majority of recipes will come from
                    Dim dinnerRecipe = _
                        From nm In recipeFile.<Recipes>.<RegEntrees>.Elements("Recipe")
                        Where CStr(nm.Element("EntreeName")) = dinnerEntree
                        Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                    'if that doesn't work, then check temp recipes
                    If dinnerRecipe.Count = 0 Then
                        dinnerRecipe = _
                            From nm In recipeFile.<Recipes>.<TempEntrees>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = dinnerEntree
                            Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                    End If
                    dinnerIngredients.Add("For " + dinnerEntree + ":")
                    dinnerInstructions.Add("For " + dinnerEntree + ":")
                    For Each element In dinnerRecipe(0).Elements
                        If element.Name = "Ingredient" Then
                            Dim ingredientString As String = ""
                            Dim ingredientItems = element.Descendants()
                            For y = 0 To 2
                                If y <> 0 Then
                                    ingredientString += " "
                                End If
                                ingredientString += ingredientItems(y).Value
                            Next
                            dinnerIngredients.Add(ingredientString)
                        Else
                            dinnerInstructions.Add(element.Value)
                        End If
                    Next

                    If dinnerSide1 <> "" Then
                        'first check reg recipes, since that's where the majority of recipes will come from
                        Dim dinnerSide1Recipe = _
                            From nm In recipeFile.<Recipes>.<RegSides>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = dinnerSide1
                            Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                        'if that doesn't work, then check temp recipes
                        If dinnerSide1Recipe.Count = 0 Then
                            dinnerSide1Recipe = _
                                From nm In recipeFile.<Recipes>.<TempSides>.Elements("Recipe")
                                Where CStr(nm.Element("EntreeName")) = dinnerSide1
                                Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                        End If
                        dinnerIngredients.Add("For " + dinnerSide1 + ":")
                        dinnerInstructions.Add("For " + dinnerSide1 + ":")
                        For Each element In dinnerSide1Recipe(0).Elements
                            If element.Name = "Ingredient" Then
                                Dim ingredientString As String = ""
                                Dim ingredientItems = element.Descendants()
                                For y = 0 To 2
                                    If y <> 0 Then
                                        ingredientString += " "
                                    End If
                                    ingredientString += ingredientItems(y).Value
                                Next
                                dinnerIngredients.Add(ingredientString)
                            Else
                                dinnerInstructions.Add(element.Value)
                            End If
                        Next
                    End If

                    If dinnerSide2 <> "" Then
                        'first check reg recipes, since that's where the majority of recipes will come from
                        Dim dinnerSide2Recipe = _
                            From nm In recipeFile.<Recipes>.<RegSides>.Elements("Recipe")
                            Where CStr(nm.Element("EntreeName")) = dinnerSide2
                            Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                        'if that doesn't work, then check temp recipes
                        If dinnerSide2Recipe.Count = 0 Then
                            dinnerSide2Recipe = _
                                From nm In recipeFile.<Recipes>.<TempSides>.Elements("Recipe")
                                Where CStr(nm.Element("EntreeName")) = dinnerSide2
                                Select {nm.Element("Ingredients"), nm.Element("Instructions")}
                        End If
                        dinnerIngredients.Add("For " + dinnerSide2 + ":")
                        dinnerInstructions.Add("For " + dinnerSide2 + ":")
                        For Each element In dinnerSide2Recipe(0).Elements
                            If element.Name = "Ingredient" Then
                                Dim ingredientString As String = ""
                                Dim ingredientItems = element.Descendants()
                                For y = 0 To 2
                                    If y <> 0 Then
                                        ingredientString += " "
                                    End If
                                    ingredientString += ingredientItems(y).Value
                                Next
                                dinnerIngredients.Add(ingredientString)
                            Else
                                dinnerInstructions.Add(element.Value)
                            End If
                        Next
                    End If
                    Dim ingredientStringToList As String = ""
                    For Each item In dinnerIngredients
                        ingredientStringToList += (item + ", ")
                    Next
                    ingredientStringToList = ingredientStringToList.Substring(0, ingredientStringToList.Length - 2)
                    para.Range.Text = ingredientStringToList
                    For Each item In dinnerInstructions
                        para.Range.Text += item
                    Next
                    para.Range.InsertParagraphAfter()
                End If
            End If


                para.Range.InsertParagraphAfter()
        Next
        word.Visible = True

        If printShoppingListFlag = True Then
            printShoppingList(startDate, endDate)
        End If

        Return Nothing
    End Function

    Private Function printShoppingList(startDate As Date, endDate As Date)
        shoppingListPopup.printShoppingList(startDate, endDate)
        Return Nothing
    End Function
End Class