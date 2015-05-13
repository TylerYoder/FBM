Imports Microsoft.Office
Imports System.IO
Imports System.Xml
Imports Microsoft.Office.Interop


Public Class main

    Friend Shared regPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\FBM\Meals\Regular Meals\"
    Friend Shared tempPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\FBM\Meals\Temp Shift Meals\"
    Friend Shared tempShiftFlag As Boolean = False
    Friend Shared mealDays(1, 6) As Boolean
    Private Property curMonth As Integer = Date.Now.Month
    Private Property curYear As Integer = Date.Now.Year
    Shared Property startPrintDay As String
    Dim settings_file As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\FBMSettings.txt"




    Private Sub main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.SuspendLayout()

        'Google Setup

        'Home tab
        generateMonth(0)

        'Settings tab
        loadSettings()

        'Budget tab

        'Help tab

    End Sub

    'handles the wageTab button click. 
    Private Sub wageButton_Click(sender As Object, e As EventArgs) Handles wageButton.Click

        'verify form is filled out
        If (yearsList.SelectedItem() = "" Or hallsList.SelectedItem() = "" Or hoursList.SelectedItem() = "") Then
            MsgBox("Please select a value for each of the three dropdown tabs on the left.")
            Exit Sub
        End If

        Dim stipend As Integer = 50
        Dim housingCost As Integer = 0
        Dim hoursWorked As Integer = 0

        'set stipend to correct amount ($10 raise each year)
        stipend += 10 * (yearsList.SelectedItem() - 1)

        'determine housing cost depending on hall
        If (hallsList.SelectedItem() = "Reiger" Or hallsList.SelectedItem() = "Krehbiel") Then
            housingCost = 6154
        ElseIf (hallsList.SelectedItem() = "Miller" Or hallsList.SelectedItem() = "Watkins") Then
            housingCost = 2656
        Else
            housingCost = 5798
        End If

        'total amount earned = stipend * 18 biweeks + housingCost
        Dim totalEarned As Integer = stipend * 18 + housingCost

        'total hours worked = hoursWorked * 36 weeks
        hoursWorked = (hoursList.SelectedIndex + 1) * 5 * 36


        Dim hourlyWage As Double = Math.Round(totalEarned / hoursWorked, 2)
        Dim hourlyWage247 As Double = Math.Round(totalEarned / 6048, 2) '6048 = total hours worked in a given academic year 24/7

        'populate fields
        totalEarningsBox.Text() = "$" + totalEarned.ToString() + ".00"
        hourlyWageBox.Text() = "$" + hourlyWage.ToString()
        Wage247Box.Text() = "$" + hourlyWage247.ToString()
        mcdonaldsBox.Text() = Math.Round((hourlyWage / 7.25) * (hoursWorked / 36), 2)
        officeJobBox.Text() = Math.Round((hourlyWage / 12.0) * (hoursWorked / 36), 2)
        ceoBox.Text() = Math.Round((hourlyWage / 1805.56) * (hoursWorked / 36), 2)

    End Sub

    'paints the calendar in homeTab. 
    Private Sub calendarPanel_CellPaint(sender As Object, e As TableLayoutCellPaintEventArgs) Handles calendarPanel.CellPaint
        If e.Row = 1 Then
            e.Graphics.FillRectangle(Brushes.Blue, e.CellBounds)
        End If
        If (e.Row = 2 Or e.Row = 4 Or e.Row = 6 Or e.Row = 8 Or e.Row = 10 Or e.Row = 12) Then
            e.Graphics.FillRectangle(Brushes.LightBlue, e.CellBounds)
        End If
        If (month.Text = DateAndTime.MonthName(Today.Month) + " " + Today.Year.ToString) Then
            If (calendarPanel.GetControlFromPosition(e.Column, e.Row).Text = Today.Day.ToString) Then
                e.Graphics.FillRectangle(Brushes.Red, e.CellBounds)
            End If
        End If
    End Sub

    'changes the dates and month name in the calendar, shifted from current month by shift argument
    Private Sub generateMonth(shift As Integer)
        curMonth += shift
        'catch december -> january shift
        If (curMonth > 12) Then
            curMonth = 1 'set to january
            curYear += 1 'increment year
        End If
        'catch january -> december shift
        If (curMonth < 1) Then
            curMonth = 12 'set to december
            curYear -= 1 'decrement year
        End If

        'set monthLabel to reflect new month/year
        month.Text = DateAndTime.MonthName(curMonth, False) + " " + curYear.ToString

        'number of days in month
        Dim daysInMonth As Integer = DateTime.DaysInMonth(curYear, curMonth)

        'day of week that month starts on. 1=sunday, 2=monday, ...7=saturday
        Dim firstDayOfMonth As Integer = DateAndTime.Weekday(Date.Parse(curMonth.ToString + "/01/" + curYear.ToString, Globalization.CultureInfo.InvariantCulture)) - 1

        'set date numbers on calendar
        Dim curDateNum As Integer = 1

        For row = 2 To 13
            For col = 0 To 6
                If (row Mod 2 = 0) Then
                    Dim x As Control = calendarPanel.GetControlFromPosition(col, row)
                    x.Text = "" 'first clear everything
                    Dim text As Control = calendarPanel.GetControlFromPosition(col, row + 1)
                    text.Text = ""
                    calendarPanel.GetControlFromPosition(col, row + 1).Text = ""
                    If (curDateNum <= daysInMonth) Then
                        If (row = 2 And col >= firstDayOfMonth) Then  'get corner case of first row
                            x.Text = curDateNum
                            curDateNum += 1
                        ElseIf (row <> 2) Then
                            x.Text = curDateNum
                            curDateNum += 1
                        End If
                    Else
                    End If
                End If
            Next
        Next
        writeMenu()
    End Sub

    'handles the next button in the calendar
    Private Sub nextButton_Click(sender As Object, e As EventArgs) Handles nextButton.Click
        generateMonth(1)
    End Sub

    'handles the prev button in the calendar
    Private Sub prevButton_Click(sender As Object, e As EventArgs) Handles prevButton.Click
        generateMonth(-1)
    End Sub

    'processes settings and updates settings file
    Private Sub changeSettingsButton_Click(sender As Object, e As EventArgs) Handles changeSettingsButton.Click
        If (System.IO.File.Exists(settings_file) <> True) Then
            System.IO.File.Create(settings_file).Dispose()
        End If
        Dim objWriter As New System.IO.StreamWriter(settings_file)
        objWriter.WriteLine(budgetLocationBox.Text)
        objWriter.WriteLine(lunchRotationBox.Text)
        objWriter.WriteLine(dinnerRotationBox.Text)
        For Each check In LunchDays.CheckedIndices
            objWriter.Write(check.ToString)
        Next
        objWriter.WriteLine()
        For Each check In DinnerDays.CheckedIndices
            objWriter.Write(check.ToString)
        Next
        objWriter.WriteLine()
        objWriter.WriteLine(foodOrderDateBox.Text)
        objWriter.WriteLine(weeklyPrintStartDateBox.Text)
        objWriter.Close()
        MsgBox("Settings Updated.")

    End Sub

    'loads settings from settings file. Currently plaintext jank. 
    Private Sub loadSettings()
        If (System.IO.File.Exists(settings_file) <> True) Then
            System.IO.File.Create(settings_file).Dispose()
            Exit Sub
        End If
        Using reader As New System.IO.StreamReader(settings_file)
            budgetLocationBox.Text = reader.ReadLine
            lunchRotationBox.Text = reader.ReadLine
            dinnerRotationBox.Text = reader.ReadLine
            Dim lunch As String = reader.ReadLine
            For x = 0 To lunch.Length - 1
                LunchDays.SetItemChecked(CInt(lunch.Substring(x, 1)), True)
                mealDays(0, CInt(lunch.Substring(x, 1))) = True
            Next
            Dim dinner As String = reader.ReadLine
            For x = 0 To dinner.Length - 1
                DinnerDays.SetItemChecked(CInt(dinner.Substring(x, 1)), True)
                mealDays(1, CInt(dinner.Substring(x, 1))) = True
            Next
            foodOrderDateBox.Text = reader.ReadLine
            weeklyPrintStartDateBox.Text = reader.ReadLine
            startPrintDay = weeklyPrintStartDateBox.Text
            reader.Close()
        End Using
    End Sub

    'makes menuPopup show when button is clicked
    Private Sub GenerateMenuButton_Click(sender As Object, e As EventArgs) Handles GenerateMenuButton.Click
        menuPopup.ShowDialog()
    End Sub

    'generates recipes file
    Private Function generateRecipes()
        'If recipe list already exists, ask if they want to replace it
        If IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\recipes.xml") = True Then
            Dim regenerateQuestion As Integer = MsgBox("The internal recipe list in this application has not been updated since " + System.IO.File.GetLastWriteTime((Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\recipes.xml")).ToString + ". Would you like to update this list now? (Any changes or new recipes added will not be reflected until you update)", MsgBoxStyle.YesNo, "Re-Generate Recipe List?")
            If regenerateQuestion = DialogResult.No Then
                Return Nothing
            End If
        End If
        Dim xmlSettings As New Xml.XmlWriterSettings
        xmlSettings.Indent = True
        xmlSettings.ConformanceLevel = ConformanceLevel.Auto
        Using writer As XmlWriter = XmlWriter.Create((Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\recipes.xml"), xmlSettings)
            Dim excel As New Excel.Application()

            Dim fileList As New List(Of String)
            Dim tempList As New List(Of String)
            tempList.AddRange(Directory.GetFiles(regPath))
            tempList.AddRange(Directory.GetFiles(tempPath))
            For Each file As String In tempList
                If (file.Contains("~$") = False) Then
                    fileList.Add(file)
                End If
            Next

            'create xml structure
            writer.WriteStartDocument()
            writer.WriteStartElement("Recipes")

            'parse through recipes for each set, creating nodes for ones that match each set
            'regular entrees
            writer.WriteStartElement("RegEntrees")
            For Each file In fileList.ToList
                Dim recipeFile As Excel.Workbook = excel.Workbooks.Open(file)
                Dim recipe As Excel.Worksheet = excel.Sheets("Meal Form")
                Dim recipeStructure As Excel.Worksheet = excel.Sheets("Structure")
                If (file.Substring(0, file.Length - recipeFile.Name.Length) = regPath And recipe.Range("C3").Value2.ToString.ToLower = "entrée") Then
                    writer.WriteStartElement("Recipe")
                    writer.WriteElementString("EntreeName", recipeFile.Name.Substring(0, recipeFile.Name.Length - 5))
                    writer.WriteElementString("Author", recipe.Range("C2").Text)
                    writer.WriteElementString("Time", recipe.Range("C4").Text)
                    writer.WriteElementString("Day", recipe.Range("C5").Text)
                    writer.WriteElementString("Side1", recipeStructure.Range("A22").Text)
                    writer.WriteElementString("Side2", recipeStructure.Range("A23").Text)
                    writer.WriteStartElement("Ingredients")
                    For x = 0 To 23 'number of ingredients rows in excel minus 1
                        Dim cellNum As String = (x + 9).ToString
                        If (recipe.Range("A" + cellNum).Value Is Nothing And recipe.Range("B" + cellNum).Value Is Nothing And recipe.Range("C" + cellNum).Value Is Nothing) Then
                            Exit For
                        Else
                            writer.WriteStartElement("Ingredient")
                            writer.WriteElementString("QTY", recipe.Range("A" + cellNum).Text)
                            writer.WriteElementString("QTYType", recipe.Range("B" + cellNum).Text)
                            writer.WriteElementString("IngredientName", recipe.Range("C" + cellNum).Text)
                            writer.WriteEndElement() 'ingredient
                        End If
                    Next
                    writer.WriteEndElement() 'ingredients
                    writer.WriteStartElement("Instructions")
                    For x = 0 To 13
                        Dim cellNum As String = (x + 36).ToString
                        If (recipe.Range("A" + cellNum).Text = "" And recipe.Range("B" + cellNum).Text = "") Then
                            Exit For
                        Else
                            writer.WriteElementString("Instruction", recipe.Range("A" + cellNum).Text + ". " + recipe.Range("B" + cellNum).Text)
                        End If
                    Next
                    writer.WriteEndElement() 'instructions
                    writer.WriteEndElement() 'recipe name
                    fileList.Remove(file)
                End If
                recipeFile.Close(False)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Next
            writer.WriteEndElement() 'regentrees

            'temporary entrees
            writer.WriteStartElement("TempEntrees")
            For Each file In fileList.ToList
                Dim recipeFile As Excel.Workbook = excel.Workbooks.Open(file)
                Dim recipe As Excel.Worksheet = excel.Sheets("Meal Form")
                Dim recipeStructure As Excel.Worksheet = excel.Sheets("Structure")
                If (file.Substring(0, file.Length - recipeFile.Name.Length) = tempPath And recipe.Range("C3").Value2.ToString.ToLower = "entrée") Then
                    writer.WriteStartElement("Recipe")
                    writer.WriteElementString("EntreeName", recipeFile.Name.Substring(0, recipeFile.Name.Length - 5))
                    writer.WriteElementString("Author", recipe.Range("C2").Text)
                    writer.WriteElementString("Time", recipe.Range("C4").Text)
                    writer.WriteElementString("Day", recipe.Range("C5").Text)
                    writer.WriteElementString("Side1", recipeStructure.Range("A22").Text)
                    writer.WriteElementString("Side2", recipeStructure.Range("A23").Text)
                    writer.WriteStartElement("Ingredients")
                    For x = 0 To 23 'number of ingredients rows in excel minus 1
                        Dim cellNum As String = (x + 9).ToString
                        If (recipe.Range("A" + cellNum).Value Is Nothing And recipe.Range("B" + cellNum).Value Is Nothing And recipe.Range("C" + cellNum).Value Is Nothing) Then
                            Exit For
                        Else
                            writer.WriteStartElement("Ingredient")
                            writer.WriteElementString("QTY", recipe.Range("A" + cellNum).Text)
                            writer.WriteElementString("QTYType", recipe.Range("B" + cellNum).Text)
                            writer.WriteElementString("IngredientName", recipe.Range("C" + cellNum).Text)
                            writer.WriteEndElement() 'ingredient
                        End If
                    Next
                    writer.WriteEndElement() 'ingredients
                    writer.WriteStartElement("Instructions")
                    For x = 0 To 13
                        Dim cellNum As String = (x + 36).ToString
                        If (recipe.Range("A" + cellNum).Text = "" And recipe.Range("B" + cellNum).Text = "") Then
                            Exit For
                        Else
                            writer.WriteElementString("Instruction", recipe.Range("A" + cellNum).Text + ". " + recipe.Range("B" + cellNum).Text)
                        End If
                    Next
                    writer.WriteEndElement() 'instructions
                    writer.WriteEndElement() 'recipe name
                    fileList.Remove(file)
                End If
                recipeFile.Close(False)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Next
            writer.WriteEndElement() 'tempentrees

            'regular sides
            writer.WriteStartElement("RegSides")
            For Each file In fileList.ToList
                Dim recipeFile As Excel.Workbook = excel.Workbooks.Open(file)
                Dim recipe As Excel.Worksheet = excel.Sheets("Meal Form")
                Dim recipeStructure As Excel.Worksheet = excel.Sheets("Structure")
                If (file.Substring(0, file.Length - recipeFile.Name.Length) = regPath And recipe.Range("C3").Value2.ToString.ToLower = "side") Then
                    writer.WriteStartElement("Recipe")
                    writer.WriteElementString("EntreeName", recipeFile.Name.Substring(0, recipeFile.Name.Length - 5))
                    writer.WriteElementString("Author", recipe.Range("C2").Text)
                    writer.WriteElementString("Time", recipe.Range("C4").Text)
                    writer.WriteElementString("Day", recipe.Range("C5").Text)
                    writer.WriteElementString("Side1", recipeStructure.Range("A22").Text)
                    writer.WriteElementString("Side2", recipeStructure.Range("A23").Text)
                    writer.WriteStartElement("Ingredients")
                    For x = 0 To 23 'number of ingredients rows in excel minus 1
                        Dim cellNum As String = (x + 9).ToString
                        If (recipe.Range("A" + cellNum).Value Is Nothing And recipe.Range("B" + cellNum).Value Is Nothing And recipe.Range("C" + cellNum).Value Is Nothing) Then
                            Exit For
                        Else
                            writer.WriteStartElement("Ingredient")
                            writer.WriteElementString("QTY", recipe.Range("A" + cellNum).Text)
                            writer.WriteElementString("QTYType", recipe.Range("B" + cellNum).Text)
                            writer.WriteElementString("IngredientName", recipe.Range("C" + cellNum).Text)
                            writer.WriteEndElement() 'ingredient
                        End If
                    Next
                    writer.WriteEndElement() 'ingredients
                    writer.WriteStartElement("Instructions")
                    For x = 0 To 13
                        Dim cellNum As String = (x + 36).ToString
                        If (recipe.Range("A" + cellNum).Text = "" And recipe.Range("B" + cellNum).Text = "") Then
                            Exit For
                        Else
                            writer.WriteElementString("Instruction", recipe.Range("A" + cellNum).Text + ". " + recipe.Range("B" + cellNum).Text)
                        End If
                    Next
                    writer.WriteEndElement() 'instructions
                    writer.WriteEndElement() 'recipe name
                    fileList.Remove(file)
                End If
                recipeFile.Close(False)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Next
            writer.WriteEndElement() 'regsides

            'temp sides
            writer.WriteStartElement("TempSides")
            For Each file In fileList.ToList
                Dim recipeFile As Excel.Workbook = excel.Workbooks.Open(file)
                Dim recipe As Excel.Worksheet = excel.Sheets("Meal Form")
                Dim recipeStructure As Excel.Worksheet = excel.Sheets("Structure")
                If (file.Substring(0, file.Length - recipeFile.Name.Length) = tempPath And recipe.Range("C3").Value2.ToString.ToLower = "side") Then
                    writer.WriteStartElement("Recipe")
                    writer.WriteElementString("EntreeName", recipeFile.Name.Substring(0, recipeFile.Name.Length - 5))
                    writer.WriteElementString("Author", recipe.Range("C2").Text)
                    writer.WriteElementString("Time", recipe.Range("C4").Text)
                    writer.WriteElementString("Day", recipe.Range("C5").Text)
                    writer.WriteElementString("Side1", recipeStructure.Range("A22").Text)
                    writer.WriteElementString("Side2", recipeStructure.Range("A23").Text)
                    writer.WriteStartElement("Ingredients")
                    For x = 0 To 23 'number of ingredients rows in excel minus 1
                        Dim cellNum As String = (x + 9).ToString
                        If (recipe.Range("A" + cellNum).Value Is Nothing And recipe.Range("B" + cellNum).Value Is Nothing And recipe.Range("C" + cellNum).Value Is Nothing) Then
                            Exit For
                        Else
                            writer.WriteStartElement("Ingredient")
                            writer.WriteElementString("QTY", recipe.Range("A" + cellNum).Text)
                            writer.WriteElementString("QTYType", recipe.Range("B" + cellNum).Text)
                            writer.WriteElementString("IngredientName", recipe.Range("C" + cellNum).Text)
                            writer.WriteEndElement() 'ingredient
                        End If
                    Next
                    writer.WriteEndElement() 'ingredients
                    writer.WriteStartElement("Instructions")
                    For x = 0 To 13
                        Dim cellNum As String = (x + 36).ToString
                        If (recipe.Range("A" + cellNum).Text = "" And recipe.Range("B" + cellNum).Text = "") Then
                            Exit For
                        Else
                            writer.WriteElementString("Instruction", recipe.Range("A" + cellNum).Text + ". " + recipe.Range("B" + cellNum).Text)
                        End If
                    Next
                    writer.WriteEndElement() 'instructions
                    writer.WriteEndElement() 'recipe name
                    fileList.Remove(file)
                End If
                recipeFile.Close(False)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Next
            writer.WriteEndElement() 'tempsides
            writer.WriteEndElement() 'recipes

            'close document
            writer.WriteEndDocument()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            excel.Quit()
        End Using
        Return Nothing
    End Function

    'generates menu based on recipes in recipe file and dates selected
    Shared Function generateMenu(startDate As Date, endDate As Date)
        main.generateRecipes()
        'If menu already exists, ask if they want to replace it
        If IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\menu.xml") = True Then
            Dim regenerateQuestion As Integer = MsgBox("You are about to create a new menu. This will overwrite your current menu. Proceed?", MsgBoxStyle.YesNo, "Re-Generate Recipe List?")
            If regenerateQuestion = DialogResult.No Then
                Return Nothing
            End If
        End If
        Dim recipeFile As XDocument = XDocument.Load((Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\recipes.xml"), LoadOptions.PreserveWhitespace)
        Dim xmlSettings As New Xml.XmlWriterSettings
        xmlSettings.Indent = True
        xmlSettings.ConformanceLevel = ConformanceLevel.Auto
        Using writer As XmlWriter = XmlWriter.Create((Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\menu.xml"), xmlSettings)
            Dim excel As New Excel.Application()
            writer.WriteStartDocument()
            writer.WriteStartElement("Calendar")
            Dim tLunches = _
                From nm In recipeFile.<Recipes>.<TempEntrees>.Elements _
                Where CStr(nm.Element("Time")) = "Lunch" _
                Select {CStr(nm.Element("EntreeName")), CStr(nm.Element("Side1")), CStr(nm.Element("Side2")), CStr(nm.Element("Day"))}
            Dim tLunchArr = tLunches.ToArray
            Dim tDinners = _
                From nm In recipeFile.<Recipes>.<TempEntrees>.Elements _
                Where CStr(nm.Element("Time")) = "Dinner" _
                Select {CStr(nm.Element("EntreeName")), CStr(nm.Element("Side1")), CStr(nm.Element("Side2")), CStr(nm.Element("Day"))}
            Dim tDinnerArr = tDinners.ToArray
            Dim rLunches = _
                From nm In recipeFile.<Recipes>.<RegEntrees>.Elements _
                Where CStr(nm.Element("Time")) = "Lunch" _
                Select {CStr(nm.Element("EntreeName")), CStr(nm.Element("Side1")), CStr(nm.Element("Side2")), CStr(nm.Element("Day"))}
            Dim rLunchArr = rLunches.ToArray
            Dim rDinners = _
                From nm In recipeFile.<Recipes>.<RegEntrees>.Elements _
                Where CStr(nm.Element("Time")) = "Dinner" _
                Select {CStr(nm.Element("EntreeName")), CStr(nm.Element("Side1")), CStr(nm.Element("Side2")), CStr(nm.Element("Day"))}
            Dim rDinnerArr = rDinners.ToArray

            'link sides
            linkSides(tLunchArr, tDinnerArr, rLunchArr, rDinnerArr)

            'fix corner case of year change
            Dim numOfDays As Integer
            If (startDate.Year = endDate.Year) Then
                numOfDays = endDate.DayOfYear - startDate.DayOfYear
            Else
                numOfDays = (365 - startDate.DayOfYear) + endDate.DayOfYear
            End If
            For x = 0 To numOfDays 'for each day
                writer.WriteStartElement("Day")
                writer.WriteStartElement("Date")
                writer.WriteElementString("Month", startDate.AddDays(x).Month.ToString)
                writer.WriteElementString("Day", startDate.AddDays(x).Day.ToString)
                writer.WriteElementString("Year", startDate.AddDays(x).Year.ToString)
                writer.WriteElementString("Weekday", startDate.AddDays(x).DayOfWeek.ToString)
                writer.WriteEndElement() 'date

                If tempShiftFlag = True Then
                    If mealDays(0, startDate.AddDays(x).DayOfWeek) = True Then
                        writer.WriteStartElement("Lunch")
                        writer.WriteElementString("Entree", tLunchArr(x Mod tLunchArr.Length)(0))
                        writer.WriteElementString("Side1", tLunchArr(x Mod tLunchArr.Length)(1))
                        writer.WriteElementString("Side2", tLunchArr(x Mod tLunchArr.Length)(2))
                        writer.WriteEndElement() 'lunch
                    End If
                    If mealDays(1, startDate.AddDays(x).DayOfWeek) = True Then
                        writer.WriteStartElement("Dinner")
                        writer.WriteElementString("Entree", tDinnerArr(x Mod tDinnerArr.Length)(0))
                        writer.WriteElementString("Side1", tDinnerArr(x Mod tDinnerArr.Length)(1))
                        writer.WriteElementString("Side2", tDinnerArr(x Mod tDinnerArr.Length)(2))
                        writer.WriteEndElement() 'dinner
                    End If

                ElseIf tempShiftFlag = False Then
                    Dim subRLunchArr(rLunchArr.Length - 1, 3) As String
                    Dim subRDinnerArr(rDinnerArr.Length - 1, 3) As String
                    Dim counter As Integer = 0
                    For y = 0 To rLunchArr.Length - 1
                        If rLunchArr(y)(3) = startDate.AddDays(x).DayOfWeek.ToString Then
                            subRLunchArr(counter, 0) = rLunchArr(y)(0)
                            subRLunchArr(counter, 1) = rLunchArr(y)(1)
                            subRLunchArr(counter, 2) = rLunchArr(y)(2)
                            subRLunchArr(counter, 3) = rLunchArr(y)(3)
                            counter += 1
                        End If
                    Next
                    counter = 0
                    For y = 0 To rDinnerArr.Length - 1
                        If rDinnerArr(y)(3) = startDate.AddDays(x).DayOfWeek.ToString Then
                            subRDinnerArr(counter, 0) = rDinnerArr(y)(0)
                            subRDinnerArr(counter, 1) = rDinnerArr(y)(1)
                            subRDinnerArr(counter, 2) = rDinnerArr(y)(2)
                            subRDinnerArr(counter, 3) = rDinnerArr(y)(3)
                            counter += 1
                        End If
                    Next
                    If mealDays(0, startDate.AddDays(x).DayOfWeek) = True Then
                        writer.WriteStartElement("Lunch")
                        writer.WriteElementString("Entree", subRLunchArr(x Mod subRLunchArr.Length, 0))
                        writer.WriteElementString("Side1", subRLunchArr(x Mod subRLunchArr.Length, 1))
                        writer.WriteElementString("Side2", subRLunchArr(x Mod subRLunchArr.Length, 2))
                        writer.WriteEndElement() 'lunch
                    End If
                    If mealDays(1, startDate.AddDays(x).DayOfWeek) = True Then
                        writer.WriteStartElement("Dinner")
                        writer.WriteElementString("Entree", subRDinnerArr(x Mod subRDinnerArr.Length, 0))
                        writer.WriteElementString("Side1", subRDinnerArr(x Mod subRDinnerArr.Length, 1))
                        writer.WriteElementString("Side2", subRDinnerArr(x Mod subRDinnerArr.Length, 2))
                        writer.WriteEndElement() 'dinner
                    End If

                End If
                writer.WriteEndElement() 'day
TODO:           'copy current recipe if over x weeks, where x is in settings

            Next
            writer.WriteEndElement() 'calendar
            writer.WriteEndDocument()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            excel.Quit()
        End Using
        main.generateMonth(0)
        Return Nothing
    End Function

    'links up to two sides to an entree if it doesn't currently have sides
    Private Shared Function sideLinker(entreeName As String, tempShift As Boolean)
        sideLinkerPopup.Label1.Text = entreeName + " does not have side dishes associated with it. Please select the side dishes that you would like to link to it."
        sideLinkerPopup.sideBox1.Items.Clear()
        sideLinkerPopup.sideBox2.Items.Clear()
        sideLinkerPopup.sideBox1.Text = ""
        sideLinkerPopup.sideBox2.Text = ""
        Dim recipes As XDocument = XDocument.Load((Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\recipes.xml"), LoadOptions.PreserveWhitespace)

        If tempShift = False Then
            Dim sidelist As IEnumerable(Of XElement) = recipes.<Recipes>.<RegSides>.Elements
            For Each side In sidelist
                sideLinkerPopup.sideBox1.Items.Add(side.Element("EntreeName").Value)
                sideLinkerPopup.sideBox2.Items.Add(side.Element("EntreeName").Value)
            Next
        Else
            Dim sidelist As IEnumerable(Of XElement) = recipes.<Recipes>.<TempSides>.Elements
            For Each side In sidelist
                sideLinkerPopup.sideBox1.Items.Add(side.Element("EntreeName").Value)
                sideLinkerPopup.sideBox2.Items.Add(side.Element("EntreeName").Value)
            Next
        End If
        sideLinkerPopup.ShowDialog()
        Dim side1 As String = sideLinkerPopup.sideBox1.Text
        Dim side2 As String = sideLinkerPopup.sideBox2.Text

        'save into excel - hidden Structure sheet, cells A22 and A23
        Dim excel As New Microsoft.Office.Interop.Excel.Application()
        Dim path As String = ""
        If tempShift = True Then
            path = tempPath
        Else
            path = regPath
        End If
        Dim file = excel.Workbooks.Open(path + entreeName + ".xlsx", , False)
        Dim recipeStructureSheet As Microsoft.Office.Interop.Excel.Worksheet = TryCast(excel.Sheets("Structure"), Microsoft.Office.Interop.Excel.Worksheet)
        recipeStructureSheet.Unprotect("FBMlocked")
        recipeStructureSheet.Range("A22").Value2 = side1
        recipeStructureSheet.Range("A23").Value2 = side2
        recipeStructureSheet.Protect("FBMlocked")
        file.Close(True)
        excel.Quit()
        GC.Collect()
        GC.WaitForPendingFinalizers()

        'return
        Dim sides As New List(Of String)
        sides.Add(side1)
        sides.Add(side2)
        Return sides
    End Function

    'links all the sides at once
    Private Shared Function linkSides(tLunchArr As String()(), tDinnerArr As String()(), rLunchArr As String()(), rDinnerArr As String()())
        For x = 0 To tLunchArr.Length - 1
            If tLunchArr(x)(1) = "" And tLunchArr(x)(2) = "" Then
                Dim sides As List(Of String) = sideLinker(tLunchArr(x)(0), True)

            End If
        Next
        For x = 0 To tDinnerArr.Length - 1
            If tDinnerArr(x)(1) = "" And tDinnerArr(x)(2) = "" Then
                Dim sides As List(Of String) = sideLinker(tDinnerArr(x)(0), True)
            End If
        Next
        For x = 0 To rLunchArr.Length - 1
            If rLunchArr(x)(1) = "" And rLunchArr(x)(2) = "" Then
                Dim sides As List(Of String) = sideLinker(rLunchArr(x)(0), False)
            End If
        Next
        For x = 0 To rDinnerArr.Length - 1
            If rDinnerArr(x)(1) = "" And rDinnerArr(x)(2) = "" Then
                Dim sides As List(Of String) = sideLinker(rDinnerArr(x)(0), False)
            End If
        Next
        Return Nothing
    End Function

    'writes menu items to the calendar
    Private Shared Function writeMenu()
        If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\menu.xml") = False Then
            Return Nothing
        End If

        Dim curMonth = main.curMonth
        Dim curYear = main.curYear
        Dim numOfDays As Integer = Date.DaysInMonth(curYear, curMonth)
        Dim firstDayOfMonth As Integer = DateAndTime.Weekday(Date.Parse(curMonth.ToString + "/01/" + curYear.ToString, Globalization.CultureInfo.InvariantCulture)) - 1

        Dim curRow = 3
        Dim curCol = firstDayOfMonth
        Dim menuFile As XDocument = XDocument.Load((Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\menu.xml"), LoadOptions.PreserveWhitespace)
        For x = 1 To numOfDays
            Dim day = x
            Dim meals = _
                From nm In menuFile.<Calendar>.<Day>.Elements("Date") _
                Where CStr(nm.Element("Year")) = curYear.ToString And CStr(nm.Element("Day")) = day.ToString And CStr(nm.Element("Month")) = curMonth.ToString
            Select {nm.Parent().Element("Lunch"), nm.Parent().Element("Dinner")}
            If meals.Count <> 0 Then
                If meals(0).ElementAt(0) IsNot Nothing Then
                    Dim lunchItems = meals(0).ElementAt(0).Value.Split(vbLf)
                    Dim lunchEntree = lunchItems(1).Substring(6)
                    Dim lunchSide1 = lunchItems(2).Substring(6)
                    Dim lunchSide2 = lunchItems(3).Substring(6)
                    If lunchSide1 <> "" Then
                        If lunchSide2 <> "" Then
                            main.calendarPanel.GetControlFromPosition(curCol, curRow).Text = "Lunch: " + lunchEntree + " with " + lunchSide1 + " and " + lunchSide2
                        Else
                            main.calendarPanel.GetControlFromPosition(curCol, curRow).Text = "Lunch: " + lunchEntree + " with " + lunchSide1
                        End If
                    Else
                        main.calendarPanel.GetControlFromPosition(curCol, curRow).Text = "Lunch: " + lunchEntree
                    End If
                End If
                If meals(0).ElementAt(1) IsNot Nothing Then
                    Dim dinnerItems = meals(0).ElementAt(1).Value.Split(vbLf)
                    Dim dinnerEntree = dinnerItems(1).Substring(6)
                    Dim dinnerSide1 = dinnerItems(2).Substring(6)
                    Dim dinnerSide2 = dinnerItems(3).Substring(6)
                    If dinnerSide1 <> "" Then
                        If dinnerSide2 <> "" Then
                            main.calendarPanel.GetControlFromPosition(curCol, curRow).Text += vbNewLine + vbNewLine + "Dinner: " + dinnerEntree + " with " + dinnerSide1 + " and " + dinnerSide2
                        Else
                            main.calendarPanel.GetControlFromPosition(curCol, curRow).Text += vbNewLine + vbNewLine + "Dinner: " + dinnerEntree + " with " + dinnerSide1
                        End If
                    Else
                        main.calendarPanel.GetControlFromPosition(curCol, curRow).Text += vbNewLine + vbNewLine + "Dinner: " + dinnerEntree
                    End If
                End If
            End If

            If curCol >= 6 Then
                curCol = 0
                curRow += 2
            Else
                curCol += 1
            End If
        Next

        Return Nothing
    End Function

    'makes printWeeklyPopup show when button clicked
    Private Sub printWeeklyButton_Click(sender As Object, e As EventArgs) Handles printWeeklyButton.Click
        printWeeklyPopup.ShowDialog()
    End Sub

    Private Sub shoppingListButton_Click(sender As Object, e As EventArgs) Handles shoppingListButton.Click
        shoppingListPopup.ShowDialog()
    End Sub

    Private Sub SetCoHallButton_Click(sender As Object, e As EventArgs) Handles SetCoHallButton.Click
        coHallListPopup.ShowDialog()
    End Sub
End Class
