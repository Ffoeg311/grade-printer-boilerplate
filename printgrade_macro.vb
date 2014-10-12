'VB Macro to arrange grades below the spreadsheet in a printable format.
'Released under the MIT license by Geoff Porter
'To view the entire project and license, go to https://github.com/Ffoeg311/grade-printer-boilerplate.git
Sub printGrade()
	MsgBox "Starting, make sure you have selected the first student's name selected"
    
	'STEP # Determine number of questions**************************************************************************
    'Go to the Cell where the questions start
	ActiveCell.Offset(0, 3).Select 
    questions = 0 'count of the total number of questions
    'Iterate over the row, incrementing number of questions
	continueQuCounting = True 'looping variable
	Do While continueQuCounting = True 'Count the questions while there is a value in the active cell        
		If ActiveCell.Value = 0 Then 'This condition occurs when you are at the end of the row
            continueQuCounting = False
        Else
            questions = questions + 1 'Add one to the count of questions
            ActiveCell.Offset(0, 1).Select 'And move to the right one cell
        End If
	Loop
	
	'STEP # Count the students ***************************************************************************************
	ActiveCell.Offset(0, -1 * (questions + 3)).Select 'Go to the first name
	nameCount = 0
	continueNameCounting = True
	'MEAT
	Do While continueNameCounting = True
        If ActiveCell.Value <> "# CORRECT" Then
            ActiveCell.Offset(1, 0).Select
            nameCount = nameCount + 1
        Else
            continueNameCounting = False
        End If
    Loop
    
    'STEP # Get the name of the test from the user******************************************************************
    testName = InputBox("Please enter a name for the test")    
    
    'STEP # Determine if the answers are right and add them to the view*********************************************    
	verticalDistance = nameCount + 10 'Give ten lines per student. Change this to provide more room
    ActiveCell.Offset(-1 * (nameCount), 0).Select 'Go to the first question
    Do While scoresPrinted < nameCount
		'Reset these variables to 0 each time you print a student's results
		questionsLooked = 0 
		questionsWrong = 0
		questionsRight = 0
		questionNumber = 0
		 
		'Go to the gradesheet and print the top line
		studentName = ActiveCell.Value
		ActiveCell.Offset(verticalDistance, 0).Select 
		printTopLine
		'Add the name to the grade sheet
		ActiveCell.Offset(0, 1).Select 
		ActiveCell.Value = studentName
		ActiveCell.Offset(0, -1).Select
		ActiveCell.Offset(-verticalDistance, 0).Select
		ActiveCell.Offset(0, 2).Select 'Go to the student score
		studentScore = ActiveCell.Value
		'Go to the first question
		ActiveCell.Offset(0, 1).Select 
		Dim questionsRightArray(0 To 100) 'An array of the problems gotten right
		'Clear the array for every student
		For i = 0 To 100 
				questionsRightArray(i) = 0
		Next i
		
		'Iterate over all of the questions and determine it they are right
		arrayPosition = 0 'variable for where you are in the questionsRightArray
		Do While questionsLooked < questions 
			questionsLooked = questionsLooked + 1
			questionNumber = questionNumber + 1
			cellColorIndex = ActiveCell.Interior.ColorIndex
			If cellColorIndex = 3 Then 'Do this code if the answer is wrong, to reverse, change 3 to the green one
				questionsWrong = questionsWrong + 1
			Else 'Add the questionNumber to the array if the answer was right
				questionsRight = questionsRight + 1
				questionsRightArray(arrayPosition) = questionNumber
				arrayPosition = arrayPosition + 1
			End If
			
			If questionsLooked < questions Then
				ActiveCell.Offset(0, 1).Select 'Keep going to the right unless you are on the last loop
			End If
		Loop
			 
		'Step # fill in the printable grade sheet*******************************************************
		'Go to the gradesheet
		ActiveCell.Offset(verticalDistance, -(questions + 1)).Select 
		ActiveCell.Offset(1, 0).Select 'Put in the test name
		ActiveCell.Value = testName
		ActiveCell.Offset(-1, 0).Select
				
		ActiveCell.Offset(2, 0).Select 'Put in the student score
		ActiveCell.Value = studentScore
		ActiveCell.Offset(0, 1).Select
		ActiveCell.Value = "%"
		ActiveCell.Offset(0, -1).Select
		ActiveCell.Offset(-2, 0).Select
		
		'print the question numbers
		ActiveCell.Offset(3, 0).Select 
		howFarOver = 0 'The column of the grade
		maxHowFarOver = 7 'furthest possible column, change for landscape
		rowOutCount = 1 'Number of rows that have been printed
		For i = 0 To (questionsRight - 1)
			If howFarOver < (maxHowFarOver - 2) Then
				ActiveCell.Value = questionsRightArray(i)
				ActiveCell.Offset(0, 1).Select
				howFarOver = howFarOver + 1
			Else
				ActiveCell.Value = questionsRightArray(i)
				ActiveCell.Offset(1, -1 * (howFarOver)).Select           
				howFarOver = 0	
				rowOutCount = rowOutCount + 1
			End If
		Next i
		
		ActiveCell.Offset(-1 * (rowOutCount), -1 * (howFarOver)).Select 'Go back to the first question printed
		ActiveCell.Offset(-3, 0).Select
		ActiveCell.Offset(-(verticalDistance), (questions + 1)).Select 'Go to the next student's data
		scoresPrinted = scoresPrinted + 1
		
		'Go to the first score of the next student          
		ActiveCell.Offset(1, -1 * (questions + 2)).Select 
		ActiveCell.Offset(1, 0).Select
			  
		tempvar = questions / maxHowFarOver
		tempvar = WorksheetFunction.RoundUp(tempvar, 0)
		'MsgBox "the number of rows should be " & tempvar
		verticalDistanceOffset = 10 'alter this integer to space the grades out further
		'TODO determine verticalDistanceOffset as a function of questions and maxHowFarOver
		verticalDistance = verticalDistance + verticalDistanceOffset
  
	Loop
End Sub

'Prints the top line of a grade sheet, executed once per student
Function printTopLine() 
    ActiveCell.Value = "Name"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = "Test name"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = "Score"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = "Questions Right"
    ActiveCell.Offset(-3, 0).Select
End Function
