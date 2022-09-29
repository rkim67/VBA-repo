Attribute VB_Name = "Module1"
'Richard Kim
'CS2034
'Winter 2020
'Christopher Brogly

Option Explicit

'sub main to call all subs
Sub Main()

Call NegWord
Call PosWord
Call LongestWordLength
Call MeanWordLength
Call UCaseWordCount
Call EndingPunctuation
Call ConsonantsVowelsRatio
Call FoodMentions

End Sub

'negative sentiment words cited from: https://gist.github.com/mkulakowski2/4289441
'   Minqing Hu and Bing Liu. "Mining and Summarizing Customer Reviews."
'   Proceedings of the ACM SIGKDD International Conference on Knowledge
'   Discovery and Data Mining (KDD-2004), Aug 22-25, 2004, Seattle,
'   Washington, USA,

'check if certain negative words exist within a text
Sub NegWord()

    'importing reviews and sentiment rating words
    Dim text As Range
    Set text = Worksheets("YELP").Range("A2:A1001")

    Dim negative As Range
    Set negative = Worksheets("NegWord").Range("A1:A4783")

    'setting strings to split into text, and an array to store the values
    Dim strNegWord As Variant
    Dim strText As Variant
    Dim words() As String
    
    'setting up line to display results
    Dim line As Integer
    line = 2
    
    'first for loop to go through each review
    For Each strText In text
        'split each text
        words = Split(strText, " ")
        
        'set comparing words and a counter
        Dim word As Variant
        Dim NegWord As Variant
        
        Dim negWordTotal As Integer
        negWordTotal = 0
            
            'second for loop to go through each word, third nested loop to go through each sentiment value
            For Each word In words
                For Each NegWord In negative
                
                    'remove punctuation from word
                    word = Replace(word, ",", "")
                    word = Replace(word, "!", "")
                    word = Replace(word, ".", "")
                    word = Replace(word, "?", "")
                    word = Replace(word, ")", "")
                    word = Replace(word, "(", "")
                
                'if word in text matches a negative word, then increment the counter
                If StrComp(LCase(word), NegWord) = 0 Then
                    negWordTotal = negWordTotal + 1
                End If
                
                Next
            Next
    
    'display results, set new line for next text
    Range("D" & line).Value = negWordTotal
    line = line + 1

    Next

End Sub

'positive sentiment words cited from: https://gist.github.com/mkulakowski2/4289437
'   Minqing Hu and Bing Liu. "Mining and Summarizing Customer Reviews."
'   Proceedings of the ACM SIGKDD International Conference on Knowledge
'   Discovery and Data Mining (KDD-2004), Aug 22-25, 2004, Seattle,
'   Washington, USA,

Sub PosWord()
    
    'importing reviews and sentiment rating words
    Dim text As Range
    Set text = Worksheets("YELP").Range("A2:A1001")
    
    Dim positive As Range
    Set positive = Worksheets("PosWord").Range("A1:A2006")

    'setting strings to split into text, and an array to store the values
    Dim strPosWord As Variant
    Dim strText As Variant
    Dim words() As String
    
    'setting up line to display results
    Dim line As Integer
    line = 2
    
    'first for loop to go through each review
    For Each strText In text
        'split each text
        words = Split(strText, " ")
        
        'set comparing words and a counter
        Dim word As Variant
        Dim PosWord As Variant
        
        Dim posWordTotal As Integer
        posWordTotal = 0
        
        'second for loop to go through each word, third nested loop to go through each sentiment value
        For Each word In words
            For Each PosWord In positive
            
                    'remove punctuation from word
                    word = Replace(word, ",", "")
                    word = Replace(word, "!", "")
                    word = Replace(word, ".", "")
                    word = Replace(word, "?", "")
                    word = Replace(word, ")", "")
                    word = Replace(word, "(", "")
                    
                    'if word in text matches a positive word, then increment the counter
                     If StrComp(LCase(word), PosWord) = 0 Then
                     posWordTotal = posWordTotal + 1
                     End If
        Next
    Next
    
    'display results, set new line for next text
    Range("E" & line).Value = posWordTotal
    line = line + 1

    Next

End Sub

'this sub calculates the length of the longest word in the review.
Sub LongestWordLength()
    
    'importing review, preparing array
    Dim text As Range
    Set text = Range("A2:A1001")
    
    Dim words() As String
    
    'line number to set to print
    Dim line As Integer
    line = 2
    
    'preparing variable to split review text into
    Dim strText As Variant
    
    For Each strText In text
    
         'remove punctuation from word, replace with text to split incase of accidental tied words
         strText = Replace(strText, ",", " ")
         strText = Replace(strText, "!", " ")
         strText = Replace(strText, ".", " ")
         strText = Replace(strText, "?", " ")
         strText = Replace(strText, ")", " ")
         strText = Replace(strText, "(", " ")
        
        'split text into variable
        words = Split(strText, " ")
        
        'set word to measure length and a counter for the longest length calculated
        Dim word As Variant
        Dim wordLength As Integer
        wordLength = 0
        
        'for loop to check every word, and if word is longer then it replaces as the next highest
        For Each word In words
                    
                If Len(word) > wordLength Then
                    wordLength = Len(word)
                End If
        
        Next
           
    'display results, set new line for next text
    Range("F" & line).Value = wordLength
    line = line + 1
    
    Next

End Sub

'this sub will calculate how many of the words in the review contain one or more uppercase letters
Sub UCaseWordCount()

    'importing review, preparing array
    Dim text As Range
    Set text = Range("A2:A1001")

    Dim words() As String
    
    'line number to set to print
    Dim line As Integer
    line = 2

    'preparing variable to split review text into
    Dim strText As Variant
    
    For Each strText In text

        'split text into variable
        words = Split(strText, " ")
        Dim capitalCount As Integer
        capitalCount = 0
        
        Dim word As Variant
            
            For Each word In words
            
                If LCase(word) <> word Then
                    capitalCount = capitalCount + 1
                End If
                            
            Next

    'display results, set new line for next text
    Range("G" & line).Value = capitalCount
    line = line + 1

    Next
    
    
End Sub

'this sub will calculate the average word length of the review.
Sub MeanWordLength()

 'importing review, preparing array
    Dim text As Range
    Set text = Range("A2:A1001")
    
    Dim words() As String
    
    'line number to set to print
    Dim line As Integer
    line = 2
    
    'preparing variable to split review text into
    Dim strText As Variant
    
    For Each strText In text
    
         'remove punctuation from word
         strText = Replace(strText, ",", "")
         strText = Replace(strText, "!", "")
         strText = Replace(strText, ".", "")
         strText = Replace(strText, "?", "")
         strText = Replace(strText, ")", "")
         strText = Replace(strText, "(", "")
        
        'split text into variable
        words = Split(strText, " ")
        
        'set word to measure length and a counter for the longest length calculated
        Dim word As Variant
        Dim wordLengthTotal As Double
        Dim wordCount As Double
        wordLengthTotal = 0
        wordCount = 0
        
        'for loop to check every word, add to length total and count each word
        For Each word In words
            
            wordLengthTotal = wordLengthTotal + Len(word)
            wordCount = wordCount + 1
        
        Next
           
    'display calculated mean, set new line for next text
    Range("H" & line).Value = wordLengthTotal / wordCount
    line = line + 1
    
    Next

End Sub

'this sub will output a 0 if the review ends in no punctuation, 1 if it ends in a period, and 2 if it ends in
'an exclamation mark
Sub EndingPunctuation()

    'importing review, preparing array
    Dim text As Range
    Set text = Range("A2:A1001")
    Dim strText As Variant

    
    'line number to set to print
    Dim line As Integer
    line = 2

    'for each review text
    For Each strText In text
        
        'if text ends in "." , then set ending value to 1
        If strText.Value Like "*[\.]" Then
            Range("I" & line).Value = 1
            line = line + 1
        'if text ends in "!", then set ending value to 2
        ElseIf strText.Value Like "*!" Then
            Range("I" & line).Value = 2
            line = line + 1
        'if text ends in anything else, value is 0
        Else
            Range("I" & line).Value = 0
            line = line + 1
        End If
        
    Next

End Sub


'this sub will calculate the number of consonants in the text string divided by the amount of vowels
Sub ConsonantsVowelsRatio()

    'importing review, preparing array
    Dim text As Range
    Set text = Range("A2:A1001")
    
    'declaring value for loop and string for comparisons
    Dim i As Integer
    Dim chrText As String
    
    'line number to set to print
    Dim line As Integer
    line = 2
    
    'preparing variable to split review text into
    Dim strText As Variant
    
    'for loop to go through each review in the spreadsheet
    For Each strText In text
    
        'setting counters for vowels and consonants, set as double to divide them later
        Dim consCount As Double
        consCount = 0
        Dim vowelCount As Double
        vowelCount = 0
    
            'for loop to go through each letter in string
            For i = 1 To Len(strText)
                
                'setting text to uppercase so we don't have to repeat each character in lower and upper
                chrText = UCase(Mid(strText, i, 1))
                
                'if statement to differentiate between vowels and consonants, Y is used more often as vowel
                If chrText Like "[AEIOUY]" Then
                    vowelCount = vowelCount + 1
                ElseIf chrText Like "[BCDFGHJKLMNPQRSTVXZW]" Then
                    consCount = consCount + 1
                End If
            
            Next i
        
        'set value to consonant divided by vowel, and go to next line
        Range("J" & line).Value = consCount / vowelCount
        line = line + 1
        
    Next


End Sub

'this sub will count the number of food mentions within each review
'keywords for foods taken from: https://github.com/CurtisGrayeBabin/List-of-all-Foods/blob/master/FOOD.txt
Sub FoodMentions()

    'importing reviews and food words
    Dim text As Range
    Set text = Worksheets("YELP").Range("A2:A1001")

    Dim food As Range
    Set food = Worksheets("FoodWord").Range("A1:A21580")

    'setting strings to split into text, and an array to store the values
    Dim strFoodWord As Variant
    Dim strText As Variant
    Dim words() As String
    
    'setting up line to display results
    Dim line As Integer
    line = 2
    
    'first for loop to go through each review
    For Each strText In text
        'split each text
        words = Split(strText, " ")
        
        'set comparing words and a counter
        Dim word As Variant
        Dim foodWord As Variant
        
        Dim foodWordTotal As Integer
        foodWordTotal = 0
            
            'second for loop to go through each word, third nested loop to go through each sentiment value
            For Each word In words
                For Each foodWord In food
                
                    'remove punctuation from word
                    word = Replace(word, ",", "")
                    word = Replace(word, "!", "")
                    word = Replace(word, ".", "")
                    word = Replace(word, "?", "")
                    word = Replace(word, ")", "")
                    word = Replace(word, "(", "")
                
                'if word in text matches a negative word, then increment the counter
                If StrComp(LCase(word), foodWord) = 0 Then
                    foodWordTotal = foodWordTotal + 1
                End If
                
                Next
            Next
    
    'display results, set new line for next text
    Range("K" & line).Value = foodWordTotal
    line = line + 1

    Next

End Sub
