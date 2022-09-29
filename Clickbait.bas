Attribute VB_Name = "Clickbait"
'This file is an example of Feature Engineering.
'All of the Subs listed below in Sub Main produce values that an ML classifier can process and used to classify distinct types of text.
'We will look at this process for A3 and in the ML lecture.

'I wrote most of this code during the lecture but I left out the last two subs with Nested Loops - I will go over these.

Option Explicit

'a sub Main can be used to run all of your subs, it is a nice way to organize things
Sub Main()
    Call StartsWithNumber
    Call CalcCBCharLength
    Call CalcCBWordLength
    Call HighlightLikelyCB
    Call CheckForStopWords
    Call CheckForPronounWords
End Sub

Sub StartsWithNumber()
    'define the range to access clickbait text
    Dim clickbaits As Range
    Set clickbaits = Range("A2:A49")
    Dim strClickbait As Variant
    
    'starting at line 2 since we have a heading on the column
    Dim line As Integer
    line = 2
    
    'NOTE: Normal For loops with LBound(...) and UBound(...) calls could have been used here.
    'loop over each line of clickbait
    For Each strClickbait In clickbaits
        'split each line into individual words
        Dim words() As String
        words = Split(strClickbait, " ")
        
        'check the first word in the line of text to see if it numeric
        If IsNumeric(words(0)) Then
            Range("F" & line).Value = 1
        Else
            Range("F" & line).Value = 0
        End If
        
        'go to the next line
        line = line + 1
    Next
    
End Sub

Sub CalcCBCharLength()
    
    'get the range of clickbait texts
    Dim clickbaits As Range
    Set clickbaits = Range("A2:A49")
    
    Dim i As Integer
    Dim strClickbait As Variant
    
    i = 2
    For Each strClickbait In clickbaits
        'get the character length of each clickbait text
        Range("C" & i).Value = Len(strClickbait)
        i = i + 1
    Next
    
End Sub

'pretty much same as above, but count the number of words instead of characters
Sub CalcCBWordLength()
    
    Dim clickbaits As Range
    Set clickbaits = Range("A2:A49")
    
    Dim words() As String
    Dim i As Integer
    Dim strClickbait As Variant
    
    i = 2
    For Each strClickbait In clickbaits
        words = Split(strClickbait, " ")
        Range("B" & i).Value = UBound(words)
        i = i + 1
    Next
    
End Sub

'go through each line of clickbait, if they start with "Here's" then highlight red
'some info on the color indices https://docs.microsoft.com/en-us/office/vba/api/excel.colorindex
Sub HighlightLikelyCB()
    Dim clickbaits As Range
    Set clickbaits = Range("A2:A49")
    
    Dim words() As String
    Dim i As Integer
    Dim line As Integer
    Dim strClickbait As Variant
    
    line = 2
    
    For Each strClickbait In clickbaits
        words = Split(strClickbait, " ")
        
        For i = LBound(words) To UBound(words) Step 1
            If StrComp(words(i), "Here's", vbTextCompare) = 0 Then
                Range("A" & line).Interior.ColorIndex = 3
            End If
        Next i
        line = line + 1
    Next
End Sub

'check if certain words exist within a clickbait text
Sub CheckForStopWords()
    Dim clickbaits As Range
    Set clickbaits = Range("A2:A49")
    
    Dim stopwords As Range
    Set stopwords = Range("J31:J33")
    
    'with Option Explicit on you must declare your index variable As Variant
    'because coming from a Range you dont know what type it is
    Dim strClickbaitLine As Variant
    Dim strStopwordLine As Variant
    
    Dim words() As String
    
    Dim line As Integer
    line = 2
    For Each strClickbaitLine In clickbaits
        'split each clickbait text
        words = Split(strClickbaitLine, " ")
        
        Dim word As Variant
        
        'use stopword for the nested For Each loop
        Dim stopword As Variant
        
        'count the number of stopwords you find in this clickbait text
        Dim stopwordTotal As Integer
        stopwordTotal = 0
        
        For Each word In words
            For Each stopword In stopwords
                'if the word in this clickbait text matches a stop word, then increment the count
                If StrComp(LCase(word), stopword) = 0 Then
                    stopwordTotal = stopwordTotal + 1
                End If
            Next
        Next
        
        Range("D" & line).Value = stopwordTotal
        line = line + 1
    Next
    
End Sub

'This is pretty much a copy + paste of the above.
'This is bad practice in Software Engineering and it usually means code quality is low
'but we can get away with it in short programs like this.
'The only difference is that CheckForPronounWords uses a different range.
Sub CheckForPronounWords()
    Dim clickbaits As Range
    Set clickbaits = Range("A2:A49")
    
    Dim stopwords As Range
    Set stopwords = Range("I32:I108")
    
    Dim strClickbaitLine As Variant
    Dim strStopwordLine As Variant
    
    Dim words() As String
    
    Dim line As Integer
    line = 2
    For Each strClickbaitLine In clickbaits
        words = Split(strClickbaitLine, " ")
        
        Dim word As Variant
        Dim stopword As Variant
        
        Dim stopwordTotal As Integer
        stopwordTotal = 0
        For Each word In words
            For Each stopword In stopwords
                If StrComp(LCase(word), stopword) = 0 Then
                    stopwordTotal = stopwordTotal + 1
                End If
            Next
        Next
        
        Range("E" & line).Value = stopwordTotal
        line = line + 1
    Next
    
End Sub
