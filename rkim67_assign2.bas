Attribute VB_Name = "Module1"
'this function will declare a true or false value as to whether two tweets are similar enough to be duplicates, through an given threshold
Function isDup(tweet1 As String, tweet2 As String, threshold As Double) As Boolean

'declare arrays to split tweet1 and tweet2 into seperate words
Dim tw1() As String, tw2() As String

'declaring a percentage variable to compare to threshold, and a duplicate word counter, initally set to 0
Dim percentage As Double, dupCount As Double
dupCount = 0

'using split function to seperate words into array values
tw1 = Split(tweet1)
tw2 = Split(tweet2)

'using a for loop to go through values of the first tweet array
For i = LBound(tw1) To UBound(tw1)
    
    'nested for loop to go through the values of the second tweet array
    For n = LBound(tw2) To UBound(tw2)
        
        'if a match is found, our duplicate counter goes up by 1, using vbTextCompare to ignore capitalization
        If StrComp(tw1(i), tw2(n), vbTextCompare) = 0 Then
                dupCount = dupCount + 1
        End If
    Next n
Next i

'percentage of duplicates is equal to amount of duplicates / amount of words based on tweet1
percentage = dupCount / (UBound(tw1) + 1)

'if the threshold is passed, function returns true, otherwise it returns false
If percentage >= threshold Then
    isDup = True

Else
    isDup = False
    
End If

End Function
'this function will calculate the sentiment value of a tweet, accessing a seperate database with sentimeny keywords and assessing each tweet a score
Function sentimentCalc(tweet As String) As Integer

'declaring counters for positive and negative scores, default value 0
Dim posCount As Integer, negCount As Integer
posCount = 0
negCount = 0

'declaring positive and negative ranges to import from the keywords worksheet
Dim positive As Range, negative As Range

'set range equal to A2:A54, B2:B54 from the keywords worksheet
Set positive = Worksheets("keywords").Range("A2:A54")
Set negative = Worksheets("keywords").Range("B2:B54")

'declaring an array and using the split function to seperate by word
Dim tweetArr() As String
tweetArr = Split(tweet)

'for loop to remove punctuation from each word using replace
For i = LBound(tweetArr) To UBound(tweetArr)
    tweetArr(i) = Replace(tweetArr(i), "!", "")
    tweetArr(i) = Replace(tweetArr(i), ".", "")
    tweetArr(i) = Replace(tweetArr(i), ",", "")
    tweetArr(i) = Replace(tweetArr(i), "?", "")
    tweetArr(i) = Replace(tweetArr(i), ":", "")
    tweetArr(i) = Replace(tweetArr(i), ";", "")
    tweetArr(i) = Replace(tweetArr(i), ")", "")
    tweetArr(i) = Replace(tweetArr(i), "(", "")
Next i

'for loop to go through every value in the tweet array
For c = LBound(tweetArr) To UBound(tweetArr)
    
    'nested for loop to compare each value in the tweet to every postiive and negative keyword
    For n = 0 To positive.Count
    
        'if the word matches a positive keyword, +10 sentiment value
        If StrComp(tweetArr(c), positive(n), vbTextCompare) = 0 Then
            posCount = posCount + 10
        End If
        
        'if the word matches a negative keyword, -10 sentiment value
        If StrComp(tweetArr(c), negative(n), vbTextCompare) = 0 Then
            negCount = negCount + 10
        End If
        
    Next n
Next c

'calculate the total sentiment value of the entire tweet, and return the final value
sentimentCalc = posCount - negCount

End Function
'this function will categorize a tweet based on their sentiment value into positive, neutral, or negative
Function sentimentCategory(sentVal As Integer) As String

'if greater than 0, positive
If sentVal > 0 Then
    sentimentCategory = "Positive"
    
'if equal to 0, neutral
ElseIf sentVal = 0 Then
    sentimentCategory = "Neutral"
    
'otherwise, negative
Else
    sentimentCategory = "Negative"
    
End If

End Function


