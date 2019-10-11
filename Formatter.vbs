'Takes two parameters: size & colIndex
'size dictates the amount of rows and colIndex dictates the column we are pulling from
'This function populates an array from a specified column and continues to populate until the size parameter has been satisfied
'returns: a string array
Function populateCustomerNumberArray(size As Integer, colIndex As Integer) As Variant
    Dim customerNumbers() As String
    ReDim customerNumbers(1 To size)
    For i = 1 To size
        customerNumbers(i) = Cells(i + 1, colIndex).Value
    Next i
    populateCustomerNumberArray = customerNumbers
End Function

'Takes two parameters: customerArray & colIndex
'customerArray is the array of customerNumbers and colIndex dictates column we are manipulating
'This function updates changes made to the column specified by the user. In this case the changes made are the number of zeros that need to be added to the string in the cell
Function updateSpreadSheetDataWithZeros(customerArray As Variant, colIndex As Integer)
    Dim size As Integer
    Dim custNum As String
    Dim length As Integer
    Dim numberOfZeros As Integer
    size = UBound(customerArray) - LBound(customerArray) + 1
    For i = 1 To size
        length = Len(customerArray(i))
        numberOfZeros = 10 - length
        custNum = customerArray(i)
        customerArray(i) = addNumberOfZeros(numberOfZeros, custNum)
        Sheet1.Cells(i + 1, colIndex).Value = customerArray(i)
    Next i
End Function
'SUPPORT METHOD TO updateSpreadSheetData (Actually adds to zeros to the string)
'
'Takes two Parameters: zeros & customerNum
'zeros will be the number of zeros needed to make the customerNum(which is the number associated with a customer) a 10 digit number
'This function will add zeros to an empty string for however many times it needs to, and will append to the customer number
'returns: a customer number with added leading zeros
Function addNumberOfZeros(zeros As Integer, customerNum As String) As String
Dim newCustomerNum As String
Dim zeroString As String
zeroString = ""
For i = 1 To zeros
    zeroString = zeroString + "0"
Next i
newCustomerNum = zeroString + customerNum
addNumberOfZeros = newCustomerNum
End Function

Function getNumberOfActiveRows(colIndex As Integer, isHeaderPresent As Boolean) As Long
    Dim lastRow As Long
    If isHeadderPResent Then
        Cells(2, colIndex).Select
    Else
        Cells(1, colIndex).Select
    End If
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Count
End Function

'Will add leading zeros to customerNumbers in column A
'Author: Cody Vollrath
'Version: 09/25/2019
Sub main()
    Dim customersArray() As String
    Dim test As Integer
    customersArray = populateCustomerNumberArray(2733, 1)
    Call updateSpreadSheetDataWithZeros(customersArray, 1)
End Sub
