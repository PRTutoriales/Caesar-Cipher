Attribute VB_Name = "B_Caesar"
'*********************************************
'Tutorial: Caesar cipher
'Code: VBA
'Platform: Microsoft Excel/PowerPoint
'Only for educational propouses
'https://www.youtube.com/channel/UCwJ8qS-Jr8h-BaCfIrpRzEQ/featured
'*********************************************
Option Explicit
Private PLAIN_TEXT As String
Private CYPHED_TEXT As String
Private KEY As Long
'*********************************************
'PUBLIC SUBS AND FUNCTIONS
'*********************************************
'---------------------------------------------
'Cypher function as string
'---------------------------------------------
Public Function Cypher(ByVal oEnode As Boolean) As String
    
    If GetValuesFromCells(oEnode) Then                                                  'Get values of Cells/Objects
        Dim chars As New Collection
        Set chars = SplitWord()                                                         'Split Message in a New Collection
        
        If chars.Count > 0 Then                                                         'Check if Collection has elements
            Dim i As Integer
            Dim cyph As String
            For i = 1 To chars.Count                                                    'Go over Collection char by char
                Dim n As Long                                                           'chr function must need a Long value
                n = ModArithmetic(chars.Item(i))                                        'Cypher each char with a Modular Arithmetic function
                cyph = cyph & chr(n)                                                    'Concatenate the new cyphed string
            Next i
            Debug.Print ">>Message processed successfully."                             'Print messages on Immediate window: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/use-the-immediate-window
            Cypher = cyph                                                               'Returns the cyphed string
        Else
            Cypher = ""
        End If
    Else
        Debug.Print ">>Message not processed."                                          'Print messages on Immediate window
        Cypher = ""
    End If
End Function
'*********************************************
'PRIVATE SUBS AND FUNCTIONS
'*********************************************
'---------------------------------------------
'Modular Arithmetic maths / Space between words
'---------------------------------------------
Private Function ModArithmetic(ByVal oAsc As Long) As Long
    If oAsc <> 32 Then                                                                  'Keep spaces between words [Alt + 32].
        Dim result As Long
        result = oAsc + KEY
        If result > fn_ASCII Then result = (result - fn_ASCII) + (st_ASCII - 1)         'When addition is positive and out of range begins again from [A]
        If result < st_ASCII Then result = Abs((st_ASCII - result) - (fn_ASCII + 1))    'When addition is negative and out of range begins from the last char [Z]
        Debug.Print " ." & chr(oAsc) & " | " & chr(result)                              'Print referenced chars by selected key
        ModArithmetic = result                                                          'Return right char by previous additions/conditionals
    Else
        ModArithmetic = 32                                                              'Return a space if ASCII char is 32
    End If
End Function
'---------------------------------------------
'Split any phrase inside a Cell/Object
'---------------------------------------------
Private Function SplitWord() As Collection
    Dim str As String                                                                   'Use a temp var for main text
    str = PLAIN_TEXT
    
    If Len(str) <= 0 Then                                                               'Check if the cell has a value
        str = "Empty Text to Encode/Decode."                                            'Init str and Key if they have not any value
        KEY = 0
        
        MsgBox (str)
        Debug.Print ">>" & str
        
        Set SplitWord = Nothing
    Else
        Dim t, i As Integer                                                             'Get total of characters
        t = Len(str)
        
        Dim w As New Collection                                                         'Create a collection an fill it with all chars in ASCII code
        For i = 1 To t
            w.Add (Asc(Mid(str, i, 1)))                                                 'Convert selected char to ASCII code by Mid [substring] function
        Next i                                                                          'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mid-function
    
        str = "Message succesfully splitted."
        MsgBox (str)
        Debug.Print ">>" & str
        
        Set SplitWord = w                                                               'Returns the Collection for Cypher(oEnode) Function
    End If
End Function
'---------------------------------------------
'Get values from Cells/Objects
'---------------------------------------------
Private Function GetValuesFromCells(ByVal oEncode As Boolean) As Boolean
    Call CustomRange                                                                    'Get our custom Range of Cells/Objects
    KEY = GetKey                                                                        'Get Key value from our custom range
    Debug.Print ">>Getting value from Private Key: " & KEY
    
    If CheckRangeASCII() Then                                                           'Check range of Key, en/decode and continue the procedure
        Call SwitchTexts(oEncode)                                                       'Only get data if key is in range
        GetValuesFromCells = True
    Else
        GetValuesFromCells = False
    End If
End Function
'---------------------------------------------
'Check range ASCII values
'---------------------------------------------
Private Function CheckRangeASCII() As Boolean
    Dim RngASCII As Integer                                                                         'Substract MAX to MIN from ASCII range
    RngASCII = fn_ASCII - st_ASCII + 1                                                              'Conditional to get correct range of ASCII code by our Private Key
    If KEY < -RngASCII Or KEY > RngASCII Then
        KEY = 0
        Debug.Print ">>Error: Key is out of range. Try between " & -RngASCII & " and " & RngASCII
        CheckRangeASCII = False
    Else
        Debug.Print ">>Key is in range."
        CheckRangeASCII = True
    End If
End Function
'---------------------------------------------
'Assign values from our custom Cells/Objects to private vars
'---------------------------------------------
Private Sub SwitchTexts(ByVal oEncode As Boolean)
    'En/Decode
    If Not oEncode Then
        PLAIN_TEXT = GetPlainText
        CYPHED_TEXT = GetCyphedText
        Debug.Print ">>Encoding Message..."
    Else
        PLAIN_TEXT = GetCyphedText
        CYPHED_TEXT = GetPlainText
        Debug.Print ">>Decoding Message..."
    End If
End Sub
