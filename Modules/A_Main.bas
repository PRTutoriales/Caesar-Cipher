Attribute VB_Name = "A_Main"
'*********************************************
'Tutorial: Caesar cipher
'Code: VBA
'Platform: Microsoft Excel/PowerPoint
'Only for educational propouses
'https://www.youtube.com/channel/UCwJ8qS-Jr8h-BaCfIrpRzEQ/featured
'*********************************************
Option Explicit
'*********************************************
'This simple code get all consecutives ASCII values from a custom range.
'If you need a customized Range, you must to create it by a Collection or Array and implement to this code
'0[48] to z[122] / A[65] to Z[90] ASCII CODE
'*********************************************
Public Const st_ASCII As Integer = 65      '48
Public Const fn_ASCII As Integer = 90      '122
Public A_FIELD As Range                     'PowerPoint >> Public A_FIELD As String
Public B_FIELD As Range                     'PowerPoint >> Public B_FIELD As String
Public K_FIELD As Range                     'PowerPoint >> Public K_FIELD As String
'*********************************************
'MAIN SUBS
'*********************************************
Public Sub EncodeText()
    'PowerPoint >> Slide1.Encoded_txt.Value = Cypher(False)
    B_FIELD.Value = Cypher(False)
End Sub
Public Sub DecodeText()
    'PowerPoint >> Slide1.Message_txt.Value = Cypher(True)
    A_FIELD.Value = Cypher(True)
End Sub
'*********************************************
'DEFINE RANGES/OBJECTS [EXCEL or POWERPOINT]
'*********************************************
Public Sub CustomRange()
                                         'PowerPoint >> You must create three Textboxes and assign a name, assuming they are on Slide 1
    Set A_FIELD = Range("CELL_MESSAGE")  'PowerPoint >> A_FIELD = Slide1.Message_txt.Value
    Set B_FIELD = Range("CELL_ENCODED")  'PowerPoint >> B_FIELD = Slide1.Encoded_txt.Value
    Set K_FIELD = Range("CELL_KEY")      'PowerPoint >> K_FIELD = Slide1.Key_txt.Value
End Sub
'*********************************************
'GET DEFINED RANGES/OBJECTS
'*********************************************
Public Function GetPlainText() As String
    GetPlainText = A_FIELD.Value             'PowerPoint >> GetPlainText = A_FIELD
End Function
Public Function GetCyphedText() As String
    GetCyphedText = B_FIELD.Value            'PowerPoint >> GetCyphedText = B_FIELD
End Function
Public Function GetKey() As String
    GetKey = K_FIELD.Value                   'PowerPoint >> GetKey = K_FIELD
End Function

