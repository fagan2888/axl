Option Private Module
Option Explicit

#If VBA7 Then
    #If Win64 Then
        Const XLPyDLLName As String = "xlpython64-2.0.9.dll"
        Declare PtrSafe Function XLPyDLLActivate Lib "xlpython64-2.0.9.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
    #Else
        Private Const XLPyDLLName As String = "xlpython32-2.0.9.dll"
        Private Declare PtrSafe Function XLPyDLLActivate Lib "xlpython32-2.0.9.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
    #End If
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#Else
    #If Win64 Then
        Const XLPyDLLName As String = "xlpython64-2.0.9.dll"
        Declare Function XLPyDLLActivate Lib "xlpython64-2.0.9.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
    #Else
        Private Const XLPyDLLName As String = "xlpython32-2.0.9.dll"
        Private Declare Function XLPyDLLActivate Lib "xlpython32-2.0.9.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
    #End If
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#End If

Dim AbortAll As Boolean
Dim StackNdx As Integer
Dim StackCell As String
Dim Stack(0 To 31) As Variant
Dim PyLoaded As Boolean
Dim PP As Variant

Private Function Push(FName As String, ParamArray EArgs() As Variant)
    Dim CC As String
    CC = Application.Caller.Address(External:=True)
    If CC <> StackCell Then
        StackNdx = 0
    ElseIf StackNdx > UBound(Stack, 1) Then
        MsgBox "Maximum Python function nesting limit exceeded"
        Push = CVErr(xlErrValue)
        Exit Function
    End If
    Dim i As Integer
    Dim Args As Variant
    Dim Entry() As Variant
    Dim UB As Integer
    Args = EArgs(0)
    UB = UBound(Args)
    ReDim Entry(0 To UB + 1)
    Entry(0) = FName
    For i = 0 To UB
        Select Case TypeName(Args(i))
        Case "Error"
            Push = Args(i)
            Exit Function
        Case "Range"
            Entry(i + 1) = Args(i).Value
        Case Else
            Entry(i + 1) = Args(i)
        End Select
    Next i
    Stack(StackNdx) = Entry
    Push = "!$" + Str(StackNdx)
    StackNdx = StackNdx + 1
    StackCell = CC
End Function

Private Sub ToExcel(Headers As Boolean)
    If Stack(StackNdx - 1)(0) = "ToExcel" Then Exit Sub
    Dim R As Integer
    Dim C As Integer
    If IsObject(Application.Caller) Then
        R = Application.Caller.Rows.Count
        C = Application.Caller.Columns.Count
    Else
        R = 1
        C = 1
    End If
    X "@ToExcel", "!$" + Str(StackNdx - 1), R, C, Headers
End Sub

Private Function Exec()
    Dim Valid As Boolean
    Dim Args() As Variant
    Dim i As Integer
    Valid = StackNdx > 0
    If Valid Then
        ReDim Args(0 To StackNdx - 1)
        Valid = True
        For i = 0 To StackNdx - 1
            If TypeName(Stack(i)) = "Error" Then Valid = False
            Args(i) = Stack(i)
            Stack(i) = 0
        Next i
        StackNdx = 0
    End If
    Valid = Valid And Not AbortAll
    If Valid And Not PyLoaded Then
        Dim LongResult As Long
        LongResult = LoadLibrary(ThisWorkbook.Path + "\" + XLPyDLLName)
        If LongResult = 0 Then
            MsgBox "Error" + Str(Err.LastDllError) + " loading " + XLPyDLLName
            Valid = False
        Else
            Dim CfgFile As String
            CfgFile = ThisWorkbook.Path + "\axl.cfg"
            LongResult = XLPyDLLActivate(PP, CfgFile)
            PyLoaded = (LongResult = 0)
            If Not PyLoaded Then
                MsgBox Err.Description
                Valid = False
            End If
        End If
    End If
    If Valid Then
        Dim TName As String
        On Error Resume Next
        Exec = PP.Call(Args)
        TName = TypeName(Exec)
        If Err <> 0 Then
            MsgBox Err.Description
            Valid = False
        ElseIf TName = "Null" Then
            Exec = CVErr(xlErrNA)
        ElseIf TName = "String" Then
           ' If Left(Exec, 8) = "#PYTHON?" Then
           '     MsgBox Mid(Exec, 10)
           '     Valid = False
           ' ElseIf Len(Exec) > 32767 Then
           '
           ' End If
           Exec = Left(Exec, 32767)
        End If
    End If
    If Not Valid Then
        Exec = CVErr(xlErrValue)
    End If
End Function

Function X(FName As String, ParamArray Args() As Variant)
    X = Push(FName, Args)
End Function

Function P(FName As String, ParamArray Args() As Variant)
    If Left(FName, 2) = "!$" And UBound(Args, 1) < LBound(Args, 1) Then
        X "@Grab", FName
    Else
        Push FName, Args
    End If
    ToExcel True
    P = Exec()
End Function

Function PyLog(ParamArray Args() As Variant)
    Push "%Log", Args
    PyLog = Exec()
End Function

Function List(ParamArray Args() As Variant)
    List = Push("@List", Args)
End Function

Function Tuple(ParamArray Args() As Variant)
    Tuple = Push("@Tuple", Args)
End Function

Function Slice(ParamArray Args() As Variant)
    Slice = Push("@Slice", Args)
End Function

Function Dict(ParamArray Args() As Variant)
    Dict = Push("@Dict", Args)
End Function

Function Matrix(ParamArray Args() As Variant)
    Matrix = Push("@Matrix", Args)
End Function

Function Vector(ParamArray Args() As Variant)
    Vector = Push("@Vector", Args)
End Function

Function RowV(ParamArray Args() As Variant)
    RowV = Push("@Row", Args)
End Function

Function ColV(ParamArray Args() As Variant)
    ColV = Push("@Column", Args)
End Function

Function RowDF(ParamArray Args() As Variant)
    RowDF = Push("@RowDF", Args)
End Function

Function ColDF(ParamArray Args() As Variant)
    ColDF = Push("@ColDF", Args)
End Function

Function VecDF(ParamArray Args() As Variant)
    VecDF = Push("@VecDF", Args)
End Function

Function MatDF(ParamArray Args() As Variant)
    MatDF = Push("@MatDF", Args)
End Function

Function Repr(Arg As String)
    Repr = P("@Repr", Arg)
End Function

Function S(FName As String, Func As String, ParamArray Args() As Variant)
    S = P("%Save", FName, Push(Func, Args))
End Function

Function PySave(FName As String, Arg As String)
    PySave = P("%Save", FName, Arg)
End Function

Function L(FName As Variant)
    If TypeName(FName) = "Range" Then FName = FName.Value2
    L = X("%Load", FName)
End Function

Function PyLoad(FName As String)
    PyLoad = X("%Load", FName)
End Function

Function DFCols(FName As Variant, Columns As Variant, Optional SortBy As Variant, Optional Ascending As Variant)
    FName = L(FName)
    If IsMissing(SortBy) Then
        X "@DFCols", FName, "columns=", Columns
    ElseIf IsMissing(Ascending) Then
        X "@DFCols", FName, "columns=", Columns, "sortby=", SortBy
    Else
        X "@DFCols", FName, "columns=", Columns, "sortby=", SortBy, "ascending=", Ascending
    End If
    ToExcel False
    DFCols = Exec()
End Function

Function DFColsExclude(FName As Variant, Exclude As Variant, Optional SortBy As Variant, Optional Ascending As Variant)
    FName = L(FName)
    If IsMissing(SortBy) Then
        X "@DFCols", FName, "exclude=", Exclude
    ElseIf IsMissing(Ascending) Then
        X "@DFCols", FName, "exclude=", Exclude, "sortby=", SortBy
    Else
        X "@DFCols", FName, "exclude=", Exclude, "sortby=", SortBy, "ascending=", Ascending
    End If
    ToExcel True
    DFColsExclude = Exec()
End Function

Function DFExtract(FName As Variant, ParamArray Args() As Variant)
    Dim i As Integer
    Dim N As Integer
    Dim Args2() As Variant
    N = UBound(Args, 1)
    ReDim Args2(0 To N + 1)
    Args2(0) = L(FName)
    For i = 0 To N
        Args2(i + 1) = Args(i)
    Next i
    Push "@Extract", Args2
    ToExcel False
    DFExtract = Exec()
End Function
