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
Dim StackValid As Boolean
Dim Stack() As Variant
Dim PyLoaded As Boolean
Dim PP As Variant

Private Function Push(FName As String, ParamArray EArgs() As Variant)
    Dim Ndx As Integer
    Dim Args As Variant
    Args = EArgs(0)
    For Ndx = 0 To UBound(Args)
        Select Case TypeName(Args(Ndx))
        Case "Error"
            StackValid = False
        Case "Range"
            Args(Ndx) = Args(Ndx).Value
        End Select
    Next Ndx
    If Not StackValid Then
        ReDim Preserve Stack(0 To 1, 0 To 0)
        StackValid = True
        Stack(0, 0) = "@"
        Stack(1, 0) = 0
    End If
    Dim CC As String
    CC = Application.Caller.Address(External:=True)
    If CC <> Stack(0, 0) Then
        Stack(0, 0) = CC
        Ndx = 1
    Else
        Ndx = UBound(Stack, 2) + 1
    End If
    ReDim Preserve Stack(0 To 1, 0 To Ndx)
    Stack(0, Ndx) = FName
    Stack(1, Ndx) = Args
    Push = "!$" + Str(Ndx - 1)
End Function

Private Sub ToExcel(Headers As Boolean)
    Dim R As Integer
    Dim C As Integer
    If IsObject(Application.Caller) Then
        R = Application.Caller.Rows.Count
        C = Application.Caller.Columns.Count
    Else
        R = 1
        C = 1
    End If
    Dim Ndx As Integer
    Ndx = UBound(Stack, 2)
    If Headers Then
        X "@ToExcel", "!$" + Str(Ndx - 1), R, C
    Else
        X "@ToExcel", "!$" + Str(Ndx - 1), R, C, False
    End If
End Sub

Private Function Exec()
    Dim Valid As Boolean
    Dim i As Integer
    Valid = StackValid And Not AbortAll
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
        Exec = PP.Call(Stack)
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
    StackValid = False
    If Not Valid Then
        Exec = CVErr(xlErrValue)
    End If
End Function

Function X(FName As String, ParamArray Args() As Variant)
    X = Push(FName, Args)
End Function

Function P(FName As String, ParamArray Args() As Variant)
    If Left(FName, 2) <> "!$" Then
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

Function Repr(Arg As Variant)
    X "@Repr", Arg
    Repr = Exec()
End Function

Function PySave(FName As String, Arg As Variant)
    If Not IsObject(Application.Caller) Then
        MsgBox "PySave must be called within a cell"
        PySave = CVErr(xlErrValue)
    ElseIf Application.Caller.Count <> 1 Then
        MsgBox "PySave cannot be used in an array function"
        PySave = CVErr(xlErrValue)
    Else
        X "%Save", Application.Caller.Address(External:=True), Arg
        Exec
        PySave = FName
    End If
End Function

Function S(FName As String, Func As String, ParamArray Args() As Variant)
    S = PySave(FName, Push(Func, Args))
End Function

Function PyLoad(FName As Variant)
    Dim FType As String
    If TypeName(FName) <> "Range" Then
        MsgBox "Argument to Load() must be a cell reference, not a " & TypeName(FName)
        PyLoad = CVErr(xlErrValue)
    ElseIf FName.Count <> 1 Then
        MsgBox "Argument to Load() must be a single cell"
        PyLoad = CVErr(xlErrValue)
    Else
        PyLoad = X("%Load", FName.Address(External:=True))
    End If
End Function

Function L(FName As Variant)
    L = PyLoad(FName)
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
