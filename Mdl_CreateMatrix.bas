Attribute VB_Name = "Mdl_CreateMatrix"
'------------------------------------------------
'Name: Mdl_CreateMatrix
'Ver: 1.0.1
'Description�FCodes to create matrix from selected range
'Created By: @cnsl_oyu (from X)
'Date: 20240102
'Note: Please use this module on your own responsibitily.
'
'Copy Right 2024 @cnsl_oyu
'For terms and conditions, please read through the following:
'https://note.com/cnsl_oyu/n/nd5859008e017
'https://github.com/cnsl-oyu/CreateMatrix_ExcelVBA
'------------------------------------------------

Option Explicit

Private Const RangeToConvert = "RangeToConvert"
Private Const ppLayoutBlank = 12

Dim unit_height As Long
Dim unit_width As Long
Dim margin_v As Long
Dim margin_h As Long
Dim top_start As Long
Dim left_start As Long

Dim ErrMsg As String

Sub CreateMatrix()
    On Error GoTo ErrHandler
    ErrMsg = "F_CreateMatrix���J���܂���B�t�H�[�����������C���|�[�g����Ă��邩�m�F���Ă��������B"
    F_CreateMatrix.Show
    Exit Sub
    
ErrHandler:
    MsgBox ErrMsg, vbExclamation
End Sub

Sub convertCellsToMatrix(argstrApp As String) 'PowerPoint or Excel
    If argstrApp <> "PowerPoint" And argstrApp <> "Excel" Then
        ErrMsg = "convertCellsToMatrix���W���[���Ăяo�����������s���ł�(�����G���[)"
        Err.Raise 513 '0-512 is reserved
    End If
    
    Dim rngs() As Range
    Dim sld As Object 'Slide for PowerPoint
    Dim sp_height As Long, sp_width As Long
    Dim sp_top As Long, sp_left As Long
    Dim ptr_top As Long, ptr_left As Long
    Dim r As Variant
    
    On Error GoTo ErrHandler
    
    'Setting range to convert (creating a named range first and setting the address of the named range)
    ErrMsg = "�I��͈͂̎擾���ɃG���[���������܂����B�I���Z�������������m�F���Ă��������B"
    Selection.CurrentRegion.Name = RangeToConvert
    ActiveWorkbook.Names(RangeToConvert).RefersToLocal = "=" & Selection.Address
    
    'Load settings from the form
    LoadSettings
    
    'Get ranges considering merged areas
    rngs() = GetMergedRanges(RangeToConvert)
    
    'Initialize the variables
    ptr_top = 0
    ptr_left = 0
    
    'Create new slide if arg is PowerPoint
    If argstrApp = "PowerPoint" Then
        ErrMsg = "PowerPoint��V�K�쐬�ł��܂���ł����BPowerPoint�̏�Ԃ��m�F���Ă��������B"
        Set sld = CreateNewPPT
    End If
    
    'Depict cells in rngs
    For Each r In rngs
        'Calc cell size
        ErrMsg = "�}�`�T�C�Y�v�Z���ɃG���[���������܂����B�ݒ�l�ɕs���Ȓl�i�ɒ[�ɑ傫�Ȓl�Ȃǁj�����͂���Ă��Ȃ����m�F���Ă��������B"
        sp_height = unit_height * r.Rows.Count + margin_v * (r.Rows.Count - 1)
        sp_width = unit_width * r.Columns.Count + margin_h * (r.Columns.Count - 1)
        
        'Set position pointer
        ptr_top = r.Row - Range(RangeToConvert).Row
        ptr_left = r.Column - Range(RangeToConvert).Column

        'Depict shape
        If argstrApp = "PowerPoint" Then
            Call InsertShape(sld, sp_height, sp_width, ptr_top * (unit_height + margin_v) + top_start, ptr_left * (unit_width + margin_h) + left_start, r.Cells(1, 1))
        ElseIf argstrApp = "Excel" Then
            Call InsertShape(ActiveSheet, sp_height, sp_width, ptr_top * (unit_height + margin_v) + top_start, ptr_left * (unit_width + margin_h) + left_start, r.Cells(1, 1))
        End If
    Next
    
    'Release sld if arg is PowerPoint
    If argstrApp = "PowerPoint" Then Set sld = Nothing

    MsgBox "�o�͊������܂����B"
    
    Exit Sub
ErrHandler:
    MsgBox ErrMsg, vbExclamation
End Sub

'Load settings regarding the values of the user form
Sub LoadSettings()
    On Error GoTo ErrHandler
    ErrMsg = "�t�H�[����̐ݒ�l�̓ǂݍ��ݒ��ɃG���[���������܂����B�ݒ�l�ɕs���Ȓl�i�ɒ[�ɑ傫�Ȓl�Ȃǁj�����͂���Ă��Ȃ����m�F���Ă��������B"
    With F_CreateMatrix
        unit_height = .t_unit_height.Value
        unit_width = .t_unit_width.Value
        margin_v = .t_margin_v.Value
        margin_h = .t_margin_h.Value
        top_start = .t_top_start.Value
        top_start = .t_top_start.Value
        left_start = .t_left_start.Value
    End With
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number
End Sub

'Create new PowerPoint slide and returns the slide
Private Function CreateNewPPT() As Object 'slide
    On Error GoTo ErrHandler
    ErrMsg = "PowerPoint�̐V�K�쐬���ɃG���[���������܂����B"
    Dim ppApp As Object 'New PowerPoint Application
    Dim ppSlide As Object 'Slide
    
    Set ppApp = CreateObject("PowerPoint.Application")
    Set ppSlide = ppApp.Presentations.Add.Slides.Add(1, ppLayoutBlank)
    ppApp.Visible = True
    
    Set CreateNewPPT = ppSlide
    Set ppApp = Nothing
    Exit Function

ErrHandler:
    Set ppApp = Nothing
    Set ppSlide = Nothing
    Err.Raise Err.Number
End Function

Sub InsertShape(ByRef target As Object, h As Long, w As Long, t As Long, l As Long, ByRef cl As Range)
    On Error GoTo ErrHandler
    ErrMsg = "Shape�}�����ɃG���[���������܂����B"

    If TypeName(target) = "Worksheet" Then
        ' Excel�V�[�g�ւ̑}��
        With target.Shapes.AddShape(msoShapeRectangle, _
            Left:=l, Top:=t, Width:=w, Height:=h)
            .Fill.ForeColor.RGB = cl.Interior.Color
            .Line.ForeColor.RGB = cl.Borders.Color
            
            With .TextFrame.Characters.Font
                .Size = cl.Font.Size
                .Color = cl.Font.Color
            End With
            With .TextFrame2.TextRange
                .Text = cl.Text
                .Font.NameAscii = cl.Font.Name
                .Font.NameFarEast = cl.Font.Name
            End With
            .Select (False)
        End With

    ElseIf TypeName(target) = "Slide" Then
        ' PowerPoint�X���C�h�ւ̑}��
        With target.Shapes.AddShape(Type:=msoShapeRectangle, _
            Left:=l, Top:=t, Width:=w, Height:=h)
            .Fill.ForeColor.RGB = cl.Interior.Color
            .Line.ForeColor.RGB = cl.Borders.Color
            
            With .TextFrame.TextRange
                .Text = cl.Text
                With .Font
                    .Name = cl.Font.Name
                    .NameFarEast = cl.Font.Name
                    .Color.RGB = cl.Font.Color
                    .Size = cl.Font.Size
                End With
            End With
            .Select (False)
        End With
    Else
        ErrMsg = "PowerPoint�X���C�h�܂���Excel�V�[�g���w�肳��Ă��܂���i�����G���[�j"
        Err.Raise 513 '0-512 is reserved
    End If
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number
End Sub

'Function that returns an array of the input ranges considering merged cells
Private Function GetMergedRanges(ByVal strRngName) As Range()
    Dim r As Range
    Dim arr() As Range
    Dim flg As Boolean
    Dim i As Long, cnt As Long
    Dim s As Integer
    
    On Error GoTo ErrHandler
    ErrMsg = "�����͈̓��X�g�̎擾���ɃG���[���������܂����B"
    
    s = 0
    cnt = 1
    
    For Each r In Range(strRngName)
        flg = True
        If cnt > 1 Then
            For i = 1 To UBound(arr)
                If Not Intersect(arr(i), r) Is Nothing Then flg = False
            Next i
        End If
        
        If flg Then
            s = s + 1
            ReDim Preserve arr(1 To s)
            Set arr(s) = r.MergeArea
        End If
        cnt = cnt + 1
    Next r

    GetMergedRanges = arr
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number
End Function


