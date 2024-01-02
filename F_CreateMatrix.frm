VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_CreateMatrix 
   Caption         =   "OYU tool(マトリクス生成)"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5100
   OleObjectBlob   =   "F_CreateMatrix.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "F_CreateMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------
'Name: F_CreateMatrix
'Ver: 1.0.1
'Description：Codes of the form to create matrix from selected range
'Created By: @cnsl_oyu (from X)
'Date: 20231231
'Note: Please use these codes on your own responsibitily.
'
'Copy Right 2024 @cnsl_oyu
'For terms and conditions, please read through the following:
'https://note.com/cnsl_oyu/n/nd5859008e017
'https://github.com/cnsl-oyu/CreateMatrix_ExcelVBA
'------------------------------------------------

Option Explicit
'Const value for the position to depict sample shapes
Const init_top = 96
Const init_left = 24

'Init form
Private Sub UserForm_Initialize(): DepictSample: End Sub

'Btn action for excel
Private Sub cbExpToExcel_Click()
    Call convertCellsToMatrix("Excel")
    Unload F_CreateMatrix 'エクスポート後にフォームを閉じたくない場合はここをコメントアウト(1/2) ※Formをモードレスに設定することを推薦
End Sub

'Btn action for ppt
Private Sub cbExpToPpt_Click()
    Call convertCellsToMatrix("PowerPoint")
    Unload F_CreateMatrix 'エクスポート後にフォームを閉じたくない場合はここをコメントアウト(2/2) ※Formをモードレスに設定することを推薦
End Sub

'Depict sample according to the values set
Sub DepictSample()
    On Error Resume Next
    
    Me.lSpl_11.Width = t_unit_width.Value
    Me.lSpl_11.Height = t_unit_height.Value
    Me.lSpl_11.Left = init_left + t_left_start.Value
    Me.lSpl_11.Top = init_top + t_top_start.Value
    
    Me.lSpl_12.Width = t_unit_width.Value
    Me.lSpl_12.Height = t_unit_height.Value
    Me.lSpl_12.Left = Me.lSpl_11.Left + Me.lSpl_11.Width + t_margin_h.Value
    Me.lSpl_12.Top = Me.lSpl_11.Top
    
    Me.lSpl_21.Width = t_unit_width.Value
    Me.lSpl_21.Height = t_unit_height.Value
    Me.lSpl_21.Left = Me.lSpl_11.Left
    Me.lSpl_21.Top = Me.lSpl_11.Top + Me.lSpl_11.Height + t_margin_v.Value
    
    Me.lSpl_22.Width = t_unit_width.Value
    Me.lSpl_22.Height = t_unit_height.Value
    Me.lSpl_22.Left = Me.lSpl_11.Left + Me.lSpl_11.Width + t_margin_h.Value
    Me.lSpl_22.Top = Me.lSpl_11.Top + Me.lSpl_11.Height + t_margin_v.Value
End Sub

'Spin btn actions (ideally implement error handling if the value exceeds the limit of long type)
Private Sub SpinButton1_SpinUp(): Me.t_unit_height.Text = Me.t_unit_height.Text + 1: End Sub
Private Sub SpinButton1_SpinDown(): Me.t_unit_height.Text = Me.t_unit_height.Text - 1: End Sub
Private Sub SpinButton2_SpinUp(): Me.t_unit_width.Text = Me.t_unit_width.Text + 1: End Sub
Private Sub SpinButton2_SpinDown(): Me.t_unit_width.Text = Me.t_unit_width.Text - 1: End Sub
Private Sub SpinButton3_SpinUp(): Me.t_margin_v.Text = Me.t_margin_v.Text + 1: End Sub
Private Sub SpinButton3_SpinDown(): Me.t_margin_v.Text = Me.t_margin_v.Text - 1: End Sub
Private Sub SpinButton4_SpinUp(): Me.t_margin_h.Text = Me.t_margin_h.Text + 1: End Sub
Private Sub SpinButton4_SpinDown(): Me.t_margin_h.Text = Me.t_margin_h.Text - 1: End Sub
Private Sub SpinButton5_SpinUp(): Me.t_top_start.Text = Me.t_top_start.Text + 1: End Sub
Private Sub SpinButton5_SpinDown(): Me.t_top_start.Text = Me.t_top_start.Text - 1: End Sub
Private Sub SpinButton6_SpinUp(): Me.t_left_start.Text = Me.t_left_start.Text + 1: End Sub
Private Sub SpinButton6_SpinDown(): Me.t_left_start.Text = Me.t_left_start.Text - 1: End Sub

'Depict sample when values changed
Private Sub t_margin_h_Change(): DepictSample: End Sub
Private Sub t_margin_v_Change(): DepictSample: End Sub
Private Sub t_unit_height_Change(): DepictSample: End Sub
Private Sub t_unit_width_Change(): DepictSample: End Sub
Private Sub t_left_start_Change(): DepictSample: End Sub
Private Sub t_top_start_Change(): DepictSample: End Sub

