VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSlideSelector 
   Caption         =   "スライド選択"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6885
   OleObjectBlob   =   "UserFormSlideSelector.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserFormSlideSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'スライドジャンパーのメインモジュール
'スライド一覧をダイアログで表示し、選択したスライドへジャンプ
Private Sub UserForm_Initialize()
    lstSlides.Clear
    Dim i As Integer
    For i = 1 To ActivePresentation.slides.Count
        lstSlides.AddItem "スライド " & i
    Next i
    
    ' 最初のスライドを選択状態にする
    If lstSlides.ListCount > 0 Then
        lstSlides.ListIndex = 0
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "初期化中にエラーが発生しました。" & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, vbExclamation, "エラー"
    Me.Hide
    
End Sub


Private Sub cmdJump_Click()
    ModuleSlideJumper.JumpToSlide lstSlides.ListIndex + 1
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub lstSlides_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '
    ' ListBoxのダブルクリックイベント
    ' ダブルクリックでもジャンプできるように
    '
    Cancel = True
    cmdJump_Click
End Sub

' リストボックスでマウススクロールを可能にするための設定
Private Sub lstSlides_MouseMove( _
             ByVal Button As Integer, ByVal Shift As Integer, _
             ByVal x As Single, ByVal y As Single)
     HookListBoxScroll
End Sub
   
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     UnhookListBoxScroll
End Sub

