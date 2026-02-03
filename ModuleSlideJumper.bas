Attribute VB_Name = "ModuleSlideJumper"
'スライド選択のメインモジュール
'スライド一覧をダイアログで表示し、選択したスライドへジャンプ

Public Sub ShowSlideDialog()
    '
    ' ShowSlideDialog - スライド選択ダイアログを表示するメインサブ
    '
    Dim frmSlides As UserFormSlideSelector
    
    ' ユーザーフォームのインスタンスを作成
    Set frmSlides = New UserFormSlideSelector
    
    ' ダイアログを表示（モーダル表示）
    frmSlides.Show vbModal
    
    ' フォームのメモリ解放
    Unload frmSlides
    Set frmSlides = Nothing
    
End Sub

Public Sub JumpToSlide(slideNumber As Integer)
    '
    ' JumpToSlide - 指定されたスライド番号にジャンプ
    ' 引数: slideNumber - 対象スライドの番号（1から始まる）
    ' スライドショー中でも通常編集画面でも実行可能
    '
    On Error GoTo ErrorHandler
    
    Dim objPresentation As Object
    Dim max_id As Long
    Dim validSlideNumber As Long
    
    ' アクティブなプレゼンテーションを取得
    Set objPresentation = ActivePresentation
    
    ' 最大スライド数を取得
    max_id = objPresentation.slides.Count
    
    ' スライド番号のバリデーション
    ' 数値チェック - 無効な値の場合は最大スライド数に設定
    If Not IsNumeric(CStr(slideNumber)) Then
        validSlideNumber = max_id
    Else
        validSlideNumber = slideNumber
    End If
    
    ' 範囲チェック - 超過時は最大値に、未満時は1に丸める
    If validSlideNumber > max_id Then
        validSlideNumber = max_id
    End If
    If validSlideNumber < 1 Then
        validSlideNumber = 1
    End If
    
    ' スライドショー実行中かどうかを判定
    If SlideShowWindows.Count > 0 Then
        ' スライドショー中の場合：GotoSlideで遷移
        ActiveWindow.View.GotoSlide Index:=validSlideNumber
    Else
        ' 通常編集画面の場合：選択スライドを指定して、その位置に移動
        ActiveWindow.View.GotoSlide Index:=validSlideNumber
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "スライド遷移中にエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "エラー"
End Sub


Public Sub StartSlideShowFromSlide(slideNumber As Integer)
    On Error GoTo ErrorHandler
    
    Dim objPresentation As Object
    Dim pptApp As Object
    
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set objPresentation = pptApp.ActivePresentation
    
    If slideNumber < 1 Or slideNumber > objPresentation.slides.Count Then
        MsgBox "スライド番号が無効です。", vbExclamation
        Exit Sub
    End If
    
    Dim settings As Object
    Set settings = objPresentation.SlideShowSettings
    settings.ShowType = 1
    settings.AdvanceMode = 1
    settings.StartingSlide = slideNumber
    settings.EndingSlide = objPresentation.slides.Count
    settings.Run
    
    ' スライドショーが終了するまで待つ
    Do While pptApp.SlideShowWindows.Count > 0
        DoEvents
    Loop
    
    ' スライドショー終了後にメッセージを表示
    MsgBox "スライドショーが終了しました。", vbInformation, "完了"
    
    Exit Sub
ErrorHandler:
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub


