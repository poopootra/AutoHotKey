#Requires AutoHotkey v2.0
#SingleInstance Force

; グローバル変数
global objectKeys := Map()
global pptApp := 0
global activeShapes := 0

#HotIf WinActive("ahk_class PPTFrameClass")
; 修飾キーなしの場合のみFキーをカスタム処理
CapsLock & f:: {    
    ; オブジェクト選択モード開始
    StartPPTObjectSelection()
}

; 各キーのホットキー設定（選択モード中のみ有効）
#HotIf WinActive("ahk_class PPTFrameClass") && objectSelectionActive
a::SelectPPTObject("a")
b::SelectPPTObject("b")
c::SelectPPTObject("c")
d::SelectPPTObject("d")
e::SelectPPTObject("e")
f::SelectPPTObject("f")
g::SelectPPTObject("g")
h::SelectPPTObject("h")
i::SelectPPTObject("i")
j::SelectPPTObject("j")
k::SelectPPTObject("k")
l::SelectPPTObject("l")
m::SelectPPTObject("m")
n::SelectPPTObject("n")
o::SelectPPTObject("o")
p::SelectPPTObject("p")
q::SelectPPTObject("q")
r::SelectPPTObject("r")
s::SelectPPTObject("s")
t::SelectPPTObject("t")
u::SelectPPTObject("u")
v::SelectPPTObject("v")
w::SelectPPTObject("w")
x::SelectPPTObject("x")
y::SelectPPTObject("y")
z::SelectPPTObject("z")
1::SelectPPTObject("1")
2::SelectPPTObject("2")
3::SelectPPTObject("3")
4::SelectPPTObject("4")
5::SelectPPTObject("5")
6::SelectPPTObject("6")
7::SelectPPTObject("7")
8::SelectPPTObject("8")
9::SelectPPTObject("9")
0::SelectPPTObject("0")
Escape::CancelObjectSelection()
#HotIf

; 選択モードフラグ
objectSelectionActive := false

; 表示GUIの配列
keyLabels := []
statusGui := 0

; PowerPointオブジェクト選択モードを開始
StartPPTObjectSelection() {
    try {
        ; PowerPointの入力モード（テキスト編集中）でないことを確認
        ppt := ComObject("PowerPoint.Application")
        selection := ppt.ActiveWindow.Selection
        
        ; グローバル変数に保存
        global pptApp := ppt
        global activeShapes := ppt.ActiveWindow.View.Slide.Shapes
        global objectSelectionActive := true
        
        ; スライド上のオブジェクトを取得し、キーを割り当てる
        DisplayObjectLabels()
    } catch as err {
        MsgBox("起動エラー: " . err.Message)
    }
}

; オブジェクトラベルを表示
DisplayObjectLabels() {
    try {
        global keyLabels := []
        global objectKeys := Map()
        
        win := pptApp.ActiveWindow
        shapes := activeShapes
        
        if shapes.Count = 0 {
            MsgBox("スライド上にオブジェクトがありません。")
            return
        }
        
        ; 選択可能なキー（アルファベット + 数字）
        selectionKeys := "abcdefghijklmnopqrstuvwxyz1234567890"
        
        ; 各オブジェクトに対して選択キーを割り当て
        Loop Min(shapes.Count, StrLen(selectionKeys)) {
            i := A_Index
            shape := shapes.Item(i)
            key := SubStr(selectionKeys, i, 1)
            
            ; 中心座標（スクリーン基準）
            absX := win.PointsToScreenPixelsX(shape.Left)
            absY := win.PointsToScreenPixelsY(shape.Top)
            
            ; キー表示用のGUIを作成
            labelGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")
            labelGui.BackColor := "FDE975"  ; 黄色の背景に変更
            labelGui.SetFont("s10 bold", "Arial")  ; フォントサイズを小さく
            textCtrl := labelGui.Add("Text", "Center BackgroundFDE975", StrUpper(key))  ; 大文字に変換
            
            ; テキストコントロールのサイズを取得
            textCtrl.GetPos(,, &width, &height)
            
            ; ラベルを中央に配置
            labelX := absX - (width / 2)
            labelY := absY - (height / 2)
            
            ; GUIを表示
            labelGui.Show("x" . labelX . " y" . labelY . " NoActivate")
            
            ; キーとオブジェクトの対応を保存
            objectKeys[key] := i
            
            ; ラベルGUIを保存
            keyLabels.Push(labelGui)
        }
        
        ; ステータスGUIを表示
        global statusGui := Gui("+AlwaysOnTop +ToolWindow")
        statusGui.Add("Text",, "キーを押してオブジェクトを選択 (ESC=キャンセル)")
        statusGui.Show("y" . (A_ScreenHeight - 100) . " NoActivate")
        
    } catch as err {
        MsgBox("ラベル表示エラー: " . err.Message)
        CancelObjectSelection()
    }
}

; 特定のキーでオブジェクトを選択
SelectPPTObject(key) {
    try {
        ; 選択モードを終了
        global objectSelectionActive := false
        
        ; オブジェクトインデックスを取得
        objIndex := objectKeys[key]
        
        ; GUIをクリーンアップ
        CleanupGuis()
        
        ; PowerPointウィンドウをアクティブに
        WinActivate("ahk_class PPTFrameClass")
        
        ; オブジェクトを選択
        try {
            ; シェイプ選択を試みる
            activeShapes.Item(objIndex).Select()
            
            ; 選択確認
            selection := pptApp.ActiveWindow.Selection
            if selection.Type != 2 {  ; 2 = ppSelectionShapes
                ; もう一度試行
                Sleep(100)  ; 少し待機
                activeShapes.Item(objIndex).Select()
            }
            
            ; 成功メッセージ（デバッグ用。必要に応じて削除可）
            ToolTip("選択成功: オブジェクト " . objIndex)
            SetTimer () => ToolTip(), -1000  ; 1秒後に消去
        } catch as err {
            MsgBox("選択エラー: " . err.Message)
        }
        
    } catch {
        CleanupGuis()
    }
}

; 選択モードをキャンセル
CancelObjectSelection() {
    global objectSelectionActive := false
    CleanupGuis()
}

; GUIリソースをクリーンアップする関数
CleanupGuis() {
    global keyLabels, statusGui
    
    ; すべてのキーラベルGUIを閉じる
    for labelGui in keyLabels
        labelGui.Destroy()
    keyLabels := []
    
    ; ステータスGUIを閉じる
    if statusGui {
        statusGui.Destroy()
        statusGui := 0
    }
}

; Shape.Typeの値から形状タイプ名を取得する関数
GetShapeTypeName(typeNum) {
    static typeNames := Map(
        1, "自動図形",
        2, "フリーフォーム",
        3, "グループ",
        4, "画像",
        5, "OLEオブジェクト",
        6, "グラフ",
        7, "表",
        8, "線",
        9, "埋め込みOLE",
        10, "リンクOLE",
        11, "メディア",
        12, "テキストボックス",
        13, "スクリプトアンカー",
        14, "Webビデオ",
        17, "SmartArt"
    )
    
    return typeNames.Has(typeNum) ? typeNames[typeNum] : "不明なタイプ(" . typeNum . ")"
}
