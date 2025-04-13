#Requires AutoHotkey v2.0

; CapsLockを無効化
CapsLock::return

; f1キー単押しを無効化
f1::return

moveAmount := 10
dragging := false

; CapsLock + 矢印キーでマウス移動
CapsLock & Up::MouseMove(0, -moveAmount, 0, "R")
CapsLock & Down::MouseMove(0, moveAmount, 0, "R")
CapsLock & Left::MouseMove(-moveAmount, 0, 0, "R")
CapsLock & Right::MouseMove(moveAmount, 0, 0, "R")

; CapsLock + D でマウス左ボタンを押す（ドラッグ開始）
CapsLock & d::
{
    global dragging
    if !dragging {
        dragging := true
        MouseClick "left", , , 1, 0, "D"  ; ← 押す
    }
    return
}

; CapsLock + D を離すとマウス左ボタンを離す（ドラッグ終了）
#HotIf GetKeyState("CapsLock", "P")
d up::
{
    global dragging
    if dragging {
        dragging := false
        MouseClick "left", , , 1, 0, "U"  ; ← 離す
    }
    return
}
#HotIf


; CapsLock + A → 左クリック
CapsLock & a::MouseClick "left"
