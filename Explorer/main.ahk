#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode("Input")

; エクスプローラー上のみホットキー有効
#HotIf WinActive("ahk_class CabinetWClass") || WinActive("ahk_class ExploreWClass")

; Alt + Shift + D: 今日の日付 (yyyymmdd形式) を入力
!+d:: {
    today := FormatTime(, "yyyyMMdd")
    SendText(today)
}

#HotIf ; ホットキー制限解除（ここで必ず終了）
