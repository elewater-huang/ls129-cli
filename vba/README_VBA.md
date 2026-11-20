# LS129Core.dll / VBA (V0.3)

這份 SDK 是給 **Windows 32-bit Excel / Access VBA** 使用的。特別適合 Access 2003、Excel 2003 的 `.mdb`／`.xls`，不用裝 Python、編譯器，也不用 `regsvr32` 註冊 DLL。

## 安裝
1. 從 GitHub Actions 下載 `LS129CLI-SDK-x86-V0.3` 並解壓縮。
2. 將 `LS129Core.dll`、`LS129.ini` 放在 Access MDB 或 Excel XLS 的**相同資料夾**。
3. 用 VBA 編輯器（Alt+F11）→ File → Import File 匯入 `M_LS129.bas`（建議在 Access/Excel 中只匯入一次）。
4. 用裝置管理員確認 USB-RS485 的 COM Port，修改 `LS129.ini` 裡的 `port=COM3`；關閉 AccessPort / CLI / Python monitor，以免佔用同一個 COM Port。

## Access 2003 VBA 範例
在現有標準模組中加入以下程序：
```vb
Public Sub Test_LS129_Access()
    Dim rc As Long
    Dim raw As Double, adjusted As Double
    Dim peak As Double, energy As Double

    rc = LS129_Begin(CurrentProject.Path)
    If rc <> 0 Then
        MsgBox "Open: " & LS129_ErrorText(rc)
        Exit Sub
    End If

    ' Start a NEW energy and peak measurement period explicitly.
    rc = LS129_NewPeriod()
    If rc = 0 Then
        rc = LS129_ReadStatistics(raw, adjusted, peak, energy)
        If rc = 0 Then
            MsgBox "Raw=" & raw & " Corrected=" & adjusted & _
                   " Peak=" & peak & " Energy=" & energy
        End If
    End If
    If rc <> 0 Then MsgBox "Read: " & LS129_ErrorText(rc)
    Call LS129_End
End Sub
```

## Excel VBA 範例
```vb
Public Sub Test_LS129_Excel()
    Dim rc As Long
    Dim raw As Double, adjusted As Double

    rc = LS129_Begin(ThisWorkbook.Path)
    If rc <> 0 Then
        MsgBox LS129_ErrorText(rc)
        Exit Sub
    End If
    rc = LS129_ReadCurrent(raw, adjusted)
    If rc = 0 Then
        ThisWorkbook.Worksheets(1).Range("A1").Value = Now
        ThisWorkbook.Worksheets(1).Range("B1").Value = raw
        ThisWorkbook.Worksheets(1).Range("C1").Value = adjusted
    Else
        MsgBox LS129_ErrorText(rc)
    End If
    Call LS129_End
End Sub
```

## 公開 VBA API

| 函式 | 用途 |
|---|---|
| `LS129_Begin(folder)` | 開啟指定目錄內的 DLL/INI、連線；傳回 0 表成功 |
| `LS129_ReadCurrent(raw, adjusted)` | 原始照度、PC 修正照度 |
| `LS129_ReadStatistics(raw, adjusted, peak, energy)` | 同時讀即時值、最大值、累積能量 |
| `LS129_NewPeriod()` | **真的向探頭下達 Recount**，清除累積能量與峰值 |
| `LS129_CurrentSmoothing(mode)` | 讀取平滑模式 |
| `LS129_Smoothing(mode)` | 寫入平滑模式 0／1／2，並讀回比對 |
| `LS129_End()` | 關閉通訊與釋放 COM Port |
| `LS129_Version()` | DLL API 版本 |

**請留意**：DLL 的 `Read`、`ReadStatistics` **不會自動清零**；應用程式開始一次新量測時先呼叫 `LS129_NewPeriod()`，與 CLI 的 `record` 相同原則。連續讀取時維持同一次連線，不要每一筆都重複 Begin/End。

### 設定說明
- `LS129.ini` 讀取連線與電腦端校正係數；不會將校正值寫回探頭。
- DLL 沒有自動計時器。讀取頻率由 Access／Excel 呼叫迴圈或外部流程決定；最高 32Hz 的 `poll_hz` 只是設定值，不代表探頭產生新的 32 筆/秒量測。
- DLL 僅允許**一個 Session**；一台 COM Port 不可同時給 CLI、Excel、Access 開啟。
- 本版不提供站號、波特率寫入（避免失聯）。之後可加入雙階段確認、重新連線及救援流程。
- DLL 為 x86；若是 **64-bit Office**，必須另編 x64 DLL 與對應 VBA 宣告。
- 本版本沒有對 Win7 實機／Access2003 實機完成測試；GitHub x86 編譯不等於實際環境驗證。
- 原廠 LS129 V3.2 協定：<https://images.linshangtech.com/user-manual/LS129_Communication_Protocol_V3.2.pdf>
