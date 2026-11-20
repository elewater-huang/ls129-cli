---
title: "LS129Core.dll 使用與開發手冊"
subtitle: "RS485 / Modbus RTU · Windows x86 · Access / Excel VBA"
author: "LS129 Communication Toolkit"
date: "2026-11-20"
lang: zh-TW
---

# LS129Core.dll 使用與開發手冊

**版本：v0.3｜文件日期：2026-11-20｜平台：Windows x86（32 位元）**

本文件介紹如何使用 `LS129Core.dll` 控制林上 LS129 紫外線照度探頭，並透過 VBA 整合至 Microsoft Access、Excel 或其他 Windows 程式。內容依照本專案 GitHub 原始碼及已上傳的實測 DLL 核對；不代表儀器原廠認證。

> **先看這裡**：這是一般 Win32 DLL，不是 ActiveX/COM 元件，**不必執行 `regsvr32`**；使用電腦**不必安裝 Python 或 C 編譯器**。但需要合適的 USB–RS485 驅動及 32 位元呼叫程式。

## 1. 三分鐘快速開始

| 步驟 | 內容 |
|---|---|
| 1 | 安裝 USB–RS485 轉換器驅動，接妥 LS129，於 Windows 裝置管理員確認 `COMx`。 |
| 2 | 將 `LS129Core.dll`、`LS129.ini` 放在 Access `.mdb` 或 Excel 活頁簿**同一資料夾**。 |
| 3 | 以 `Alt+F11` 進入 VBA，匯入 `M_LS129.bas` 標準模組；已成功匯入的檔案不需重複匯入。 |
| 4 | 修改 `LS129.ini` 的 `port`、`baud_rate`、`station_id`，確認與探頭設定一致。 |
| 5 | 執行 `LS129_Begin(...)` → `LS129_ReadCurrent(...)` → `LS129_End()`；每個函式回傳 `0` 表示成功。 |

**Access 最小範例**（直接放入另一個標準模組）：

```vb
Public Sub LS129_QuickTest_Access()
    Dim rc As Long
    Dim rawValue As Double, correctedValue As Double

    rc = LS129_Begin(CurrentProject.Path)
    If rc <> 0 Then
        MsgBox "連線失敗：" & LS129_ErrorText(rc)
        Exit Sub
    End If

    rc = LS129_ReadCurrent(rawValue, correctedValue)
    If rc = 0 Then
        MsgBox "原始值 = " & rawValue & vbCrLf & _
               "軟體修正值 = " & correctedValue
    Else
        MsgBox "讀取失敗：" & LS129_ErrorText(rc)
    End If

    Call LS129_End
End Sub
```

Excel 使用方式完全相同，唯獨開啟時改用 `LS129_Begin(ThisWorkbook.Path)`。**單次讀值不會清零**。

## 2. DLL 的角色與系統架構

```text
LS129 UV 探頭
   │  RS485 / Modbus RTU（8N1）
   ▼
USB–RS485 轉換器 → Windows COM Port
   │
   ▼
LS129Core.dll
   ├─ 串列埠管理、Modbus CRC 與回應檢查
   ├─ 照度／最大值／累積能量讀取
   ├─ 重新統計（Recount）與平滑模式設定
   └─ PC 端照度修正係數
   │
   ├─ Access VBA（MDB）
   ├─ Excel VBA（XLS / XLSM，依 Office 版本）
   └─ 其他 32-bit Win32 呼叫程式
```

`LS129CLI.exe` 也使用同一份 `device.c` / `protocol.c` 通訊核心；CLI 是獨立程序，**不是**呼叫 DLL 的另一種介面。請不要讓 CLI、Access、Excel 同時占用同一個 COM Port。

### 2.1 相容性矩陣

| 環境 | 狀態／說明 |
|---|---|
| Windows 32-bit + 32-bit Office | 本版目標環境；DLL 為 PE32 x86。 |
| Access 2003（MDB） | 已由使用者回報實機 TestBench 成功操作。 |
| Excel（含巨集活頁簿） | 已由使用者回報實機 TestBench 成功操作；上傳成品是 `.xlsm`。 |
| Excel 2003 `.xls` | 可透過舊版 VBA 匯入方式建立；**不能直接開啟 `.xlsm`**。 |
| 64-bit Office | **不能直接載入 x86 DLL**；需另編 x64 DLL 與改寫 VBA `PtrSafe` 宣告。 |
| Windows 7 | 專案以 Win7 API 為目標，仍應在實際機台檢驗驅動及穩定性。 |

這個 DLL 不會自行建立 Access 資料表、Excel 圖表或計時器；它只提供儀器通訊 API，資料庫與 UI 的行為由上層 VBA 範本負責。

## 3. 部署與設定檔

### 3.1 檔案配置

```text
D:\LS129_Test\
  db1.mdb                   ← 或 LS129_TestBench.xlsm
  LS129Core.dll             ← 必要
  LS129.ini                 ← 必要
  M_LS129.bas               ← 用來匯入 VBA；匯入後非執行必要
  LS129CLI.exe              ← 選配，供命令列測試
```

DLL 是非註冊式函式庫，`M_LS129.bas` 中的 `LS129_Begin(folder)` 會先以完整路徑載入 DLL，再傳入同資料夾的 `LS129.ini`。請勿把 DLL 放在不明來源的系統目錄。含巨集檔案只應來自可信來源。

### 3.2 `LS129.ini` 範例

```ini
[Communication]
port=COM3
baud_rate=9600
station_id=1
timeout_ms=1000

[Acquisition]
poll_hz=5

[Calibration]
enabled=0
factor=1.0000

[System]
language=zh-TW
```

| 鍵值 | 可用值／作用 |
|---|---|
| `port` | `COM1`～`COM256`；須與裝置管理員一致。 |
| `baud_rate` | `4800`、`9600`、`19200`、`38400`；只設定 PC 通訊，不主動改寫探頭。 |
| `station_id` | `1`～`247`；要與探頭站號一致。 |
| `timeout_ms` | `100`～`60000`；通訊回應逾時控制。 |
| `poll_hz` | 大於 0、最大 `32`，可輸入 `0.2` 等小數；**DLL 僅驗證及讀取此設定，不會自行依頻率輪詢**。 |
| `enabled` | `0`：不修正；`1`：使用 PC 端係數。 |
| `factor` | 大於 0、最多 `1000000`；`corrected = raw × factor`（`enabled=1` 才生效）。 |
| `language` | 目前為保留設定，不影響 DLL 通訊或回傳值。 |

> **設定注意**：`port=COM10` 是某台實測電腦的連接埠，**不應**當作公開範例的固定預設；每台電腦都應自行修改。INI 中的 `poll_hz` 並非 LS129 內部取樣頻率，設定為 32 不代表取得 32 筆不同的感測器新資料。

## 4. VBA 開發者介面

使用 `M_LS129.bas` 時，優先呼叫下列封裝函式，不需自己宣告 `kernel32.dll` 或處理 Modbus 封包。

| VBA 函式 | 參數／輸出 | 行為 |
|---|---|---|
| `LS129_Begin(folder) As Long` | 傳入專案所在資料夾 | 載入 DLL/INI、開啟 COM，成功傳回 0。 |
| `LS129_ReadCurrent(raw, adjusted) As Long` | 兩個 `Double`，ByRef | 讀取原始照度與 PC 修正照度。 |
| `LS129_ReadStatistics(raw, adjusted, peak, energy) As Long` | 四個 `Double`，ByRef | 一次取得原始值、修正值、最大值、累積能量。 |
| `LS129_NewPeriod() As Long` | 無 | **向探頭寫入 Recount**，重新計算最大值與累積能量。 |
| `LS129_CurrentSmoothing(mode) As Long` | `Long`，ByRef | 讀回目前平滑模式。 |
| `LS129_Smoothing(mode) As Long` | `0/1/2` | 向探頭寫入平滑模式，並讀回確認。 |
| `LS129_End() As Long` | 無 | 關閉 COM 連線、釋放 DLL Session。 |
| `LS129_Version() As Long` | 無 | 已載入 DLL 時回傳 `3`，否則為 `0`。 |
| `LS129_ErrorText(code) As String` | 傳入回傳碼 | 提供 VBA 層的錯誤文字。 |

### 4.1 單次查詢與統計讀值

```vb
Public Sub LS129_Read_All_Access()
    Dim rc As Long
    Dim raw As Double, corrected As Double
    Dim peak As Double, energy As Double

    rc = LS129_Begin(CurrentProject.Path)
    If rc <> 0 Then
        MsgBox LS129_ErrorText(rc)
        Exit Sub
    End If

    rc = LS129_ReadStatistics(raw, corrected, peak, energy)
    If rc = 0 Then
        Debug.Print "照度原始值="; raw
        Debug.Print "修正照度="; corrected
        Debug.Print "最大值欄位="; peak
        Debug.Print "累積能量="; energy
    Else
        MsgBox LS129_ErrorText(rc)
    End If
    Call LS129_End
End Sub
```

### 4.2 新量測週期必須先 Recount

```vb
rc = LS129_Begin(CurrentProject.Path)
If rc <> 0 Then Exit Sub

rc = LS129_NewPeriod()     ' 新 Run 開始時只執行一次
If rc = 0 Then
    ' 在使用者的計時／記錄迴圈中重複呼叫 ReadStatistics
    rc = LS129_ReadStatistics(raw, corrected, peak, energy)
End If

Call LS129_End
```

`LS129_NewPeriod()` 會真正重置探頭的統計週期；**不要每讀一筆就呼叫**。若紫外光源持續點亮，重置完成後的累積能量會立即再增加，所以第一筆不一定等於 0。

對照 CLI，`LS129CLI.exe record --count 100 --out data.csv` 會自動在開始前重設一次；DLL 的 `LS129_Read()`／`LS129_ReadAll()` **不會自動重設**，由呼叫者管理量測週期。

### 4.3 讀取及設定 50／60 Hz 平滑模式

```vb
Dim rc As Long, mode As Long
rc = LS129_CurrentSmoothing(mode)
If rc = 0 Then Debug.Print "目前平滑模式="; mode

' 0=無平滑；1=50 Hz；2=60 Hz
rc = LS129_Smoothing(2)
If rc <> 0 Then MsgBox LS129_ErrorText(rc)
```

平滑模式會改寫**探頭內部**暫存器，不只是 UI 的顯示值。DLL 先等待 Modbus 寫入回應，再讀回設定比對；修改前應確認適合目前 UV 光源與量測條件。

## 5. 原生 DLL API（C／其他語言）

DLL 匯出名稱採 **Win32 x86 `__stdcall`**，使用一般 C 介面與 `Long`/`Double` 對應類型。函式名稱已透過 `.def` 匯出為不含修飾的形式；與 COM 元件不同，無 CLSID / ProgID。

```c
/* x86 / Win32，供已載入 LS129Core.dll 的呼叫端使用 */
LONG __stdcall LS129_Open(const char *ini_path, LONG *session);
LONG __stdcall LS129_Read(LONG session, double *raw, double *adjusted);
LONG __stdcall LS129_ReadAll(LONG session, double *raw,
                             double *adjusted, double *maximum,
                             double *energy);
LONG __stdcall LS129_Reset(LONG session);
LONG __stdcall LS129_GetSmoothing(LONG session, LONG *mode);
LONG __stdcall LS129_SetSmoothing(LONG session, LONG mode);
LONG __stdcall LS129_Close(LONG session);
LONG __stdcall LS129_APIVersion(void);
```

函式宣告展示**參數型別及呼叫慣例**；原生 C/C++ 專案仍需自行提供匯入宣告或使用 `GetProcAddress`，並在適用時連結提供的 `.lib`。

| 函式 | 成功時結果 | 補充說明 |
|---|---|---|
| `LS129_Open` | `0`，並寫入有效 `session` | `ini_path` 應為有效路徑；同一 DLL 程序同時只接受一個 Session。 |
| `LS129_Read` | `0`，寫入兩個 `double` | 原始值及 PC 校正值。 |
| `LS129_ReadAll` | `0`，寫入四個 `double` | `maximum`/`energy` 是整數暫存器解碼結果。 |
| `LS129_Reset` | `0` | 確認 Modbus 06H 回應；不保證下一個時刻讀值正好為零。 |
| `LS129_GetSmoothing` | `0`，寫入 `mode` | 可讀 `0`、`1`、`2`。 |
| `LS129_SetSmoothing` | `0` | 設定後讀回一致才成功；其他值回傳參數錯誤。 |
| `LS129_Close` | `0` | 關閉串列埠，先前的 Session 隨即失效。 |
| `LS129_APIVersion` | `3` | 表示 API 版本 0.3；不是儀器韌體版本。 |

**執行緒與程序**：此版本使用 DLL 內的單一全域 COM Handle / Session，**未提供多執行緒同步保護**。不可由多執行緒交錯呼叫，也不可讓兩個 Windows 程序同時開啟同一個 COM Port。已開啟的 Session 結束前，另一個 `Open` 會回傳 `5`。

## 6. 量測資料、單位及原始暫存器

### 6.1 本版本實作的暫存器

| Modbus 操作 | 位址 | DLL 解讀 | 資料型態 |
|---|---|---|---|
| FC03 讀取 | `201` | 即時照度／功率原始值 | `uint16` |
| FC03 讀取 | `202` | 最大值欄位（見以下注意） | `uint16` |
| FC03 讀取 | `203–204` | 累積能量原始值 | `uint32`，整數組合 |
| FC06 寫入 | `50 = 1` | 重新統計最大值與累積能量 | 需確認回應 |
| FC03／FC06 | `320` | 平滑：`0` 無、`1` 50Hz、`2` 60Hz | `uint16` |

**協定文件疑義**：原廠《LS129 Digital Probe Communication Protocol V3.2》的整數表格把暫存器 `201` 與 `202` 都標成「Real-time value」，但相同文件的浮點／長整數資料格式將第二組數值定義為最大值；目前 DLL 將 `202` 解讀為 `maximum`。因此，在對外報告之前，請以 LS129 實際清零／峰值測試確認 `202` 的語意。這是文件標示與實作之間的**待確認差異**，不是已認證的量測定義。

### 6.2 精度與校正

- 本版本採用**整數暫存器**，沒有啟用 Float 資料模式；不能憑 VBA 的 `Double` 宣告取得超過儀器整數資料的新精度。
- `raw` 是從 LS129 讀到的值；`adjusted` 只在 PC 端依 `enabled` 與 `factor` 計算。
- `peak` 與 `energy` 目前直接回傳原始整數，不套用照度修正係數。
- `energy` 在 C 中是無號 32 位整數，在 DLL/VBA 介面轉成 `Double`，避免使用有號 VBA `Long` 造成溢位。
- **不要直接假設輸出數字的單位一定是 mW/cm² 或 mJ/cm²**。應依探頭型號、原廠刻度及現場對照量測確認。

## 7. 連續量測與資料擷取原則

**DLL 不負責排程，也不自行儲存 CSV。** 你的 Access／Excel 應用程式控制何時呼叫 `ReadStatistics`、如何計時、如何處理資料與錯誤。

建議完整 Run 流程：

1. 開啟連線，驗證函式回傳碼。
2. 建立 Run ID、記錄設定與開始時間。
3. 新 Run 開始時呼叫一次 `LS129_NewPeriod()`；失敗就停止，不宣稱本次能量從零開始。
4. 依程式設定的輪詢節奏呼叫 `LS129_ReadStatistics()`，附加本機時間戳及實際經過時間。
5. 結束後關閉 Session，儲存完整資料及失敗筆數；必要時匯出 CSV。

`poll_hz` 最高允許 `32`，但這只是上位機欲發送 Modbus 指令的頻率；實際回應時間、Windows/VBA 計時機制以及探頭更新週期都可能使資料重複或低於設定頻率。**不得把 32Hz 輪詢直接宣稱為 32Hz 感測器有效新數據。**

## 8. 回傳碼與故障排除

所有正常 API（不含 `APIVersion`）都以 **`0 = 成功`** 作為統一判斷依據。

### 8.1 DLL 正值回傳碼

| 回傳碼 | 含義 | 建議處理 |
|---|---|---|
| `0` | 成功 | 繼續下一步。 |
| `1` | 參數錯誤 | 檢查指標、`mode` 範圍。 |
| `2` | INI 檔案不存在／非一般檔案 | 核對傳入路徑。 |
| `3` | INI 格式或設定不合法 | 檢查 COM、鮑率、站號、逾時、係數。 |
| `4` | 開啟／設定 COM Port 失敗 | 檢查驅動、端口號、埠占用、權限。 |
| `5` | DLL 已開啟一個 Session | 先 `LS129_Close`。 |
| `6` | 無效 Session | 確保先 `Open`，並傳入正確 Session。 |

### 8.2 VBA 封裝額外回傳碼

| 回傳碼 | 意義 |
|---|---|
| `101` | 沒有傳入工作簿／資料庫資料夾。 |
| `102` | 指定資料夾找不到 `LS129Core.dll`。 |
| `103` | 無法載入 DLL（含位元數不匹配等原因）。 |
| `104` | 指定資料夾找不到 `LS129.ini`。 |
| `105` | VBA 已經有進行中的 Session。 |

### 8.3 Modbus／串列埠負值回傳碼

本版底層的負值依正在執行的讀取／寫入函式而異，不宜把它當成完全固定的一對一錯誤字典。

| 範圍／範例 | 通常代表 | 建議處理 |
|---|---|---|
| `-10`～`-13` | 建立命令、清除緩衝區、寫入或送出失敗 | 檢查 COM 開啟狀態、驅動和轉換器。 |
| `-20`～`-25` | 逾時、站號不匹配、回應格式或長度異常 | 檢查站號、鮑率、供電、A/B 接線。 |
| `-3`（資料解碼時） | Modbus CRC 不符 | 檢查線材、干擾、接地、通訊品質。 |
| `-26`～`-28`（部分設定操作） | CRC／回應比對失敗，或寫入後讀回不一致 | 暫停寫入並核對儀器設定。 |
| `-100 - N` | 儀器回傳 Modbus Exception Code `N`（例 `-102` = 代碼 `2`） | 核對原廠例外碼及暫存器權限。 |

### 8.4 排查順序

1. 先關閉 AccessPort、CLI、其他 Excel／Access 專案，避免占用同一個 COM Port。
2. 在裝置管理員檢查 USB–RS485 裝置、COM 號及驅動。
3. 比對 `LS129.ini` 中的站號、鮑率及現場 RS485 A/B 接線。
4. 用 CLI `LS129CLI.exe selftest` 確認本機協定測試；**這不是實機連線成功的證據**。
5. 用 `LS129CLI.exe read` 進行真正單次讀取，再測試 Excel／Access。
6. 若 DLL 可連線但不能讀取，記錄原始回傳碼、使用程式、Windows／Office 位元數及操作順序。

## 9. 設計邊界、發布與安全建議

- DLL **不提供**修改 LS129 的站號、探頭鮑率、內部校正係數及其他不在公開 API 的功能。
- `LS129_SetSmoothing()` 和 `LS129_Reset()` 是實際儀器寫入操作，請勿在無提示的自動開啟流程中呼叫。
- 這份原生 DLL 沒有數位簽章認證；巨集安全設定應遵循使用單位規範。
- 目前已上傳的 Access `.mdb` 含實測紀錄、Excel `.xlsm` 含巨集及文件屬性；若公開 GitHub，應先清理個人資料與內部量測資料，保留未修改的私有實測備份。
- 上傳的 Excel 實測成品格式是 `.xlsm`（Excel 2007+）；**請勿把它標成 Excel 2003 `.xls` 下載品**。

## 10. 版本、驗證清單與參考資料

**版本識別**

| 項目 | 內容 |
|---|---|
| DLL 檔案 | `LS129Core.dll`（Win32 PE32 / x86） |
| API 版本 | `LS129_APIVersion()` → `3`，代表 SDK v0.3 |
| 通訊核心 | `src/device.c` + `src/protocol.c` |
| VBA 模組 | `vba/M_LS129.bas` |
| GitHub 原始碼 | https://github.com/elewater-huang/ls129-cli |
| 原廠協定 | LS129 Digital Probe Communication Protocol (Customer), V3.2, 2020-07-23 |

**建議的發行前驗收：**

- [ ] Windows 與 Office 位元數已確認，能載入 DLL。
- [ ] COM Port、鮑率、站號與探頭一致。
- [ ] Access／Excel 可以單次讀值、正常關閉串列埠。
- [ ] `NewPeriod` 後累積能量從新的統計週期開始；並確認暫存器 `202` 的實際意義。
- [ ] 平滑模式 `0/1/2` 可寫入與讀回。
- [ ] 連續記錄含真實時間戳，沒有把重複讀值誤認為新資料。
- [ ] CSV、MDB／XLSM 發布檔已清除敏感資料與不必要的文件屬性。

**原廠協定文件：** https://images.linshangtech.com/user-manual/LS129_Communication_Protocol_V3.2.pdf  
**專案程式碼：** https://github.com/elewater-huang/ls129-cli  
**本手冊定位：** 社群製作的非官方 DLL 操作與開發說明，不取代儀器原廠規格、校正證明或安全作業文件。