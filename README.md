# LS129CLI / VBA SDK V0.3 — LS129 RS485 / Modbus RTU 命令列讀取工具

**DLL 完整操作與開發說明：** [LS129Core.dll 使用與開發手冊（繁體中文）](docs/LS129Core_DLL_Manual_zh-TW.md)。


**目的：讓沒有安裝 Python、沒有 C/C++ 編譯器的 Windows 電腦，透過 COM Port 讀取林上 LS129。**

使用電腦只需要 USB-RS485 驅動程式和這個 EXE、INI，不需要編譯器。

下載 `LS129CLI-SDK-x86-V0.3`，解壓縮取得 `LS129CLI.exe`、`LS129Core.dll`、`M_LS129.bas`、`LS129.ini`、`README_VBA.md` 和 `RUN_FIRST.bat`。
編輯 INI 中 `port=COM3`，改為裝置管理員實際的 COM Port，接上 LS129；在 CMD 中執行 `LS129CLI.exe selftest`，再執行 `LS129CLI.exe read`。你也可以直接執行 `RUN_FIRST.bat`，依序進行自我測試和實機讀取。

## 指令範例

```bat
LS129CLI.exe selftest
LS129CLI.exe read
LS129CLI.exe status
LS129CLI.exe reset
LS129CLI.exe reset
LS129CLI.exe record --count 100 --out data.csv
LS129CLI.exe read --port COM4 --station 1 --baud 9600
LS129CLI.exe record --count 40 --hz 0.2 --out test.csv
```

`read`：只輸出一個無單位的原始整數，適合讓其他 BAT 或 VBA 擷取。`status`：顯示原始即時值、最大值、能量與 PC 端修正值。`record`：每次開始記錄前，會先用 Modbus 06H 向選定站號的暫存器 **50 寫入 1**，要求探頭重新統計累積能量及最大值。**必須收到正確回應才開始寫入 CSV**；失敗就終止，不寫入量測資料列。這是真正重設 LS129 內部統計，不是從 CSV 數值扣除舊值。`reset` 可獨立執行重新統計；`read`、`status` 不會清除統計。若 UV 燈持續照射，重設後能量立刻再次累積，因此第一筆不一定剛好為 0。記錄依照 PC 輪詢頻率輸出 CSV；**不要把重複讀值視為儀器內部產生的新量測**。`selftest`：只測試 CRC 與解碼，不需要探頭，不能當作實機通訊成功的證據。

### 設定檔

`LS129.ini` 放在 EXE **同一目錄**。`poll_hz` 可使用 0.2、1、5、32 等正數，最高 32；此參數只控制 PC 發送讀取命令的頻率，不改寫 LS129 的實際取樣頻率。PC 端校正係數由 `calibration.enabled` 和 `calibration.factor` 管理，**絕不寫入探頭**。

讀值支援暫存器 **201–204**（03H）；`record`/`reset` 支援以 **06H 寫入暫存器 50 = 1**。不會修改通訊站號、波特率、平滑設定或校正係數。讀取指令在站號 1 時為 `01 03 00 C9 00 04 94 37`。這是**整數格式**；單位必須依探頭型號與原始資料確認，勿直接假設所有型號都採相同物理單位。如果想保留小數，可以在下一版使用相應的 Float 暫存器。

## 常見問題

- `Cannot open/configure COMx`：檢查 USB-RS485 驅動、COM Port、是否被 AccessPort 或 Python 軟體佔用。不可同時讓兩個程式操作同一個 COM Port。
- `read failed (-20)`：沒有收到足夠的回應；檢查站號、波特率、供電、A/B 線與 USB-RS485 轉換器的收發方向控制。
- `read failed (-3)`：回應的 CRC 不符合，應檢查雜訊、接線與串列埠設定。
- `Record ABORTED: recount failed`：LS129 沒有確認重新統計，記錄被取消，CSV 不會產生任何資料列；請檢查串列埠設定及 Modbus 回應。
- 若測試失敗，先用 AccessPort 比較原本已經成功的封包；不要急著修改探頭內部站號或波特率。
- 本版本有 Windows 32-bit 原生 DLL 和 VBA 模組，但 CLI 不提供一般 `set` 指令；DLL 的平滑設定必須由呼叫者明確執行。

## 系統與相容性

- GitHub Actions 使用 `windows-2022` + Visual Studio 2022 編譯 **32-bit x86** 可執行檔，並以 `/MT` 靜態連結 MSVC C Runtime，避免部署時額外安裝 Visual C++ Redistributable。
- 原始碼限制使用 Windows 7 已存在的 Win32 API；但 **Windows 7 32-bit 相容性必須在實機測試確認**，雲端編譯成功不代表 Win7 一定能啟動。
- 支援標準 USB 虛擬 COM Port，RS485 轉換器須具備自動收發方向切換；若轉換器需要專用 RTS 控制，必須調整串列埠設定。
- `LS129CLI.exe` 不使用 Python，也沒有預設進行自動寫入或外部校正。

## 資料依據

林上官方文件：《LS129 Digital Probe Communication Protocol (Customer) V3.2》：
https://images.linshangtech.com/user-manual/LS129_Communication_Protocol_V3.2.pdf

## DLL 與 Excel / Access VBA

V0.3 新增 `LS129Core.dll` 和 `M_LS129.bas`，兩者可放在 MDB/XLS 專案同一資料夾。
依 `vba/README_VBA.md` 的範例在 Access 2003 用 `CurrentProject.Path`、
在 Excel 用 `ThisWorkbook.Path` 呼叫。

核心 API：開啟、讀取、讀全部、重新統計、讀/設定平滑模式、關閉。
DLL 預設不清零；一次新的量測要明確呼叫 `LS129_NewPeriod()`。
儀器平滑設定 0=無、1=50Hz、2=60Hz，寫入後會讀回驗證。
站號、波特率寫入尚未開放，避免產線錯誤設定造成失聯。
C 通訊核心在 `src/device.c`，CLI/DLL 共用。

## 下一階段

在 Win7 x86 與 Access2003 / Excel2003 實機測試 DLL 和 VBA；
再加入站號、波特率寫入的安全保護與失聯救援機制。
校正因子始終維持 PC 端處理。
