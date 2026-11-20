# LS129CLI V0.1 — LS129 RS485 / Modbus RTU 命令列讀取工具

**目的：讓沒有安裝 Python、沒有 C/C++ 編譯器的 Windows 電腦，透過 COM Port 讀取林上 LS129。**

目前是**第一版原始碼與 GitHub Actions 自動編譯配置**，不是已完成的實機認證產品。GitHub 編譯後，會產生 `LS129CLI.exe`，使用電腦只需要 USB-RS485 驅動程式和這個 EXE、INI，不需要編譯器。

## 第一次使用 — 只需要 5 個步驟

1. 在 GitHub 建立一個新的 repository，例如 `ls129-cli`。
2. 將這個 ZIP 解壓縮後，**保留 `.github/workflows/build.yml` 目錄結構**，把所有檔案上傳到 repository 根目錄。請不要只上傳 ZIP 本身。
3. 在 GitHub 開啟 **Actions → Build LS129CLI x86 → Run workflow**（如上傳時已自動觸發 push，則可以直接觀看執行結果）。
4. 成功後在該次 Actions 執行頁的 **Artifacts** 下載 `LS129CLI-x86-V0.1`，解壓縮取得 `LS129CLI.exe`、`LS129.ini` 和 `RUN_FIRST.bat`。
5. 編輯 INI 中 `port=COM3`，改為裝置管理員實際的 COM Port，接上 LS129；在 CMD 中執行 `LS129CLI.exe selftest`，再執行 `LS129CLI.exe read`。你也可以直接執行 `RUN_FIRST.bat`，依序進行自我測試和實機讀取。

## 指令範例

```bat
LS129CLI.exe selftest
LS129CLI.exe read
LS129CLI.exe status
LS129CLI.exe record --count 100 --out data.csv
LS129CLI.exe read --port COM4 --station 1 --baud 9600
LS129CLI.exe record --count 40 --hz 0.2 --out test.csv
```

`read`：只輸出一個無單位的原始整數，適合讓其他 BAT 或 VBA 擷取。`status`：顯示原始即時值、最大值、能量與 PC 端修正值。`record`：依照 PC 輪詢頻率輸出 CSV；**不要把重複讀值視為儀器內部產生的新量測**。`selftest`：只測試 CRC 與解碼，不需要探頭，不能當作實機通訊成功的證據。

### 設定檔

`LS129.ini` 放在 EXE **同一目錄**。`poll_hz` 可使用 0.2、1、5、32 等正數，最高 32；此參數只控制 PC 發送讀取命令的頻率，不改寫 LS129 的實際取樣頻率。PC 端校正係數由 `calibration.enabled` 和 `calibration.factor` 管理，**絕不寫入探頭**。

第一版僅支援讀取暫存器 **201–204**（03H）。讀取指令在站號 1 時為 `01 03 00 C9 00 04 94 37`。這是**整數格式**；單位必須依探頭型號與原始資料確認，勿直接假設所有型號都採相同物理單位。如果想保留小數，可以在下一版使用相應的 Float 暫存器。

## 常見問題

- `Cannot open/configure COMx`：檢查 USB-RS485 驅動、COM Port、是否被 AccessPort 或 Python 軟體佔用。不可同時讓兩個程式操作同一個 COM Port。
- `read failed (-20)`：沒有收到足夠的回應；檢查站號、波特率、供電、A/B 線與 USB-RS485 轉換器的收發方向控制。
- `read failed (-3)`：回應的 CRC 不符合，應檢查雜訊、接線與串列埠設定。
- 若測試失敗，先用 AccessPort 比較原本已經成功的封包；不要急著修改探頭內部站號或波特率。
- 本版本尚無 DLL、VBA 模組或 `set` 指令；先驗證 CMD 通訊，再做 SDK 第 2 階段。

## 系統與相容性

- GitHub Actions 使用 `windows-2022` + Visual Studio 2022 編譯 **32-bit x86** 可執行檔，並以 `/MT` 靜態連結 MSVC C Runtime，避免部署時額外安裝 Visual C++ Redistributable。
- 原始碼限制使用 Windows 7 已存在的 Win32 API；但 **Windows 7 32-bit 相容性必須在實機測試確認**，雲端編譯成功不代表 Win7 一定能啟動。
- 支援標準 USB 虛擬 COM Port，RS485 轉換器須具備自動收發方向切換；若轉換器需要專用 RTS 控制，必須調整串列埠設定。
- `LS129CLI.exe` 不使用 Python，也沒有預設進行自動寫入或外部校正。

## 資料依據

林上官方文件：《LS129 Digital Probe Communication Protocol (Customer) V3.2》：
https://images.linshangtech.com/user-manual/LS129_Communication_Protocol_V3.2.pdf

## 下一階段

CLI 與實機通訊驗證後，建立共用核心 DLL（Win32 x86）、VBA `Declare` 介面，以及保護性參數寫入（平滑 50/60Hz、站號、波特率），寫入前後均做確認與讀回驗證。校正因子維持 PC 端處理。
