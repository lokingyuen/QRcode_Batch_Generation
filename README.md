# QR Code Generator 使用說明

## 介紹

這個工具允許您根據一個 Excel 文件中的數據批量生成 QR 代碼圖片。工具會根據 Excel 文件的內容生成對應的 QR 代碼並將其保存為 PNG 格式的圖片。每個圖片將根據 Excel 第一列的名稱來命名。

## 系統要求

- **操作系統**：Windows 10/11（適用於打包的 [EXE](https://github.com/lokingyuen/QRcode_Batch_Generation/releases) 版本）
- **無需安裝 Python**：EXE 版本無需安裝 Python 環境，直接運行即可。

## 安裝與運行

1. **下載 QR Code Generator EXE 文件**：
   - 下載 `QR Code Generator` 的 [EXE](https://github.com/lokingyuen/QRcode_Batch_Generation/releases) 版本。

2. **運行 EXE 文件**：
   - 直接雙擊運行下載的 `QR Code Generator.exe` 文件。
   - 您無需安裝 Python 或其他依賴項，EXE 文件是獨立運行的。

## 使用方法

### 1. 準備 Excel 文件

創建一個 Excel 文件（例如 `input.xlsx`），並按以下格式準備數據：

| 圖片名稱   | 網址                                |
|------------|-------------------------------------|
| QR_Code_1  | https://example.com/1               |
| QR_Code_2  | https://example.com/2               |
| QR_Code_3  | https://example.com/3               |
| QR_Code_4  | https://example.com/4               |

- **第一列**：圖片名稱，將用作生成的 QR 代碼圖片的文件名。
- **第二列**：對應的 URL，該 URL 會被編碼成 QR 代碼。

> **注意**：如果您的 Excel 文件中包含標題行（例如 "圖片名稱" 和 "網址"），請確保這些行位於第一行（程序會跳過這一行）。

### 2. 運行程序

####  使用方法：

1. 雙擊運行 `QR Code Generator.exe`。
2. 點擊 "Open Excel File" 按鈕，選擇您的 `input.xlsx` 文件。
3. 選擇保存 QR 代碼圖片的文件夾。
4. 點擊 "Batch Generation" 按鈕，程序會根據 Excel 文件中的數據生成 QR 代碼，並將每個 QR 代碼保存為 PNG 文件，命名為圖片名稱（例如 `QR_Code_1.png`）。

### 3. 生成的文件

所有生成的 QR 代碼圖片將保存為 PNG 格式，並存儲在您選擇的文件夾中。每個圖片將根據 Excel 中的第一列命名（例如 `QR_Code_1.png`）。

> Friendships never go out of style, But betrayal will
