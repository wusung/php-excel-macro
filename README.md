# Write excel with macro in PHP

### 系統需求

- PHP 5
- Microsoft Excel

### 安裝方式

1. 修改 php.ini extension_dir 非必要，若 php 安裝於 c:\php 不需要此步驟

```
extension_dir = "D:\php-5.6.40\ext"
```

2. 修改 php.ini 加上 COM_DOT_NET

```
[COM_DOT_NET]
extension=php_com_dotnet.dll
```

3. Excel 啟用 選項 > 信任中心 > 信任中心設定 > 巨集設定 > 信任存取VBA 專案物件模型

### 執行
```
php.exe app.php
```
