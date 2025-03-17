<?php
// index.php

// 設定檔案上傳目錄（請確保該目錄存在且具有寫入權限）
$uploadDir = __DIR__ . '/uploads/';
if (!is_dir($uploadDir)) {
    mkdir($uploadDir, 0777, true);
}

/**
 * 將儲存格參考字母轉為 0-based 欄位索引
 * 例如："A" => 0, "B" => 1, "AA" => 26
 *
 * @param string $letters 儲存格字母部分
 * @return int 欄位索引（0-based）
 */
function colIndexFromLetter($letters) {
    $letters = strtoupper($letters);
    $result = 0;
    $len = strlen($letters);
    for ($i = 0; $i < $len; $i++) {
        $result = $result * 26 + (ord($letters[$i]) - ord('A') + 1);
    }
    return $result - 1;
}

/**
 * 解析 XLSX 文件（僅適用於格式較簡單的 XLSX）
 * 利用 shell_exec 與 unzip 命令提取 xl/sharedStrings.xml 與 xl/worksheets/sheet1.xml，
 * 並使用正則式與 cell reference（r 屬性）還原正確欄位順序，補齊缺失欄位。
 *
 * @param string $filePath XLSX 檔案路徑
 * @return array|false 返回二維陣列（第一列為表頭），失敗返回 false
 */
function readXLSXWithoutExtensions($filePath) {
    // 取出 sharedStrings.xml 與 sheet1.xml
    $sharedStringsXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/sharedStrings.xml");
    $sheetXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/worksheets/sheet1.xml");

    if (!$sheetXML) {
        return false;
    }

    // 解析 sharedStrings.xml，取得所有 <t> 標籤內容
    $sharedStrings = [];
    if ($sharedStringsXML && preg_match_all('/<t[^>]*>(.*?)<\/t>/s', $sharedStringsXML, $matches)) {
        $sharedStrings = $matches[1];
    }

    $rows = [];
    // 解析所有 <row>…</row> 區塊
    if (preg_match_all('/<row[^>]*>(.*?)<\/row>/s', $sheetXML, $rowMatches)) {
        foreach ($rowMatches[1] as $rowContent) {
            $rowDataTemp = [];
            // 取得每個儲存格，捕捉 r 屬性與內容
            if (preg_match_all('/<c\s+[^>]*r="([A-Z]+)\d+"[^>]*>(.*?)<\/c>/s', $rowContent, $cellMatches, PREG_SET_ORDER)) {
                foreach ($cellMatches as $cellMatch) {
                    $colLetters = $cellMatch[1];
                    $cellXmlContent = $cellMatch[2];
                    $colIndex = colIndexFromLetter($colLetters);
                    $cellType = "";
                    if (preg_match('/t="([^"]+)"/', $cellMatch[0], $tMatch)) {
                        $cellType = $tMatch[1];
                    }
                    $cellValue = "";
                    if (preg_match('/<v>(.*?)<\/v>/s', $cellXmlContent, $vMatch)) {
                        $cellValue = $vMatch[1];
                    } elseif ($cellType == "inlineStr") {
                        if (preg_match('/<is>.*?<t[^>]*>(.*?)<\/t>.*?<\/is>/s', $cellXmlContent, $inlineMatch)) {
                            $cellValue = $inlineMatch[1];
                        }
                    }
                    if ($cellType === 's') {
                        $index = intval($cellValue);
                        $cellValue = isset($sharedStrings[$index]) ? $sharedStrings[$index] : $cellValue;
                    }
                    $rowDataTemp[$colIndex] = $cellValue;
                }
            }
            if (!empty($rowDataTemp)) {
                $maxIndex = max(array_keys($rowDataTemp));
                $row = [];
                for ($i = 0; $i <= $maxIndex; $i++) {
                    $row[] = isset($rowDataTemp[$i]) ? $rowDataTemp[$i] : "";
                }
                $rows[] = $row;
            }
        }
    }
    return $rows;
}

/**
 * 將表頭正規化：去除前置的 "*" 與空白字符
 *
 * @param string $str 原始表頭
 * @return string 正規化後的表頭
 */
function normalizeHeader($str) {
    return trim(preg_replace('/^[\*\s]+/', '', $str));
}

/**
 * Excel 檔案後續解析處理：
 * 根據預先定義的欄位對應規則，對表頭進行正規化並匹配，
 * 再處理價格、貨幣及含税資訊，返回最終資料陣列
 *
 * @param array $rows 解析後的 Excel 二維陣列
 * @return array 最終資料陣列
 */
function processExcelRows($rows) {
    // 取出第一列表頭並正規化
    $headers = $rows[0];
    $normHeaders = [];
    foreach ($headers as $h) {
        $normHeaders[] = normalizeHeader($h);
    }

    // 定義欄位對應規則（包含簡繁體）
    $fieldsMap = [
        'part_no'            => ['P/N', 'Part No.', 'PartNo', '型号', 'Your internal Part id', 'Manufacturer Part Number', 'PART NO'],
        'manufacturer_name'  => ['MFG', 'MNF', 'Mfg', '厂商', '廠商', 'Manufacturer Name', 'BRAND'],
        'available_qty'      => ['QTY', 'Quantity', '数量', '數量', 'Quantity (free on Hand)', 'QUANTITY'],
        'lead_time'          => ['L/T', 'LeadTime'],
        'price'              => ['PRICE', 'Cost', '销售价', '銷售價', '人民币', '美金'],
        'currency'           => ['Currency', 'USD', 'usd', 'rmb', 'RMB', 'CNY', 'cny', 'cn', 'us'],
        'moq'                => ['MOQ', '起订量', '起訂量', 'Minimum Order Quantity'],
        'spq'                => ['SPQ'],
        'order_increment'    => ['Order Increment / Pack Qty', 'Pack Qty', 'Order Increment'],
        'qty_1'              => ['Qty 1'],
        'qty_1_price'        => ['Qty 1 price'],
        'qty_2'              => ['Qty 2'],
        'qty_2_price'        => ['Qty 2 price'],
        'qty_3'              => ['Qty 3'],
        'qty_3_price'        => ['Qty 3 price'],
        'supplier_code'      => ['supplier code', '供应商代码', '供应商编码', '供应商代号', '供應商代號'],
        // update_time 直接填入當下時間，不匹配
        'warranty'           => ['Warranty / Pedigree Rating', 'Warranty', 'Pedigree Rating'],
        'rohs_compliant'     => ['RoHS Compliant'],
        'eccn_code'          => ['ECCN Code'],
        'hts_code'           => ['HTS Code'],
        'warehouse_code'     => ['仓库位置', 'Warehouse Code', '倉庫位置'],
        'certificate_origin' => ['Country Of Origin', 'CO,'],
        'packing'            => ['PACKING'],
        'date_code_range'    => ['DC', 'DateCode', '批号', '批號', 'Date Code Range'],
        'package'            => ['Package', '封装', '封裝', 'PACKAGE'],
        'package_type'       => ['Package Type'],
        'price_validity'     => ['Price validity'],
        'contact'            => ['联络人', '聯絡人', '业务', '業務', 'contact'],
        'part_description'   => ['产品参数', '產品參數', 'Part Description']
        // tax_included 不直接匹配，後續特殊判斷
    ];

    // 建立表頭與欄位對應關係（使用正規化後的表頭索引）
    $colMapping = [];
    foreach ($normHeaders as $index => $normHeader) {
        foreach ($fieldsMap as $field => $aliases) {
            foreach ($aliases as $alias) {
                if (stripos($normHeader, $alias) !== false) {
                    $colMapping[$field] = $index;
                    break 2;
                }
            }
        }
    }

    // 檢查是否存在含税、含稅或 Tax Included 的欄位
    $taxIncludedCol = null;
    foreach ($normHeaders as $index => $normHeader) {
        if (stripos($normHeader, 'Tax Included') !== false ||
            stripos($normHeader, '含税') !== false ||
            stripos($normHeader, '含稅') !== false) {
            $taxIncludedCol = $index;
            break;
        }
    }

    // 判斷貨幣：優先取 currency 欄位，若無則從 price 標題中解析；若仍無則預設 USD
    $detectedCurrency = "";
    if (isset($colMapping['currency']) && isset($normHeaders[$colMapping['currency']])) {
        $detectedCurrency = strtoupper(trim($normHeaders[$colMapping['currency']]));
    }
    if (!$detectedCurrency && isset($colMapping['price'])) {
        $priceHeader = $normHeaders[$colMapping['price']];
        if (preg_match('/[\/\(]\s*([A-Za-z]+)\s*[\)\/]?/', $priceHeader, $match)) {
            $detectedCurrency = strtoupper(trim($match[1]));
        }
    }
    if (!$detectedCurrency) {
        $detectedCurrency = "USD";
    }

    // 根據貨幣決定 tax_included 預設值：USD 預設 0，若為 RMB/CNY 預設 0（可依需求調整）
    if ($detectedCurrency === "USD") {
        $computedTaxIncluded = 0;
    } elseif (in_array($detectedCurrency, ["RMB", "CNY"])) {
        $computedTaxIncluded = 0;
    } else {
        $computedTaxIncluded = 0;
    }

    // 將每一行資料轉換成最終格式，每筆自動填入 update_time、currency 與 tax_included
    $finalData = [];
    foreach (array_slice($rows, 1) as $row) {
        $data = [];
        foreach ($fieldsMap as $field => $aliases) {
            $data[$field] = (isset($colMapping[$field]) && isset($row[$colMapping[$field]]))
                            ? trim($row[$colMapping[$field]])
                            : '';
        }
        $data['update_time'] = date('Y-m-d H:i:s');
        $data['currency'] = $detectedCurrency;
        if ($taxIncludedCol !== null && isset($row[$taxIncludedCol]) && trim($row[$taxIncludedCol]) !== '') {
            $data['tax_included'] = trim($row[$taxIncludedCol]);
        } else {
            $data['tax_included'] = $computedTaxIncluded;
        }
        $finalData[] = $data;
    }
    return $finalData;
}

/**
 * 處理上傳檔案：合併分段後解析 Excel 並回傳結果（JSON 格式）
 */
function processUploadedExcel($mergedFilePath) {
    $rows = readXLSXWithoutExtensions($mergedFilePath);
    if (!$rows || count($rows) < 2) {
        return "解析 Excel 文件失敗或資料不足。";
    }
    $finalData = processExcelRows($rows);
    return $finalData;
}

/**
 * 以下為分段上傳處理：
 * 當接收到 POST 請求且有 fileId 參數時，表示為分段上傳
 */
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['fileId'])) {
    $fileId = $_POST['fileId'];
    $chunkIndex = isset($_POST['chunkIndex']) ? intval($_POST['chunkIndex']) : 0;
    $totalChunks = isset($_POST['totalChunks']) ? intval($_POST['totalChunks']) : 0;
    $fileName = isset($_POST['fileName']) ? $_POST['fileName'] : '';

    $targetFile = $uploadDir . $fileId . '_' . $fileName;
    $out = fopen($targetFile, "ab");
    if (!$out) {
        http_response_code(500);
        echo "無法開啟目標檔案";
        exit;
    }
    $in = fopen($_FILES['fileChunk']['tmp_name'], "rb");
    if ($in) {
        while ($buff = fread($in, 4096)) {
            fwrite($out, $buff);
        }
        fclose($in);
    }
    fclose($out);

    if ($chunkIndex + 1 == $totalChunks) {
        $result = processUploadedExcel($targetFile);
        // 選擇性：合併後可刪除臨時檔案
        // unlink($targetFile);
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode($result, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    } else {
        echo "區塊 $chunkIndex 已上傳";
    }
} else {
    // 若非分段上傳請求，則顯示上傳表單
    $html = <<<HTML
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>分段上傳 Excel 檔案</title>
</head>
<body>
    <h2>分段上傳 Excel 檔案 (每段約 1MB)</h2>
    <input type="file" id="fileInput" accept=".xlsx">
    <button id="uploadBtn">開始上傳</button>
    <div id="progress"></div>
    <script>
        document.getElementById('uploadBtn').addEventListener('click', function () {
            var fileInput = document.getElementById('fileInput');
            if (!fileInput.files.length) {
                alert("請選擇檔案");
                return;
            }
            var file = fileInput.files[0];
            var chunkSize = 1024 * 1024; // 每塊 1MB
            var totalChunks = Math.ceil(file.size / chunkSize);
            var currentChunk = 0;
            var fileId = Date.now() + "_" + file.name;
            var progressDiv = document.getElementById('progress');

            function uploadNextChunk() {
                var start = currentChunk * chunkSize;
                var end = Math.min(start + chunkSize, file.size);
                var blob = file.slice(start, end);
                var formData = new FormData();
                formData.append('fileId', fileId);
                formData.append('chunkIndex', currentChunk);
                formData.append('totalChunks', totalChunks);
                formData.append('fileName', file.name);
                formData.append('fileChunk', blob);

                var xhr = new XMLHttpRequest();
                xhr.open('POST', 'index.php', true);
                xhr.onload = function () {
                    if (xhr.status === 200) {
                        currentChunk++;
                        progressDiv.innerText = "已上傳 " + currentChunk + " / " + totalChunks + " 區塊";
                        if (currentChunk < totalChunks) {
                            uploadNextChunk();
                        } else {
                            progressDiv.innerText = "檔案上傳完成，正在處理 Excel 檔案...";
                            progressDiv.innerText = xhr.responseText;
                        }
                    } else {
                        alert("上傳失敗，請重試");
                    }
                };
                xhr.send(formData);
            }
            uploadNextChunk();
        });
    </script>
</body>
</html>
HTML;
    echo $html;
}
?>
