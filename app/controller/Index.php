<?php
// 引入 Parts 模型（請依據您的專案命名空間調整）
use app\model\Parts;

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
 * 解析 XLSX 文件（僅適用於格式較簡單的 XLSX），
 * 利用 shell_exec 搭配 unzip 命令提取 xl/sharedStrings.xml 與 xl/worksheets/sheet1.xml，
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
            // 用正則捕捉每個儲存格，並取得 r 屬性（例如 r="B1"）與其內容
            if (preg_match_all('/<c\s+[^>]*r="([A-Z]+)\d+"[^>]*>(.*?)<\/c>/s', $rowContent, $cellMatches, PREG_SET_ORDER)) {
                foreach ($cellMatches as $cellMatch) {
                    $colLetters = $cellMatch[1];
                    $cellXmlContent = $cellMatch[2];
                    $colIndex = colIndexFromLetter($colLetters);
                    // 取得儲存格的型別 t 屬性（若有）
                    $cellType = "";
                    if (preg_match('/t="([^"]+)"/', $cellMatch[0], $tMatch)) {
                        $cellType = $tMatch[1];
                    }
                    $cellValue = "";
                    // 優先從 <v> 標籤中取得數值
                    if (preg_match('/<v>(.*?)<\/v>/s', $cellXmlContent, $vMatch)) {
                        $cellValue = $vMatch[1];
                    }
                    // 若無 <v> 且型別為 inlineStr，則從 <is><t> 中取值
                    elseif ($cellType == "inlineStr") {
                        if (preg_match('/<is>.*?<t[^>]*>(.*?)<\/t>.*?<\/is>/s', $cellXmlContent, $inlineMatch)) {
                            $cellValue = $inlineMatch[1];
                        }
                    }
                    // 若型別為共享字串 (s)，則用 sharedStrings 表取得真正值
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

//==================================================
// 主程式：上傳 XLSX、解析並映射資料
//==================================================
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excel_file'])) {
    if ($_FILES['excel_file']['error'] === UPLOAD_ERR_OK) {
        $filePath = $_FILES['excel_file']['tmp_name'];
        $rows = readXLSXWithoutExtensions($filePath);
        if (!$rows) {
            exit("解析 Excel 文件失敗。");
        }
        if (count($rows) < 2) {
            exit("Excel 文件中沒有足夠的數據。");
        }
        
        // 取出第一列作為表頭，正規化後存入 $normHeaders
        $headers = $rows[0];
        $normHeaders = [];
        foreach ($headers as $h) {
            $normHeaders[] = normalizeHeader($h);
        }
        
        /*
         定義資料庫欄位對應的 Excel 表頭關鍵字（只要表頭中包含該關鍵字即可匹配），
         此處同時包含簡體與繁體版本。
        */
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
            // update_time 為自動填入，不做匹配
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
        ];
        
        // 建立 Excel 表頭（索引）與資料庫欄位對應關係（使用正規化後的表頭）
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
        
        // 檢查是否有「含税」、「含稅」或「Tax Included」的欄位，若有則記錄該欄索引
        $taxIncludedCol = null;
        foreach ($normHeaders as $index => $normHeader) {
            if (stripos($normHeader, 'Tax Included') !== false ||
                stripos($normHeader, '含税') !== false ||
                stripos($normHeader, '含稅') !== false) {
                $taxIncludedCol = $index;
                break;
            }
        }
        
        //=============================================
        // 判斷貨幣：先看是否有 currency 欄位，若無則從 price 標題中解析
        //=============================================
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
        
        //=============================================
        // 根據貨幣決定 tax_included 的預設值
        //=============================================
        if ($detectedCurrency === "USD") {
            $computedTaxIncluded = 0;
        } elseif (in_array($detectedCurrency, ["RMB", "CNY"])) {
            $computedTaxIncluded = 0;
        } else {
            $computedTaxIncluded = 0;
        }
        
        //=============================================
        // 將每一行資料依據 mapping 轉換成最終格式，並自動填入 currency 與 tax_included
        // 注意：此處不再手動設定 update_time，讓模型自動填入時間戳
        //=============================================
        $insertedData = [];
        foreach (array_slice($rows, 1) as $row) {
            $data = [];
            foreach ($fieldsMap as $field => $aliases) {
                $data[$field] = (isset($colMapping[$field]) && isset($row[$colMapping[$field]]))
                                ? trim($row[$colMapping[$field]])
                                : '';
            }
            $data['currency'] = $detectedCurrency;
            if ($taxIncludedCol !== null && isset($row[$taxIncludedCol]) && trim($row[$taxIncludedCol]) !== '') {
                $data['tax_included'] = trim($row[$taxIncludedCol]);
            } else {
                $data['tax_included'] = $computedTaxIncluded;
            }
            $insertedData[] = $data;
        }
        
        // 輸出前 5 筆資料 (僅供參考)
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(array_slice($insertedData, 0, 5), JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
        
        // 將 $insertedData 寫入資料庫 parts 表，使用 Parts 模型批量寫入
        try {
            (new Parts())->saveAll($insertedData);
            echo " 資料已成功寫入資料庫。";
        } catch (\Exception $e) {
            echo "資料庫錯誤：" . $e->getMessage();
        }
        
    } else {
        echo "檔案上傳失敗。";
    }
} else {
?>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Excel 上傳導入（完整匹配+正規化+CellRef+含税導入）</title>
</head>
<body>
    <h2>上傳 Excel 文件 (僅限 XLSX 格式)</h2>
    <form action="" method="post" enctype="multipart/form-data">
        <input type="file" name="excel_file" accept=".xlsx" required>
        <br><br>
        <input type="submit" value="上傳並模擬導入數據">
    </form>
</body>
</html>
<?php
}
?>
