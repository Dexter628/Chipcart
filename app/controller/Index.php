<?php
/**
 * 將 Excel 儲存格參考字母轉換為 0-based 欄位索引
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
 * 解析 XLSX 文件（僅限格式較簡單的 XLSX），使用 shell_exec 搭配 unzip 命令
 * 並根據每個儲存格的 r 屬性確定其正確欄位索引，補齊缺失欄位。
 *
 * @param string $filePath XLSX 檔案路徑
 * @return array|false 成功返回二維陣列（第一列為表頭），失敗返回 false
 */
function readXLSXWithoutExtensions($filePath) {
    // 透過 shell_exec 取得 sharedStrings.xml 與 sheet1.xml
    $sharedStringsXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/sharedStrings.xml");
    $sheetXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/worksheets/sheet1.xml");

    if (!$sheetXML) {
        return false;
    }

    // 解析 sharedStrings.xml，利用正則取得所有 <t> 標籤內容
    $sharedStrings = [];
    if ($sharedStringsXML && preg_match_all('/<t[^>]*>(.*?)<\/t>/s', $sharedStringsXML, $matches)) {
        $sharedStrings = $matches[1];
    }

    // 解析 sheet1.xml，使用正則取得所有 <row> ... </row> 區塊
    $rows = [];
    if (preg_match_all('/<row[^>]*>(.*?)<\/row>/s', $sheetXML, $rowMatches)) {
        foreach ($rowMatches[1] as $rowContent) {
            // 使用 cell reference r 屬性確定正確欄位
            $rowDataTemp = [];
            if (preg_match_all('/<c\s+[^>]*>(.*?)<\/c>/s', $rowContent, $cellMatches, PREG_SET_ORDER)) {
                foreach ($cellMatches as $cellMatch) {
                    $cellXml = $cellMatch[0];
                    $cellType = "";
                    if (preg_match('/t="([^"]+)"/', $cellXml, $tMatch)) {
                        $cellType = $tMatch[1];
                    }
                    $cellValue = "";
                    if (preg_match('/<v>(.*?)<\/v>/s', $cellXml, $vMatch)) {
                        $cellValue = $vMatch[1];
                    } elseif ($cellType == "inlineStr") {
                        if (preg_match('/<is>.*?<t[^>]*>(.*?)<\/t>.*?<\/is>/s', $cellXml, $inlineMatch)) {
                            $cellValue = $inlineMatch[1];
                        }
                    }
                    if ($cellType === 's') {
                        $index = intval($cellValue);
                        $cellValue = isset($sharedStrings[$index]) ? $sharedStrings[$index] : $cellValue;
                    }
                    // 取得 cell reference r 屬性，例如 r="B1"，取出 "B" 並轉換為欄位索引
                    $colIndex = null;
                    if (preg_match('/r="([A-Z]+)\d+"/', $cellXml, $rMatch)) {
                        $colIndex = colIndexFromLetter($rMatch[1]);
                    }
                    if ($colIndex === null) {
                        $colIndex = count($rowDataTemp);
                    }
                    $rowDataTemp[$colIndex] = $cellValue;
                }
            }
            if (!empty($rowDataTemp)) {
                $maxIndex = max(array_keys($rowDataTemp));
                $rowData = [];
                for ($i = 0; $i <= $maxIndex; $i++) {
                    $rowData[] = isset($rowDataTemp[$i]) ? $rowDataTemp[$i] : "";
                }
                $rows[] = $rowData;
            }
        }
    }
    return $rows;
}

/**
 * 將表頭正規化：去除前置的 "*" 與空白
 *
 * @param string $str 原始表頭
 * @return string 正規化後的表頭
 */
function normalizeHeader($str) {
    return trim(preg_replace('/^[\*\s]+/', '', $str));
}

//===============================================
// 主程式：上傳 XLSX、解析並映射資料
//===============================================
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

        // 第一列作為表頭，正規化後存入 $normHeaders
        $headers = $rows[0];
        $normHeaders = [];
        foreach ($headers as $h) {
            $normHeaders[] = normalizeHeader($h);
        }

        /*
         定義資料庫欄位對應的 Excel 表頭關鍵字（只要表頭中包含該關鍵字即可匹配），
         包含簡體與繁體版本。
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
            // update_time 為自動填入當下時間，不做匹配
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

        // 建立 Excel 表頭（索引）與資料庫欄位對應關係（利用正規化後的表頭）
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

        // 檢查是否存在「含税」、「含稅」或「Tax Included」的欄位
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
        // 判斷貨幣：先看是否有 currency 欄位，若無則從 price 標題中取
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
        // 如果貨幣為 USD，則預設 0 (未稅)
        // 如果為 RMB/CNY，預設 0，但若 Excel 中有含税欄位則以其資料為準
        if ($detectedCurrency === "USD") {
            $computedTaxIncluded = 0;
        } elseif (in_array($detectedCurrency, ["RMB", "CNY"])) {
            $computedTaxIncluded = 0;
        } else {
            $computedTaxIncluded = 0;
        }

        //=============================================
        // 將每一行資料依據 mapping 轉換成最終格式，並自動填入 update_time、currency 與 tax_included
        //=============================================
        $insertedData = [];
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
            $insertedData[] = $data;
        }

        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(array_slice($insertedData, 0, 5), JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);

        // 正式應用中，您可將 $insertedData 寫入資料庫，例如：
        // foreach ($insertedData as $data) {
        //     Db::name('chip_db')->insert($data);
        // }
    } else {
        echo "檔案上傳失敗。";
    }
} else {
    ?>
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Excel 上傳導入（完整匹配+正規化+含税導入）</title>
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
