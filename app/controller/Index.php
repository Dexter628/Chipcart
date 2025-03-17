<?php
/**
 * 利用 shell_exec 與 unzip 命令提取 XLSX 檔案內的 XML，
 * 再利用正則表達式解析 sharedStrings.xml 與 sheet1.xml，
 * 返回一個二維陣列（第一列視為表頭）。
 *
 * 此實作僅適用於格式較簡單的 XLSX 文件。
 *
 * @param string $filePath XLSX 檔案路徑
 * @return array|false 解析成功返回二維陣列，失敗返回 false
 */
function readXLSXWithoutExtensions($filePath) {
    // 透過 shell_exec 提取 sharedStrings.xml 與 sheet1.xml
    $sharedStringsXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/sharedStrings.xml");
    $sheetXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/worksheets/sheet1.xml");

    if (!$sheetXML) {
        return false;
    }

    // 解析 sharedStrings.xml：用正則取得所有 <t> 標籤內容
    $sharedStrings = [];
    if ($sharedStringsXML) {
        if (preg_match_all('/<t[^>]*>(.*?)<\/t>/s', $sharedStringsXML, $matches)) {
            $sharedStrings = $matches[1];
        }
    }

    // 解析 sheet1.xml：先用正則抓取所有 <row>…</row> 區塊
    $rows = [];
    if (preg_match_all('/<row[^>]*>(.*?)<\/row>/s', $sheetXML, $rowMatches)) {
        foreach ($rowMatches[1] as $rowContent) {
            $rowData = [];
            // 針對每個儲存格 <c ...>...</c> 進行處理
            if (preg_match_all('/<c[^>]*>(.*?)<\/c>/s', $rowContent, $cellMatches, PREG_SET_ORDER)) {
                foreach ($cellMatches as $cellMatch) {
                    $cellXml = $cellMatch[0];
                    $cellType = "";
                    // 取得 t 屬性，可能為 s (shared string) 或 inlineStr
                    if (preg_match('/t="([^"]+)"/', $cellXml, $tMatch)) {
                        $cellType = $tMatch[1];
                    }
                    $cellValue = "";
                    // 先從 <v> 標籤中取值
                    if (preg_match('/<v>(.*?)<\/v>/s', $cellXml, $vMatch)) {
                        $cellValue = $vMatch[1];
                    }
                    // 若無 <v> 且屬性為 inlineStr，則嘗試從 <is><t> 中取得文字
                    elseif ($cellType == "inlineStr") {
                        if (preg_match('/<is>.*?<t[^>]*>(.*?)<\/t>.*?<\/is>/s', $cellXml, $inlineMatch)) {
                            $cellValue = $inlineMatch[1];
                        }
                    }
                    // 若屬性為 s（shared string），則用 sharedStrings 表取真正值
                    if ($cellType === 's') {
                        $index = intval($cellValue);
                        $cellValue = isset($sharedStrings[$index]) ? $sharedStrings[$index] : $cellValue;
                    }
                    $rowData[] = $cellValue;
                }
            }
            $rows[] = $rowData;
        }
    }
    return $rows;
}

/**
 * 正規化表頭函式：去除前置特殊字元（如 * 與空白）
 * @param string $str 原始表頭
 * @return string 正規化後的表頭
 */
function normalizeHeader($str) {
    return trim(preg_replace('/^[\*\s]+/', '', $str));
}

//==================================================
// 以下為上傳、解析並映射資料的完整範例程式
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
        
        // 第一列作為 Excel 表頭，正規化後存入 $normHeaders
        $headers = $rows[0];
        $normHeaders = [];
        foreach ($headers as $h) {
            $normHeaders[] = normalizeHeader($h);
        }
        
        /*
         定義資料庫欄位對應的 Excel 表頭關鍵字（包含簡體與繁體）
         這裡列出所有需要匹配的欄位，不包括 tax_included（後續特殊判斷）
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
            // tax_included 不在此處直接匹配
        ];
        
        // 建立 Excel 表頭（索引） 與 資料庫欄位對應關係（利用正規化後的表頭）
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
        
        // 單獨檢查是否有「含税」、「含稅」或「Tax Included」的欄位，若有則記錄該欄位索引
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
        // 判斷 currency 欄位與價格標題中是否含有貨幣資訊
        //=============================================
        $detectedCurrency = "";
        // 若 mapping 中有 currency 欄位，優先使用該欄位資料
        if (isset($colMapping['currency']) && isset($normHeaders[$colMapping['currency']])) {
            $detectedCurrency = strtoupper(trim($normHeaders[$colMapping['currency']]));
        }
        // 若未取得，再檢查 price 欄位標題中是否附帶貨幣資訊 (如 "PRICE/USD" 或 "Qty 1 price (USD)")
        if (!$detectedCurrency && isset($colMapping['price'])) {
            $priceHeader = $normHeaders[$colMapping['price']];
            if (preg_match('/[\/\(]\s*([A-Za-z]+)\s*[\)\/]?/', $priceHeader, $match)) {
                $detectedCurrency = strtoupper(trim($match[1]));
            }
        }
        // 若仍未取得，預設為 USD
        if (!$detectedCurrency) {
            $detectedCurrency = "USD";
        }
        
        //=============================================
        // 根據 currency 決定 tax_included 欄位值
        // 如果貨幣為 USD，則預設為 0 (未稅)
        // 如果為 RMB/CNY，則預設為 0；但若 Excel 有提供含稅欄位則以該欄資料為準
        if ($detectedCurrency === "USD") {
            $computedTaxIncluded = 0;
        } elseif (in_array($detectedCurrency, ["RMB", "CNY"])) {
            $computedTaxIncluded = 0;
            // 此處您可根據需求調整預設邏輯
        } else {
            $computedTaxIncluded = 0;
        }
        
        //=============================================
        // 將每一行資料依據 mapping 轉換成最終格式，並自動填入 update_time、currency 與 tax_included
        //=============================================
        $insertedData = [];
        // 從第二列開始（第一列為表頭）
        foreach (array_slice($rows, 1) as $row) {
            $data = [];
            foreach ($fieldsMap as $field => $aliases) {
                $data[$field] = (isset($colMapping[$field]) && isset($row[$colMapping[$field]]))
                                ? trim($row[$colMapping[$field]])
                                : '';
            }
            // 自動填入 update_time
            $data['update_time'] = date('Y-m-d H:i:s');
            // 填入 currency
            $data['currency'] = $detectedCurrency;
            // 決定 tax_included：若 Excel 有提供含税欄位，則以該欄資料（不為空時）；否則以預設 computedTaxIncluded
            if ($taxIncludedCol !== null && isset($row[$taxIncludedCol]) && trim($row[$taxIncludedCol]) !== '') {
                $data['tax_included'] = trim($row[$taxIncludedCol]);
            } else {
                $data['tax_included'] = $computedTaxIncluded;
            }
            
            $insertedData[] = $data;
        }
        
        // 輸出 JSON 結果（僅顯示前 5 筆）
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(array_slice($insertedData, 0, 5), JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
        
        // 正式應用中，您可以將 $insertedData 寫入資料庫，例如使用 ThinkPHP 的 Db::name('chip_db')->insert($data)
        
    } else {
        echo "檔案上傳失敗。";
    }
} else {
    // 顯示上傳表單
    ?>
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Excel 上傳導入（完整匹配+正規化+含稅導入）</title>
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
