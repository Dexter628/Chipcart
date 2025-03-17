<?php
/**
 * 利用 shell_exec 與 unzip 命令取出 XLSX 檔案內的 XML，
 * 再以正則表達式解析 sharedStrings.xml 與 sheet1.xml，
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
    $sharedStrings = array();
    if ($sharedStringsXML) {
        if (preg_match_all('/<t[^>]*>(.*?)<\/t>/s', $sharedStringsXML, $matches)) {
            $sharedStrings = $matches[1];
        }
    }

    // 解析 sheet1.xml：先用正則抓取所有 <row>…</row> 區塊
    $rows = array();
    if (preg_match_all('/<row[^>]*>(.*?)<\/row>/s', $sheetXML, $rowMatches)) {
        foreach ($rowMatches[1] as $rowContent) {
            $rowData = array();
            // 針對每個儲存格 <c ...>...</c>
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
        
        // 第一列視為 Excel 表頭
        $headers = $rows[0];

        //=================================================================
        // 定義資料庫欄位對應的 Excel 表頭關鍵字（只要表頭中包含關鍵字即可匹配）
        //=================================================================
        $fieldsMap = array(
            'part_no'            => array('P/N', 'Part No.', 'PartNo', '型号', 'Your internal Part id', 'Manufacturer Part Number', 'PART NO'),
            'manufacturer_name'  => array('MFG', 'MNF', 'Mfg', '厂商', 'Manufacturer Name', 'BRAND'),
            'available_qty'      => array('QTY', 'Quantity', '数量', 'Quantity (free on Hand)', 'QUANTITY'),
            'lead_time'          => array('L/T', 'LeadTime'),
            'price'              => array('PRICE', 'Cost', '销售价', '人民币', '美金'),
            'currency'           => array('Currency', 'USD', 'usd', 'rmb', 'RMB', 'CNY', 'cny', 'cn', 'us'),
            'moq'                => array('MOQ', '起订量', 'Minimum Order Quantity'),
            'spq'                => array('SPQ'),
            'order_increment'    => array('Order Increment / Pack Qty', 'Pack Qty', 'Order Increment'),
            'qty_1'              => array('Qty 1'),
            'qty_1_price'        => array('Qty 1 price'),
            'qty_2'              => array('Qty 2'),
            'qty_2_price'        => array('Qty 2 price'),
            'qty_3'              => array('Qty 3'),
            'qty_3_price'        => array('Qty 3 price'),
            'supplier_code'      => array('supplier code', '供应商代码', '供应商编码', '供应商代号'),
            // update_time 為自動填入當下時間，不需匹配
            'warranty'           => array('Warranty / Pedigree Rating', 'Warranty', 'Pedigree Rating'),
            'rohs_compliant'     => array('RoHS Compliant'),
            'eccn_code'          => array('ECCN Code'),
            'hts_code'           => array('HTS Code'),
            'warehouse_code'     => array('仓库位置', 'Warehouse Code'),
            'certificate_origin' => array('Country Of Origin', 'CO,'),
            'packing'            => array('PACKING'),
            'date_code_range'    => array('DC', 'DateCode', '批号', 'Date Code Range'),
            'package'            => array('Package', '封装', 'PACKAGE'),
            'package_type'       => array('Package Type'),
            'price_validity'     => array('Price validity,'),
            'contact'            => array('聯絡人', '業務', 'contact'),
            'part_description'   => array('产品参数', 'Part Description')
        );
        // tax_included 會根據特殊邏輯計算，不在 mapping 中

        //===============================================================
        // 建立 Excel 表頭（索引） 與 資料庫欄位對應關係
        //===============================================================
        $colMapping = array();
        foreach ($headers as $index => $header) {
            foreach ($fieldsMap as $field => $aliases) {
                foreach ($aliases as $alias) {
                    if (stripos($header, $alias) !== false) {
                        $colMapping[$field] = $index;
                        break 2;
                    }
                }
            }
        }
        
        //=======================================
        // 判斷 currency 欄位與特定 header 資訊
        //=======================================
        $detectedCurrency = "";
        // 若 Excel 中有匹配到 "currency" 欄位，則優先使用
        if (isset($colMapping['currency']) && isset($headers[$colMapping['currency']])) {
            $detectedCurrency = strtoupper(trim($headers[$colMapping['currency']]));
        }
        // 若未從 "currency" 欄位取得，則檢查 price 欄位標題中是否附帶貨幣資訊（例如 PRICE/USD 或 Qty 1 price (USD)）
        if (!$detectedCurrency && isset($colMapping['price'])) {
            $priceHeader = $headers[$colMapping['price']];
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
        // 如果貨幣為 USD，自動設定為 0 (未稅)
        // 如果為 RMB/CNY，則檢查所有表頭是否含有 "Tax Included"、"含税" 或 "含稅"
        // 其他情況則預設 0
        if ($detectedCurrency === "USD") {
            $tax_included = 0;
        } elseif (in_array($detectedCurrency, array("RMB", "CNY"))) {
            $taxHeaderFound = false;
            foreach ($headers as $header) {
                if (stripos($header, 'Tax Included') !== false || stripos($header, '含税') !== false || stripos($header, '含稅') !== false) {
                    $taxHeaderFound = true;
                    break;
                }
            }
            $tax_included = $taxHeaderFound ? 1 : 0;
        } else {
            $tax_included = 0;
        }
        
        //=====================================================
        // 將每一行資料依據 mapping 轉換成最終格式，並自動填入 update_time 與 tax_included
        //=====================================================
        $insertedData = array();
        // 從第二列開始（第一列為表頭）
        foreach (array_slice($rows, 1) as $row) {
            $data = array();
            foreach ($fieldsMap as $field => $aliases) {
                $data[$field] = (isset($colMapping[$field]) && isset($row[$colMapping[$field]]))
                                ? trim($row[$colMapping[$field]])
                                : '';
            }
            // 自動填入 update_time
            $data['update_time'] = date('Y-m-d H:i:s');
            // 填入 tax_included（由上面邏輯決定，對所有資料一致）
            $data['tax_included'] = $tax_included;
            // 填入 currency（全局判斷）
            $data['currency'] = $detectedCurrency;
            
            // 將結果存入模擬陣列
            $insertedData[] = $data;
        }
        
        // 模擬輸出最終結果（僅顯示前 5 筆資料）
        echo "<h3>模擬最終要插入 chip_db 表的資料（前 5 筆）：</h3>";
        echo "<pre>" . print_r(array_slice($insertedData, 0, 5), true) . "</pre>";
        
        // 若要正式插入資料，可使用 ThinkPHP Db 操作，例如：
        // foreach ($insertedData as $data) {
        //     Db::name('chip_db')->insert($data);
        // }
        
    } else {
        echo "檔案上傳失敗。";
    }
} else {
    // 未上傳檔案時，顯示上傳表單
    ?>
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Excel 上傳導入（完整匹配）</title>
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
