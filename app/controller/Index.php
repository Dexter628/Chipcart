<?php
// 這個範例不依賴外部解析庫，而是用內建 ZipArchive 及 SimpleXML 解析 XLSX

/**
 * 解析 XLSX 檔案，返回二维陣列，每個元素為一行資料（第一列為表頭）。
 * 此函式為簡易實作，僅適用於基本的 XLSX 檔案。
 *
 * @param string $filePath XLSX 檔案路徑
 * @return array|false 成功返回二維陣列，失敗返回 false
 */
function readXLSX($filePath) {
    $zip = new ZipArchive;
    if ($zip->open($filePath) === true) {
        // 讀取 sharedStrings.xml
        $sharedStrings = [];
        $sharedStringsXML = $zip->getFromName('xl/sharedStrings.xml');
        if ($sharedStringsXML) {
            $xml = simplexml_load_string($sharedStringsXML);
            // 注意：有些檔案可能不含 <si> 標籤
            if ($xml && isset($xml->si)) {
                foreach ($xml->si as $si) {
                    // 有時候 <si> 直接有 <t>，有時候有多個 <r>
                    if (isset($si->t)) {
                        $sharedStrings[] = (string)$si->t;
                    } else {
                        $text = '';
                        foreach ($si->r as $r) {
                            $text .= (string)$r->t;
                        }
                        $sharedStrings[] = $text;
                    }
                }
            }
        }
        // 讀取第一個工作表 xl/worksheets/sheet1.xml
        $sheetXML = $zip->getFromName('xl/worksheets/sheet1.xml');
        if (!$sheetXML) {
            $zip->close();
            return false;
        }
        $xml = simplexml_load_string($sheetXML);
        $rows = [];
        if (isset($xml->sheetData->row)) {
            foreach ($xml->sheetData->row as $row) {
                $rowData = [];
                foreach ($row->c as $c) {
                    $value = (string)$c->v;
                    $type = (string)$c['t']; // 若為 s 表示使用 shared string
                    if ($type === 's') {
                        $index = intval($value);
                        $value = isset($sharedStrings[$index]) ? $sharedStrings[$index] : $value;
                    }
                    $rowData[] = $value;
                }
                $rows[] = $rowData;
            }
        }
        $zip->close();
        return $rows;
    }
    return false;
}

// -------------------------
// 下面為模擬 Excel 解析並映射到資料庫資料格式的完整程式
// -------------------------
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excel_file'])) {
    if ($_FILES['excel_file']['error'] === UPLOAD_ERR_OK) {
        $filePath = $_FILES['excel_file']['tmp_name'];
        $rows = readXLSX($filePath);
        if (!$rows) {
            exit("解析 Excel 文件失敗。");
        }
        if (count($rows) < 2) {
            exit("Excel 文件中沒有足夠的數據。");
        }
        // 第一列視為表頭
        $headers = $rows[0];
        
        /* 
          定義資料庫欄位對應的 Excel 表頭關鍵字（部分欄位根據需求匹配）
          可根據需要調整或擴充
        */
        $fieldsMap = array(
            'part_no'            => array('Your internal Part id', 'Part No.', 'PartNo', '型号', 'Manufacturer Part Number', 'PART NO'),
            'manufacturer_name'  => array('Manufacturer Name', 'MFG', 'MNF', '厂商', 'BRAND'),
            'part_description'   => array('Part Description', '产品参数'),
            'available_qty'      => array('Quantity (free on Hand)', 'QTY', 'Quantity', '数量'),
            'moq'                => array('Minimum Order Quantity', 'MOQ', '起订量'),
            'order_increment'    => array('Order Increment / Pack Qty', 'Pack Qty', 'Order Increment'),
            'date_code_range'    => array('Date Code Range', 'DC', 'DateCode', '批号'),
            'price'              => array('Resale (web price)', 'Cost (USD)', 'Cost'),
            'certificate_origin' => array('Country Of Origin', 'CO,'),
            'warranty'           => array('Warranty / Pedigree Rating', 'Warranty', 'Pedigree Rating'),
            'warehouse_code'     => array('Warehouse Code', 'Warehouse Code (if applicable)', '仓库位置'),
            'eccn_code'          => array('ECCN Code'),
            'hts_code'           => array('HTS Code'),
            'rohs_compliant'     => array('RoHS Compliant (Y/N)', 'RoHS Compliant'),
            'package_type'       => array('Package Type'),
            'qty_1'              => array('Qty 1 (pcs)', 'Qty 1'),
            'qty_1_price'        => array('Qty 1 price (USD)', 'Qty 1 price'),
            'qty_2'              => array('Qty 2 (pcs)', 'Qty 2'),
            'qty_2_price'        => array('Qty 2 price (USD)', 'Qty 2 price'),
            'qty_3'              => array('Qty 3 (pcs)', 'Qty 3'),
            'qty_3_price'        => array('Qty 3 price (USD)', 'Qty 3 price')
        );
        
        // 建立 Excel 表頭（索引） 與 資料庫欄位的對應關係
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
        
        // 檢查是否有價格欄位帶 "(USD)" 的資訊，以判斷貨幣
        $detectedCurrency = 'USD'; // 此處預設為 USD，您也可根據需要進行調整
        foreach ($headers as $header) {
            if (preg_match('/\((.*?)\)/', $header, $matches)) {
                $curr = strtoupper(trim($matches[1]));
                if ($curr === 'USD') {
                    $detectedCurrency = 'USD';
                    break;
                }
            }
        }
        
        // 模擬將每一行數據轉換為資料庫格式
        $insertedData = array();
        // 從第二列開始（第一列為表頭）
        foreach (array_slice($rows, 1) as $row) {
            $data = array();
            // 根據 mapping 取得各欄位數值；若無對應則設為空字串
            foreach ($fieldsMap as $field => $aliases) {
                $data[$field] = isset($colMapping[$field]) && isset($row[$colMapping[$field]])
                                  ? trim($row[$colMapping[$field]])
                                  : '';
            }
            
            // 處理 currency 與 tax_included 邏輯：
            // 如果價格欄位帶有 "(USD)"，則 currency 為 USD 且 tax_included = 0 (未稅)
            $data['currency'] = $detectedCurrency;
            if ($detectedCurrency === 'USD') {
                $data['tax_included'] = 0;
            } else {
                $data['tax_included'] = 1;
            }
            
            // 自動填入更新時間
            $data['update_time'] = date('Y-m-d H:i:s');
            
            // 模擬「插入」：實際應用中可使用 ThinkPHP Db::name('chip_db')->insert($data)
            $insertedData[] = $data;
        }
        
        // 輸出模擬結果 (顯示前 5 筆)
        echo "<h3>模擬最終要插入 chip_db 表的資料（前 5 筆）：</h3>";
        echo "<pre>" . print_r(array_slice($insertedData, 0, 5), true) . "</pre>";
        
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
        <title>Excel 上傳導入（原生解析 XLSX）</title>
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
