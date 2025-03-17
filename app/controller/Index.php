<?php
/**
 * 透過 shell_exec 與 unzip 命令取得 XLSX 檔內 XML，
 * 並利用正則表達式解析 sharedStrings.xml 與 sheet1.xml，
 * 最終返回一個二維陣列（第一列為表頭）。
 *
 * @param string $filePath XLSX 檔案路徑
 * @return array|false 成功返回二維陣列，失敗返回 false
 */
function readXLSXWithoutExtensions($filePath) {
    // 使用 shell_exec 提取 sharedStrings.xml 與 sheet1.xml
    $sharedStringsXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/sharedStrings.xml");
    $sheetXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/worksheets/sheet1.xml");

    if (!$sheetXML) {
        return false;
    }

    // 解析 sharedStrings.xml：利用正則抓取所有 <t> 標籤內容
    $sharedStrings = array();
    if ($sharedStringsXML) {
        if (preg_match_all('/<t[^>]*>(.*?)<\/t>/s', $sharedStringsXML, $matches)) {
            $sharedStrings = $matches[1];
        }
    }

    // 解析 sheet1.xml：先抓取所有 <row>…</row> 區塊
    $rows = array();
    if (preg_match_all('/<row[^>]*>(.*?)<\/row>/s', $sheetXML, $rowMatches)) {
        foreach ($rowMatches[1] as $rowContent) {
            $rowData = array();
            // 解析每一個 <c> (cell) 節點
            // 捕捉：可有 t 屬性（表示型別），以及 <v> 內的數值
            if (preg_match_all('/<c[^>]*(?:t="([^"]+)")?[^>]*>.*?<v>(.*?)<\/v>/s', $rowContent, $cellMatches, PREG_SET_ORDER)) {
                foreach ($cellMatches as $cell) {
                    $cellType  = isset($cell[1]) ? $cell[1] : "";
                    $cellValue = $cell[2];
                    // 若型別為 s，則表示是 shared string，需要查表
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

//-----------------------------------------------------
// 以下為完整範例：上傳 XLSX 檔案、解析、映射、模擬資料庫插入 chip_db 表
//-----------------------------------------------------
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

        // 第一列作為表頭
        $headers = $rows[0];

        /*
         定義資料庫欄位對應的 Excel 表頭關鍵字，
         您可根據實際情況調整或擴充
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

        // 建立 Excel 表頭（索引） 與 資料庫欄位對應關係
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

        // 檢查表頭中是否包含 "(USD)" 來判斷貨幣資訊（預設為 USD）
        $detectedCurrency = 'USD';
        foreach ($headers as $header) {
            if (preg_match('/\((.*?)\)/', $header, $matches)) {
                $curr = strtoupper(trim($matches[1]));
                if ($curr === 'USD') {
                    $detectedCurrency = 'USD';
                    break;
                }
            }
        }

        // 模擬將每一行資料轉換為最終要插入 chip_db 表的格式
        $insertedData = array();
        // 從第二列開始（第一列為表頭）
        foreach (array_slice($rows, 1) as $row) {
            $data = array();
            // 根據 mapping 取得對應欄位數值；若無對應則設為空字串
            foreach ($fieldsMap as $field => $aliases) {
                $data[$field] = (isset($colMapping[$field]) && isset($row[$colMapping[$field]]))
                                ? trim($row[$colMapping[$field]])
                                : '';
            }

            // 處理 currency 與 tax_included：
            // 若檢測到 "(USD)"，則 currency 為 USD 且 tax_included 設為 0 (未稅)
            $data['currency'] = $detectedCurrency;
            $data['tax_included'] = ($detectedCurrency === 'USD') ? 0 : 1;

            // 自動填入當前更新時間
            $data['update_time'] = date('Y-m-d H:i:s');

            // 模擬「插入」：實際應用中可用 ThinkPHP Db::name('chip_db')->insert($data)
            $insertedData[] = $data;
        }

        // 輸出模擬結果 (僅顯示前 5 筆)
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
        <title>Excel 上傳導入（無 XML 擴展）</title>
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
