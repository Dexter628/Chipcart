<?php
// 請先下載 SimpleXLSX.php 並放在此檔案相同目錄中
require_once __DIR__ . '/SimpleXLSX.php';
if (!class_exists('SimpleXLSX')) {
    die("SimpleXLSX 類別未正確載入");
} else {
    echo "SimpleXLSX 類別載入成功";
}
// 因為使用 ThinkPHP 的資料庫連線設定，所以引入 Db 類
use think\Db;

// 若要模擬，這裡直接輸出最終轉換結果；正式環境中可改為執行 Db::name('chip_db')->insert($data)

if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['excel_file'])) {
    if ($_FILES['excel_file']['error'] == UPLOAD_ERR_OK) {
        $filePath = $_FILES['excel_file']['tmp_name'];
        
        // 解析 Excel 檔案（支援 xls 與 xlsx）
        if (!$xlsx = SimpleXLSX::parse($filePath)) {
            exit("解析 Excel 文件失敗: " . SimpleXLSX::parseError());
        }
        
        $rows = $xlsx->rows();
        if (count($rows) < 2) {
            exit("Excel 文件中沒有足夠的數據。");
        }
        
        // 第一列作為表頭
        $headers = $rows[0];
        
        /* 
          定義資料庫欄位對應的 Excel 表頭別名（可根據需求調整）  
          這裡只針對您這份 Excel 主要欄位進行映射，缺少的欄位預設為空值。
        */
        $fieldsMap = array(
            'part_no'            => array('Your internal Part id', 'Part No.', 'PartNo', '型号', 'Manufacturer Part Number', 'PART NO'),
            'manufacturer_name'  => array('Manufacturer Name', 'MFG', 'MNF', '厂商', 'BRAND'),
            'part_description'   => array('Part Description', '产品参数'),
            'available_qty'      => array('Quantity (free on Hand)', 'QTY', 'Quantity', '数量'),
            'moq'                => array('Minimum Order Quantity', 'MOQ', '起订量'),
            'order_increment'    => array('Order Increment / Pack Qty', 'Pack Qty', 'Order Increment'),
            'date_code_range'    => array('Date Code Range', 'DC', 'DateCode', '批号'),
            // 價格欄位：可從 "Resale (web price)" 或 "Cost (USD)" 判斷，這裡統一映射為 price
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
            'qty_3_price'        => array('Qty 3 price (USD)', 'Qty 3 price'),
            // 其它欄位可根據需要擴充...
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
        
        // 處理價格欄位，檢查是否帶有貨幣資訊 (例如 "(USD)")
        $detectedCurrency = '';
        foreach ($headers as $header) {
            if (preg_match('/\((.*?)\)/', $header, $matches)) {
                $curr = strtoupper(trim($matches[1]));
                if ($curr == 'USD') {
                    $detectedCurrency = 'USD';
                    break;
                }
            }
        }
        if (!$detectedCurrency) {
            $detectedCurrency = 'USD'; // 預設為 USD
        }
        
        // 模擬將每一行數據轉換為資料庫格式
        $insertedData = array();
        // 從第二列開始（第一列為表頭）
        foreach (array_slice($rows, 1) as $row) {
            $data = array();
            // 根據 mapping 取得欄位數值；若無對應則設為空字串
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
                // 若為人民幣等，可根據需求判斷，此處預設 1 表示含稅
                $data['tax_included'] = 1;
            }
            
            // 自動填入更新時間
            $data['update_time'] = date('Y-m-d H:i:s');
            
            // 實際應用中，您可使用 ThinkPHP 的 Db 類將 $data 插入 chip_db 表，例如：
            // Db::name('chip_db')->insert($data);
            // 這裡僅模擬插入，將結果存入陣列
            $insertedData[] = $data;
        }
        
        // 輸出模擬結果 (顯示前 5 筆)
        echo "<h3>模擬最終要插入 chip_db 表的資料（前 5 筆）：</h3>";
        echo "<pre>" . print_r(array_slice($insertedData, 0, 5), true) . "</pre>";
        
    } else {
        echo "檔案上傳失敗。";
    }
} else {
    // 未上傳檔案時顯示上傳表單
    ?>
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Excel 上傳導入模擬</title>
    </head>
    <body>
        <h2>上傳 Excel 文件 (xls 或 xlsx)</h2>
        <form action="" method="post" enctype="multipart/form-data">
            <input type="file" name="excel_file" accept=".xls,.xlsx" required>
            <br><br>
            <input type="submit" value="上傳並模擬導入數據">
        </form>
    </body>
    </html>
    <?php
}
?>
