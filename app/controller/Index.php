<?php
namespace app\controller;

use think\Controller;
use think\Request;
use think\facade\View;

class Upload extends Controller
{
    /**
     * index 方法：回傳上傳介面（upload.html 模板或直接回傳 HTML）
     */
    public function index()
    {
        // 若使用模板，請確認 view 目錄中有 upload.html 模板文件
        // return View::fetch('upload'); 
        // 或直接回傳 HTML 字串：
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
                xhr.open('POST', '/upload/upload', true); // 注意此路由對應 Upload/upload 方法
                xhr.onload = function () {
                    if (xhr.status === 200) {
                        currentChunk++;
                        progressDiv.innerText = "已上傳 " + currentChunk + " / " + totalChunks + " 區塊";
                        if (currentChunk < totalChunks) {
                            uploadNextChunk();
                        } else {
                            progressDiv.innerText = "檔案上傳完成，正在處理 Excel 檔案...";
                            // 上傳完成後回傳的結果即為 Excel 解析的 JSON 結果
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
        return $html;
    }

    /**
     * 以下為 Excel 解析相關函式與分段上傳合併處理
     */

    // 設定檔案上傳目錄（注意：這裡使用 dirname(__DIR__) 回到 app 目錄，再指定 uploads/）
    protected $uploadDir;

    public function __construct(Request $request = null)
    {
        parent::__construct($request);
        $this->uploadDir = dirname(__DIR__) . '/uploads/';
        if (!is_dir($this->uploadDir)) {
            mkdir($this->uploadDir, 0777, true);
        }
    }

    /**
     * 將儲存格參考字母轉為 0-based 欄位索引
     */
    protected function colIndexFromLetter($letters) {
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
     */
    protected function readXLSXWithoutExtensions($filePath) {
        $sharedStringsXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/sharedStrings.xml");
        $sheetXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/worksheets/sheet1.xml");
        if (!$sheetXML) {
            return false;
        }
        $sharedStrings = [];
        if ($sharedStringsXML && preg_match_all('/<t[^>]*>(.*?)<\/t>/s', $sharedStringsXML, $matches)) {
            $sharedStrings = $matches[1];
        }
        $rows = [];
        if (preg_match_all('/<row[^>]*>(.*?)<\/row>/s', $sheetXML, $rowMatches)) {
            foreach ($rowMatches[1] as $rowContent) {
                $rowDataTemp = [];
                if (preg_match_all('/<c\s+[^>]*r="([A-Z]+)\d+"[^>]*>(.*?)<\/c>/s', $rowContent, $cellMatches, PREG_SET_ORDER)) {
                    foreach ($cellMatches as $cellMatch) {
                        $colLetters = $cellMatch[1];
                        $cellXmlContent = $cellMatch[2];
                        $colIndex = $this->colIndexFromLetter($colLetters);
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
     * 正規化表頭：去除前置 "*" 與空白
     */
    protected function normalizeHeader($str) {
        return trim(preg_replace('/^[\*\s]+/', '', $str));
    }

    /**
     * Excel 後續處理：根據欄位對應規則匹配表頭、處理價格、貨幣與含税資訊，返回最終資料陣列
     */
    protected function processExcelRows($rows) {
        $headers = $rows[0];
        $normHeaders = [];
        foreach ($headers as $h) {
            $normHeaders[] = $this->normalizeHeader($h);
        }
        // 定義欄位對應規則（包含簡繁體版本）
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
            // tax_included 不直接匹配，後續特殊判斷
        ];
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
        $taxIncludedCol = null;
        foreach ($normHeaders as $index => $normHeader) {
            if (stripos($normHeader, 'Tax Included') !== false ||
                stripos($normHeader, '含税') !== false ||
                stripos($normHeader, '含稅') !== false) {
                $taxIncludedCol = $index;
                break;
            }
        }
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
        if ($detectedCurrency === "USD") {
            $computedTaxIncluded = 0;
        } elseif (in_array($detectedCurrency, ["RMB", "CNY"])) {
            $computedTaxIncluded = 0;
        } else {
            $computedTaxIncluded = 0;
        }
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
     * 處理上傳檔案：合併分段後解析 Excel 並輸出結果
     */
    protected function processUploadedExcel($mergedFilePath) {
        $rows = $this->readXLSXWithoutExtensions($mergedFilePath);
        if (!$rows || count($rows) < 2) {
            return "解析 Excel 文件失敗或資料不足。";
        }
        $finalData = $this->processExcelRows($rows);
        return $finalData;
    }

    /**
     * upload 方法：分段上傳處理，接收每個區塊並合併成完整檔案，
     * 最後解析 Excel 並回傳結果 (JSON 格式)
     */
    public function upload()
    {
        // 取得分段上傳參數
        $request = request();
        $fileId = $request->post('fileId');
        $chunkIndex = intval($request->post('chunkIndex'));
        $totalChunks = intval($request->post('totalChunks'));
        $fileName = $request->post('fileName');

        $targetFile = $this->uploadDir . $fileId . '_' . $fileName;
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
            $result = $this->processUploadedExcel($targetFile);
            // 可選：上傳完成後刪除臨時檔案
            // unlink($targetFile);
            return json($result);
        } else {
            return "區塊 $chunkIndex 已上傳";
        }
    }
}
