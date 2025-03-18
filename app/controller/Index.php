<?php
namespace app\controller;

use think\facade\Request;
use think\Response;

class Index
{
    /**
     * 顯示首頁（返回 index.html 的內容）
     */
    public function index(): Response
    {
        // 讀取 index.html 文件內容
        $htmlContent = file_get_contents(dirname(dirname(__FILE__)) . '/view/index.html');
        
        // 返回 think\response\Html 對象
        return Response::create($htmlContent, 'html');
    }

    /**
     * 處理分段上傳
     */
    public function upload(): Response
    {
        // 處理分段上傳邏輯
        $result = $this->handleUpload();
        
        // 返回 JSON 響應
        return json($result);
    }

    /**
     * 處理分段上傳的具體邏輯
     */
    private function handleUpload(): array
    {
        // 設定檔案上傳目錄（請確保該目錄存在且具有寫入權限）
        $uploadDir = dirname(dirname(__FILE__)) . '/uploads/';
        if (!is_dir($uploadDir)) {
            mkdir($uploadDir, 0777, true);
        }

        // 取得分段上傳參數
        $fileId = Request::post('fileId');
        $chunkIndex = Request::post('chunkIndex', 0);
        $totalChunks = Request::post('totalChunks', 0);
        $fileName = Request::post('fileName', '');

        // 目標檔案名稱（加入 fileId 確保唯一性）
        $targetFile = $uploadDir . $fileId . '_' . $fileName;

        // 開啟目標檔案，若存在則附加寫入，不存在則創建
        $out = fopen($targetFile, "ab");
        if (!$out) {
            return ['status' => 'error', 'message' => '無法開啟目標檔案'];
        }
        $in = fopen($_FILES['fileChunk']['tmp_name'], "rb");
        if ($in) {
            while ($buff = fread($in, 4096)) {
                fwrite($out, $buff);
            }
            fclose($in);
        }
        fclose($out);

        // 如果這是最後一個區塊，則進行 Excel 處理
        if ($chunkIndex + 1 == $totalChunks) {
            // 解析 Excel 並回傳結果
            $result = $this->processUploadedExcel($targetFile);
            // (選擇性) 合併後可刪除臨時檔案
            // unlink($targetFile);
            return ['status' => 'success', 'data' => $result];
        } else {
            // 回傳當前區塊上傳成功訊息
            return ['status' => 'success', 'message' => "區塊 $chunkIndex 已上傳"];
        }
    }

    /**
     * 解析上傳的 Excel 文件
     */
    private function processUploadedExcel(string $filePath): array
    {
        $rows = $this->readXLSXWithoutExtensions($filePath);
        if (!$rows || count($rows) < 2) {
            return ['status' => 'error', 'message' => '解析 Excel 文件失敗或資料不足。'];
        }
        $finalData = $this->processExcelRows($rows);
        return $finalData;
    }

    /**
     * 解析 XLSX 文件
     */
    private function readXLSXWithoutExtensions(string $filePath): array
    {
        // 取出 sharedStrings.xml 與 sheet1.xml
        $sharedStringsXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/sharedStrings.xml");
        $sheetXML = shell_exec("unzip -p " . escapeshellarg($filePath) . " xl/worksheets/sheet1.xml");

        if (!$sheetXML) {
            return [];
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
     * 將儲存格參考字母轉為 0-based 欄位索引
     */
    private function colIndexFromLetter(string $letters): int
    {
        $letters = strtoupper($letters);
        $result = 0;
        $len = strlen($letters);
        for ($i = 0; $i < $len; $i++) {
            $result = $result * 26 + (ord($letters[$i]) - ord('A') + 1);
        }
        return $result - 1;
    }

    /**
     * 將表頭正規化：去除前置的 "*" 與空白字符
     */
    private function normalizeHeader(string $str): string
    {
        return trim(preg_replace('/^[\*\s]+/', '', $str));
    }

    /**
     * 處理 Excel 資料
     */
    private function processExcelRows(array $rows): array
    {
        // 取出第一列表頭並正規化
        $headers = $rows[0];
        $normHeaders = [];
        foreach ($headers as $h) {
            $normHeaders[] = $this->normalizeHeader($h);
        }

        // 定義欄位對應規則
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
        ];

        // 建立表頭與欄位對應
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

        // 檢查是否存在含税/含稅/Tax Included 的欄位
        $taxIncludedCol = null;
        foreach ($normHeaders as $index => $normHeader) {
            if (stripos($normHeader, 'Tax Included') !== false ||
                stripos($normHeader, '含税') !== false ||
                stripos($normHeader, '含稅') !== false) {
                $taxIncludedCol = $index;
                break;
            }
        }

        // 判斷貨幣
        $detectedCurrency = "";
        if (isset($colMapping['currency']) && isset($normHeaders[$colMapping['currency']])) {
            $detectedCurrency = strtoupper(trim($normHeaders[$colMapping['currency']]));
        }
        if (!$detectedCurrency && isset($colMapping['price'])) {
            $priceHeader = $normHeaders[$colMapping['price']];
            if (preg_match('/[\/$]\s*([A-Za-z]+)\s*[$\/]?/', $priceHeader, $match)) {
                $detectedCurrency = strtoupper(trim($match[1]));
            }
        }
        if (!$detectedCurrency) {
            $detectedCurrency = "USD";
        }

        // 根據貨幣決定 tax_included 預設值
        if ($detectedCurrency === "USD") {
            $computedTaxIncluded = 0;
        } elseif (in_array($detectedCurrency, ["RMB", "CNY"])) {
            $computedTaxIncluded = 0;
        } else {
            $computedTaxIncluded = 0;
        }

        // 將資料轉換成最終格式
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
}