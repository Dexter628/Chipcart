<?php
namespace app\controller;

use app\model\PartsMain;
use think\facade\Request;

class Search
{
    /**
     * 關鍵字查詢主要芯片信息（支援分頁）
     */
    public function index()
    {
        $keyword = Request::param('keyword', '', 'trim');
        if (empty($keyword)) {
            return json(['error' => '請輸入查詢關鍵字'], 400);
        }

        // 處理關鍵字
        $cleanedKeyword = preg_replace('/[　\s]+/u', ' ', $keyword); // 將全形/半形空格轉為單一半形空格
        $terms = array_filter(explode(' ', $cleanedKeyword));       // 分詞
        $terms = array_unique($terms);                              // 去除重複詞

        // 分頁參數
        $page = max(1, (int)Request::param('page', 1));
        $pageSize = max(1, (int)Request::param('page_size', 50));
        $offset = ($page - 1) * $pageSize;

        // 查詢條件封裝
        $condition = function ($query) use ($terms) {
            foreach ($terms as $term) {
                $query->whereOr(function ($q) use ($term) {
                    $q->where('part_no', 'like', "%$term%")
                      ->whereOr('manufacturer_name', 'like', "%$term%")
                      ->whereOr('contact', 'like', "%$term%");
                });
            }
        };

        // 總筆數查詢
        $total = PartsMain::where($condition)->count();

        // 分頁資料查詢
        $result = PartsMain::where($condition)
            ->field('
                id, part_no, manufacturer_name, available_qty, lead_time, price, currency, 
                tax_included as tax_include, moq, spq, order_increment, qty_1, qty_1_price, 
                qty_2, qty_2_price, qty_3, qty_3_price, warranty, rohs_compliant, eccn_code, 
                hts_code, warehouse_code, certificate_origin, packing, date_code_range, 
                package, package_type, price_validity, contact, part_description, country
            ')
            ->limit($offset, $pageSize)
            ->select();

        if ($result->isEmpty()) {
            return json([
                'data' => [],
                'total' => $total,
                'page' => $page,
                'page_size' => $pageSize,
                'message' => '未找到符合條件的芯片'
            ], 200);
        }

        return json([
            'data' => $result,
            'total' => $total,
            'page' => $page,
            'page_size' => $pageSize
        ]);
    }
}
