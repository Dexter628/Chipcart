<?php
namespace app\controller;

use app\model\PartsMain;
use think\facade\Request;

class Search {
    /**
     * 關鍵字查詢主要芯片信息（支持分頁）
     */
    public function index() {
        $keyword = Request::param('keyword', '', 'trim');
        $page = max(1, (int)Request::param('page', 1));
        $pageSize = max(1, min(100, (int)Request::param('page_size', 50))); // 限制最大每頁 100 筆

        if (empty($keyword)) {
            return json(['error' => '請輸入查詢關鍵字'], 400);
        }

        // 關鍵字預處理
        $cleanedKeyword = preg_replace('/[　\s]+/u', ' ', $keyword); // 全/半形空格統一
        $terms = array_filter(explode(' ', $cleanedKeyword));
        $terms = array_unique($terms);

        // 查詢構造器
        $query = PartsMain::where(function ($query) use ($terms) {
            foreach ($terms as $term) {
                $query->whereOr(function ($q) use ($term) {
                    $q->where('part_no', 'like', "%$term%")
                      ->whereOr('manufacturer_name', 'like', "%$term%")
                      ->whereOr('contact', 'like', "%$term%");
                });
            }
        });

        // 查總數（用來回傳 total）
        $total = $query->count();

        if ($total === 0) {
            return json(['data' => [], 'total' => 0]);
        }

        // 再次複製查詢（避免 count() 後接 limit 出錯）
        $data = $query
            ->field('
                id, part_no, manufacturer_name, available_qty, lead_time, price, currency, 
                tax_included as tax_include, moq, spq, order_increment, qty_1, qty_1_price, 
                qty_2, qty_2_price, qty_3, qty_3_price, warranty, rohs_compliant, eccn_code, 
                hts_code, warehouse_code, certificate_origin, packing, date_code_range, 
                package, package_type, price_validity, contact, part_description, country
            ')
            ->limit($pageSize)
            ->offset(($page - 1) * $pageSize)
            ->select();

        return json([
            'data' => $data,
            'total' => $total
        ]);
    }
}
