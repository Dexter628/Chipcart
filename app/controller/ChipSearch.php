<?php
namespace app\controller;

use think\facade\Request;
use think\facade\Db;

class ChipSearch
{
    public function search()
    {
        $keyword = strtolower(trim(Request::param('keyword', '')));
        
        if (empty($keyword)) {
            return json(['code' => 400, 'msg' => '请输入搜索关键字']);
        }

        $cleanedKeyword = preg_replace('/[　\s]+/u', ' ', $keyword); // 將全形/半形空格統一為半形空格
        $terms = array_filter(explode(' ', $cleanedKeyword)); // 分詞
        $terms = array_unique($terms); // 去重複

        try {
            $result = Db::table('parts')
            ->where(function ($query) use ($terms) {
                foreach ($terms as $term) {
                    $query->whereOr(function ($q) use ($term) {
                        $q->where('manufacturer_name', 'like', "%{$term}%")
                          ->whereOr('part_no', 'like', "%{$term}%")
                          ->whereOr('contact', 'like', "%{$term}%")
                          ->whereOr('part_description', 'like', "%{$term}%");
                    });
                }
            })
            ->field("id, part_no, manufacturer_name, available_qty, lead_time, price, currency, tax_included as tax_include, moq, spq, order_increment, qty_1, qty_1_price, qty_2, qty_2_price, qty_3, qty_3_price, warranty, rohs_compliant, eccn_code, hts_code, warehouse_code, certificate_origin, packing, date_code_range, package, package_type, price_validity, contact, part_description")
            ->select();


            return json([
                'code' => 0,
                'data' => $result ?: [],
                'msg' => $result ? '查詢成功' : '未找到符合條件的芯片'
            ]);

        } catch (\Exception $e) {
            return json(['code' => 500, 'msg' => '服务器繁忙']);
        }
    }
}