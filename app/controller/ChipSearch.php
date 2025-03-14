<?php
namespace app\controller;

use think\facade\Request;
use think\facade\Db;

class ChipSearch
{
    public function search()
    {
        $keyword = trim(Request::param('keyword', ''));
        
        if (empty($keyword)) {
            return json(['code' => 400, 'msg' => '请输入搜索关键字']);
        }

        try {
            $result = Db::table('parts')
                ->where('part_no|manufacturer_name|part_description', 'like', "%{$keyword}%")
                ->field('
                    id, part_no, manufacturer_name, available_qty, lead_time, price, currency, 
                    tax_included as tax_include, moq, spq, order_increment, qty_1, qty_1_price, 
                    qty_2, qty_2_price, qty_3, qty_3_price, warranty, rohs_compliant, eccn_code, 
                    hts_code, warehouse_code, certificate_origin, packing, date_code_range, 
                    package, package_type, price_validity, contact, part_description
                ')
                ->select();

            return json([
                'code' => 0,
                'data' => $result ?: []
            ]);

        } catch (\Exception $e) {
            return json(['code' => 500, 'msg' => '服务器繁忙']);
        }
    }
}