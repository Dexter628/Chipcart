<?php
namespace app\controller;
use app\model\PartsMain;
use app\model\PartsExtra;
use think\facade\Request;

class Search {
    public function index() {
        $keyword = Request::param('keyword');
        $result = PartsMain::where('part_no|manufacturer_name', 'like', "%$keyword%")
                         ->field('part_no, manufacturer_name, available_qty, lead_time, price, tax_included, moq, spq, order_increment, qty1_price, qty2_price, qty3_price, pcs')
                         ->select();
        return json($result);
    }
    public function detail() {
    $partNo = Request::param('part_no');
    $detail = \app\model\PartsExtra::where('part_no', $partNo)
                                  ->field('warranty, rohs_compliant, eccn_code, hts_code, warehouse_code, certificate_origin, packing, date_code_range')
                                  ->find();
    return json($detail);
}
}