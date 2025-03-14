<?php
namespace app\controller;

use app\model\PartsMain;
use app\model\PartsExtra;
use think\facade\Request;
use think\facade\Db;

class Search {
    /**
     * 關鍵字查詢主要芯片信息
     */
    public function index() {
        $keyword = Request::param('keyword', '', 'trim');
        if (empty($keyword)) {
            return json(['error' => '請輸入查詢關鍵字'], 400);
        }

        $result = PartsMain::where('part_no', 'like', "%$keyword%")
                           ->whereOr('manufacturer_name', 'like', "%$keyword%")
                           ->field('part_no, manufacturer_name, available_qty, lead_time, price, currency, tax_included, moq, spq, order_increment, qty_1, qty_1_price, qty_2, qty_2_price, qty_3, qty_3_price')
                           ->select();

        if ($result->isEmpty()) {
            return json(['error' => '未找到符合條件的芯片'], 404);
        }

        return json($result);
    }
}
