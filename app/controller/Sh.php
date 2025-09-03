<?php
namespace app\controller;

use app\model\PartsMain;
use think\facade\Request;

class Sh
{
    /**
     * ����r�d�ߥD�n����H���]�䴩�����^
     */
    public function index()
    {
        $keyword = Request::param('keyword', '', 'trim');
        if (empty($keyword)) {
            return json(['error' => '�п�J�d������r'], 400);
        }

        // �B�z����r
        $cleanedKeyword = preg_replace('/[�@\s]+/u', ' ', $keyword); // �N����/�b�ΪŮ��ର��@�b�ΪŮ�
        $terms = array_filter(explode(' ', $cleanedKeyword));       // ����
        $terms = array_unique($terms);                              // �h�����Ƶ�

        // �����Ѽ�
        $page = max(1, (int)Request::param('page', 1));
        $pageSize = max(1, (int)Request::param('page_size', 50));
        $offset = ($page - 1) * $pageSize;

        // �d�߱���ʸ�
        $condition = function ($query) use ($terms) {
            foreach ($terms as $term) {
                $query->whereOr(function ($q) use ($term) {
                    $q->where('part_no', 'like', "%$term%")
                      ->whereOr('manufacturer_name', 'like', "%$term%")
                      ->whereOr('contact', 'like', "%$term%");
                });
            }
        };

        // �`���Ƭd��
        $total = PartsMain::where($condition)->count();

        // ������Ƭd��
        $result = PartsMain::where($condition)
            ->field('
                id, part_no, manufacturer_name, available_qty, lead_time, price, currency, 
                tax_included as tax_include, moq, spq, order_increment, qty_1, qty_1_price, 
                qty_2, qty_2_price, qty_3, qty_3_price, warranty, rohs_compliant, eccn_code, 
                hts_code, warehouse_code, certificate_origin, packing, date_code_range, 
                package, package_type, price_validity, contact, part_description, country,supplier_code,update_time
            ')
            ->limit($offset, $pageSize)
            ->select();

        if ($result->isEmpty()) {
            return json([
                'data' => [],
                'total' => $total,
                'page' => $page,
                'page_size' => $pageSize,
                'message' => '�����ŦX���󪺪��'
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
