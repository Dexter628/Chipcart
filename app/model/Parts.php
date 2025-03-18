<?php
namespace app\model;

use think\Model;

class Parts extends Model
{
    // 指定對應的資料表名稱
    protected $table = 'parts';

    // 定義允許批量賦值的欄位，請根據您的資料庫欄位調整
    protected $field = [
        'part_no', 
        'manufacturer_name', 
        'available_qty', 
        'lead_time', 
        'price', 
        'currency', 
        'moq', 
        'spq', 
        'order_increment', 
        'qty_1', 
        'qty_1_price', 
        'qty_2', 
        'qty_2_price', 
        'qty_3', 
        'qty_3_price', 
        'supplier_code', 
        'warranty', 
        'rohs_compliant', 
        'eccn_code', 
        'hts_code', 
        'warehouse_code', 
        'certificate_origin', 
        'packing', 
        'date_code_range', 
        'package', 
        'package_type', 
        'price_validity', 
        'contact', 
        'part_description', 
        'tax_included'
    ];
    
    // 啟用自動寫入時間戳，並指定格式為 datetime
    protected $autoWriteTimestamp = 'datetime';

    // 如果不需要建立時間戳，則關閉 createTime 自動填寫（或根據需求啟用）
    protected $createTime = false;

    // 指定更新時間欄位為 update_time，該欄位將自動填入當前日期時間
    protected $updateTime = 'update_time';
}
