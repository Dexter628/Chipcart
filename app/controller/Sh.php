<?php
namespace app\controller;

use app\model\PartsMain;
use think\facade\Request;

class Sh
{
    /**
     * 關鍵字查詢主要芯片信息（支援分頁）
     * 
     * @return \think\Response
     */
    public function index()
    {
        // 1. 获取并验证输入参数
        $keyword = Request::param('keyword', '', 'trim');
        if (empty($keyword)) {
            return json(['error' => 'Please enter search keywords'], 400);
        }

        // 2. 处理关键词（安全过滤和分词）
        $processedTerms = $this->processSearchKeyword($keyword);
        if (empty($processedTerms)) {
            return json(['error' => 'Invalid search keywords'], 400);
        }

        // 3. 处理分页参数
        $pageParams = $this->getPaginationParams();
        
        // 4. 构建查询条件
        $condition = $this->buildSearchCondition($processedTerms);

        // 5. 执行查询
        $queryResult = $this->executeSearchQuery($condition, $pageParams);

        // 6. 返回格式化结果
        return $this->formatSearchResult($queryResult, $pageParams);
    }

    /**
     * 处理搜索关键词
     * 
     * @param string $keyword
     * @return array
     */
    protected function processSearchKeyword(string $keyword): array
    {
        // 编码转换和验证
        $cleanedKeyword = mb_convert_encoding($keyword, 'UTF-8', 'UTF-8');
        
        // 替换各种空格为单个半角空格
        $cleanedKeyword = preg_replace('/[　\s]+/u', ' ', $cleanedKeyword);
        
        // 分词并去重
        $terms = array_filter(explode(' ', $cleanedKeyword));
        $terms = array_unique($terms);
        
        // 过滤危险字符
        return array_map(function($term) {
            return addslashes(trim($term));
        }, $terms);
    }

    /**
     * 获取分页参数
     * 
     * @return array
     */
    protected function getPaginationParams(): array
    {
        $page = max(1, (int)Request::param('page', 1));
        $pageSize = min(max(1, (int)Request::param('page_size', 50)), 200); // 限制最大200条
        $offset = ($page - 1) * $pageSize;
        
        return [
            'page' => $page,
            'page_size' => $pageSize,
            'offset' => $offset
        ];
    }

    /**
     * 构建搜索条件
     * 
     * @param array $terms
     * @return callable
     */
    protected function buildSearchCondition(array $terms): callable
    {
        return function ($query) use ($terms) {
            foreach ($terms as $term) {
                $query->whereOr(function ($q) use ($term) {
                    $q->where('part_no', 'like', "%{$term}%")
                      ->whereOr('manufacturer_name', 'like', "%{$term}%")
                      ->whereOr('contact', 'like', "%{$term}%")
                      ->whereOr('part_description', 'like', "%{$term}%");
                });
            }
        };
    }

    /**
     * 执行搜索查询
     * 
     * @param callable $condition
     * @param array $pageParams
     * @return array
     */
    protected function executeSearchQuery(callable $condition, array $pageParams): array
    {
        // 查询总数量
        $total = PartsMain::where($condition)->count();
        
        // 查询分页数据
        $result = PartsMain::where($condition)
            ->field([
                'id', 'part_no', 'manufacturer_name', 'available_qty', 
                'lead_time', 'price', 'currency', 'tax_included as tax_include',
                'moq', 'spq', 'order_increment', 
                'qty_1', 'qty_1_price', 'qty_2', 'qty_2_price', 
                'qty_3', 'qty_3_price', 'warranty', 'rohs_compliant', 
                'eccn_code', 'hts_code', 'warehouse_code', 
                'certificate_origin', 'packing', 'date_code_range', 
                'package', 'package_type', 'price_validity', 
                'contact', 'part_description', 'country',
                'supplier_code', 'update_time'
            ])
            ->limit($pageParams['offset'], $pageParams['page_size'])
            ->select();
            
        return [
            'data' => $result,
            'total' => $total
        ];
    }

    /**
     * 格式化搜索结果
     * 
     * @param array $queryResult
     * @param array $pageParams
     * @return \think\Response
     */
    protected function formatSearchResult(array $queryResult, array $pageParams)
    {
        $response = [
            'data' => $queryResult['data']->isEmpty() ? [] : $queryResult['data'],
            'total' => $queryResult['total'],
            'page' => $pageParams['page'],
            'page_size' => $pageParams['page_size']
        ];

        if ($queryResult['data']->isEmpty()) {
            $response['message'] = 'No matching chips found';
        }

        return json($response);
    }
}