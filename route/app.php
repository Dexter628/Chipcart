<?php
// +----------------------------------------------------------------------
// | ThinkPHP [ WE CAN DO IT JUST THINK ]
// +----------------------------------------------------------------------
// | Copyright (c) 2006~2018 http://thinkphp.cn All rights reserved.
// +----------------------------------------------------------------------
// | Licensed ( http://www.apache.org/licenses/LICENSE-2.0 )
// +----------------------------------------------------------------------
// | Author: liu21st <liu21st@gmail.com>
// +----------------------------------------------------------------------
use think\facade\Route;

// 搜索接口
Route::get('/search', 'Search/index');   

// 详情接口
Route::get('/detail', 'Search/detail');    

// 获取当前计数
Route::get('/api/count', 'index/getCount');

// 更新计数，自增或者清零
Route::post('/api/count', 'index/updateCount');

Route::get('parts/search', 'parts/search');

Route::get('parts/search', 'ChipSearch/search');