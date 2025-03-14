<?php
namespace app\controller;
use think\Controller;
use think\Db;

class Parts extends Controller {
    public function search() {
        $part_no = input('get.part_no');

        if (!$part_no) {
            return json(['error' => '缺少 part_no 參數']);
        }

        $main = Db::name('parts_main')->where('part_no', $part_no)->find();
        if (!$main) {
            return json(['error' => '未找到此芯片']);
        }

        $extra = Db::name('parts_extra')->where('part_no', $part_no)->find();
        return json(['main' => $main, 'extra' => $extra]);
    }
}
