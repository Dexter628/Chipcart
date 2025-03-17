<?php
namespace app\controller;
use think\Controller;
use PhpOffice\PhpSpreadsheet\IOFactory;

class Import extends Controller
{
    public function index()
    {
        // 获取上传文件
        $file = $this->request->file('file');
        
        // 验证文件
        if (!$file->checkExt('xlsx')) {
            return json(['code' => 500, 'msg' => '仅支持XLSX格式']);
        }

        // 保存文件
        $savePath = 'public/uploads/' . $file->hashName();
        $file->move('public/uploads', $file->hashName());

        // 处理Excel
        $spreadsheet = IOFactory::load($savePath);
        $sheet = $spreadsheet->getActiveSheet();
        
        // 数据入库逻辑
        // ...（参考之前的处理代码）

        return json(['code' => 200, 'msg' => '导入成功']);
    }
}