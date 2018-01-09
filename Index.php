<?php
namespace app\index\controller;

use think\Loader;
use think\Controller;
use think\Db;
use think\PHPMailer\PHPMailer\PHPMailer;

class Index extends Controller
{
//    public function index()
//    {
//        return "<a href='".url('excel')."'>导出</a>";
//    }
//    public function excel()
//    {
//        $path = dirname(__FILE__); //找到当前脚本所在路径
//        Loader::import('PHPExcel.PHPExcel'); //手动引入PHPExcel.php
//        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 文件里面的PHPExcel_IOFactory这个类
//        $PHPExcel = new \PHPExcel(); //实例化
//        $iclasslist=db('iclass')->select();
//        foreach($iclasslist as $key=> $v){
//            $PHPExcel->createSheet();
//            $PHPExcel->setactivesheetindex($key);
//            $PHPSheet = $PHPExcel->getActiveSheet();
//            $PHPSheet->setTitle($v['classname']); //给当前活动sheet设置名称
//            $PHPSheet->setCellValue("A1", "编号")
//                     ->setCellValue("B1", "姓名")
//                     ->setCellValue("C1", "性别")
//                     ->setCellValue("D1", "身份证号")
//                     ->setCellValue("E1", "宿舍编号")
//                     ->setCellValue("F1", "班级");//表格数据
//            $userlist=db('users')->where("iclass=".$v['id'])->select();
//            //echo db('users')->getLastSql();
//            $i=2;
//            foreach($userlist as $t)
//            {
//                $PHPSheet->setCellValue("A".$i, $t['id'])
//                         ->setCellValue("B".$i, $t['us ername'])
//                        ->setCellValue("C".$i, $t['sex'])
//                        ->setCellValue("D".$i, $t['idcate'])
//                        ->setCellValue("E".$i, $t['dorm_id'])
//                        ->setCellValue("F".$i, $t['iclass']);
//                        //表格数据
//                $i++;
//            }
//
//        }
//       // exit;
//        $PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, "Excel2007"); //创建生成的格式
//        header('Content-Disposition: attachment;filename="学生列表'.time().'.xlsx"'); //下载下来的表格名
//        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//        $PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
//    }
    public function index()
    {
//        return "<a href='".url('excel')."'>导出</a>";
        return $this->fetch();
    }
    //sql导出
    public function excel()
    {
        $path = dirname(__FILE__); //找到当前脚本所在路径
        Loader::import('PHPExcel.PHPExcel'); //手动引入PHPExcel.php
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 文件里面的PHPExcel_IOFactory这个类
        $PHPExcel = new \PHPExcel(); //实例化
        $iclasslist=db('iclass')->select();
        $letarr=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
        // echo "<pre>";
        // print_r($iclasslist);exit;
        foreach($iclasslist as $key=>$v)
        {
            $lists=Db::query('SHOW FULL COLUMNS from wx_users');
//             echo "<pre>";
//             print_r($lists);exit;
            $PHPExcel->createSheet();
            $PHPExcel->setactivesheetindex($key);
            $PHPSheet = $PHPExcel->getActiveSheet();
            $PHPSheet->setTitle($v['classname']); //给当前活动sheet设置名称
            foreach ($lists as $titles)
            {
                $titles=$titles['Comment']?$titles['Comment']:$titles['Field'];
                $PHPSheet->setCellValue($letarr[$key].'1', $titles);
            }
            $userlist=db('users')->where("iclass=".$v['id'])->select();
            $i=2;
            foreach($userlist as $key=>$use){
                $j=0;
                foreach($use as $u){
                    $PHPSheet->setCellValue($letarr[$j].$i,$u);
                    $j++;
                }
                $i++;
            }
        }
        $PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, "Excel2007"); //创建生成的格式
        header('Content-Disposition: attachment;filename="学生列表'.time().'.xlsx"'); //下载下来的表格名
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件

    }
    public  function excelgo()
    {
        return $this->fetch();
    }
    //sql导入
    public function do_excelImport() {
        $file = request()->file('file');
        $pathinfo = pathinfo($file->getInfo()['name']);
        $extension = $pathinfo['extension'];
        $savename = time().'.'.$extension;
        if($upload = $file->move('./upload',$savename)) {
            $savename = './upload/'.$upload->getSaveName();
            Loader::import('PHPExcel.PHPExcel');
            Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory');
            $objReader = \PHPExcel_IOFactory::createReader('Excel2007');
            $objPHPExcel = $objReader->load($savename,$encode = 'utf8');
            $sheetCount = $objPHPExcel->getSheetCount();
            for($i=0 ; $i<$sheetCount ; $i++) {    //循环每一个sheet
                $sheet = $objPHPExcel->getSheet($i)->toArray();
                unset($sheet[0]);
                foreach ($sheet as $v) {
//                    $data['id'] = $v[0];
                    $data['username'] = $v[1];
                    $data['sex'] = $v[2];
                    $data['idcate'] = $v[3];
                    $data['dorm_id'] = $v[4];
                    $data['iclass'] = $v[5];
                    $data['adress'] = $v[6];
                    $data['nation'] = $v[7];
                    $data['major'] = $v[8];
                    $data['birthday'] = $v[9];
                    $data['photo'] = $v[10];
                    $data['famname'] = $v[11];
                    $data['hujiadress'] = $v[12];
                    $data['stutel'] = $v[13];
                    $data['weixin'] = $v[14];
                    $data['qq'] = $v[15];
                    $data['email'] = $v[16];
                    $data['famtel'] = $v[17];
                    $data['pro'] = $v[18];
                    $data['city'] = $v[19];
                    $data['area'] = $v[20];
                    $data['rili'] = $v[21];
                    $data['bed'] = $v[22];
                    $data['openid'] = $v[23];
                    $data['status'] = $v[24];
                    try {
                        db('users2')->insert($data);
                    } catch(\Exception $e) {
                        echo "<pre>";
                       echo $e;
                        return '插入失败';
                    }
                }
            }
            echo "succ";
        } else {
            return $upload->getError();
        }
    }





    //php meail 发送 common.php
    public function reg()
    {
        $email=input('post.email');
        $username=input('post.username');
//        print_r($email);exit;
        $title="你好,".$username.'欢迎注册相亲网';
        $body="你好，".$username.',相亲网欢迎你的加入，以下是激活链接：http://localhost/ttp';
        sendmail($email,$title,$body);
    }

}
