<?php
namespace app\index\controller;

use think\Loader;
use think\Controller;
use think\cache\dirver\redis;


class Index extends Controller
{
    public function index()
    {
        return "<a href='".url('excel')."'>导出</a>";
    }
    public function excel()
    {
        Loader::import('PHPExcel.PHPExcel');
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory');
        $phpexcel = new \PHPExcel();
        $class = db('iclass')->select();
        $column  = db()->query('show full columns from wx_users');
        $letter = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
        $k = 0;
        foreach ($class as $v) {
            $stuinfo = db('users')->where("iclass = {$v['id']}")->select();
            if(!$stuinfo) {
                continue;
            }
            $phpexcel->createSheet();
            $phpexcel->setActiveSheetIndex($k);
            $sheet = $phpexcel->getActiveSheet();
            $sheet->setTitle($v['classname']);
            foreach ($column as $ckey=>$cval) {    //设置第一行
                $comment = $cval['Comment'] ? $cval['Comment'] : $cval['Field'];
                $sheet->setCellValue("{$letter[$ckey]}1",$comment);
            }

            $line = 2;
            foreach ($stuinfo as $skey=>$sval) { //设置下面几行
                $row = 0;
                foreach($sval as $sskey=>$ssval) {
                    $sheet->setCellValue("{$letter[$row]}{$line}",$ssval);
                    $row++;
                }
                $line++;
            }
            $k++;
        }
        $phpwrite = \PHPExcel_IOFactory::createWriter($phpexcel,'Excel2007');
        header('Content-Type: application/octet-stream');
        $filename = urlencode('学生列表');
        header("Content-Disposition:attachment;filename={$filename}.xlsx;charset=utf8");
        $phpwrite->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
    }


/**
     * 导入
     */
    public function excelImport() {
        return $this->fetch();
    }
    
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
                    $data['id'] = $v[0];
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
                        db('users1')->insert($data);
                    } catch(\Exception $e) {
                        return '插入失败';
                    }

                }
            }
            echo "succ";
        } else {
            return $upload->getError();
        }

    }
    //发送邮件
    public function email()
    {
        return $this->fetch();
    }
     public function reg()
    {
        $email=input('post.email');
        $username=input('post.username');
        $title="你好,".$username.'欢迎注册相亲网';
        $body="你好,".$username.',相亲网欢迎你的加入，以下是激活链接：http://localhost/tp5';
        sendmail($email,$title,$body);
    }

    //上传文件 其上传的大小可在php.ini 中修改  查upload 大约791行
    
   public function shangchuan()
    {
        return $this->fetch();
    }
    public function do_shangchuan()
    {
        $data=input('post.');
        $file=request()->file('files');
        $dir= ROOT_PATH."public/Uploads";
        if(is_dir($dir)){
            echo "已存在正在上传文件...";
            $files=$file->move($dir);
        }else{
            mkdir($dir);
            echo "已创建完文件夹，并上传";
            $files=$file->move($dir);
        }
    }


    //分页 
    public function userlist()
    {

//        echo "<pre>";
//        print_r($info);exit;
        $page=input('get.page')?input('get.page'):1;
        $num=db('users')->count();
        $tiao=5;
        $pages=ceil($num/$tiao);
        if($page==$pages)
        {
            $xia=$page;
        }else{
            $xia=$page+1;
        }if($page==1) {
            $shang=$page;
    }else {
            $shang=$page-1;
    }
        $info=db('users')->page($page,$tiao)->select();
        $this->assign('page',$page);
        $this->assign('pages',$pages);
        $this->assign('shang',$shang);
        $this->assign('xia',$xia);
        $this->assign('info',$info);
        return $this->fetch('userlist');
    }

    //验证码
     public function yanzhengma()
    {
        return $this->fetch();
    }
    public function do_yanzhengma()
    {
        $data = input('post.');
        $code = $data['code'];
        unset($data['code']);

        if (captcha_check($code)) {
            $info=db('member')->insert($data);
            if ($info) {
                $result = [
                    'msg' => '添加成功', 'status' => 1

                ];
                return json($result);
            } else {
                $result = [
                    'msg' => '添加失败', 'status' => 2
                ];
                return json($result);
            }
        } else {
            $result = [
                'msg' => '验证码错误', 'status' => 3
            ];
            return json($result);
        }
    }
    //缓存秒杀
    public function miaoshas()
    {
        return $this->fetch('miaosha');
    }

    public function do_miaosha()
    {
        $ms=db('goods')->find();
        $this->assign('ms',$ms);
        return $this->fetch();
    }

    public function setnums()//导入缓存库
    {
        $nums=3;//设置商品数量
        $redis=new \Redis();//实例化
        $redis->connect('127.0.0.1','6379');//连接主机
        for($i=0;$i<$nums;$i++)
        {
            $redis->lpush('goods_order:1',1);
        }
    }

    public function do_ms()//处理立即秒杀
{
    $redis=new \Redis();
    $redis->connect('127.0.0.1','6379');
    $count=$redis->lpop('goods_order:1');//商品个数

    if(!$count)
    {
        echo "抢光了！";
    }else{
        $info=$redis->lpush('miaosha_order:1',session('userid'));//用户id和个数
        if($info)
        {
            echo "恭喜你抢到了！";
        }else{
            echo "没抢到！";
        }
    }

   $goods_number=db('goods')->where('goods_id=1')->value('goods_number');//查询商品数量
   if( $goods_number)
   {
       $data['goods_id']=1;
       $data['userid']=session('userid');
       $info=db('order')->insert($data);
       db('goods')->where('goods_id=1')->setDec('goods_number');//自减数量
       echo"恭喜你抢到了";
   }else{
       echo"抢光了！！！";
   }
}




    
    
     
}
