<?php
namespace app\index\controller;

use think\Loader;
use think\Db;
use think\Controller;
use PHPMailer\PHPMailer\PHPMailer;

class Index extends Controller
{
  
	public function excel_export()
	{

		Loader::import('PHPExcel.PHPExcel');
		Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 文件里面的PHPExcel_IOFactory这个类
		$excel=new \PHPExcel();
		$clalist=db('iclass')->select();
		foreach($clalist  as $key=> $cla){
			$excel->createSheet();
			$excel->setActiveSheetIndex($key);
			$sheet=$excel->getActiveSheet();
			$sheet->setTitle($cla['classname']);
			$userlist=db('backhome')->where('iclass='.$cla['id'])->select();
			$column=Db::query("SHOW FULL COLUMNS from wx_backhome");
			$letarr=array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');

			foreach($column as $key=>$col)
			{
				$Comment=$col['Comment']?$col['Comment']:$col['Field'];
				$sheet->setCellValue($letarr[$key].'1',$Comment);
			}
			$i=2;
			foreach($userlist as  $ulist)
			{
				$j=0;
				foreach($ulist as $key=> $v){
					$sheet->setCellValue($letarr[$j].$i,$v);
					$j++;
				}
				$i++;
			}
		}


		// exit;
		$PHPWriter = \PHPExcel_IOFactory::createWriter($excel, "Excel2007"); //创建生成的格式
		header('Content-Disposition: attachment;filename="周末回家申请'.'.xlsx"'); //下载下来的表格名
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		$PHPWriter->save("php://output"); //表示在$path路径下面
	}


	public function index()
	{
		return $this->fetch();
	}
	public function reg()
	{
		$email=input('post.email');
		$username=input('post.username');
		$title="你好,".$username.'欢迎注册相亲网';
		$body="你好，".$username.',相亲网欢迎你的加入，以下是激活链接：http://localhost/tp5';
		sendmail($email,$title,$body);
	}
	public function excel()
	{
		$path = dirname(__FILE__); //找到当前脚本所在路径
		Loader::import('PHPExcel.PHPExcel'); //手动引入PHPExcel.php
		Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 文件里面的PHPExcel_IOFactory这个类
		$PHPExcel = new \PHPExcel(); //实例化
		$iclasslist=db('iclass')->select();
		foreach($iclasslist as $key=> $v){
			$PHPExcel->createSheet();
			$PHPExcel->setactivesheetindex($key);
			$PHPSheet = $PHPExcel->getActiveSheet();
			$PHPSheet->setTitle($v['classname']); //给当前活动sheet设置名称
			$PHPSheet->setCellValue("A1", "编号")
					->setCellValue("B1", "姓名")
					->setCellValue("C1", "性别")
					->setCellValue("D1", "身份证号")
					->setCellValue("E1", "宿舍编号")
					->setCellValue("F1", "班级");//表格数据
			$userlist=db('users')->where("iclass=".$v['id'])->select();
			//echo db('users')->getLastSql();
			$i=2;
			foreach($userlist as $t)
			{
				$PHPSheet->setCellValue("A".$i, $t['id'])
						->setCellValue("B".$i, $t['username'])
						->setCellValue("C".$i, $t['sex'])
						->setCellValue("D".$i, $t['idcate'])
						->setCellValue("E".$i, $t['dorm_id'])
						->setCellValue("F".$i, $t['iclass']);
				//表格数据
				$i++;
			}

		}
		// exit;
		$PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, "Excel2007"); //创建生成的格式
		header('Content-Disposition: attachment;filename="学生列表'.time().'.xlsx"'); //下载下来的表格名
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		$PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
	}
}
