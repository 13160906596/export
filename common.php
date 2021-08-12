<?php
// +----------------------------------------------------------------------
// | ThinkPHP [ WE CAN DO IT JUST THINK ]
// +----------------------------------------------------------------------
// | Copyright (c) 2006-2016 http://thinkphp.cn All rights reserved.
// +----------------------------------------------------------------------
// | Licensed ( http://www.apache.org/licenses/LICENSE-2.0 )
// +----------------------------------------------------------------------
// | Author: 流年 <liu21st@gmail.com>
// +----------------------------------------------------------------------

// 应用公共文件
use app\common\helper\Session;
use app\common\helper\Getui;
use think\Db;
use \PhpOffice\PhpSpreadsheet\Spreadsheet;
use \PhpOffice\PhpSpreadsheet\IOFactory;
use \PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use app\common\helper\File as FileCtrl;
//返回格式化数据
function returnMessage($code, $message, $data = [])
{
    $message = urlencode($message);
    header("X-Info: {$message}");
    return \think\Response::create($data, 'json', $code, ["HTTP/1.1 {$code} {$message}" => null]);
}

/**
 * Excel导出数据
 * @param $fields
 * @param $title
 * @param $data
 * @return string
 */
function exportExcel($fields, $data, $title = '能源管理系统导出文件')
{
    vendor('PHPExcel.PHPExcel');
    $phpExcel  = new PHPExcel();
    $phpSheet  = $phpExcel->getActiveSheet();
    $phpWriter = PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');
    $cellName  = array(
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
        'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA',
        'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN',
        'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'
    );

    if (!empty($data)) {
        $rowNum    = count($data); //行数
        $columnNum = count($fields); //列数
        $values    = array_values($fields);
        $keys      = array_keys($fields);
        $phpSheet->setCellValue('A1', $title);

        for ($i = 0; $i < $rowNum; $i++) {
            for ($j = 0; $j < $columnNum; $j++) {
                if ($i == 0) {
                    $phpSheet->setCellValue($cellName[$j] . ($i + 2), $values[$j]);
                } else {
                    $phpSheet->setCellValue($cellName[$j] . ($i + 2), $data[$i][$keys[$j]]);
                }
            }
        }
    }

    //标题处理
    if (!empty($data)) {
        $phpSheet->mergeCells('A1:' . $cellName[$columnNum - 1] . '1'); //合并单元格
    }
    $phpSheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $phpSheet->getStyle('A1')->getFont()->setBold(true);

    //用户下载
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="' . $title . '.xlsx"');
    header('Cache-Control: max-age=0');
    // 用户下载excel
    $phpWriter->save('php://output');

    //保存到服务器
    //    $filename = md5($fileData['user_id']) . '.xls';
    //    $phpWriter->save($fileDir . $filename);
    //    return '/uploads/xls/' . $filename;
}

function createNodesConfig()
{
    $file        = __DIR__ . '/extra/Nodes.php';
    $node        = new \app\admin\model\Node();
    $list        = $node->where('status', 1)->where('project', 3)->select();
    $public      = ['index' => ['save' => '登陆系统', 'delete' => '退出系统', 'resetPassword' => '重置密码']];
    $controllers = [];
    $actions     = [];
    foreach ($list as $value) {
        if ($value['level'] == 2) { //controller
            $controllers[] = $value;
        } else { //action
            $actions[] = $value;
        }
    }
    $logList = [];
    foreach ($controllers as $key => $value) {
        foreach ($actions as $subValue) {
            if ($value['id'] == $subValue['pid']) {
                $logList[strtolower($value['name'])][$subValue['name']] = $subValue['title'];
            }
        }
    }
    $logList = array_merge($logList, $public);

    if (file_exists($file)) {
        unlink($file);
    }
    $userLogConfig = fopen($file, "w");
    fwrite($userLogConfig, '<?php' . "\n");
    fwrite($userLogConfig, 'return' . "\n");
    fwrite($userLogConfig, var_export($logList, true));
    fwrite($userLogConfig, ';' . "\n");
    fclose($userLogConfig);
}

function getCode($length)
{
    $number = range(0, 9);

    $code = '';
    for ($i = 0; $i < $length; $i++) {
        shuffle($number);
        $code .= $number[0];
    }
    return $code;
}
function stmp_mail($sendto_email, $subject = null, $body = null, $sendto_name = null, $account = '15622145620@163.com')
{
    // vendor("phpmailer.phpmailer.PHPMailerAutoload"); //导入函数包的类class.phpmailer.php
    $mail = new \PHPMailer\PHPMailer\PHPMailer(); //新建一个邮件发送类对象
    $mail->IsSMTP(); // send via SMTP
    $mail->Port = 25; //发送端口
    $mail->Host = "smtp.163.com"; // SMTP 邮件服务器地址，这里需要替换为发送邮件的邮箱所在的邮件服务器地址
    //  $mail->Host     = "ssl://smtp.exmail.qq.com:465"; // SMTP 邮件服务器地址，这里需要替换为发送邮件的邮箱所在的邮件服务器地址
    $mail->SMTPAuth = true; // turn on SMTP authentication 邮件服务器验证开
    $mail->Username = $account; // SMTP服务器上此邮箱的用户名，有的只需要@前面的部分，有的需要全名。请替换为正确的邮箱用户名
    $mail->Password = "aa12345678"; // SMTP服务器上该邮箱的密码，请替换为正确的密码
    $mail->From     = $account; // SMTP服务器上发送此邮件的邮箱，请替换为正确的邮箱，$mail->Username 的值是对应的。
    $mail->FromName = "盛杰软件"; // 真实发件人的姓名等信息，这里根据需要填写
    $mail->CharSet  = "UTF-8"; // 这里指定字符集！
    $mail->Encoding = "base64";
    if (is_string($sendto_email)) {
        $mail->AddAddress($sendto_email, $sendto_name); // 收件人邮箱和姓名
    } else if (is_array($sendto_email)) {
        foreach ($sendto_email as $one_mail) {
            $mail->AddAddress($one_mail, $sendto_name);
        }
    }
    //$mail->AddReplyTo('sdaping@mail.ustc.edu.cn',"管理员");//这一项根据需要而设
    //$mail->WordWrap = 50; // set word wrap
    //$mail->AddAttachment("/var/tmp/file.tar.gz"); // 附件处理
    //$mail->AddAttachment("/tmp/image.jpg", "new.jpg");
    $mail->IsHTML(true); // send as HTML
    $mail->Subject = "=?utf-8?B?" . base64_encode($subject) . "?="; // 邮件主题
    // 邮件内容
    $mail->Body = '<html><head>
                        <meta http-equiv="Content-Language" content="zh-cn">
                        <meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
                        <body>' . $body . '</body></html>';
    $mail->AltBody = "text/html";
    if (!$mail->Send()) {
        //邮件发送失败
        return false;
    } else {
        //邮件发送成功

        return true;
    }
}

function exportf2($filename, $tableheader, $data, $info = [])
{
    set_time_limit(0);
    ini_set("memory_limit", "1024M");

    header("Content-Type:text/html;charset=utf-8");
    vendor('PHPExcel.PHPExcel');
    error_reporting(E_ALL);
    ini_set('display_errors', true);
    ini_set('display_startup_errors', true);
    //创建对象
    $excel = new \PHPExcel();

    //Excel表格式
    $letter = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ');
    //填充表头信息
    if (!empty($info)) {
        if ($info['id'] == 1) { //房东数据导出
            //处理表头标题
            // $title = "户籍管理 查询日期：" . $info['start_time'] . " 至 " . $info['end_time'];
            $title = "房东管理";
            $title2 = "报表生成时间：" . date('Y-m-d H:i', time());
            $dataCount = count($data) + 3;
            $excel->getActiveSheet()->mergeCells('A1:' . $letter[count($tableheader) - 1] . '1');
            $excel->getActiveSheet()->mergeCells('A2:' . $letter[count($tableheader) - 1] . '2');
            //设置边框
            $style_array = array(
                'borders' => array(
                    'allborders' => array(
                        'style' => \PHPExcel_Style_Border::BORDER_THIN
                    )
                )
            );
            $excel->getActiveSheet()->getStyle('A1:' . $letter[count($tableheader) - 1] . $dataCount)->applyFromArray($style_array);

            $excel->setActiveSheetIndex(0)->setCellValue('A1', $title);
            $excel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
            $excel->getActiveSheet()->getStyle('A1')->getFont()->setSize(24);
            $excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT); //文字居左
            $excel->getActiveSheet()->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
            $excel->setActiveSheetIndex(0)->setCellValue('A2', $title2);
            $excel->getActiveSheet()->getStyle('A2')->getFont()->setBold(false);
            $excel->getActiveSheet()->getStyle('A2')->getFont()->setSize(14);
            $excel->getActiveSheet()->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT); //文字居左
            $excel->getActiveSheet()->getStyle('A2')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
            // 第一行的默认高度
            $excel->getActiveSheet()->getRowDimension('1')->setRowHeight(50);
            $excel->getActiveSheet()->getRowDimension('2')->setRowHeight(30);
            // 设置填充颜色
            $excel->getActiveSheet()->getstyle('A1')->getFill()->setFillType(PHPExcel_style_Fill::FILL_SOLID);
            $excel->getActiveSheet()->getstyle('A1')->getFill()->getStartColor()->setARGB('FFCCFFFF');
            $excel->getActiveSheet()->getstyle('A2')->getFill()->setFillType(PHPExcel_style_Fill::FILL_SOLID);
            $excel->getActiveSheet()->getstyle('A2')->getFill()->getStartColor()->setARGB('FFCCFFFF');

            $startRow = 3; //第几行开始表头

            for ($i = 0; $i < count($tableheader); $i++) {
                $excel->getActiveSheet()->setCellValue("$letter[$i]$startRow", "$tableheader[$i]");
                $excel->getActiveSheet()->getstyle("$letter[$i]$startRow")->getFill()->setFillType(PHPExcel_style_Fill::FILL_SOLID);
                $excel->getActiveSheet()->getstyle("$letter[$i]$startRow")->getFill()->getStartColor()->setARGB('FFFFFFCC');
                $excel->getActiveSheet()->getStyle("$letter[$i]$startRow")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //文字居中
                $excel->getActiveSheet()->getStyle("$letter[$i]$startRow")->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
                $excel->getActiveSheet()->getRowDimension($startRow)->setRowHeight(20);
                // 设置宽width
                // $excel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);//自适应宽
                $excel->getActiveSheet()->getColumnDimension("$letter[$i]")->setWidth(20);
            }
            $excel->getActiveSheet()->getColumnDimension("B")->setWidth(30); //身份证
            $excel->getActiveSheet()->getColumnDimension("F")->setWidth(60);

            // 冻结窗口
            $excel->getActiveSheet()->freezePaneByColumnAndRow(1, 4);

            //填充表格信息
            $i = $startRow + 1; //行
            foreach ($data as $key => $value) {
                $j = 0; //列
                foreach ($value as $d => $k) {
                    $excel->getActiveSheet()->setCellValueExplicit($letter[$j] . $i, $k);
                    $excel->getActiveSheet()->getStyle("$letter[$j]$i")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //文字居中
                    $excel->getActiveSheet()->getStyle("$letter[$j]$i")->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
                    $j++;
                }
                $excel->getActiveSheet()->getRowDimension($i)->setRowHeight(20);
                $i++;
            }
            //创建Excel输入对象
            $write = new \PHPExcel_Writer_Excel2007($excel);
            $filename = urlencode($filename);
            header("Pragma: public");
            header("Expires: 0");
            header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
            header("Content-Type:application/force-download");
            header("Content-Type:application/vnd.ms-execl");
            header("Content-Type:application/octet-stream");
            header("Content-Type:application/download");
            header("Content-Disposition:attachment;filename=" . $filename . ".xlsx");
            header("Content-Transfer-Encoding:binary");
            $write->save('php://output');

            // $session = Session::getSessionId();
            // if (file_exists("./temp/{$session}")) {
            // } else {
            //     mkdir("./temp/{$session}");
            // }
            // $write->save("./temp/{$session}/{$filename}");

            // // FileCtrl::commit(['name' => $filename]);
            // // return $filename;
            //  $write->save('php://output');

        } elseif ($info['id'] == 2) { //房屋管理数据导出
            //处理表头标题
            $title = "房屋管理";
            $title2 = "报表生成时间：" . date('Y-m-d H:i', time());
            $dataCount = count($data) + 3;
            $excel->getActiveSheet()->mergeCells('A1:' . $letter[count($tableheader) - 1] . '1');
            $excel->getActiveSheet()->mergeCells('A2:' . $letter[count($tableheader) - 1] . '2');
            //设置边框
            $style_array = array(
                'borders' => array(
                    'allborders' => array(
                        'style' => \PHPExcel_Style_Border::BORDER_THIN
                    )
                )
            );
            $excel->getActiveSheet()->getStyle('A1:' . $letter[count($tableheader) - 1] . $dataCount)->applyFromArray($style_array);

            $excel->setActiveSheetIndex(0)->setCellValue('A1', $title);
            $excel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
            $excel->getActiveSheet()->getStyle('A1')->getFont()->setSize(24);
            $excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT); //文字居左
            $excel->getActiveSheet()->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
            $excel->setActiveSheetIndex(0)->setCellValue('A2', $title2);
            $excel->getActiveSheet()->getStyle('A2')->getFont()->setBold(false);
            $excel->getActiveSheet()->getStyle('A2')->getFont()->setSize(14);
            $excel->getActiveSheet()->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT); //文字居左
            $excel->getActiveSheet()->getStyle('A2')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
            // 第一行的默认高度
            $excel->getActiveSheet()->getRowDimension('1')->setRowHeight(50);
            $excel->getActiveSheet()->getRowDimension('2')->setRowHeight(30);
            // 设置填充颜色
            $excel->getActiveSheet()->getstyle('A1')->getFill()->setFillType(PHPExcel_style_Fill::FILL_SOLID);
            $excel->getActiveSheet()->getstyle('A1')->getFill()->getStartColor()->setARGB('FFCCFFFF');
            $excel->getActiveSheet()->getstyle('A2')->getFill()->setFillType(PHPExcel_style_Fill::FILL_SOLID);
            $excel->getActiveSheet()->getstyle('A2')->getFill()->getStartColor()->setARGB('FFCCFFFF');

            $startRow = 3; //第几行开始表头

            for ($i = 0; $i < count($tableheader); $i++) {
                $excel->getActiveSheet()->setCellValue("$letter[$i]$startRow", "$tableheader[$i]");
                $excel->getActiveSheet()->getstyle("$letter[$i]$startRow")->getFill()->setFillType(PHPExcel_style_Fill::FILL_SOLID);
                $excel->getActiveSheet()->getstyle("$letter[$i]$startRow")->getFill()->getStartColor()->setARGB('FFFFFFCC');
                $excel->getActiveSheet()->getStyle("$letter[$i]$startRow")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //文字居中
                $excel->getActiveSheet()->getStyle("$letter[$i]$startRow")->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
                $excel->getActiveSheet()->getRowDimension($startRow)->setRowHeight(20);
                // 设置宽width
                // $excel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);//自适应宽
                $excel->getActiveSheet()->getColumnDimension("$letter[$i]")->setWidth(20);
            }
            $excel->getActiveSheet()->getColumnDimension("A")->setWidth(60);
            $excel->getActiveSheet()->getColumnDimension("E")->setWidth(60);

            // 冻结窗口
            $excel->getActiveSheet()->freezePaneByColumnAndRow(1, 4);

            //填充表格信息
            $i = $startRow + 1; //行
            foreach ($data as $key => $value) {
                $j = 0; //列
                foreach ($value as $d => $k) {
                    $excel->getActiveSheet()->setCellValueExplicit($letter[$j] . $i, $k);
                    $excel->getActiveSheet()->getStyle("$letter[$j]$i")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //文字居中
                    $excel->getActiveSheet()->getStyle("$letter[$j]$i")->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
                    $j++;
                }
                $excel->getActiveSheet()->getRowDimension($i)->setRowHeight(20);
                $i++;
            }
            //创建Excel输入对象
            $write = new \PHPExcel_Writer_Excel2007($excel);
            $filename = urlencode($filename);
            header("Pragma: public");
            header("Expires: 0");
            header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
            header("Content-Type:application/force-download");
            header("Content-Type:application/vnd.ms-execl");
            header("Content-Type:application/octet-stream");
            header("Content-Type:application/download");
            header("Content-Disposition:attachment;filename=" . $filename . ".xlsx");
            header("Content-Transfer-Encoding:binary");
            $write->save('php://output');

            // $session = Session::getSessionId();
            // if (file_exists("./temp/{$session}")) {
            // } else {
            //     mkdir("./temp/{$session}");
            // }
            // header("Content-Type: text/xlsx");
            // header("Content-Disposition:filename=" . $filename);
            // $write->save("./temp/{$session}/{$filename}");
            // // FileCtrl::commit(['name' => $filename]);
            // // return $filename;
            //  $write->save('php://output');

        } elseif ($info['id'] == 3) { //租客管理数据导出
            //处理表头标题
            $title = "租客管理";
            $title2 = "报表生成时间：" . date('Y-m-d H:i', time());
            $dataCount = count($data) + 3;
            $excel->getActiveSheet()->mergeCells('A1:' . $letter[count($tableheader) - 1] . '1');
            $excel->getActiveSheet()->mergeCells('A2:' . $letter[count($tableheader) - 1] . '2');
            //设置边框
            $style_array = array(
                'borders' => array(
                    'allborders' => array(
                        'style' => \PHPExcel_Style_Border::BORDER_THIN
                    )
                )
            );
            $excel->getActiveSheet()->getStyle('A1:' . $letter[count($tableheader) - 1] . $dataCount)->applyFromArray($style_array);

            $excel->setActiveSheetIndex(0)->setCellValue('A1', $title);
            $excel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
            $excel->getActiveSheet()->getStyle('A1')->getFont()->setSize(24);
            $excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT); //文字居左
            $excel->getActiveSheet()->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
            $excel->setActiveSheetIndex(0)->setCellValue('A2', $title2);
            $excel->getActiveSheet()->getStyle('A2')->getFont()->setBold(false);
            $excel->getActiveSheet()->getStyle('A2')->getFont()->setSize(14);
            $excel->getActiveSheet()->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT); //文字居左
            $excel->getActiveSheet()->getStyle('A2')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
            // 第一行的默认高度
            $excel->getActiveSheet()->getRowDimension('1')->setRowHeight(50);
            $excel->getActiveSheet()->getRowDimension('2')->setRowHeight(30);
            // 设置填充颜色
            $excel->getActiveSheet()->getstyle('A1')->getFill()->setFillType(PHPExcel_style_Fill::FILL_SOLID);
            $excel->getActiveSheet()->getstyle('A1')->getFill()->getStartColor()->setARGB('FFCCFFFF');
            $excel->getActiveSheet()->getstyle('A2')->getFill()->setFillType(PHPExcel_style_Fill::FILL_SOLID);
            $excel->getActiveSheet()->getstyle('A2')->getFill()->getStartColor()->setARGB('FFCCFFFF');

            $startRow = 3; //第几行开始表头

            for ($i = 0; $i < count($tableheader); $i++) {
                $excel->getActiveSheet()->setCellValue("$letter[$i]$startRow", "$tableheader[$i]");
                $excel->getActiveSheet()->getstyle("$letter[$i]$startRow")->getFill()->setFillType(PHPExcel_style_Fill::FILL_SOLID);
                $excel->getActiveSheet()->getstyle("$letter[$i]$startRow")->getFill()->getStartColor()->setARGB('FFFFFFCC');
                $excel->getActiveSheet()->getStyle("$letter[$i]$startRow")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //文字居中
                $excel->getActiveSheet()->getStyle("$letter[$i]$startRow")->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
                $excel->getActiveSheet()->getRowDimension($startRow)->setRowHeight(20);
                // 设置宽width
                // $excel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);//自适应宽
                $excel->getActiveSheet()->getColumnDimension("$letter[$i]")->setWidth(20);
            }
            $excel->getActiveSheet()->getColumnDimension("D")->setWidth(60);
            $excel->getActiveSheet()->getColumnDimension("H")->setWidth(60);
            $excel->getActiveSheet()->getColumnDimension("I")->setWidth(60);
            $excel->getActiveSheet()->getColumnDimension("K")->setWidth(60);

            // 冻结窗口
            $excel->getActiveSheet()->freezePaneByColumnAndRow(1, 4);

            //填充表格信息
            $i = $startRow + 1; //行
            // print_r($data);
            // exit;
            foreach ($data as $key => $value) {
                $j = 0; //列
                foreach ($value as $d => $k) {
                    if (in_array($d, ['head_photo', 'identity_photo'])) {
                        if (count($value[$d]) > 0) {
                            $image_count = count($value[$d]);
                            for ($m = 0; $m < $image_count; $m++) {
                                $img_path = $value[$d][$m]['url'];
                                $state = get_photo($img_path, $value[$d][$m]['image']);
                                $new_img_path = './caiji/' . $value[$d][$m]['image'];
                                if (file_exists($new_img_path)) {
                                    $imgs = new \PHPExcel_Worksheet_Drawing();
                                    $imgs->setPath($new_img_path); //写入图片路径
                                    $excel->getActiveSheet()->getColumnDimension($letter[$j])->setWidth(60);

                                    $imgs->setHeight(100); //写入图片高度
                                    // $imgs->setWidth(120);//写入图片宽度
                                    $imgs->setOffsetX(2 + 180 * $m); //写入图片在指定格中的X坐标值
                                    $imgs->setOffsetY(2); //写入图片在指定格中的Y坐标值
                                    $imgs->setRotation(0); //设置旋转角度
                                    // $img->getShadow()->setVisible(true);
                                    // $img->getShadow()->setDirection(50);
                                    $imgs->setCoordinates($letter[$j] . $i); //设置图片所在表格位置
                                    $imgs->setWorksheet($excel->getActiveSheet()); //把图片写到当前的表格中
                                    // unlink($new_img_path);
                                }
                            }
                        }
                    } else {
                        $excel->getActiveSheet()->setCellValueExplicit($letter[$j] . $i, $k);
                        $excel->getActiveSheet()->getStyle("$letter[$j]$i")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //文字居中
                        $excel->getActiveSheet()->getStyle("$letter[$j]$i")->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); //垂直居中
                    }
                    $j++;
                }
                $excel->getActiveSheet()->getRowDimension($i)->setRowHeight(80);
                $i++;
            }
            //创建Excel输入对象
            $write = new \PHPExcel_Writer_Excel2007($excel);
            $filename = urlencode($filename);
            header("Pragma: public");
            header("Expires: 0");
            header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
            header("Content-Type:application/force-download");
            header("Content-Type:application/vnd.ms-execl");
            header("Content-Type:application/octet-stream");
            header("Content-Type:application/download");
            header("Content-Disposition:attachment;filename=" . $filename . ".xlsx");
            header("Content-Transfer-Encoding:binary");
            $write->save('php://output');

            // $session = Session::getSessionId();
            // if (file_exists("./temp/{$session}")) {
            // } else {
            //     mkdir("./temp/{$session}");
            // }
            // header("Content-Type: text/xlsx");
            // header("Content-Disposition:filename=" . $filename);
            // $write->save("./temp/{$session}/{$filename}");
            // // FileCtrl::commit(['name' => $filename]);
            // // return $filename;
            //  $write->save('php://output');

        }
    }
    exit();
}

function get_photo($url, $filename = '', $savefile = './caiji/')
{
    $imgArr = array('gif', 'bmp', 'png', 'ico', 'jpg', 'jepg');

    if (!$url) return false;

    if (!$filename) {
        $ext = strtolower(end(explode('.', $url)));
        if (!in_array($ext, $imgArr)) return false;
        $filename = date("dMYHis") . '.' . $ext;
    }

    if (!is_dir($savefile)) mkdir($savefile, 0777);
    if (!is_readable($savefile)) chmod($savefile, 0777);

    $filename = $savefile . $filename;
    if (file_exists($filename)) {
        unlink($filename);
    }
    ob_start();
    readfile($url);
    $img = ob_get_contents();
    ob_end_clean();
    $size = strlen($img);

    $fp2 = @fopen($filename, "a");
    fwrite($fp2, $img);
    fclose($fp2);

    return $filename;
}


function import($excel, $sign = 1, $area_id = 0)
{
    set_time_limit(0);
    vendor('PHPExcel.PHPExcel');
    $PHPexcel = new \PHPExcel();
    $phpSheet = $PHPexcel->getActiveSheet();
    $reader = null;
    if (preg_match('/xls/', $excel->getInfo()['name'])) {
        $reader = \PHPExcel_IOFactory::createReader('Excel5');;
    }
    if (preg_match('/xlsx/', $excel->getInfo()['name'])) {
        $reader = new \PHPExcel_Reader_Excel2007();
    }

    if (!$reader) {
        return tips(400, "尚未支持当前Excel表格式，请使用Excel 2007或以上版本");
    }

    $excel = $reader->load($excel->getRealPath());
    $sheet = $excel->getSheet(0);
    $allColumn = $sheet->getHighestRow();
    $arr = [];
    // Db::startTrans();
    // try {
    if ($sign == 1) {
        for ($row = 2; $row <= $allColumn; $row++) {
            // $identity = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/","",$sheet->getCell("C".$row)->getCalculatedValue());
            // $name = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/","",$sheet->getCell("B".$row)->getCalculatedValue());
            // $phone = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/","",$sheet->getCell("F".$row)->getValue());

            // $address = $sheet->getCell("A".$row)->getCalculatedValue() ?: "";

            // $check = Db::name('renthouse_new_landlord')->where(['status' => 1, 'identity_card_number' => $identity])->value('id');
            // $check = Db::name('renthouse_new_house')->where(['status' => 1, 'address' => $address])->value('id');
            // if ($check) {//避免重复添加
            //     continue;
            // }
            // $landlord_id = Db::name('renthouse_new_landlord')->where(['name' => $name])->value('id');
            // if ($landlord_id > 0) {
            // }else {
            //     continue;
            // }

            // if (empty($identity) || strlen($identity) != 18) {//身份证不规范跳过
            //     continue;
            // }

            //租户导入
            $room = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/", "", $sheet->getCell("D" . $row)->getCalculatedValue());
            $name = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/", "", $sheet->getCell("E" . $row)->getCalculatedValue());

            $identity = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/", "", $sheet->getCell("F" . $row)->getCalculatedValue());
            $gender = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/", "", $sheet->getCell("G" . $row)->getCalculatedValue());
            $mobile = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/", "", $sheet->getCell("M" . $row)->getCalculatedValue());
            if (empty($identity) || strlen($identity) != 18) { //身份证不规范跳过
                continue;
            }
            // $check = Db::name('renthouse_new_tenant')->where(['identity' => $identity])->value('id');
            // if ($check) {//避免重复添加
            //     continue;
            // }
            $huji_address = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/", "", $sheet->getCell("L" . $row)->getCalculatedValue());
            $house_id = 436;
            $landlord_id = 404;
            // $house = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/","",$sheet->getCell("H".$row)->getCalculatedValue());
            // if ($room != '') {
            //     $address = explode($room, $house)[0];
            //     $house = Db::name('renthouse_new_house')->where(['address' => $address])->field('id, landlord_id')->find();
            //     $house_id = $house['id'];
            //     $landlord_id = $house['landlord_id'];
            // } else {
            //     $house = Db::name('renthouse_new_house')->where(['address' => $house])->field('id, landlord_id')->find();
            //     $house_id = $house['id'];
            //     $landlord_id = $house['landlord_id'];
            // }
            // if ($house_id > 0) {

            // }else {
            //     continue;
            // }

            $_t = [
                // 'name' => $name,
                // 'area_id' => 24,//苏溪：98 上元：24
                // 'address' => $sheet->getCell("A".$row)->getCalculatedValue() ?: "",
                // 'mobile' => $phone,
                // 'identity_card_number' => $identity,
                // 'create_time' => time(),
                // 'role' => 0,
                // 'status' => 1,
                // 'home_mobile' => $sheet->getCell("D".$row)->getCalculatedValue() ?: "",

                // 'address' => $address,
                // 'landlord_id' => $landlord_id,
                // 'status' => 1,
                // 'create_time' => time()

                'name' => $name,
                'identity' => $identity,
                'room' => $room,
                'gender' => $gender,
                'mobile' => $mobile,
                'huji_address' => $huji_address,
                // 'address' => $address
                'house_id' => $house_id,
                'landlord_id' => $landlord_id,
                'status' => 1,
                'create_time' => time()
            ];

            if ($identity && strlen($identity) == 18) {
                // 年龄和性别计算
                $y = intval(substr($_t['identity'], 6, 4));
                $_t['age'] = 2021 - $y;

                if (intval(substr($_t['identity'], 14, 3)) % 2 == 0) {
                    $_t['gender'] = 2;
                } else {
                    $_t['gender'] = 1;
                }

                // $_t['year'] = $y;
                // $_t['month'] = intval(substr($_t['identity'], 10, 2));
                // $_t['day'] = intval(substr($_t['identity'], 12, 2));
            }

            $arr[] = $_t;
            // Db::name('renthouse_new_tenant')->insert($_t);
        }
        // Db::name('renthouse_new_tenant')->insertAll($arr);
        print_r($arr);
        exit;
    } elseif ($sign == 2) {
        for ($row = 2; $row <= $allColumn; $row++) {
            // $identity = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/","",$sheet->getCell("C".$row)->getCalculatedValue());
            // $name = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/","",$sheet->getCell("B".$row)->getCalculatedValue());
            // $phone = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/","",$sheet->getCell("F".$row)->getValue());

            // $address = $sheet->getCell("A".$row)->getCalculatedValue() ?: "";

            // $check = Db::name('renthouse_new_landlord')->where(['status' => 1, 'identity_card_number' => $identity])->value('id');
            // $check = Db::name('renthouse_new_house')->where(['status' => 1, 'address' => $address])->value('id');
            // if ($check) {//避免重复添加
            //     continue;
            // }
            // $landlord_id = Db::name('renthouse_new_landlord')->where(['name' => $name])->value('id');
            // if ($landlord_id > 0) {
            // }else {
            //     continue;
            // }

            // if (empty($identity) || strlen($identity) != 18) {//身份证不规范跳过
            //     continue;
            // }

            //租户导入
            $house = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/", "", $sheet->getCell("B" . $row)->getCalculatedValue());

            $house_num = preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/", "", $sheet->getCell("C" . $row)->getCalculatedValue());



            $_t = [
                'address' => $house . $house_num

            ];



            $arr[] = $_t;
            // Db::name('renthouse_new_tenant')->insert($_t);
        }
        // Db::name('renthouse_new_tenant')->insertAll($arr);
        $brr = [];
        $crr = [];
        foreach ($arr as $key => $val) {
            $arr[$key]['house_id'] = Db::name('renthouse_new_house')->where(['address' => $arr[$key]['address']])->value('id');
            if (isset($arr[$key]['house_id'])) {
                $brr[] = $arr[$key];
            } else {
                $crr[] = $arr[$key];
            }
        }

        print_r($brr);
        exit;
    }
    // }
    // catch (\Exception $e) {
    //     $this->error = $e->getMessage();
    //     Db::rollback();

    //     return false;
    // }
    // Db::commit();

    // print_r($arr);
    // exit;
    // if ($res){
    //     return true;
    // }
}

/**
 * push 推送服务
 *
 * @param array $data
 * @return void
 */
function push(array $devices, string $title)
{
    $getui = new Getui();
    $pushData = [
        'title' => $title
    ];
    if (!empty($devices)) {
        $getui->pushMessageToSingleBatch($devices, $pushData);
    }
}

function import2($sign = 0, $area_id = 0, $user_id = 0)
{
    header("content-type:text/html;charset=utf-8");

    //上传excel文件
    $file_name = request()->post('name');

    //获取文件路径
    $filePath = dirname(dirname(__DIR__)) . '/【出租屋管理系统】房屋数据导入模版.xlsx';
    // $session = Session::getSessionId();
    // $filePath = dirname(__DIR__) . '/public/temp/' . $session . '/' . $file_name;
    // if (!file_exists($filePath)) {
    //     $return['status'] = '4';
    //     $return['message'] = '请重新上传';
    //     return $return;
    // }

    $spreadsheet = IOFactory::load($filePath);
    $spreadsheet->setActiveSheetIndex(0);
    $sheetData = $spreadsheet->getActiveSheet()->toArray(false, true, true, true, true);
    $row_num = count($sheetData);
    $now_time = time();
    // print_r($sheetData);
    // exit;
    $import_data = []; //数组形式获取表格数据
    if ($sign == 2) {
        $error_arr = [];
        $no_landlord_arr = [];
        $is_grid = [];
        $excel_mobile_arr = [];
        // 获取已添加的房东和网格员 以手机号码为唯一值 身份证是能重复添加的
        $landlord_list = Db::name('renthouse_new_landlord')->where(['status' => 1])->field('id, mobile, role')->select();
        $mobile_arr = array_column($landlord_list, 'mobile', 'id');
        $role_arr = array_column($landlord_list, 'role', 'id');
        // 获取已添加的网格员
        $grid_list = Db::name('renthouse_new_landlord')->where(['status' => 1, 'role' => 1, 'area_id' => $area_id])->field('id, name')->select();
        $grid_name_arr = array_column($grid_list, 'name', 'id');
        // 需要新增的网格员信息
        $grid_house_add = [];
        // 网格员异常数组
        $grid_err_arr = [];
        // 房东手机号码异常数组
        $landlord_err_arr = [];
    }
    if ($sign == 3) {
        // 缺少姓名或手机号码
        $lack_of_key = [];
        // 获取已添加的房东
        $landlord_list = Db::name('renthouse_new_landlord')->where(['status' => 1, 'role' => 0])->field('id, mobile')->select();
        $landlord_mobile_arr = array_column($landlord_list, 'mobile', 'id');
        // 获取已添加的网格员
        $grid_list = Db::name('renthouse_new_landlord')->where(['status' => 1, 'role' => 1])->field('id, mobile')->select();
        $grid_mobile_arr = array_column($grid_list, 'mobile', 'id');
        // 已添加为房东
        $added_landlord = [];
        // 已添加为网格员
        $added_grid = [];
    }
    for ($i = 5; $i <= $row_num; $i++) {

        if ($sign == 1) { //导入租客
            $name            = $sheetData[$i]['A'];
            $identity        = $sheetData[$i]['B'];
            $mobile          = $sheetData[$i]['C'];
            $address         = trim($sheetData[$i]['D']);
            $landlord_name   = $sheetData[$i]['E'];
            $room            = $sheetData[$i]['F'];
            $live_num        = $sheetData[$i]['G'];
            $check_in_time   = $sheetData[$i]['H'];
            $check_out_time  = $sheetData[$i]['I'];
            $status          = $sheetData[$i]['J'];
            $huji_address    = $sheetData[$i]['K'];
            $nation          = $sheetData[$i]['L'];

            if (!empty($name)) {
                $import_data[$i]['name']           = $name;
                $import_data[$i]['identity']       = $identity;
                $import_data[$i]['mobile']         = $mobile;
                $import_data[$i]['address']        = $address;
                $import_data[$i]['landlord_name']  = $landlord_name;
                $import_data[$i]['room']           = $room;
                $import_data[$i]['live_num']       = $live_num;
                $import_data[$i]['check_in_time']  = strtotime($check_in_time);
                $import_data[$i]['check_out_time'] = strtotime($check_out_time);
                $import_data[$i]['status']         = $status;
                $import_data[$i]['huji_address']   = $huji_address;
                $import_data[$i]['create_time']    = time();
                $import_data[$i]['row_name']       = '第' . $i . '行:' . $name;
                $import_data[$i]['nation']         = $nation;
            }
        }
        if ($sign == 2) {
            $name           = trim($sheetData[$i]['A']);
            $identity       = trim($sheetData[$i]['B']);
            $address        = trim($sheetData[$i]['C']);
            $mobile         = trim($sheetData[$i]['D']);
            $owner_name     = trim($sheetData[$i]['E']);
            $owner_identity = trim($sheetData[$i]['F']);
            // $owner_address  = trim($sheetData[$i]['G']);
            $owner_address  = '';
            $owner_mobile   = trim($sheetData[$i]['H']);
            $grid_members   = trim($sheetData[$i]['I']);
            // $grid_members   = '关日锐,关国盖,关祖铨,何桂锵,关结玲,何福汉,陆乐,罗凤雲,关炜怡,罗燕芬,关桂学';
            $house_address  = trim($sheetData[$i]['J']);
            $row_name       = '第' . $i . '行:' . $name;

            if (!empty($name) || !empty($owner_name) || !empty($house_address)) {
                if (!empty($mobile)) { // 如果有房东
                    if (in_array($mobile, $mobile_arr)) { // 该房东已存在
                        $landlord_id = array_search($mobile, $mobile_arr);
                        $role = $role_arr[$landlord_id];
                        if ($role == 0) {
                            $import_data[$i]['landlord_id'] = $landlord_id;
                        } elseif ($role == 1) {
                            $is_grid[] = $row_name;
                            continue;
                        }
                    } else { // 该房东不存在
                        if (!in_array($mobile, $excel_mobile_arr)) {
                            if (strlen($mobile) == 11) {
                                $new_landlord['name']                 = $name;
                                $new_landlord['area_id']              = $area_id;
                                $new_landlord['identity_card_number'] = $identity;
                                $new_landlord['address']              = $address;
                                $new_landlord['mobile']               = $mobile;
                                $new_landlord['create_time']          = time();
                                $new_landlord['status']               = 1;
                                $import_data[$i]['landlord_id'] = Db::name('renthouse_new_landlord')->insertGetId($new_landlord);
                                $excel_mobile_arr[] = $mobile;
                            } else {
                                $landlord_err_arr[] = $row_name;
                                continue;
                            }
                        } else {
                            $import_data[$i]['landlord_id'] = Db::name('renthouse_new_landlord')->where(['mobile' => $mobile, 'status' => 1])->value('id');
                        }
                    }
                } else {
                    if (!empty($owner_mobile)) { // 如果没有填房东只填了屋主
                        if (in_array($owner_mobile, $mobile_arr)) { // 该房东已存在
                            $landlord_id = array_search($owner_mobile, $mobile_arr);
                            $role = $role_arr[$landlord_id];
                            if ($role == 0) {
                                $import_data[$i]['landlord_id'] = $landlord_id;
                            } elseif ($role == 1) {
                                $is_grid[] = $row_name;
                            }
                        } else { // 该房东不存在
                            if (!in_array($owner_mobile, $excel_mobile_arr)) {
                                if (strlen($mobile) == 11) {
                                    $new_landlord['name']                 = $owner_name;
                                    $new_landlord['area_id']              = $area_id;
                                    $new_landlord['identity_card_number'] = $owner_identity;
                                    $new_landlord['address']              = $owner_address;
                                    $new_landlord['mobile']               = $owner_mobile;
                                    $new_landlord['create_time']          = time();
                                    $new_landlord['status']               = 1;
                                    $import_data[$i]['landlord_id'] = Db::name('renthouse_new_landlord')->insertGetId($new_landlord);
                                    $excel_mobile_arr[] = $owner_mobile;
                                } else {
                                    $landlord_err_arr[] = $row_name;
                                    continue;
                                }
                            } else {
                                $import_data[$i]['landlord_id'] = Db::name('renthouse_new_landlord')->where(['mobile' => $owner_mobile, 'status' => 1])->value('id');
                            }
                        }
                    } else { //房东和屋主都没有填
                        $no_landlord_arr[] = $row_name;
                        continue;
                    }
                }
                $import_data[$i]['owner_name']     = $owner_name;
                $import_data[$i]['owner_identity'] = $owner_identity;
                $import_data[$i]['owner_address']  = $owner_address;
                $import_data[$i]['owner_mobile']   = $owner_mobile;
                $import_data[$i]['house_address']  = $house_address;
                $import_data[$i]['row_name']       = $row_name;
                $import_data[$i]['rental_type']    = 0;
                $import_data[$i]['status']         = 1;
                $import_data[$i]['create_time']    = time();
                // 1 自治 2 承包
                if (empty($name) && !empty($owner_name)) {
                    $import_data[$i]['rental_type'] = 1;
                }
                if (!empty($name) && !empty($owner_name) && $name == $owner_name) {
                    $import_data[$i]['rental_type'] = 1;
                }
                if (!empty($name) && empty($owner_name)) {
                    $import_data[$i]['rental_type'] = 2;
                }
                if (!empty($name) && !empty($owner_name) && $name != $owner_name) {
                    $import_data[$i]['rental_type'] = 2;
                }
                // 网格员
                if (!empty($grid_members)) {
                    if (strpos($grid_members, "，")) {
                        $grid_members = str_replace("，", ",", $grid_members);
                    }
                    $grid_members = explode(",", $grid_members);
                    $grid_ids = [];
                    if (count($grid_members) > 0) {
                        $count1 = count($grid_members);
                        foreach ($grid_members as $key => $val) {
                            if (in_array($grid_members[$key], $grid_name_arr)) {
                                $sub['landlord_id'] = $import_data[$i]['landlord_id'];
                                $sub['house_address'] = $import_data[$i]['house_address'];
                                $sub['grid_id'] = array_search($grid_members[$key], $grid_name_arr);
                                $grid_house_add[] = $sub;
                                $grid_ids[] = $sub['grid_id'];
                            }
                        }
                        $count2 = count($grid_ids);
                        array_unique($grid_ids);
                        $import_data[$i]['grid_members'] = implode(",", $grid_ids);
                        if ($count1 != $count2) {
                            $grid_err_arr[] = $row_name;
                        }
                    }
                }
            }
        }
        if ($sign == 3) {
            $name           = trim($sheetData[$i]['A']);
            $identity       = trim($sheetData[$i]['B']);
            $address        = trim($sheetData[$i]['C']);
            $mobile         = trim($sheetData[$i]['D']);
            $row_name       = '第' . $i . '行:' . $name;

            if (!empty($name)) {
                if (!empty($name) && !empty($mobile) && strlen($mobile) == 11) {
                    if (in_array($mobile, $landlord_mobile_arr)) {
                        $added_landlord[] = $row_name;
                    } else {
                        if (in_array($mobile, $grid_mobile_arr)) {
                            $added_grid[] = $row_name;
                        } else {
                            $import_data[$i]['name']                 = $name;
                            $import_data[$i]['identity_card_number'] = $identity;
                            $import_data[$i]['address']              = $address;
                            $import_data[$i]['mobile']               = $mobile;
                            $import_data[$i]['status']               = 1;
                            $import_data[$i]['create_time']          = time();
                            $import_data[$i]['area_id']              = $area_id;
                        }
                    }
                } else {
                    $lack_of_key[] = $row_name;
                }
            }
        }
    }
    if ($sign == 1) {
        $error_arr = [];
        // 1、获取房屋id
        sort($import_data);
        $houses_address = array_column($import_data, 'address');
        $house_ids = Db::name('renthouse_new_house')
            ->join('renthouse_new_landlord', 'renthouse_new_landlord.id = renthouse_new_house.landlord_id', 'left')
            ->where([
                'area_id' => $area_id,
                'renthouse_new_landlord.status' => 1,
                'renthouse_new_house.status' => 1,
                'renthouse_new_house.address' => ['in', $houses_address]
            ])
            ->field('renthouse_new_house.id, renthouse_new_house.address, landlord_id')
            ->select();
        $name_arr = [];
        $houseRoom = [];
        foreach ($import_data as $key => $val) {
            $import_data[$key]['house_id'] = 0;
            foreach ($house_ids as $key2 => $val2) {
                if ($import_data[$key]['address'] == $house_ids[$key2]['address']) {
                    $import_data[$key]['house_id'] = $house_ids[$key2]['id'];
                    $import_data[$key]['landlord_id'] = $house_ids[$key2]['landlord_id'];
                }
            }
            if ($import_data[$key]['house_id'] == 0) {
                $name_arr[] = $import_data[$key]['row_name'];
                unset($import_data[$key]);
            } else {
                if (isset($houseRoom[$import_data[$key]['house_id']])) {
                    if (!in_array($import_data[$key]['room'], $houseRoom[$import_data[$key]['house_id']])) {
                        $houseRoom[$import_data[$key]['house_id']][] = $import_data[$key]['room'];
                    }
                } else {
                    $houseRoom[$import_data[$key]['house_id']][] = $import_data[$key]['room'];
                }
            }
        }
        if (!empty($name_arr)) {
            $sub_error['reason']  = '房屋列表中获取对应房屋失败，请检查房屋地址';
            $sub_error['count']   = count($name_arr);
            $sub_error['persons'] = implode(",", $name_arr);
            $error_arr[] = $sub_error;
        }
        // 添加新房间
        if (!empty($houseRoom)) {
            foreach ($houseRoom as $key => $val) {
                $get_floor_box = Db::name('renthouse_new_house')->where(['id' => $key])->value('floor_box');
                if (!empty($get_floor_box)) {
                    $floor_box = json_decode($get_floor_box, true);
                    $room_list = [];
                    foreach ($floor_box as $key2 => $val2) {
                        if (!empty($floor_box[$key2]['detail'])) {
                            foreach ($floor_box[$key2]['detail'] as $key3 => $val3) {
                                $room_list[] = $floor_box[$key2]['detail'][$key3]['room'];
                            }
                        }
                    }
                    foreach ($houseRoom[$key] as $key2 => $val2) {
                        if (!in_array($houseRoom[$key][$key2], $room_list)) {
                            $new_room = [];
                            $new_room['room'] = $houseRoom[$key][$key2];
                            $num = substr($new_room['room'], 0, 1);
                            if (is_numeric($num) && $num > 0 && $num < 6) {
                                array_push($floor_box[$num - 1]['detail'], $new_room);
                            } else {
                                array_push($floor_box[0]['detail'], $new_room);
                            }
                        }
                    }
                } else { // 没有楼层管理的情况
                    $floor_box = [
                        [
                            "storey" => "第一层房间设置：",
                            "detail" => []
                        ],
                        [
                            "storey" => "第二层房间设置：",
                            "detail" => []
                        ],
                        [
                            "storey" => "第三层房间设置：",
                            "detail" => []
                        ],
                        [
                            "storey" => "第四层房间设置：",
                            "detail" => []
                        ],
                        [
                            "storey" => "第五层房间设置：",
                            "detail" => []
                        ]
                    ];
                    foreach ($houseRoom[$key] as $key2 => $val2) {
                        $new_room = [];
                        $new_room['room'] = $houseRoom[$key][$key2];
                        $num = substr($new_room['room'], 0, 1);
                        if (is_numeric($num) && $num > 0 && $num < 6) {
                            array_push($floor_box[$num - 1]['detail'], $new_room);
                        } else {
                            array_push($floor_box[0]['detail'], $new_room);
                        }
                    }
                }
                Db::name('renthouse_new_house')->where(['id' => $key])->update(['floor_box' => json_encode($floor_box)]);
            }
        }
        // 2、获取房东id
        // sort($import_data);
        // $landlord_names = array_column($import_data, 'landlord_name');
        // $landlord_ids = Db::name('renthouse_new_landlord')
        //     ->where([
        //         // 'area_id' => $area_id,
        //         'status' => 1,
        //         'name' => ['in', $landlord_names]
        //     ])
        //     ->field('id, name')
        //     ->select();
        // $name_arr = [];
        // foreach ($import_data as $key => $val) {
        //     $import_data[$key]['landlord_id'] = 0;
        //     foreach ($landlord_ids as $key2 => $val2) {
        //         if ($import_data[$key]['landlord_name'] == $landlord_ids[$key2]['name']) {
        //             $import_data[$key]['landlord_id'] = $landlord_ids[$key2]['id'];
        //         }
        //     }
        //     if ($import_data[$key]['landlord_id'] == 0) {
        //         $name_arr[] = $import_data[$key]['row_name'];
        //         unset($import_data[$key]);
        //     }
        // }
        // if (!empty($name_arr)) {
        //     $sub_error['reason'] = '房东列表中获取对应房东失败，请检查房东姓名';
        //     $sub_error['persons'] = implode(",", $name_arr);
        //     $error_arr[] = $sub_error;
        // }

        // 3、检查有无重复数据，个别地区不允许一人多屋
        sort($import_data);
        $czw_village_registration = Db::name('area')->where(['id' => $area_id])->value('czw_village_registration');
        $identity_array = [];
        foreach ($import_data as $key => $val) {
            if (!empty($import_data[$key]['identity'])) {
                $identity_array[] = $import_data[$key]['identity'];
            }
        }
        $result_data = Db::name('renthouse_new_tenant')
            ->join('renthouse_new_landlord', 'renthouse_new_tenant.landlord_id = renthouse_new_landlord.id', 'left')
            ->field('renthouse_new_tenant.id, identity, house_id, room')
            ->where([
                'renthouse_new_tenant.identity' => ['in', $identity_array],
                'renthouse_new_landlord.status' => 1,
                'area_id' => $area_id
            ])
            ->select();

        if (!empty($result_data)) {
            foreach ($result_data as $key => $val) {
                $result_data[$key]['check'] = $result_data[$key]['identity'] . '-' . $result_data[$key]['house_id'] . '-' . $result_data[$key]['room'];
            }
            $result_data_array = array_column($result_data, 'check');
            $identity_data_array = array_column($result_data, 'identity');
            $name_arr = [];
            $name_arr2 = [];
            $new_arr = [];
            foreach ($import_data as $key => $val) {
                $import_data[$key]['check'] = $import_data[$key]['identity'] . '-' . $import_data[$key]['house_id'] . '-' . $import_data[$key]['room'];
                // 去掉重复登记
                if (in_array($import_data[$key]['check'], $result_data_array)) {
                    $name_arr[] = $import_data[$key]['row_name'];
                    unset($import_data[$key]);
                }
            }
            if (!empty($name_arr)) {
                $sub_error['reason']  = '此租客信息已添加，无需重复导入';
                $sub_error['count']   = count($name_arr);
                $sub_error['persons'] = implode(",", $name_arr);
                $error_arr[] = $sub_error;
            }

            if ($czw_village_registration == 1) { //不允许登记多屋
                // 不允许登记多屋
                foreach ($import_data as $key => $val) {
                    if (in_array($import_data[$key]['identity'], $identity_data_array)) {
                        $name_arr2[] = $import_data[$key]['row_name'];
                        unset($import_data[$key]);
                    }
                }
                if (!empty($name_arr2)) {
                    $sub_error['reason'] = '在当前地区已登记其他住址';
                    $sub_error['count']   = count($name_arr2);
                    $sub_error['persons'] = implode(",", $name_arr2);
                    $error_arr[] = $sub_error;
                }
            }
        }

        // 4、最后处理多余数据
        $insert_field = ['name', 'gender', 'identity', 'mobile', 'room', 'live_num', 'age', 'huji_address', 'house_id', 'landlord_id', 'check_in_time', 'check_out_time', 'status', 'create_time', 'nation', 'relationship'];
        foreach ($import_data as $key => $val) {
            // 电话
            if (!empty($import_data[$key]['mobile'])) {
                if ($import_data[$key]['mobile'] == '无') {
                    $import_data[$key]['mobile'] = '';
                }
            }
            // 状态
            if (!empty($import_data[$key]['status'])) {
                if ($import_data[$key]['status'] == '在租') {
                    $import_data[$key]['status'] = 1;
                } elseif ($import_data[$key]['status'] == '退租') {
                    $import_data[$key]['status'] = 0;
                } else {
                    $import_data[$key]['status'] = 1;
                }
            } else {
                $import_data[$key]['status'] = 1;
            }
            // 性别、年龄处理
            if (intval(substr($import_data[$key]['identity'], 14, 3)) % 2 == 0) {
                $import_data[$key]['gender'] = 2;
            } else {
                $import_data[$key]['gender'] = 1;
            }
            $import_data[$key]['year'] = intval(substr($import_data[$key]['identity'], 6, 4));
            $import_data[$key]['month'] = intval(substr($import_data[$key]['identity'], 10, 2));
            $import_data[$key]['day'] = intval(substr($import_data[$key]['identity'], 12, 2));
            $m_count = diffDate(date('Y-m-d', time()), $import_data[$key]['year'] . '-' . $import_data[$key]['month'] . '-' . $import_data[$key]['day']);
            $import_data[$key]['age'] = ceil($m_count / 12);

            // 关系默认设置为本人
            $import_data[$key]['relationship'] = 1;

            // 多余字段处理
            $diff = array_diff(array_keys($import_data[$key]), $insert_field);
            foreach ($diff as $key2 => $val2) {
                unset($import_data[$key][$val2]);
            }
        }
        sort($import_data);

        // print_r($import_data);
        // exit;

        if (!empty($import_data)) {
            FileCtrl::flushTemp(); //清缓存
            // 将数据保存到数据库
            $res = Db::name('renthouse_new_tenant')->insertAll($import_data);
            if ($res) {
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 1;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '1';
                $return['message'] = '导入成功';
                $return['result'] = $error_arr;
                return $return;
            } else {
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 2;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '2';
                $return['message'] = '导入失败';
                $return['result'] = $error_arr;
                return $return;
            }
        }
        // 导入日志
        $log = [];
        $log['user_id']     = $user_id;
        $log['err_data']    = json_encode($error_arr);
        $log['create_time'] = time();
        $log['sign']        = $sign;
        $log['result']      = 3;
        Db::name('renthouse_new_import_log')->insert($log);
        $return['status'] = '3';
        $return['message'] = '无新数据添加';
        $return['result'] = $error_arr;
        return $return;
    }
    if ($sign == 2) {
        sort($import_data);
        $house_addresses = [];
        $repeat_data = [];
        foreach ($import_data as $key => $val) {
            $import_data[$key]['landlord_address'] = $import_data[$key]['landlord_id'] . '-' . $import_data[$key]['house_address'];
            $house_addresses[] = $import_data[$key]['house_address'];
        }
        $get_houses = Db::name('renthouse_new_house')
            ->where([
                'address' => ['in', $house_addresses]
            ])
            ->field('id, landlord_id, address')
            ->select();
        foreach ($get_houses as $key => $val) {
            $get_houses[$key]['landlord_address'] = $get_houses[$key]['landlord_id'] . '-' . $get_houses[$key]['address'];
        }
        $landlord_address_arr = array_column($get_houses, 'landlord_address');
        foreach ($import_data as $key => $val) {
            if (in_array($import_data[$key]['landlord_address'], $landlord_address_arr)) {
                $repeat_data[] = $import_data[$key]['row_name'];
                if (in_array($import_data[$key]['row_name'], $grid_err_arr)) {
                    unset($grid_err_arr[array_search($import_data[$key]['row_name'], $grid_err_arr)]);
                }
                unset($import_data[$key]);
            }
        }
        // 错误提示
        // 1、房东和屋主都没有填
        if (!empty($no_landlord_arr)) {
            $sub_error['reason']  = '房屋管理员和屋主手机号码都为空，无法添加房东及房屋数据';
            $sub_error['count']   = count($no_landlord_arr);
            $sub_error['persons'] = implode(",", $no_landlord_arr);
            $error_arr[] = $sub_error;
        }
        // 2、房东手机号码异常
        if (!empty($landlord_err_arr)) {
            $sub_error['reason']  = '房屋管理员手机号码为空或手机号码长度不等于11，无法添加房东数据';
            $sub_error['count']   = count($landlord_err_arr);
            $sub_error['persons'] = implode(",", $landlord_err_arr);
            $error_arr[] = $sub_error;
        }
        // 3、已是网格员
        if (!empty($is_grid)) {
            $sub_error['reason']  = '此手机号码已添加为网格员，无法设置为房屋管理员，请更换手机号码';
            $sub_error['count']   = count($is_grid);
            $sub_error['persons'] = implode(",", $is_grid);
            $error_arr[] = $sub_error;
        }
        // 4、此房子已添加
        if (!empty($repeat_data)) {
            $sub_error['reason']  = '此房屋信息已添加，已省略添加，如需更改请前往房屋列表';
            $sub_error['count']   = count($repeat_data);
            $sub_error['persons'] = implode(",", $repeat_data);
            $error_arr[] = $sub_error;
        }
        // 5、存在未添加的网格员
        if (!empty($grid_err_arr)) {
            $sub_error['reason']  = '存在未添加的网格员，请先前往网格员列表添加网格员，然后前往房屋详情选择网格员';
            $sub_error['count']   = count($grid_err_arr);
            $sub_error['persons'] = implode(",", $grid_err_arr);
            $error_arr[] = $sub_error;
        }
        // 6、最后处理多余数据
        $insert_field = ['landlord_id', 'address', 'create_time', 'status', 'owner_name', 'owner_mobile', 'owner_identity', 'owner_address', 'rental_type', 'grid_members'];
        foreach ($import_data as $key => $val) {
            $import_data[$key]['address'] = $import_data[$key]['house_address'];
            // 多余字段处理
            $diff = array_diff(array_keys($import_data[$key]), $insert_field);
            foreach ($diff as $key2 => $val2) {
                unset($import_data[$key][$val2]);
            }
        }
        sort($import_data);
        // print_r($import_data);
        // exit;
        if (!empty($import_data)) {
            FileCtrl::flushTemp(); //清缓存
            // 将数据保存到数据库
            $res = Db::name('renthouse_new_house')->insertAll($import_data);
            if ($res) {
                // 处理网格员
                if (count($grid_house_add) > 0) {
                    $get_new_houses = Db::name('renthouse_new_house')
                        ->where([
                            'address' => ['in', $house_addresses]
                        ])
                        ->field('id, grid_members')
                        ->select();
                    foreach ($get_new_houses as $key => $val) {
                        if (!empty($get_new_houses[$key]['grid_members'])) {
                            $grid_members = explode(",", $get_new_houses[$key]['grid_members']);
                            if (count($grid_members) > 0) {
                                Db::name('renthouse_new_grid_house')->where(['house_id' => $get_new_houses[$key]['id']])->delete();
                                $insert_data = [];
                                foreach ($grid_members as $key2 => $val2) {
                                    $sub_data['grid_id'] = $grid_members[$key2];
                                    $sub_data['house_id'] = $get_new_houses[$key]['id'];
                                    $insert_data[] = $sub_data;
                                }
                                Db::name('renthouse_new_grid_house')->insertAll($insert_data);
                            }
                        }
                    }
                }
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 1;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '1';
                $return['message'] = '导入成功';
                $return['result'] = $error_arr;
                return $return;
            } else {
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 2;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '2';
                $return['message'] = '导入失败';
                $return['result'] = $error_arr;
                return $return;
            }
        }
        // 导入日志
        $log = [];
        $log['user_id']     = $user_id;
        $log['err_data']    = json_encode($error_arr);
        $log['create_time'] = time();
        $log['sign']        = $sign;
        $log['result']      = 3;
        Db::name('renthouse_new_import_log')->insert($log);
        $return['status'] = '3';
        $return['message'] = '无新数据添加';
        $return['result'] = $error_arr;
        return $return;
    }
    if ($sign == 3) {
        sort($import_data);
        // 错误提示
        $error_arr = [];
        // 1、房东电话号码或者姓名没有填
        if (!empty($lack_of_key)) {
            $sub_error['reason']  = '房屋管理员手机号码为空或手机号码长度不等于11，无法添加房东数据';
            $sub_error['count']   = count($lack_of_key);
            $sub_error['persons'] = implode(",", $lack_of_key);
            $error_arr[] = $sub_error;
        }
        // 2、已是房东
        if (!empty($added_landlord)) {
            $sub_error['reason']  = '此手机号码已添加为房东，无法设置为房东，请更换手机号码';
            $sub_error['count']   = count($added_landlord);
            $sub_error['persons'] = implode(",", $added_landlord);
            $error_arr[] = $sub_error;
        }
        // 2、已是网格员
        if (!empty($added_grid)) {
            $sub_error['reason']  = '此手机号码已添加为网格员，无法设置为房东，请更换手机号码';
            $sub_error['count']   = count($added_grid);
            $sub_error['persons'] = implode(",", $added_grid);
            $error_arr[] = $sub_error;
        }

        // 5、最后处理多余数据
        $insert_field = ['name', 'address', 'create_time', 'status', 'identity_card_number', 'mobile', 'area_id'];
        foreach ($import_data as $key => $val) {
            // 多余字段处理
            $diff = array_diff(array_keys($import_data[$key]), $insert_field);
            foreach ($diff as $key2 => $val2) {
                unset($import_data[$key][$val2]);
            }
        }
        sort($import_data);
        // print_r($error_arr);
        // exit;
        if (!empty($import_data)) {
            FileCtrl::flushTemp(); //清缓存
            // 将数据保存到数据库
            $res = Db::name('renthouse_new_landlord')->insertAll($import_data);
            if ($res) {
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 1;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '1';
                $return['message'] = '导入成功';
                $return['result'] = $error_arr;
                return $return;
            } else {
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 2;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '2';
                $return['message'] = '导入失败';
                $return['result'] = $error_arr;
                return $return;
            }
        }
        // 导入日志
        $log = [];
        $log['user_id']     = $user_id;
        $log['err_data']    = json_encode($error_arr);
        $log['create_time'] = time();
        $log['sign']        = $sign;
        $log['result']      = 3;
        Db::name('renthouse_new_import_log')->insert($log);
        $return['status'] = '3';
        $return['message'] = '无新数据添加';
        $return['result'] = $error_arr;
        return $return;
    }
}

function diffDate($date1, $date2)
{
    $datestart = date('Y-m-d', strtotime($date1));
    if (strtotime($datestart) > strtotime($date2)) {
        $tmp = $date2;
        $date2 = $datestart;
        $datestart = $tmp;
    }
    list($Y1, $m1, $d1) = explode('-', $datestart);
    list($Y2, $m2, $d2) = explode('-', $date2);
    $Y = $Y2 - $Y1;

    $m = $m2 - $m1;

    $d = $d2 - $d1;

    if ($d < 0) {
        $d += (int)date('t', strtotime("-1 month $date2"));
        $m = $m - 1;
    }
    if ($m < 0) {
        $m += 12;
        $Y = $Y - 1;
    }
    $m = $Y * 12 + $m;
    return $m;
}

function get_excel($url, $filename = '', $savefile = './caiji/')
{
    $imgArr = array('xls', 'xlsx');

    if (!$url) return false;

    if (!$filename) {
        $ext = strtolower(end(explode('.', $url)));
        if (!in_array($ext, $imgArr)) return false;
        $filename = date("dMYHis") . '.' . $ext;
    }

    if (!is_dir($savefile)) mkdir($savefile, 0777);
    if (!is_readable($savefile)) chmod($savefile, 0777);

    $filename = $savefile . $filename;
    if (file_exists($filename)) {
        unlink($filename);
    }
    ob_start();
    readfile($url);
    $img = ob_get_contents();
    ob_end_clean();
    $size = strlen($img);

    $fp2 = @fopen($filename, "a");
    fwrite($fp2, $img);
    fclose($fp2);

    return $filename;
}

function importTenant($area_id = 0, $user_id = 0)
{
    $sign = 1;
    header("content-type:text/html;charset=utf-8");

    //上传excel文件
    $file_name = request()->post('name');

    //获取文件路径
    // $filePath = dirname(dirname(__DIR__)) . '/副本【出租屋管理系统】租客导入.xlsx';
    $session = Session::getSessionId();
    $filePath = dirname(__DIR__) . '/public/temp/' . $session . '/' . $file_name;
    if (!file_exists($filePath)) {
        $return['status'] = '4';
        $return['message'] = '请重新上传';
        return $return;
    }
    Db::startTrans();
    try {
        // 加载文件
        $spreadsheet = IOFactory::load($filePath);
        $spreadsheet->setActiveSheetIndex(0);
        $sheetData = $spreadsheet->getActiveSheet()->toArray(false, true, true, true, true);
        $row_num = count($sheetData);
        $now_time = time();
        $without_name = [];
        // print_r($sheetData);
        // exit;
        $import_data = []; //数组形式获取表格数据

        for ($i = 5; $i <= $row_num; $i++) {

            $name            = trim($sheetData[$i]['A']);
            $identity        = trim($sheetData[$i]['B']);
            $mobile          = trim($sheetData[$i]['C']);
            $address         = trim($sheetData[$i]['D']);
            $landlord_name   = trim($sheetData[$i]['E']);
            $room            = trim($sheetData[$i]['F']);
            $live_num        = trim($sheetData[$i]['G']);
            $check_in_time   = trim($sheetData[$i]['H']);
            $check_out_time  = trim($sheetData[$i]['I']);
            $status          = trim($sheetData[$i]['J']);
            $huji_address    = trim($sheetData[$i]['K']);
            $nation          = trim($sheetData[$i]['L']);
            $row_name        = '第' . $i . '行:' . $name;

            if (!empty($name) && strlen($identity) == 18) {
                $import_data[$i]['name']           = $name;
                $import_data[$i]['identity']       = $identity;
                $import_data[$i]['mobile']         = $mobile;
                $import_data[$i]['address']        = $address;
                $import_data[$i]['landlord_name']  = $landlord_name;
                $import_data[$i]['room']           = $room;
                $import_data[$i]['live_num']       = $live_num;
                $import_data[$i]['check_in_time']  = strtotime($check_in_time);
                $import_data[$i]['check_out_time'] = strtotime($check_out_time);
                $import_data[$i]['status']         = $status;
                $import_data[$i]['huji_address']   = $huji_address;
                $import_data[$i]['create_time']    = time();
                $import_data[$i]['row_name']       = $row_name;
                $import_data[$i]['nation']         = $nation;
                $import_data[$i]['input_from']     = 1;// 标记来源是导入的
            } else {
                if (!empty($name)) {
                    $without_name[] = $row_name;
                }
            }
        }
        $error_arr = [];
        sort($import_data);
        // 没有姓名或身份证长度不是18位
        if (!empty($without_name)) {
            $sub_error['reason']  = '身份证长度不为18位';
            $sub_error['count']   = count($without_name);
            $sub_error['persons'] = implode(",", $without_name);
            $error_arr[] = $sub_error;
        }
        // 1、获取房屋id
        $houses_address = array_column($import_data, 'address');
        $house_ids = Db::name('renthouse_new_house')
            ->join('renthouse_new_landlord', 'renthouse_new_landlord.id = renthouse_new_house.landlord_id', 'left')
            ->where([
                'area_id' => $area_id,
                'renthouse_new_landlord.status' => 1,
                'renthouse_new_house.status' => 1,
                'renthouse_new_house.address' => ['in', $houses_address]
            ])
            ->field('renthouse_new_house.id, renthouse_new_house.address, landlord_id')
            ->select();
        $name_arr = [];
        $address_arr = [];
        // $houseRoom = [];
        foreach ($import_data as $key => $val) {
            $import_data[$key]['house_id'] = 0;
            foreach ($house_ids as $key2 => $val2) {
                if ($import_data[$key]['address'] == $house_ids[$key2]['address']) {
                    $import_data[$key]['house_id'] = $house_ids[$key2]['id'];
                    $import_data[$key]['landlord_id'] = $house_ids[$key2]['landlord_id'];
                }
            }
            if ($import_data[$key]['house_id'] == 0) {
                $name_arr[] = $import_data[$key]['row_name'];
                if (!in_array($import_data[$key]['address'], $address_arr)) {
                    $address_arr[] = $import_data[$key]['address'];
                }
                unset($import_data[$key]);
            } else {
                // if (isset($houseRoom[$import_data[$key]['house_id']])) {
                //     if (!in_array($import_data[$key]['room'], $houseRoom[$import_data[$key]['house_id']])) {
                //         $houseRoom[$import_data[$key]['house_id']][] = $import_data[$key]['room'];
                //     }
                // } else {
                //     $houseRoom[$import_data[$key]['house_id']][] = $import_data[$key]['room'];
                // }
            }
        }
        if (!empty($name_arr)) {
            $sub_error['reason']  = '房屋列表中获取对应房屋失败，请检查房屋地址';
            $sub_error['count']   = count($name_arr);
            $sub_error['persons'] = implode(",", $name_arr);
            $sub_error['address'] = implode(",", $address_arr);
            $error_arr[] = $sub_error;
        }
        // 添加新房间
        // if (!empty($houseRoom)) {
        //     foreach ($houseRoom as $key => $val) {
        //         $get_floor_box = Db::name('renthouse_new_house')->where(['id' => $key])->value('floor_box');
        //         if (!empty($get_floor_box)) {
        //             $floor_box = json_decode($get_floor_box, true);
        //             $room_list = [];
        //             foreach ($floor_box as $key2 => $val2) {
        //                 if (!empty($floor_box[$key2]['detail'])) {
        //                     foreach ($floor_box[$key2]['detail'] as $key3 => $val3) {
        //                         $room_list[] = $floor_box[$key2]['detail'][$key3]['room'];
        //                     }
        //                 }
        //             }
                    
        //             foreach ($houseRoom[$key] as $key2 => $val2) {
        //                 if (!in_array($houseRoom[$key][$key2], $room_list)) {
        //                     $new_room = [];
        //                     $new_room['room'] = $houseRoom[$key][$key2];
        //                     $num = substr($new_room['room'], 0, 1);
        //                     if (is_numeric($num) && $num > 0 && $num < 6) {
        //                         array_push($floor_box[$num - 1]['detail'], $new_room);
        //                     } else {
        //                         array_push($floor_box[0]['detail'], $new_room);
        //                     }
        //                 }
        //             }
        //         } else { // 没有楼层管理的情况
        //             $floor_box = [
        //                 [
        //                     "storey" => "第一层房间设置：",
        //                     "detail" => []
        //                 ],
        //                 [
        //                     "storey" => "第二层房间设置：",
        //                     "detail" => []
        //                 ],
        //                 [
        //                     "storey" => "第三层房间设置：",
        //                     "detail" => []
        //                 ],
        //                 [
        //                     "storey" => "第四层房间设置：",
        //                     "detail" => []
        //                 ],
        //                 [
        //                     "storey" => "第五层房间设置：",
        //                     "detail" => []
        //                 ]
        //             ];
        //             foreach ($houseRoom[$key] as $key2 => $val2) {
        //                 $new_room = [];
        //                 $new_room['room'] = $houseRoom[$key][$key2];
        //                 $num = substr($new_room['room'], 0, 1);
        //                 if (is_numeric($num) && $num > 0 && $num < 6) {
        //                     array_push($floor_box[$num - 1]['detail'], $new_room);
        //                 } else {
        //                     array_push($floor_box[0]['detail'], $new_room);
        //                 }
        //             }
        //         }
        //         Db::name('renthouse_new_house')->where(['id' => $key])->update(['floor_box' => json_encode($floor_box)]);
        //     }
        // }

        // 2、检查有无重复数据，个别地区不允许一人多屋
        sort($import_data);
        $czw_village_registration = Db::name('area')->where(['id' => $area_id])->value('czw_village_registration');
        $identity_array = [];
        foreach ($import_data as $key => $val) {
            if (!empty($import_data[$key]['identity'])) {
                $identity_array[] = $import_data[$key]['identity'];
            }
        }
        $result_data = Db::name('renthouse_new_tenant')
            ->join('renthouse_new_landlord', 'renthouse_new_tenant.landlord_id = renthouse_new_landlord.id', 'left')
            ->field('renthouse_new_tenant.id, identity, house_id, room')
            ->where([
                'renthouse_new_tenant.identity' => ['in', $identity_array],
                'renthouse_new_landlord.status' => 1,
                'area_id' => $area_id
            ])
            ->select();

        if (!empty($result_data)) {
            foreach ($result_data as $key => $val) {
                $result_data[$key]['check'] = $result_data[$key]['identity'] . '-' . $result_data[$key]['house_id'] . '-' . $result_data[$key]['room'];
                // $result_data[$key]['check'] = $result_data[$key]['identity'] . '-' . $result_data[$key]['house_id'];
            }
            $result_data_array = array_column($result_data, 'check');
            $identity_data_array = array_column($result_data, 'identity');
            $name_arr = [];
            $name_arr2 = [];
            $new_arr = [];
            foreach ($import_data as $key => $val) {
                $import_data[$key]['check'] = $import_data[$key]['identity'] . '-' . $import_data[$key]['house_id'] . '-' . $import_data[$key]['room'];
                // $import_data[$key]['check'] = $import_data[$key]['identity'] . '-' . $import_data[$key]['house_id'];
                // 去掉重复登记
                if (in_array($import_data[$key]['check'], $result_data_array)) {
                    $name_arr[] = $import_data[$key]['row_name'];
                    unset($import_data[$key]);
                }
            }
            if (!empty($name_arr)) {
                $sub_error['reason']  = '此租客信息已添加，无需重复导入';
                $sub_error['count']   = count($name_arr);
                $sub_error['persons'] = implode(",", $name_arr);
                $error_arr[] = $sub_error;
            }

            if ($czw_village_registration == 1) { //不允许登记多屋
                // 不允许登记多屋
                foreach ($import_data as $key => $val) {
                    if (in_array($import_data[$key]['identity'], $identity_data_array)) {
                        $name_arr2[] = $import_data[$key]['row_name'];
                        unset($import_data[$key]);
                    }
                }
                if (!empty($name_arr2)) {
                    $sub_error['reason'] = '在当前地区已登记其他住址';
                    $sub_error['count']   = count($name_arr2);
                    $sub_error['persons'] = implode(",", $name_arr2);
                    $error_arr[] = $sub_error;
                }
            }
        }

        // 3、最后处理多余数据
        $insert_field = ['name', 'gender', 'identity', 'mobile', 'room', 'live_num', 'age', 'huji_address', 'house_id', 'landlord_id', 'check_in_time', 'check_out_time', 'status', 'create_time', 'nation', 'relationship', 'input_from'];
        foreach ($import_data as $key => $val) {
            // 电话
            if (!empty($import_data[$key]['mobile'])) {
                if ($import_data[$key]['mobile'] == '无') {
                    $import_data[$key]['mobile'] = '';
                }
            }
            // 状态
            if (!empty($import_data[$key]['status'])) {
                if ($import_data[$key]['status'] == '在租') {
                    $import_data[$key]['status'] = 1;
                } elseif ($import_data[$key]['status'] == '退租') {
                    $import_data[$key]['status'] = 0;
                } else {
                    $import_data[$key]['status'] = 1;
                }
            } else {
                $import_data[$key]['status'] = 1;
            }
            // 如果填了退租时间，那就标记为已退租
            if ($import_data[$key]['check_out_time'] > 0) {
                $import_data[$key]['status'] = 0;
            }
            // 性别、年龄处理
            if (intval(substr($import_data[$key]['identity'], 14, 3)) % 2 == 0) {
                $import_data[$key]['gender'] = 2;
            } else {
                $import_data[$key]['gender'] = 1;
            }
            $import_data[$key]['year'] = intval(substr($import_data[$key]['identity'], 6, 4));
            $import_data[$key]['month'] = intval(substr($import_data[$key]['identity'], 10, 2));
            $import_data[$key]['day'] = intval(substr($import_data[$key]['identity'], 12, 2));
            $m_count = diffDate(date('Y-m-d', time()), $import_data[$key]['year'] . '-' . $import_data[$key]['month'] . '-' . $import_data[$key]['day']);
            $import_data[$key]['age'] = ceil($m_count / 12);

            // 关系默认设置为本人
            $import_data[$key]['relationship'] = 1;

            // 多余字段处理
            $diff = array_diff(array_keys($import_data[$key]), $insert_field);
            foreach ($diff as $key2 => $val2) {
                unset($import_data[$key][$val2]);
            }
        }
        sort($import_data);

        // print_r($import_data);
        // exit;

        if (!empty($import_data)) {
            FileCtrl::flushTemp(); //清缓存
            // 将数据保存到数据库
            $before_insert_id = Db::name('renthouse_new_tenant')->order('id desc')->value('id');
            $res = Db::name('renthouse_new_tenant')->insertAll($import_data);
            $after_insert_id = Db::name('renthouse_new_tenant')->order('id desc')->value('id');
            if ($res) {
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 1;
                $insert_info_arr = [];
                $insert_info['form'] = 'renthouse_new_tenant';
                $insert_info['before_insert_id'] = $before_insert_id;
                $insert_info['after_insert_id'] = $after_insert_id;
                $insert_info_arr[] = $insert_info;
                $log['new_data_info'] = json_encode($insert_info_arr);
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '1';
                $return['message'] = '导入成功';
                $return['result'] = $error_arr;
                Db::commit();
                return $return;
            } else {
                Db::rollback();
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 2;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '2';
                $return['message'] = '导入失败';
                $return['result'] = $error_arr;
                return $return;
            }
        }
        Db::rollback();
        // 导入日志
        $log = [];
        $log['user_id']     = $user_id;
        $log['err_data']    = json_encode($error_arr);
        $log['create_time'] = time();
        $log['sign']        = $sign;
        $log['result']      = 3;
        Db::name('renthouse_new_import_log')->insert($log);
        $return['status'] = '3';
        $return['message'] = '无新数据添加';
        $return['result'] = $error_arr;
        return $return;
    } catch (\Exception $e) {
        Db::rollback();
        // 导入日志
        $log = [];
        $log['user_id']     = $user_id;
        $log['err_data']    = json_encode($error_arr);
        $log['create_time'] = time();
        $log['sign']        = $sign;
        $log['result']      = 4;
        Db::name('renthouse_new_import_log')->insert($log);
        $return['status'] = '4';
        $return['message'] = 'error';
        $return['result'] = $error_arr;
        return $return;
    }
}

function importHouse($area_id = 0, $user_id = 0)
{
    $sign = 2;
    header("content-type:text/html;charset=utf-8");

    //上传excel文件
    $file_name = request()->post('name');

    //获取文件路径
    // $filePath = dirname(dirname(__DIR__)) . '/【出租屋管理系统】房屋数据导入模版.xlsx';
    $session = Session::getSessionId();
    $filePath = dirname(__DIR__) . '/public/temp/' . $session . '/' . $file_name;
    if (!file_exists($filePath)) {
        $return['status'] = '4';
        $return['message'] = '请重新上传';
        return $return;
    }
    Db::startTrans();
    try {
        // 加载文件
        $spreadsheet = IOFactory::load($filePath);
        $spreadsheet->setActiveSheetIndex(0);
        $sheetData = $spreadsheet->getActiveSheet()->toArray(false, true, true, true, true);
        $row_num = count($sheetData);
        $now_time = time();
        // print_r($sheetData);
        // exit;
        $import_data = []; //数组形式获取表格数据
        $error_arr = [];
        $no_landlord_arr = [];
        $is_grid = [];
        $excel_mobile_arr = [];
        // 获取已添加的房东和网格员 以手机号码为唯一值 身份证是能重复添加的
        $landlord_list = Db::name('renthouse_new_landlord')->where(['status' => 1])->field('id, mobile, role')->select();
        $mobile_arr = array_column($landlord_list, 'mobile', 'id');
        $role_arr = array_column($landlord_list, 'role', 'id');
        // 获取已添加的网格员
        $grid_list = Db::name('renthouse_new_landlord')->where(['status' => 1, 'role' => 1, 'area_id' => $area_id])->field('id, name')->select();
        $grid_name_arr = array_column($grid_list, 'name', 'id');
        // 需要新增的网格员信息
        $grid_house_add = [];
        // 网格员异常数组
        $grid_err_arr = [];
        // 房东手机号码异常数组
        $landlord_err_arr = [];

        for ($i = 5; $i <= $row_num; $i++) {

            $name           = trim($sheetData[$i]['A']);
            $identity       = trim($sheetData[$i]['B']);
            $address        = trim($sheetData[$i]['C']);
            $mobile         = trim($sheetData[$i]['D']);
            $owner_name     = trim($sheetData[$i]['E']);
            $owner_identity = trim($sheetData[$i]['F']);
            $owner_address  = trim($sheetData[$i]['G']);
            // $owner_address  = '';
            $owner_mobile   = trim($sheetData[$i]['H']);
            $grid_members   = trim($sheetData[$i]['I']);
            // $grid_members   = '关日锐,关国盖,关祖铨,何桂锵,关结玲,何福汉,陆乐,罗凤雲,关炜怡,罗燕芬,关桂学';
            $house_address  = trim($sheetData[$i]['J']);
            $row_name       = '第' . $i . '行:' . $name;

            if (!empty($name) || !empty($owner_name) || !empty($house_address)) {
                if (!empty($mobile)) { // 如果有房东
                    if (in_array($mobile, $mobile_arr)) { // 该房东已存在
                        $landlord_id = array_search($mobile, $mobile_arr);
                        $role = $role_arr[$landlord_id];
                        if ($role == 0) {
                            $import_data[$i]['landlord_id'] = $landlord_id;
                        } elseif ($role == 1) {
                            $is_grid[] = $row_name;
                            continue;
                        }
                    } else { // 该房东不存在
                        if (!in_array($mobile, $excel_mobile_arr)) {
                            if (strlen($mobile) == 11) {
                                $new_landlord['name']                 = $name;
                                $new_landlord['area_id']              = $area_id;
                                $new_landlord['identity_card_number'] = $identity;
                                $new_landlord['address']              = $address;
                                $new_landlord['mobile']               = $mobile;
                                $new_landlord['create_time']          = time();
                                $new_landlord['status']               = 1;
                                $new_landlord['input_from']           = 1;
                                $import_data[$i]['landlord_id'] = Db::name('renthouse_new_landlord')->insertGetId($new_landlord);
                                $excel_mobile_arr[] = $mobile;
                            } else {
                                $landlord_err_arr[] = $row_name;
                                continue;
                            }
                        } else {
                            $import_data[$i]['landlord_id'] = Db::name('renthouse_new_landlord')->where(['mobile' => $mobile, 'status' => 1])->value('id');
                        }
                    }
                } else {
                    if (!empty($owner_mobile)) { // 如果没有填房东只填了屋主
                        if (in_array($owner_mobile, $mobile_arr)) { // 该房东已存在
                            $landlord_id = array_search($owner_mobile, $mobile_arr);
                            $role = $role_arr[$landlord_id];
                            if ($role == 0) {
                                $import_data[$i]['landlord_id'] = $landlord_id;
                            } elseif ($role == 1) {
                                $is_grid[] = $row_name;
                            }
                        } else { // 该房东不存在
                            if (!in_array($owner_mobile, $excel_mobile_arr)) {
                                if (strlen($mobile) == 11) {
                                    $new_landlord['name']                 = $owner_name;
                                    $new_landlord['area_id']              = $area_id;
                                    $new_landlord['identity_card_number'] = $owner_identity;
                                    $new_landlord['address']              = $owner_address;
                                    $new_landlord['mobile']               = $owner_mobile;
                                    $new_landlord['create_time']          = time();
                                    $new_landlord['status']               = 1;
                                    $new_landlord['input_from']           = 1;
                                    $import_data[$i]['landlord_id'] = Db::name('renthouse_new_landlord')->insertGetId($new_landlord);
                                    $excel_mobile_arr[] = $owner_mobile;
                                } else {
                                    $landlord_err_arr[] = $row_name;
                                    continue;
                                }
                            } else {
                                $import_data[$i]['landlord_id'] = Db::name('renthouse_new_landlord')->where(['mobile' => $owner_mobile, 'status' => 1])->value('id');
                            }
                        }
                    } else { //房东和屋主都没有填
                        $no_landlord_arr[] = $row_name;
                        continue;
                    }
                }
                $import_data[$i]['owner_name']     = $owner_name;
                $import_data[$i]['owner_identity'] = $owner_identity;
                $import_data[$i]['owner_address']  = $owner_address;
                $import_data[$i]['owner_mobile']   = $owner_mobile;
                $import_data[$i]['house_address']  = $house_address;
                $import_data[$i]['row_name']       = $row_name;
                $import_data[$i]['rental_type']    = 0;
                $import_data[$i]['status']         = 1;
                $import_data[$i]['create_time']    = time();
                $import_data[$i]['input_from']     = 1;
                // 1 自治 2 承包
                if (empty($name) && !empty($owner_name)) {
                    $import_data[$i]['rental_type'] = 1;
                }
                if (!empty($name) && !empty($owner_name) && $name == $owner_name) {
                    $import_data[$i]['rental_type'] = 1;
                }
                if (!empty($name) && empty($owner_name)) {
                    $import_data[$i]['rental_type'] = 2;
                }
                if (!empty($name) && !empty($owner_name) && $name != $owner_name) {
                    $import_data[$i]['rental_type'] = 2;
                }
                // 网格员
                if (!empty($grid_members)) {
                    if (strpos($grid_members, "，")) {
                        $grid_members = str_replace("，", ",", $grid_members);
                    }
                    $grid_members = explode(",", $grid_members);
                    $grid_ids = [];
                    if (count($grid_members) > 0) {
                        $count1 = count($grid_members);
                        foreach ($grid_members as $key => $val) {
                            if (in_array($grid_members[$key], $grid_name_arr)) {
                                $sub['landlord_id'] = $import_data[$i]['landlord_id'];
                                $sub['house_address'] = $import_data[$i]['house_address'];
                                $sub['grid_id'] = array_search($grid_members[$key], $grid_name_arr);
                                $grid_house_add[] = $sub;
                                $grid_ids[] = $sub['grid_id'];
                            }
                        }
                        $count2 = count($grid_ids);
                        array_unique($grid_ids);
                        $import_data[$i]['grid_members'] = implode(",", $grid_ids);
                        if ($count1 != $count2) {
                            $grid_err_arr[] = $row_name;
                        }
                    }
                }
            }
        }

        sort($import_data);
        $house_addresses = [];
        $repeat_data = [];
        foreach ($import_data as $key => $val) {
            $import_data[$key]['landlord_address'] = $import_data[$key]['landlord_id'] . '-' . $import_data[$key]['house_address'];
            $house_addresses[] = $import_data[$key]['house_address'];
        }
        $get_houses = Db::name('renthouse_new_house')
            ->where([
                'address' => ['in', $house_addresses]
            ])
            ->field('id, landlord_id, address')
            ->select();
        foreach ($get_houses as $key => $val) {
            $get_houses[$key]['landlord_address'] = $get_houses[$key]['landlord_id'] . '-' . $get_houses[$key]['address'];
        }
        $landlord_address_arr = array_column($get_houses, 'landlord_address');
        foreach ($import_data as $key => $val) {
            if (in_array($import_data[$key]['landlord_address'], $landlord_address_arr)) {
                $repeat_data[] = $import_data[$key]['row_name'];
                if (in_array($import_data[$key]['row_name'], $grid_err_arr)) {
                    unset($grid_err_arr[array_search($import_data[$key]['row_name'], $grid_err_arr)]);
                }
                unset($import_data[$key]);
            }
        }
        // 错误提示
        // 1、房东和屋主都没有填
        if (!empty($no_landlord_arr)) {
            $sub_error['reason']  = '房屋管理员和屋主手机号码都为空，无法添加房东及房屋数据';
            $sub_error['count']   = count($no_landlord_arr);
            $sub_error['persons'] = implode(",", $no_landlord_arr);
            $error_arr[] = $sub_error;
        }
        // 2、房东手机号码异常
        if (!empty($landlord_err_arr)) {
            $sub_error['reason']  = '房屋管理员手机号码为空或手机号码长度不等于11，无法添加房东数据';
            $sub_error['count']   = count($landlord_err_arr);
            $sub_error['persons'] = implode(",", $landlord_err_arr);
            $error_arr[] = $sub_error;
        }
        // 3、已是网格员
        if (!empty($is_grid)) {
            $sub_error['reason']  = '此手机号码已添加为网格员，无法设置为房屋管理员，请更换手机号码';
            $sub_error['count']   = count($is_grid);
            $sub_error['persons'] = implode(",", $is_grid);
            $error_arr[] = $sub_error;
        }
        // 4、此房子已添加
        if (!empty($repeat_data)) {
            $sub_error['reason']  = '此房屋信息已添加，已省略添加，如需更改请前往房屋列表';
            $sub_error['count']   = count($repeat_data);
            $sub_error['persons'] = implode(",", $repeat_data);
            $error_arr[] = $sub_error;
        }
        // 5、存在未添加的网格员
        if (!empty($grid_err_arr)) {
            $sub_error['reason']  = '存在未添加的网格员，请先前往网格员列表添加网格员，然后前往房屋详情选择网格员';
            $sub_error['count']   = count($grid_err_arr);
            $sub_error['persons'] = implode(",", $grid_err_arr);
            $error_arr[] = $sub_error;
        }
        // 6、最后处理多余数据
        $insert_field = ['landlord_id', 'address', 'create_time', 'status', 'owner_name', 'owner_mobile', 'owner_identity', 'owner_address', 'rental_type', 'grid_members'];
        foreach ($import_data as $key => $val) {
            $import_data[$key]['address'] = $import_data[$key]['house_address'];
            // 多余字段处理
            $diff = array_diff(array_keys($import_data[$key]), $insert_field);
            foreach ($diff as $key2 => $val2) {
                unset($import_data[$key][$val2]);
            }
        }
        sort($import_data);
        // print_r($error_arr);
        // exit;
        if (!empty($import_data)) {
            FileCtrl::flushTemp(); //清缓存
            // 将数据保存到数据库
            $res = Db::name('renthouse_new_house')->insertAll($import_data);
            if ($res) {
                // 处理网格员
                if (count($grid_house_add) > 0) {
                    $get_new_houses = Db::name('renthouse_new_house')
                        ->where([
                            'address' => ['in', $house_addresses]
                        ])
                        ->field('id, grid_members')
                        ->select();
                    foreach ($get_new_houses as $key => $val) {
                        if (!empty($get_new_houses[$key]['grid_members'])) {
                            $grid_members = explode(",", $get_new_houses[$key]['grid_members']);
                            if (count($grid_members) > 0) {
                                Db::name('renthouse_new_grid_house')->where(['house_id' => $get_new_houses[$key]['id']])->delete();
                                $insert_data = [];
                                foreach ($grid_members as $key2 => $val2) {
                                    $sub_data['grid_id'] = $grid_members[$key2];
                                    $sub_data['house_id'] = $get_new_houses[$key]['id'];
                                    $insert_data[] = $sub_data;
                                }
                                Db::name('renthouse_new_grid_house')->insertAll($insert_data);
                            }
                        }
                    }
                }
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 1;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '1';
                $return['message'] = '导入成功';
                $return['result'] = $error_arr;
                Db::commit();
                return $return;
            } else {
                Db::rollback();
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 2;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '2';
                $return['message'] = '导入失败';
                $return['result'] = $error_arr;
                return $return;
            }
        }
        Db::rollback();
        // 导入日志
        $log = [];
        $log['user_id']     = $user_id;
        $log['err_data']    = json_encode($error_arr);
        $log['create_time'] = time();
        $log['sign']        = $sign;
        $log['result']      = 3;
        Db::name('renthouse_new_import_log')->insert($log);
        $return['status'] = '3';
        $return['message'] = '无新数据添加';
        $return['result'] = $error_arr;
        return $return;
    } catch (\Exception $e) {
        Db::rollback();
        // 导入日志
        $log = [];
        $log['user_id']     = $user_id;
        $log['err_data']    = json_encode($error_arr);
        $log['create_time'] = time();
        $log['sign']        = $sign;
        $log['result']      = 4;
        Db::name('renthouse_new_import_log')->insert($log);
        $return['status'] = '4';
        $return['message'] = 'error';
        $return['result'] = $error_arr;
        return $return;
    }
}

function importLandlord($area_id = 0, $user_id = 0)
{
    $sign = 3;
    header("content-type:text/html;charset=utf-8");

    //上传excel文件
    $file_name = request()->post('name');

    //获取文件路径
    $filePath = dirname(dirname(__DIR__)) . '/【出租屋管理系统】房屋数据导入模版.xlsx';
    // $session = Session::getSessionId();
    // $filePath = dirname(__DIR__) . '/public/temp/' . $session . '/' . $file_name;
    // if (!file_exists($filePath)) {
    //     $return['status'] = '4';
    //     $return['message'] = '请重新上传';
    //     return $return;
    // }
    Db::startTrans();
    try {
        // 加载文件
        $spreadsheet = IOFactory::load($filePath);
        $spreadsheet->setActiveSheetIndex(0);
        $sheetData = $spreadsheet->getActiveSheet()->toArray(false, true, true, true, true);
        $row_num = count($sheetData);
        $now_time = time();
        // print_r($sheetData);
        // exit;
        $import_data = []; //数组形式获取表格数据

        // 缺少姓名或手机号码
        $lack_of_key = [];
        // 获取已添加的房东
        $landlord_list = Db::name('renthouse_new_landlord')->where(['status' => 1, 'role' => 0])->field('id, mobile')->select();
        $landlord_mobile_arr = array_column($landlord_list, 'mobile', 'id');
        // 获取已添加的网格员
        $grid_list = Db::name('renthouse_new_landlord')->where(['status' => 1, 'role' => 1])->field('id, mobile')->select();
        $grid_mobile_arr = array_column($grid_list, 'mobile', 'id');
        // 已添加为房东
        $added_landlord = [];
        // 已添加为网格员
        $added_grid = [];

        for ($i = 5; $i <= $row_num; $i++) {

            $name           = trim($sheetData[$i]['A']);
            $identity       = trim($sheetData[$i]['B']);
            $address        = trim($sheetData[$i]['C']);
            $mobile         = trim($sheetData[$i]['D']);
            $row_name       = '第' . $i . '行:' . $name;

            if (!empty($name)) {
                if (!empty($name) && !empty($mobile) && strlen($mobile) == 11) {
                    if (in_array($mobile, $landlord_mobile_arr)) {
                        $added_landlord[] = $row_name;
                    } else {
                        if (in_array($mobile, $grid_mobile_arr)) {
                            $added_grid[] = $row_name;
                        } else {
                            $import_data[$i]['name']                 = $name;
                            $import_data[$i]['identity_card_number'] = $identity;
                            $import_data[$i]['address']              = $address;
                            $import_data[$i]['mobile']               = $mobile;
                            $import_data[$i]['status']               = 1;
                            $import_data[$i]['create_time']          = time();
                            $import_data[$i]['area_id']              = $area_id;
                            $import_data[$i]['input_from']           = 1;
                        }
                    }
                } else {
                    $lack_of_key[] = $row_name;
                }
            }
        }
        sort($import_data);
        // 错误提示
        $error_arr = [];
        // 1、房东电话号码或者姓名没有填
        if (!empty($lack_of_key)) {
            $sub_error['reason']  = '房屋管理员手机号码为空或手机号码长度不等于11，无法添加房东数据';
            $sub_error['count']   = count($lack_of_key);
            $sub_error['persons'] = implode(",", $lack_of_key);
            $error_arr[] = $sub_error;
        }
        // 2、已是房东
        if (!empty($added_landlord)) {
            $sub_error['reason']  = '此手机号码已添加为房东，无法设置为房东，请更换手机号码';
            $sub_error['count']   = count($added_landlord);
            $sub_error['persons'] = implode(",", $added_landlord);
            $error_arr[] = $sub_error;
        }
        // 2、已是网格员
        if (!empty($added_grid)) {
            $sub_error['reason']  = '此手机号码已添加为网格员，无法设置为房东，请更换手机号码';
            $sub_error['count']   = count($added_grid);
            $sub_error['persons'] = implode(",", $added_grid);
            $error_arr[] = $sub_error;
        }

        // 5、最后处理多余数据
        $insert_field = ['name', 'address', 'create_time', 'status', 'identity_card_number', 'mobile', 'area_id'];
        foreach ($import_data as $key => $val) {
            // 多余字段处理
            $diff = array_diff(array_keys($import_data[$key]), $insert_field);
            foreach ($diff as $key2 => $val2) {
                unset($import_data[$key][$val2]);
            }
        }
        sort($import_data);
        // print_r($error_arr);
        // exit;
        if (!empty($import_data)) {
            FileCtrl::flushTemp(); //清缓存
            // 将数据保存到数据库
            $res = Db::name('renthouse_new_landlord')->insertAll($import_data);
            if ($res) {
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 1;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '1';
                $return['message'] = '导入成功';
                $return['result'] = $error_arr;
                Db::commit();
                return $return;
            } else {
                Db::rollback();
                // 导入日志
                $log = [];
                $log['user_id']     = $user_id;
                $log['err_data']    = json_encode($error_arr);
                $log['create_time'] = time();
                $log['sign']        = $sign;
                $log['result']      = 2;
                Db::name('renthouse_new_import_log')->insert($log);
                $return['status'] = '2';
                $return['message'] = '导入失败';
                $return['result'] = $error_arr;
                return $return;
            }
        }
        Db::rollback();
        // 导入日志
        $log = [];
        $log['user_id']     = $user_id;
        $log['err_data']    = json_encode($error_arr);
        $log['create_time'] = time();
        $log['sign']        = $sign;
        $log['result']      = 3;
        Db::name('renthouse_new_import_log')->insert($log);
        $return['status'] = '3';
        $return['message'] = '无新数据添加';
        $return['result'] = $error_arr;
        return $return;
    } catch (\Exception $e) {
        Db::rollback();
        // 导入日志
        $log = [];
        $log['user_id']     = $user_id;
        $log['err_data']    = json_encode($error_arr);
        $log['create_time'] = time();
        $log['sign']        = $sign;
        $log['result']      = 4;
        Db::name('renthouse_new_import_log')->insert($log);
        $return['status'] = '4';
        $return['message'] = 'error';
        $return['result'] = $error_arr;
        return $return;
    }
}
