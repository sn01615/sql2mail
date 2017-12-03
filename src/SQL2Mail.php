<?php
namespace PhpUtils;

use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\PHPMailer;

/**
 *
 * @author YangLong
 * Date: 2017-10-09
 */
class SQL2Mail
{

    private $host, $username, $password, $port;

    private $debugLevel = 0;

    private $mail;

    private $mailto, $cc, $bcc, $replyto, $attachments, $ishtml;

    private $subject, $body, $altbody;

    public function __construct($host, $username, $password, $port = 587)
    {
        $this->host = $host;
        $this->username = $username;
        $this->password = $password;
        $this->port = $port;
        
        $this->mail = new PHPMailer(true); // Passing `true` enables exceptions
        
        //Server settings
        $this->mail->SMTPDebug = $this->debugLevel; // Enable verbose debug output
        $this->mail->isSMTP(); // Set mailer to use SMTP
        $this->mail->Host = $this->host; // Specify main and backup SMTP servers
        $this->mail->SMTPAuth = true; // Enable SMTP authentication
        $this->mail->Username = $this->username; // SMTP username
        $this->mail->Password = $this->password; // SMTP password
        $this->mail->SMTPSecure = 'tls'; // Enable TLS encryption, `ssl` also accepted
        $this->mail->Port = $this->port; // TCP port to connect to
    }

    public function setDebug($level)
    {
        $this->mail->SMTPDebug = $this->debugLevel = $level;
        return $this;
    }

    public function setFrom($address, $name)
    {
        $this->mail->setFrom($address, $name);
        return $this;
    }

    public function setMailTo(array $mailto)
    {
        $this->mailto = $mailto;
        return $this;
    }

    public function setCc(array $cc)
    {
        $this->cc = $cc;
        return $this;
    }

    public function setBcc(array $bcc)
    {
        $this->bcc = $bcc;
        return $this;
    }

    public function setReplyTo(array $replyto)
    {
        $this->replyto = $replyto;
        return $this;
    }

    public function setAttachments(array $attachments)
    {
        $this->attachments = $attachments;
        return $this;
    }

    public function isHtml(bool $isHtml)
    {
        $this->ishtml = $isHtml;
        return $this;
    }

    public function setSubject(string $subject)
    {
        $this->subject = $subject;
        return $this;
    }

    public function setBody(string $body)
    {
        $this->body = $body;
        return $this;
    }

    private function updateSet()
    {
        //Recipients
        foreach ($this->mailto as $name => $address) {
            if (is_string($name)) {
                $this->mail->addAddress($address, $name); // Add a recipient
            } else {
                $this->mail->addAddress($address); // Name is optional
            }
        }
        foreach ($this->cc as $name => $address) {
            $this->mail->addCC($address);
        }
        foreach ($this->bcc as $name => $address) {
            $this->mail->addBCC($address);
        }
        foreach ($this->replyto as $name => $address) {
            $this->mail->addReplyTo($address);
        }
        
        //Attachments
        foreach ($this->replyto as $name => $path) {
            if (is_string($name)) {
                $this->mail->addAttachment($path, $name); // Optional name
            } else {
                $this->mail->addAttachment($path); // Add attachments
            }
        }
        
        //Content
        $this->mail->isHTML($this->ishtml); // Set email format to HTML
        $this->mail->Subject = $this->subject;
        $this->mail->Body = $this->body;
        $this->mail->AltBody = $this->altbody;
    }

    public function sendTable($table, $title = null, $filename = '����')
    {
        $this->updateSet();
        try {
            $index = 0;
            
            $title or $title = $this->subject;
            
            // Create new PHPExcel object
            $objPHPExcel = new \PHPExcel();
            
            // �������
            $first_row = reset($table);
            $y = 'A';
            $first_arr = [];
            foreach ($first_row as $key => $row) {
                $first_arr[$y] = $key;
                $y ++;
            }
            unset($y);
            unset($first_row);
            
            $x = 1;
            $index and $objPHPExcel->createSheet();
            $as = $objPHPExcel->setActiveSheetIndex($index ++);
            foreach ($first_arr as $key => $value) {
                $as->setCellValueExplicit("{$key}{$x}", $value);
            }
            
            // �������
            foreach ($table as $row) {
                $x ++;
                foreach ($first_arr as $key => $value) {
                    if (is_array($row->{$value})) {
                        if ($row->{$value}['type'] == 'comment') {
                            $as->setCellValueExplicit("{$key}{$x}", $row->{$value}['value']);
                            // $as->getComment("{$key}{$x}")->setAuthor("{$day}");
                            $as->getComment("{$key}{$x}")
                                ->getText()
                                ->createTextRun($row->{$value}['comment']);
                        }
                    } else {
                        $as->setCellValueExplicit("{$key}{$x}", $row->{$value});
                    }
                }
            }
            
            // Rename worksheet
            $objPHPExcel->getActiveSheet()->setTitle($title);
            
            // Set active sheet index to the first sheet, so Excel opens this as the first sheet
            $objPHPExcel->setActiveSheetIndex(0);
            
            $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
            $filename = "{$filename}.xls";
            $filepath = sys_get_temp_dir() . '/' . $filename;
            $objWriter->save($filepath);
            
            $this->mail->send();
            echo 'Message has been sent';
        } catch (\Exception $e) {
            echo 'Message could not be sent.';
            echo 'Mailer Error: ' . $this->mail->ErrorInfo;
        }
    }
}

