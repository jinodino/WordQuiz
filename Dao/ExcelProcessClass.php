<?php

require_once "../PHPExcel/Classes/PHPExcel.php";

class ExcelProcess
{
    private $filePath;

    public function __construct()
    {
        $this->filePath = "C:\Users\sonjh\Desktop\単語.xlsx";
    }

    // 一つの単語を選んでクイズを提出
    public function submitAnExamination()
    {
        // 基本単語配列
        $wordArr = $this->readToWord();   
        
        // 매개변수로 셔플 선택 유무 확인
        shuffle($wordArr);
        
        return $wordArr;
        
    }
    
    // 単語がある行を読む
    public function readToWord()
    {
        $objPHPExcel = new PHPExcel();

        $wordArr = [];

        // $filename = iconv("UTF-8", "EUC-KR", $this->filePath);

        $objPHPExcel = PHPExcel_IOFactory::load($this->filePath);

        $sheetsCount = $objPHPExcel->getSheetCount();


        for($sheet = 0; $sheet < $sheetsCount; $sheet++) {

            $objPHPExcel->setActiveSheetIndex($sheet);
            $activesheet   = $objPHPExcel -> getActiveSheet();
            $highestRow    = $activesheet -> getHighestRow();             // 마지막 행
            $highestColumn = $activesheet -> getHighestColumn();    // 마지막 컬럼
  
            // 한줄읽기 -> 인덱스가 0부터 시작아님 1부터 시작(한자 히라가나 뜻) 그래서 인덱스 1은 제외한다
            for($row = 2; $row <= $highestRow; $row++) {
  
                // $rowData가 한줄의 데이터를 셀별로 배열처리 된다.
                $rowData = $activesheet -> rangeToArray("A" . $row . ":" . $highestColumn . $row, NULL, TRUE, FALSE);
    
                // $rowData에 들어가는 값은 계속 초기화 되기때문에 값을 담을 새로운 배열을 선안하고 담는다.
                $wordArr[$row] = $rowData[0];
            }
        }
        
        return $wordArr;
    }
}


?>