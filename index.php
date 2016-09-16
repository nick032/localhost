<?php
require_once 'config.php';
require_once 'Classes/PHPExcel.php';

//$excelObject = new PHPExcel();
//echo __DIR__;
$cells = array('A' => 'model',
               'B' => 'name',
               'C' => 'category',
               'D' => 'sub_category',
               'E' => 'brend',
               'F' => 'price', 
               'G' => 'quantity');

if(!empty($_FILES['file']['tmp_name'])) {

    $file = uploadFile($_FILES['file']['tmp_name']);
    
}else {
    echo 'No file for upload';
}

$xlsObject = PHPExcel_IOFactory::load($file);

$xlsObject->setActiveSheetIndex(0);

$sheet = $xlsObject->getActiveSheet();

$rowIterator = $sheet->getRowIterator();

foreach($rowIterator as $row) {
    if($row->getRowIndex() != 1) {
        $cellIterator = $row->getCellIterator();
        foreach($cellIterator as $cell) {
            $cellPath = $cell->getColumn();
            if(isset($cells[$cellPath])) {
                $data[$row->getRowIndex()][$cells[$cellPath]] = $cell->getCalculatedValue(); 
            }
        }
    }
}


$mysqli = new mysqli(DBHOST, DBUSER, DBPASSWORD, DBNAME);
if($mysqli->connect_errno){
    echo "Не удалось подключться с MySQL" . $mysqli->connect_error;
}

$db_prefix = 'oc_';

$category_id = getMaxId($mysqli, 'category_id', 'category');
$product_id = getMaxId($mysqli, 'product_id', 'product');



foreach ($data as $item) {


    $category_name = $mysqli->query("SELECT * FROM oc_category_description WHERE name = '{$item['category']}'");

    $date = date('Y-m-d H-m-s');

    if($category_name->num_rows == 0) {

        $category_id++;

        if(!$mysqli->query("INSERT INTO oc_category(category_id, 
                                                    parent_id, 
                                                    top, 
                                                    `column`, 
                                                    status, 
                                                    date_added, 
                                                    date_modified) 
                            VALUE({$category_id}, 0, 0, 0, 1, '$date', '$date')")) {
            die($mysqli->errno . " - " . $mysqli->error);

        }

        //TODO Записываем новую категорию и описание категории
        if(!$mysqli->query("INSERT INTO oc_category_description(category_id, language_id, name, description,
                            meta_title, meta_description, meta_keyword) 
                            VALUE('{$category_id}', 3, '{$item['category']}', '{$item['category']}', '{$item['category']}', '{$item['category']}', '{$item['category']}')")){
            echo $mysqli->errno . " - " . $mysqli->error;
        }
        
        
    }
        //TODO Проверяем наличие подкатегории
    $sub_category_name = $mysqli->query("SELECT * FROM oc_category_description WHERE name = '{$item['sub_category']}'");
    
    if($sub_category_name->num_rows == 0) {
        $parent_id = $category_id;
        $category_id++;

        if(!$mysqli->query("INSERT INTO oc_category(category_id, 
                                                    parent_id, 
                                                    top, 
                                                    `column`, 
                                                    status, 
                                                    date_added, 
                                                    date_modified) 
                                VALUE({$category_id}, $parent_id, 0, 0, 1, '$date', '$date')")) {
            die($mysqli->errno . " - " . $mysqli->error);

        }

        //TODO Записываем новую категорию и описание категории
        if(!$mysqli->query("INSERT INTO oc_category_description(category_id, language_id, name, description,
                                meta_title, meta_description, meta_keyword) 
                                VALUE('{$category_id}', 3, '{$item['sub_category']}', '{$item['sub_category']}', '{$item['sub_category']}', '{$item['sub_category']}', '{$item['sub_category']}')")){
            echo $mysqli->errno . " - " . $mysqli->error;
        }
    }

}




//var_dump($res);
///prnt($data[2]['category']);



function prnt($data) {
    echo "<pre>";
    if(is_object($data)) {
        var_dump($data);
    }else {
        print_r($data);
    }
    echo "</pre>";
}

function uploadFile($filename) {
    $uploaddir = __DIR__ . "/files";
    $uploadfile = $uploaddir . "/" . (int)microtime(true) . '.xlsx';

    if(move_uploaded_file($filename, $uploadfile)) {
        return  $uploadfile;
    }else {
        die('bad!');
    }
}

function getMaxId($mysqli, $field_id, $table) {
    //prnt();
    $res = $mysqli->query("SELECT MAX($field_id) as id FROM " . DBPREFIX . "$table");
    $row = $res->fetch_assoc();
    return $row['id'];

}
?>

<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Document</title>
</head>
<body>
    <form enctype="multipart/form-data" action="/" method="post">
        <input type="hidden" name="MAX_FILE_SIZE" value="4194304">
        <input type="file" name="file"><br>
        <input type="submit" value="обработать">
    </form>
</body>
</html>
