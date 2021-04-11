<?php

if(!isset($_GET['sutun1']) ||!isset($_GET['sutun2']) ){
    Echo "<pre>
    GET Deşişkenleri<br>
     sutun1= Hangi Sutundaki veriyi okuyacaksanız Onun Harf Kodu<br>
     sutun2= Hangi Sutundaki veriyi okuyacaksanız Onun Harf Kodu<br>
     cikti=Çıktı Formatı json veya txt olabilir<br>
     kaydet= veriyi kaydetmek için 1  varsayılan değeri False yani 0 dır<br>
     isim=kaydedilecek dosya dosya adi varsayilan file dir<br>
    Örnek Kullanım:<a href='index.php?sutun1=A&sutun2=B&cikti=json&isim=dosya_Adi&kaydet=0'>index.php?sutun1={A}&sutun2={B}&cikti={json}&isim={dosya_Adi}&kaydet={0}</a></pre>";
}else{
require_once "Classes/PHPExcel.php";
        $tmpfname = "tablo.xlsx";
        $excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
        $excelObj = $excelReader->load($tmpfname);
        $worksheet = $excelObj->getSheet(0);//
        $lastRow = $worksheet->getHighestRow();

        $data = [];
        for ($row = 2; $row <= $lastRow; $row++) {
             $data[] = [
                'KOD' => $worksheet->getCell($_GET['sutun1'].$row)->getValue(),
                'STRING' => $worksheet->getCell($_GET['sutun2'].$row)->getValue()
             ];
        }
$jsonObj= json_encode($data,JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);

$uzanti="";

        if(isset($_GET['cikti'])){
      if($_GET['cikti']=="json"){
$uzanti=".json";
if(isset($_GET['kaydet'])){
 if($_GET['kaydet']==1){
      if(isset($_GET['isim'])){
   $isim=$_GET['isim'];
      }else{
            $isim="file";
          
      }
header("Content-Disposition: attachment; filename=".$isim.$uzanti);

}
}
header('Content-Type: application/json');
echo $jsonObj;
}else{
 $uzanti=".txt";
 if(isset($_GET['kaydet'])){
 if($_GET['kaydet']==1){
if(isset($_GET['isim'])){
   $isim=$_GET['isim'];
      }else{
 $isim="file";
      }
header("Content-Disposition: attachment; filename=".$isim.$uzanti);
}
}
 header("Content-type: text/plain");
  // echo "txt formata çevir"; 
  $data = json_decode($jsonObj, TRUE);

   foreach ($data as $itinerary) {
    echo $itinerary["KOD"].'           "'.$itinerary["STRING"].'"'.PHP_EOL;
}
   
}
}
//Kaydetme
if(isset($_GET['kaydet'])){
 if($_GET['kaydet']==1){
      if(isset($_GET['isim'])){
   $isim=$_GET['isim'];
      }else{
            $isim="file";
          
      }
header("Content-Disposition: attachment; filename=".$isim.$uzanti);
}
}

}// Ana Koşul
?>