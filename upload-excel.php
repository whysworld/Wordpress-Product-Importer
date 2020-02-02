<?php
$target_dir = "upload/";
// $target_file = $target_dir . basename($_FILES["file"]["name"]);
$target_file = $target_dir . $_FILES["file"]["name"];
if (file_exists($target_file))
	unlink($target_file);
if (move_uploaded_file($_FILES["file"]["tmp_name"], $target_file)) {
   $status = 1;
   $result = array(
   	"status" => 1,
   	"fileName" => $_FILES["file"]["name"]
   );
   echo json_encode($result);
}