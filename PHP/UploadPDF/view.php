<?php
$dbh = new PDO('mysql:dbname=testdb;host=localhost;charset=utf8','root','sakura');
$id = isset($_GET['id'])? $_GET['id'] : "";

$stat = $dbh->prepare("select * from myblob where id=?");
$stat->bindParam(1,$id);

$stat->execute();
$row = $stat->fetch();
header('Content-Type:'.$row['mime']);
echo $row['data'];

?>