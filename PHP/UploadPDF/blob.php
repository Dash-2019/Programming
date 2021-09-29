<!DOCTYPE html>
<html>
<head>
	<meta charset="utf-8">
	<title>Bolb file mysql </title>
</head>
<body>

    <?php
   
    $dbh = new PDO('mysql:dbname=testdb;host=localhost;charset=utf8','root','sakura');

    if(isset($_POST['btn'])){
        $name = $_FILES['myfile']['name'];
        $type = $_FILES['myfile']['type'];
        $data = file_get_contents($_FILES['myfile']['tmp_name']);
        $stmt = $dbh->prepare("insert into myblob values('',?,?,?)");
        $stmt->bindParam(1,$name);
        $stmt->bindParam(2,$type);
        $stmt->bindParam(3,$data);
        $stmt->execute();

    }
    ?>


    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="myfile"/>
        <button name="btn">Upload</button>
    </form>

<p></p>
<ol>
    <?php

    $stat = $dbh->prepare("select * from myblob");
    $stat->execute();
    while($row = $stat->fetch()){
        echo "<li><a target='_blank' href='view.php?id=".$row['id']."'>".$row['name']."</a></li>";
    }


    ?>



</body>
</html>