<?php
  $myfile = fopen("cal.txt", "a") or die("Unable to open file!");
  $subject = $_POST["subject"];
  $start = $_POST["start"];
  $end = $_POST["end"];
  $location = $_POST["location"];
  fwrite($myfile, $subject);
  fwrite($myfile, $start);
  fwrite($myfile, $end);
  fwrite($myfile, $location);
  fclose($myfile);
?>
