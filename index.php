<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <link rel="stylesheet" href="css/style.css">
    <script src="https://code.jquery.com/jquery-3.7.1.min.js" integrity="sha256-/JqT3SQfawRcv/BIHPThkBvs0OEvtFFmqPF/lYI/Cxo=" crossorigin="anonymous"></script>
    <script src="./js/script.js" type="text/javascript"></script>
</head>

<body>
<div class="divA">
    Mini Spotlight - Tie-out<br>
    v0.0.1
</div>    

<div class="divB">
    FTP를 통해 서버상 전후 DSD(HTML) 파일 업로드 후 작동
    <button id="call">Tie-out 툴 작동</button>
</div>

<div class="divC">
    <div class="divD">
    작업상태(3초마다갱신)>> <br />
    DONE으로 변경되면 산출물 다운로드 <br />
    </div>

    <div class="divStatus">
    <?php include("./php/readStatus.php") ?>
    </div>

</div>    

</body>
</html>

