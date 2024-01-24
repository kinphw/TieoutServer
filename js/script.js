$(document).ready(function(){ // JQuery

    //테스트 함수
    // function msg() {
    //     alert("Hello");        
    // };

    // 이벤트리스너

    $("#call").click(function() {
        $.ajax({
            type:"GET",
            url:"/php/exec.php",
            success:function(res){
                // alert(res);
                if(res=="COOKIE_SUCCESS") {
                    console.log("COOKIE SET");
                    doMain();
            }},
            error:function(){
                alert("CONNECTION FAIL");
            }
        }

        )
    })

    // 쿠키설정 성공시 진행하는 함수

    function doMain() {
        $.ajax({
            type:"GET",
            url:"/php/execMain.php",
            data: {
                'ExecMain':1
            },
            success:function(res){
                alert(res);
            },
            error:function(){
                alert("CONNECTION FAIL");
            }
        }
        )

    }


    // 주기적으로 업데이트

    $(function() {
        timer = setInterval( function () {        
            $.ajax ({        
                type:"GET",
                data:{
                    'how':'direct'
                },
                url : "/php/readStatus.php",                
                success : function (res) {
                    console.log(res);
                    $(".divStatus").text(res);
                }        
            });        
            }, 3000);        
        });
});
