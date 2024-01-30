$(document).ready(function(){ // JQuery

    // 이벤트리스너

    $("#call").click(function() {

        // console.log("called");        
        //deleteCookie('tieout');

        if(getCookie('tieout')) {
            alert("과부하를 방지하기 위해 10초에 1번 클릭 가능합니다.");
        } else {
            setCookie('tieout', true, 10); 
            console.log("first click");
            $.ajax({
                type:"GET",
                url:"/tieout/php/exec.php",
                success:function(res){
                    alert(res);
                    },
                error:function(){
                    alert("Main tool 호출실패");
                }
            })

        };
        }
        );

    // // 쿠키설정 성공시 진행하는 함수

    // function doMain() {
    //     $.ajax({
    //         type:"GET",
    //         url:"/tieout/php/execMain.php",
    //         data: {
    //             'ExecMain':1
    //         },
    //         success:function(res){
    //             alert(res);
    //         },
    //         error:function(){
    //             alert("CONNECTION FAIL");
    //         }
    //     }
    //     )

    // }


    // 주기적으로 업데이트

    $(function() {
        timer = setInterval( function () {        
            $.ajax ({        
                type:"GET",
                data:{
                    // 'how':'direct'
                },
                url : "/tieout/php/readStatus.php",                
                success : function (res) {
                    console.log(res);
                    $(".divStatus").text(res);
                }        
            });        
            }, 3000);        
        });

    function setCookie(name, value, exp) {
        //var date = new Date();
        //date.setTime(date.getTime() + exp**1000); //초단위
        //document.cookie = name + '=' + value + ';expires=' + date.toUTCString() + ';path=/';
        document.cookie = name + '=' + value + ';max-age=' + exp + ';path=/';
    };

    function getCookie(name) {
        var value = document.cookie.match('(^|;) ?' + name + '=([^;]*)(;|$)');
        return value? value[2] : null;
    };

    function deleteCookie(name) {
        // document.cookie = name + '=; expires=Thu, 01 Jan 1970 00:00:01 GMT; domain=127.0.0.1;path=/;';
        document.cookie = name + '=; expires=Thu, 01 Jan 1970 00:00:01 GMT; path=/;';
    }
      
    

        
});
