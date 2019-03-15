jQuery.noConflict();

(function($) {

    "use strict";
    kintone.events.on("app.record.index.show", function(event) {

        //https://fdossena.com/?p=html5cool/buttons/i.frag

        var appId = kintone.app.getId();
        //alert(event.viewId); 
 
        if (event.viewId !== 5525025) {
            return;
        } 
 
        //var canvas = $("<canvas></canvas>", {id: "canvas", style: "position:fixed; top:0; left:0; width:100%; height:100%;"} ) 
        var button_space = kintone.app.getHeaderSpaceElement(); 
        var headerDiv = $("<div></div>", {class: "header-contents"});
        var bulkButton = $("<button></button>", {class: "animated bounce infinite click example_c"} )//https://www.magnet4blogging.net/button-styles/ 
            .html("批量審批").css({'color': 'rgb(5, 222, 30)', 'font-family': 'STXingkai'}).click(function() {
                if (event.records.length > 0) {
                    window.swal({
                        title: "確定要批准此列表的所有申請嗎？",
                        text: "將批准所有申請",
                        type: "warning",
                        showCancelButton: true,
                        confirmButtonColor: "#DD6B55",
                        confirmButtonText: "確定",
                        cancelButtonText: "還是算了",
                        closeOnConfirm: false},

                    function() { 
                        var records = [];
                        for (var i = 0; i < event.records.length; i++) {
                            var obj = {};
                            obj["id"] = event.records[i].$id.value; 
                            obj["action"] = "審批";  
                            //obj["assignee"] = " "; 
                            records.push(obj); 
                        }

                        var requestObj = {
                            "app": appId,
                            "records": records
                        };


                        kintone.api("/k/v1/records/status", "PUT", requestObj, function() {
                            window.swal({title: "批量審批成功",
                                    text: "申請已完成審批",
                                    type: "success"}, function() {
                                            location.reload();
                                        });
                        },
                        function(error) {
                            // error
                            console.log(error);
                        });
                    });
                }
            });

            $("bulkButton").html('批量審批').addClass("myClass");

        //headerDiv.append(canvas);
        headerDiv.append(bulkButton);
        headerDiv.appendTo(button_space); 
    });
})(jQuery);