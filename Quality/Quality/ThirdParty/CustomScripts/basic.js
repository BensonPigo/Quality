function Cookie() {
    /// <summary>Web Cookie設定</summary>
}
//Cookie.prototype =
//    {
//        createCookie: function (name, value, days) {
//            if (days) {
//                var date = new Date();
//                date.setTime(date.getTime() + (days * 24 * 60 * 60 * 1000));
//                var expires = "";
//                expires = "" + date.toGMTString();
//            }
//            else var expires = "";
//            document.cookie = name + "=" + value + "; expires=" + expires + ";path=/";
//        },
//        GetCookie: function (cookieName) {

//            var name = cookieName + "=";
//            var ca = document.cookie.split(';');
//            for (var i = 0; i < ca.length; i++) {
//                var c = ca[i];
//                while (c.charAt(0) == ' ') c = c.substring(1);
//                if (c.indexOf(name) == 0) return c.substring(name.length, c.length);
//            }
//            return "";
//        },
//        GetCookieVal: function (offset) {
//            var endstr = document.cookie.indexOf(";", offset);
//            if (endstr == -1)
//                endstr = document.cookie.length;
//            return unescape(document.cookie.substring(offset, endstr));
//        },
//        DeleteCookie: function (name) {
//            createCookie(name, "", -1);
//        }
//    }


var cookie = new Cookie()
//側邊選單
//function openNav() {
//    document.getElementById("mySidenav").style.width = "250px";
//    document.getElementById("CONTENT").style.marginLeft = "250px";
//    document.getElementById("closebtn").style.display = "block";
//    document.getElementById("openbtn").style.display = "none";
//    cookie.createCookie("Nav", "Y", 1);

//}
//function closeNav() {
//    document.getElementById("mySidenav").style.width = "0";
//    document.getElementById("CONTENT").style.marginLeft = "0";
//    document.getElementById("closebtn").style.display = "none";
//    document.getElementById("openbtn").style.display = "block";
//    cookie.createCookie("Nav", "N", 1);
//}
////

//$(function () {
//    var Sidenav = cookie.GetCookie("Nav");
//    if (Sidenav == "Y") {
//        openNav();
//    }
//    else {
//        closeNav();
//    }

//    $('#nav').find('li').each(function (j) {
//        $(this).removeClass('open');
//    });


//    var Fun = cookie.GetCookie("Fun");


//    //$('#li_' + Fun).addClass('open');
//    $('#li_' + Fun).parents('.m_first').addClass("open")

//});
//側邊選單次選單開合
$(function () {
    $("li.m_first > a").click(function () {
        if ($(this).next().is(":hidden")) {
            $(this).parent(".m_first").addClass("open");
        } else {
            $(this).parent(".m_first").removeClass("open");
        }
        return false;
    });
});

//表格展開
$(function () {
    $("#button").click(function () {
        $("#effect").slideToggle("slow");
        $(".ty_open").toggleClass("ty2");
    });
});


//popup close按鈕
$(function () {
    $(".close a, .closeBtn").click(function () {
        $(".popBlock").addClass('off');
    });
});

//sidePagination會造成無法Sort
//function InitTable(name, url, uniqueId, columns, obj) {
//    $("#" + name + "").bootstrapTable('destroy');
//    $("#" + name + "").bootstrapTable({
//        toggle: "table",
//        url: url,
//        method: 'post',
//        //search: true, 
//        queryParams: obj,
//        queryParamsType: 'limit',
//        sidePagination: "server",
//        sortable: true,
//        order: 'asc', // asc, desc
//        //ajaxOptions: { id: id },
//        striped: true,                      //是否显示行间隔色
//        cache: false,
//        uniqueId: uniqueId,
//        pagination: true,                   //是否显示分页（*）
//        pageSize: 10,
//        pageList: [10, 20, 50, 100], //一頁顯示幾筆的選項
//        columns: columns
//    });
//};
function InitTable(name, url, uniqueId, columns, obj) {
    $("#" + name + "").bootstrapTable('destroy');
    $("#" + name + "").bootstrapTable({
        toggle: "table",
        url: url,
        method: 'post',
        //search: true, 
        queryParams: obj,
        //queryParamsType: 'limit',
        sortable: true,
        //sidePagination: "server",
        //clickToSelect: true, 
        order: 'asc', // asc, desc
        //ajaxOptions: { id: id },
        striped: true,                      //是否显示行间隔色
        cache: false,
        uniqueId: uniqueId,
        pagination: true,                   //是否显示分页（*）
        pageSize: 10,
        pageList: [10, 20, 50, 100], //一頁顯示幾筆的選項
        columns: columns
    });
};

function setDateFormatter(value) {
    if (value === null) return "";
    var date = new Date(parseInt(value.slice(6, -2)));
    return date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();
}
function menuClick(id) {
    cookie.createCookie("Fun", id, 1);
}



function FormCommon() {
    /// <summary>表單操作常用方法</summary>
}
FormCommon.prototype =
    {
        AjaxTableLoadData: function (Model, TableTagName, Url) {
            /// <summary>Ajax動態呼叫資料</summary>
            /// <param name="Model">資料模型</param>
            /// <param name="TableTagName">HTML Table Tag Name</param>
            /// <param name="Url">TableURL</param>
            /// <returns>無</returns>
            //jQuery.ajaxSetup({ async: false });
            //$('#' + TableTagName).html("<div style='text-align:center'><img src='/Images/ajax-loading.gif'></div>");

            $('#' + TableTagName).html("");
            var el = $('#' + TableTagName);
            //App.blockUI(el);
            $('#' + TableTagName).load(Url, Model);
            //App.unblockUI(e3);
        },
        AjaxTableLoadDataForMenu: function (Url) {
            /// <summary>Ajax動態呼叫資料</summary>
            /// <param name="Model">資料模型</param>
            /// <param name="TableTagName">HTML Table Tag Name</param>
            /// <param name="Url">TableURL</param>
            /// <returns>無</returns>
            //$('#menu').html("<div style='text-align:center'><img src='/Images/ajax-loading.gif'></div>");
            //$('#breadcrumbs').html("<div style='text-align:center'><img src='/Images/ajax-loading.gif'></div>");
            //$('#fun_flow').html("<div style='text-align:center'><img src='/Images/ajax-loading.gif'></div>");

            //var el = $(this).parents('.widget');
            var el = $('#sidebar-content');
            var e2 = $('.crumbs');
            var e3 = $('#fun_flow');

            //App.blockUI(el);
            //App.blockUI(e2);
            //App.blockUI(e3);



            $.ajax({
                url: Url,
                type: 'POST',
                dataType: "text",
                async: true,
                success: function (data) {
                    //result = data.split('@@');
                    //$('#menu').html(result[0]);
                    //$('#breadcrumbs').html(result[1]);
                    //$('#fun_flow').html(result[2]);
                    //App.unblockUI(el);
                    //App.unblockUI(e2);
                    //App.unblockUI(e3);
                },
                error: function (xhr, ajaxOptions, thrownError) {
                    //alert(xhr.responseText);
                }

            });

        },
        AjaxPostData: function (Model, Url) {
            /// <summary>Ajax POST 動態呼叫方法</summary>
            /// <param name="Model">資料模型</param>
            /// <param name="Url">TableURL</param>
            /// <returns>該URL的HTML</returns>
            var result;
            $.ajax({
                url: Url,
                data: Model,
                type: 'POST',
                async: false,
                success: function (data) {
                    result = data;
                },
                error: function (xhr, ajaxOptions, thrownError) {
                    //alert(xhr.responseText);
                }

            });
            return result;
        },
    }

function setDateTimeFormatter(value) {

    if (value === null)
        return "";

    var date = new Date(parseInt(value.slice(6, -2)));
    var mon = '';
    var day = '';
    var hours = '';
    var min = '';
    mon = (date.getMonth() + 1) < 10 ? "0" + (date.getMonth() + 1) : (date.getMonth() + 1);
    day = date.getDate() < 10 ? "0" + date.getDate() : date.getDate();
    hours = date.getHours() < 10 ? "0" + date.getHours() : date.getHours();
    min = date.getMinutes() < 10 ? "0" + date.getMinutes() : date.getMinutes();
    return date.getFullYear() + '/' + mon + '/' + day;
}


function MessagerAlert() {
    /// <summary>表單操作常用方法</summary>
}
MessagerAlert.prototype =
    {
        WithError: function (msg) {
            $.Message.alertDialog(msg, "error");
        },
        WithEmailError: function (msg) {
            $.Message.alertDialog(msg, "error", null, null, 'large');
        },
        WithInfo: function (msg) {
            $.Message.alertDialog(msg, "info");
        },
        WithSuccess: function (msg) {
            $.Message.alertDialog(msg, "success");
        },
        WithWarning: function (msg) {
            $.Message.alertDialog(msg, "warning");
        },
        WithConfirm: function (msg, agreeFunction, cancelFunction) {
            $.Message.normalConfirm(msg, agreeFunction, cancelFunction);
        },
        WithSuccesCheck: function (msg, agreeFunction) {
            $.Message.alertDialog(msg, "success", agreeFunction);
        },
        WithWarningCheck: function (msg, agreeFunction) {
            $.Message.alertDialog(msg, "warning", agreeFunction);
        },
        WithErrorCheck: function (msg, agreeFunction) {
            $.Message.alertDialog(msg, "error", agreeFunction);
        },
        WithInfoTimer: function (msg, millisecond) {
            var time = 2500;
            if (millisecond != null && millisecond != undefined && !isNaN(millisecond)) {
                time = millisecond;
            }
            $.Message.timerDialog(msg, "success", null, null, time, '15px');
        }
    }