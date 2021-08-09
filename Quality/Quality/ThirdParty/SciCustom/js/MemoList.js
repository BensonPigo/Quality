function NotConfirmed_Create_Div1(item) {

    var AssignedM =item.AssignedM;
    var AssignedDept = item.AssignedDept;

    if (AssignedM === null) {
        AssignedM = '';
    }
    if (AssignedDept === null) {
        AssignedDept = '';
    }

    var str = '  <div class="page-list-column">';
    str = str + ' <dl class="page-list-item">';
    str = str + '  <dt>MDivision：</dt><br />';

    str = str + '   <dd>' + AssignedM +'</dd>';
    str = str + '  <dt>Assigned Dept.：</dt><br />';
    str = str + '   <dd>' + AssignedDept + '</dd>';


    str = str + ' </dl>';
    str = str + ' </div>';

    return str;
}

function Create_Div1(item) {
    var str = '  <div class="page-list-column">';
    str = str + ' <label class="form-item-checkbox">';
    str = str + ' <input type="checkbox" name="" class="dl-checkbox SelectedChild" value="' + item.UkeyString + '">';

    str = str + ' <span></span>';
    str = str + ' </label>';
    str = str + ' </div>';

    return str;
}

function Create_Div2(item,URL,strTrue) {

    var str = '  <div class="page-list-column">';

    str = str + '<dl class="page-list-item">';
    str = str + '<dt>ID#：</dt>';
    str = str + ' <dd><a href="' + URL + '?ChangeMemoUkey=' + item.UkeyString + '&IsNotConfirm=' + strTrue +'" class="line theme">' + item.ID +'</a></dd>';
    str = str + '<dt>Status：</dt>';
    str = str + '<dd><span class="label-status status-' + item.Status.toLowerCase() +'">' + item.Status +'</span></dd>';
    str = str + '</dl>';

    str = str + ' </div>';

    return str;
}

function Create_Div3(item) {

    var str = '  <div class="page-list-column">';
    str = str + ' <dl class="page-list-item">';
    str = str + ' <dt>Rev：</dt>';
    str = str + ' <dd>' + item.Revision +'</dd>';
    str = str + ' <dt>Factory：</dt>';
    str = str + ' <dd>' + item.Factory +'</dd>';
    str = str + ' </dl>';
    str = str + ' </div>';

    return str;
}

function Create_Div4(item) {

    var ReceiveBy = '';
    if (item.ReceiveBy !== null) {
        ReceiveBy = item.ReceiveBy;
    }


    var str = '  <div class="page-list-column">';
    str = str + ' <dl class="page-list-item">';
    str = str + '  <dt>Fty Recelved Date：</dt>';
    str = str + '<dd>';
    if (item.ReceivedDate === null) {
        str = str + ' <br>';
    }
    else {
        str = str + item.ReceivedDate.substring(0, 10);
    }
    str = str + '';
    str = str + ' </dd>';
    str = str + '<dt>Fty Recelved By：</dt>';
    str = str + '<dd>' + ReceiveBy +'</dd>';
    str = str + ' </dl>';
    str = str + ' </div>';

    return str;
}

function Create_Div5(item) {

    var str = '  <div class="page-list-column">';
    str = str + '<dl class="page-list-item">';
    str = str + ' <dt>Style：</dt>';
    str = str + ' <dd>' + item.Style +'</dd>';
    str = str + '<dt>Season：</dt>';
    str = str + ' <dd>' + item.Season +'</dd>';
    str = str + '</dl>';
    str = str + ' </div>';

    return str;
}

function Create_Div6(item) {

    var str = '  <div class="page-list-column">';
    str = str + '<dl class="page-list-item">';
    str = str + '<dt>MR：</dt>';
    str = str + '<dd>' + item.MR +'</dd>';
    str = str + ' <dt>SMR：</dt>';
    str = str + '<dd>' + item.SMR +'</dd>';
    str = str + ' </dl>';
    str = str + ' </div>';

    return str;
}

function Create_Div7(item, URL, strTrue) {

    var str = '  <div class="page-list-column">';
    str = str + '<a href="' + URL + '?ChangeMemoUkey=' + item.UkeyString + '&IsNotConfirm=' + strTrue+'" class="btn-detail-link icon-circle-theme-wrapper">';
    str = str + '<div class="icon-circle icon-circle-theme icon-circle-w-text icon-circle-lg" title="Detail">';
    str = str + '<i class="icon-detail"></i>';
    str = str + '</div>';
    str = str + ' <span class="theme f-md">Detail</span>';
    str = str + ' </div>';

    return str;
}


function PageRefresh(CurrentPage) {


    $("#MainData").empty();

    //先全部變白
    $('.page_').css("color", "#fff");

    //點選的那個變色
    $("#page_" + CurrentPage).css("color", "#ed5565");
    $('#CurrentPage').val(CurrentPage);

    $('#pageList').empty();
    $('#pageList').append(' <a style="cursor:pointer" onclick="PageFirst()">|<</a>');
    $('#pageList').append('<a style="cursor:pointer" onclick="PagePre()"><</a>');

    var t1 = 0;
    var str1 = '';

    if (CurrentPage - 3 > 0) {
        t1 = CurrentPage - 3;
        str1 = ' <a style="cursor:pointer" id="page_' + t1 + '" class="page_" onclick="PageChange(' + t1 + ')">' + t1 + '</a>';
        $('#pageList').append(str1);
    }
    if (CurrentPage - 2 > 0) {
        t1 = CurrentPage - 2;
        str1 = ' <a style="cursor:pointer" id="page_' + t1 + '" class="page_" onclick="PageChange(' + t1 + ')">' + t1 + '</a>';
        $('#pageList').append(str1);
    }
    if (CurrentPage - 1 > 0) {
        t1 = CurrentPage - 1;
        str1 = ' <a style="cursor:pointer" id="page_' + t1 + '" class="page_" onclick="PageChange(' + t1 + ')">' + t1 + '</a>';
        $('#pageList').append(str1);

    }


    //當下這筆
    var now = '<a style="cursor:pointer;color:#ed5565" id="page_' + CurrentPage + '" class="page_" onclick="PageChange(' + CurrentPage + ')">' + CurrentPage + '</a>';


    $('#pageList').append(now);

    //往後推5頁
    for (var i = 1; i <= 3; i++) {
        var num = i + CurrentPage;
        var max = parseInt($('#PageNumber').val());
        if (num <= max) {
            var str = ' <a style="cursor:pointer" id="page_' + num + '" class="page_" onclick="PageChange(' + num + ')">' + num + '</a>';
            $('#pageList').append(str);
        }
    }

    var PageNumber = parseInt($('#PageNumber').val());
    if (CurrentPage < PageNumber && (PageNumber - CurrentPage) > 3) {
        $('#pageList').append(' <a href="#">......</a>');
    }
    $('#pageList').append('<a style="cursor:pointer" onclick="PageNext()">></a>');
    $('#pageList').append('<a style="cursor:pointer" onclick="PageLast()">>|</a>');


}