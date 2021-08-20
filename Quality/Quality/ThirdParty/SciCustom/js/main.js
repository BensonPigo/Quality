
/*menu-trigger*/
(function ($) {
        /* Store sidebar state */
        $('.menu-trigger').click(function(event) {
            event.preventDefault();
            if (Boolean(localStorage.getItem('sidebar-toggle-collapsed'))) {
                localStorage.setItem('sidebar-toggle-collapsed', '');
             } else {
                localStorage.setItem('sidebar-toggle-collapsed', '1');
             }
         });
    })(jQuery);
    /* Recover sidebar state */
     (function () {
        if (Boolean(localStorage.getItem('sidebar-toggle-collapsed'))) {
            var body = document.getElementsByTagName('body')[0];
            body.className = body.className + ' sidebar-collapse';
        }
    })();

$(".menu-trigger").click(function() {
  $("body").toggleClass("sidebar-collapse");
});


//
 $('.sidebar-menu > li > h6').click(function(){
      $(this).next('.sidebar-submenu').slideToggle();
    });

/*search-trigger*/
$( ".search-trigger" ).click(function() {
	  $( ".search-container" ).toggleClass( "open" );
	});
//表格展開
$( ".expend , .btn-detail-edit , .btn-detail-save " ).click(function() {
	  $(this).parent().parent().parent().toggleClass( "plus" );
	});

$(".RootLi").on('click', function (event) {
    $(".RootLi").removeClass("current");
    $(this).addClass("current");
    if ($("body").hasClass("sidebar-collapse")) {
        $(".menu-trigger").click();
    } 
});

//新增區塊
        window.onload = function(){
		    $('.add-edit-list:not(:has(li))').parent().hide(); // Will not alter layout
		    // $('.nested:not(:has(li))').css("display", "none"); // Will alter layout
		    }
		//$(".btn-add-incharge").click(function () {
  //             $(this).parent().parent().parent().find(".inform-dept-list-expend-inner").show();
		//	   $(this).parent().parent().parent().find(".add-list-incarge").append('<li class="form-item"><label>ID：</label><input type="text" name="id" value=""><button class="site-btn btn-red btn-md icon-save btn-save btn-w-icon">Save</button></li>');
  //       });
		//$(".btn-add-inform").click(function () {
		//	   $(this).parent().parent().parent().find(".inform-dept-list-expend-inner").show();
		//	   $(this).parent().parent().parent().find(".add-list-inform").append('<li class="form-item"><label>ID：</label><input type="text" name="id" value=""><button class="site-btn btn-red btn-md icon-save btn-save btn-w-icon">Save</button></li>');
  //       });
		//
		
// 對話框模組
$ (document).ready(function(){

	

// var $target = $('.inform-dept-list-expend-inner');
// if ($target.find("ul:not(:contains(li))") ) {
//     $target.parent().parent().show();
// }



	$( ".site-dialog" ).dialog({
		   autoOpen: false, 
		   modal: true,
		   closeText: ' ',
		   width: 680,
		   minHeight: 200,
		   dialogClass: "site-box",
		   open: function(event, ui) 
				    { 
               $('.ui-widget-overlay').bind('click', function ()
               {
                             location.reload();//強制Refresh頁面 Benson
				            $(".ui-dialog-content").dialog('close'); 
				        }); 
				        $(this).parent().siblings().find(".ui-dialog-content").dialog('close'); 

				    },
			close: function() {
		        },
		});
	    //
	    $( ".to-done ,.to-remove" ).dialog({
	    	width: 520,
		      buttons: {
                  "1": {
                      text: 'OK', click: function () {
                          location.reload();//強制Refresh頁面 Benson
                          $(this).dialog("close");
                      }, "class": "site-btn btn-red btn-lg"
                  },
		      }
		    });
		//
		$(".close").click(function () {
		  $(".site-dialog").dialog('close');
		    });
		//Benson 
        //$(".btn-detail-remove").click(function(e) {
        //   $(".to-remove").dialog( "open" );
        //});

});
//Benson 
// //移除項目
// $('.btn-detail-remove').on('click', function(){
//     $(this).parent().parent().remove();
// });

//tooltip
    $( document ).tooltip({
      position: {
        my: "center bottom-5",
        at: "center top",
        using: function( position, feedback ) {
          $( this ).css( position );
          $( "<div>" )
            .addClass( "arrow" )
            .addClass( feedback.vertical )
            .addClass( feedback.horizontal )
            .appendTo( this );
        }
      }
    });
    //表格全選
		// $(document).ready(function () {
		//     $("input[type='checkbox']").prop("checked", false);
		// });
		$('.selecctall').click(function (event) {
	        if (this.checked) {
	            $('.dl-checkbox').each(function () {  
	                $(this).prop('checked', true); 
	            });
	        } else {
	            $('.dl-checkbox').each(function () {  
	                $(this).prop('checked', false);             
	            });
	           
	        }
	    });
	    // $('.dl-checkbox').click(function (event) {
	    //     if (this.checked) {
	    //         $(this).parent().parent().parent().addClass( "select" );
	           
	    //     } else {
	    //         $(this).parent().parent().parent().removeClass( "select" );
	    //     }
	    // });

  	//搜尋表單
  $( function() {
    var dateFormat = "yy/mm/dd",
        SendDate_s = $("#SendDate_s").datepicker({
            defaultDate: "+1w",
            dateFormat: "yy/mm/dd", //自行定義格式//Benson 
            numberOfMonths: 1
        })
            .on("change", function () {
                SendDate_e.datepicker("option", "minDate", getDate(this));
            }),
        SendDate_e = $( "#SendDate_e" ).datepicker({
                                    defaultDate: "+1w",
                                    dateFormat: "yy/mm/dd",
                                        changeMonth: true,
                                        numberOfMonths: 1
                                        })
                                        .on( "change", function() {
                                            SendDate_s.datepicker( "option", "maxDate", getDate( this ) );
            }),
        ReceiveDate = $("#ReceiveDate").datepicker({
            defaultDate: "+1w",
            dateFormat: "yy/mm/dd", //自行定義格式//Benson
            numberOfMonths: 1
        });
 
    function getDate( element ) {
      var date;
      try {
        date = $.datepicker.parseDate( dateFormat, element.value );
      } catch( error ) {
        date = null;
      }
 
      return date;
    }

  } );


function padLeft(str, lenght) {
    if (str.length >= lenght)
        return str;
    else
        return padLeft("0" + str, lenght);
}

function parseIntEX(x, base = 10) {
    const parsed = parseInt(x, base);
    if (isNaN(parsed)) { return 0; }
    return parsed;
}

function ByteArrayToBase64(d, a, e, b, c, f) {
    c = "";
    for (a = e = b = 0; a < 4 * d.length / 3; f = b >> 2 * (++a & 3) & 63, c += String.fromCharCode(f + 71 - (f < 26 ? 6 : f < 52 ? 0 : f < 62 ? 75 : f ^ 63 ? 90 : 87)) + (75 == (a - 1) % 76 ? "" : ""))
        a & 3 ^ 3 && (b = b << 8 ^ d[e++]);
    for (; a++ & 3;)
        c += "=";
    return c
};

function onlyUnique(value, index, self) {
    return self.indexOf(value) === index;
}

function gcd(a, b) {
    if (b == 0) {
        return a;
    }
    else {
        return gcd(b, a % b);
    }
}


function FractionDiff(standard, val) {
    var si = 0, i = 0
    var sm, m
    var sd, d
    var finVal = "";
    if (standard.indexOf(' ') > -1) {
        si = standard.substr(0, standard.indexOf(' '));
    }
    else if (standard.indexOf(" ") < 0 && standard.indexOf('/') < 0) {
        si = standard;
    }

    if (val.indexOf(" ") > -1) {
        i = val.substr(0, val.indexOf(' '));
    }
    else if (val.indexOf(" ") < 0 && val.indexOf('/') < 0) {
        i = val;
    }

    if (standard.indexOf('/') > -1) {
        sd = standard.substr(standard.indexOf(' ') + 1, standard.indexOf('/') - 1 - standard.indexOf(' '));
        sm = standard.substr(standard.indexOf('/') + 1, standard.length - 1 - standard.indexOf('/'));
    }
    else {
        sd = 1;
        sm = 1;
        si = si - 1;
    }

    if (val.indexOf('/') > -1) {
        d = val.substr(val.indexOf(' ') + 1, val.indexOf('/') - 1 - val.indexOf(' '));
        m = val.substr(val.indexOf('/') + 1, val.length - 1 - val.indexOf('/'));
    }
    else {
        d = 1;
        m = 1;
        i = i - 1;
    }

    var fi = si - i;
    if (fi <= 0) {
        var fm = sm * m;
        var sd1 = sd * m;
        var d1 = d * sm + Math.abs(fi) * m * sm;
        var fd = sd1 - d1;
        var divisor = gcd(Math.abs(fd), fm);
        fi = 0;
        fm = fm / divisor;
        fd = fd / divisor;

        if (Math.abs(fd) >= fm) {
            fi = Math.floor(Math.abs(fd) / fm);
            if (fd < 0) fi = fi * -1;
            fd = Math.abs(fd) - (Math.abs(fi) * fm);

            if (fd == 0) {
                finVal = fi;
            }
            else {
                finVal = fi + ' ' + fd + '/' + fm;
            }
        }
        else {
            if (fd == 0) {
                finVal = 0;
            }
            else {
                finVal = fd + '/' + fm;
            }
        }

        return finVal;
    }
    else {
        var fm = sm * m;
        var sd1 = sd * m;
        var d1 = d * sm;
        var fd = sd1 - d1;
        if (fd < 0) {
            fd = fd + fm;
            fi = fi - 1;
        }
        var divisor = gcd(Math.abs(fd), fm);
        fm = fm / divisor;
        fd = fd / divisor;

        if (fi <= 0) {
            if (fd == 0) {
                finVal = 0;
            }
            else {
                finVal = fd + '/' + fm;
            }
        }
        else {
            if (fd == 0) {
                finVal = fi;
            }
            else {
                finVal = fi + ' ' + fd + '/' + fm;
            }
        }

        return finVal;
    }
}

function FractionAdd(standard, val) {
    var si = 0, i = 0
    var sm, m
    var sd, d
    var finVal = "";
    if (standard.indexOf(' ') > -1) {
        si = standard.substr(0, standard.indexOf(' '));
    }
    else if (standard.indexOf(" ") < 0 && standard.indexOf('/') < 0) {
        si = standard;
    }

    if (val.indexOf(" ") > -1) {
        i = val.substr(0, val.indexOf(' '));
    }
    else if (val.indexOf(" ") < 0 && val.indexOf('/') < 0) {
        i = val;
    }

    if (standard.indexOf('/') > -1) {
        sd = standard.substr(standard.indexOf(' ') + 1, standard.indexOf('/') - 1 - standard.indexOf(' '));
        sm = standard.substr(standard.indexOf('/') + 1, standard.length - 1 - standard.indexOf('/'));
    }
    else {
        sd = 1;
        sm = 1;
        si = si - 1;
    }

    if (val.indexOf('/') > -1) {
        d = val.substr(val.indexOf(' ') + 1, val.indexOf('/') - 1 - val.indexOf(' '));
        m = val.substr(val.indexOf('/') + 1, val.length - 1 - val.indexOf('/'));
    }
    else {
        d = 1;
        m = 1;
        i = i - 1;
    }

    var fi = accAdd(si, i);
    var fm = accMul(sm, m);
    var sd1 = accMul(sd, m);
    var d1 = accMul(d, sm);
    fd = accAdd(sd1, d1);

    if (Math.abs(fd) >= fm) {
        fi = accAdd(fi, Math.floor(Math.abs(fd) / fm));
        if (fd < 0) fi = accMul(fi, -1);
        fd = accSubtr(Math.abs(fd), accMul(Math.floor(Math.abs(fd) / fm), fm));
    }

    if (fd == 0) {
        finVal = fi;
    }
    else {
        finVal = fi + ' ' + fd + '/' + fm;
    }

    return finVal;
}

//除法
function accDiv(arg1, arg2) {
    var t1 = 0, t2 = 0, r1, r2;
    try {
        t1 = arg1.toString().split(".")[1].length;
    } catch (e) { }
    try {
        t2 = arg2.toString().split(".")[1].length;
    } catch (e) { }
    with (Math) {
        r1 = Number(arg1.toString().replace(".", ""));
        r2 = Number(arg2.toString().replace(".", ""));
        return (r1 / r2) * pow(10, t2 - t1);
    }
}

//乘法
function accMul(arg1, arg2) {
    var m = 0, s1 = arg1.toString(), s2 = arg2.toString();
    try {
        m += s1.split(".")[1].length;
    } catch (e) { }
    try {
        m += s2.split(".")[1].length;
    } catch (e) { }
    return Number(s1.replace(".", "")) * Number(s2.replace(".", "")) / Math.pow(10, m);
}

//加法
function accAdd(arg1, arg2) {
    var r1, r2, m;
    try { r1 = arg1.toString().split(".")[1].length } catch (e) { r1 = 0 }
    try { r2 = arg2.toString().split(".")[1].length } catch (e) { r2 = 0 }
    m = Math.pow(10, Math.max(r1, r2));
    return (arg1 * m + arg2 * m) / m;
}

//減法
function accSubtr(arg1, arg2) {
    var r1, r2, m, n;
    try {
        r1 = arg1.toString().split(".")[1].length;
    } catch (e) { r1 = 0 }
    try {
        r2 = arg2.toString().split(".")[1].length;
    } catch (e) { r2 = 0 }
    m = Math.pow(10, Math.max(r1, r2));
    n = (r1 >= r2) ? r1 : r2;
    return ((arg1 * m - arg2 * m) / m).toFixed(n);
}

// Json To HtmlTable
function generateTable(jArray) {
    let tbody = document.createElement('tbody');
    let thead = document.createElement('thead');
    let table = document.createElement('table');

    // 將所有資料列的資料轉成tbody
    jArray.forEach(row => {
        let tr = document.createElement('tr');

        Object.keys(row).forEach(tdName => {
            let td = document.createElement('td');
            td.textContent = row[tdName];

            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    // 將所有資料列的欄位轉成thead
    let headerTr = document.createElement('tr')

    Object.keys(jArray[0]).forEach(header => {
        let th = document.createElement('th')
        th.textContent = header

        headerTr.appendChild(th)
    });

    // 新增thead到table上
    thead.appendChild(headerTr);
    table.appendChild(thead);

    return table;
}


function decodeHtml(html) {
    var txt = document.createElement("textarea");
    txt.innerHTML = html;
    return txt.value;
}