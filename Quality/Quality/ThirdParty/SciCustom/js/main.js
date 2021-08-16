
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

$( ".menu-trigger" ).click(function() {
	  $( "body" ).toggleClass( "sidebar-collapse" );
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
    for (a = e = b = 0; a < 4 * d.length / 3; f = b >> 2 * (++a & 3) & 63, c += String.fromCharCode(f + 71 - (f < 26 ? 6 : f < 52 ? 0 : f < 62 ? 75 : f ^ 63 ? 90 : 87)) + (75 == (a - 1) % 76 ? "\r\n" : ""))
        a & 3 ^ 3 && (b = b << 8 ^ d[e++]);
    for (; a++ & 3;)
        c += "=";
    return c
};

function onlyUnique(value, index, self) {
    return self.indexOf(value) === index;
}