(function($){
	
	var css = {
				border: 'none',
				padding: '5px',
				backgroundColor: '#000',
				'-webkit-border-radius': '10px',
				'-moz-border-radius': '10px',
				opacity: .5,
				color: '#fff',
			};
	
	jQuery.extend({
		// 開啓遮罩
		showBlock : function(msg, m_css) { 
			msg = msg ? msg : '<h5>Loading...</h5>';
			jQuery.blockUI({
				message: msg,
				baseZ: 99999,
				css: m_css ? m_css : css
			})
		},
		// 關閉遮罩
		unBlock : function(){
			jQuery.unblockUI();
		},
		// Ajax 自動彈出遮罩
		autoBlock: function (msg, m_css) {
			msg = msg ? msg : '<h5>Loading...</h5>';
			jQuery(document).ajaxStart(jQuery.blockUI({
				message: msg,
				baseZ: 99999,
				css: m_css ? m_css : css
			})).ajaxStop(jQuery.unblockUI());
		},
	    // 驗證大於等於0的整数
		isIntMinZero: function (number) {
		    var reg = /^([1-9]\d*|[0]{1,1})$/;
		    return reg.test(number);
		},
		// 民國年轉西元年
		toWestDate: function (dateStr, symbol) {
			var symbol = symbol ? symbol : "/";
			if (dateStr) {
				var splitIndex = dateStr.indexOf(symbol);
				if (splitIndex != -1) {
					var year = dateStr.substring(0, splitIndex);
					var newDateStr = parseInt(year) + 1911 + dateStr.substring(splitIndex, dateStr.length);
					return newDateStr;
				}
			}
			return "";
		}
	});
})(jQuery);

function formatMoney() {

    // 金額格式化︰三位一撇
    $.each($('.accountFormat'), function (key, val) {
        $(this).setThousandsSymbolToNumber();
    });

    $('.accountFormat').blur(function () { formatNumber(this); $(this).setThousandsSymbolToNumber(); });
    $('.accountFormat').focus(function () { $(this).deleteThousandsSymbol(); });

    $('.accountFormat').keydown(function (e) {
        
        var key;
        if (window.event) {
            key = e.keyCode;
        } else if (e.which) {
            key = e.which;
        } else {
            return true;
        }

        // 只能有一個小數點
        if ($(this).val().indexOf(".") > 0 && 110 == key) {
            return false;
        }

        var intlen = $(this).attr("intLength");
        if (intlen != "" && intlen != undefined) {
            
            // 整數部份的最大長度限制
            var nums = $(this).val().split(".");
            if (nums[0].length == intlen
            && $(this).val().indexOf(".") < 0
            && ((key >= 48 && key <= 57) || (key >= 96 && key <= 105))) {
                return false;
            }

            if (nums.length == 2) {

                // 整數及小數部份都為最大長度後，不允許輸入
                var dotlen = $(this).attr("dotLength");

                if (nums[0].length == intlen && $(this).val().indexOf(".") > 0 && nums[1].length == dotlen
                && ((key >= 48 && key <= 57) || (key >= 96 && key <= 105))) {
                    return false;
                }
            }
        }

        // 開始數字為0且第二位不為小數點時，清除開始的0
        if ($(this).val().charAt(0) == '0' && $(this).val().charAt(1) != '.') {
            $(this).val(parseFloat($(this).val()));
        }

        if ((key >= 48 && key <= 57)
            || (key >= 96 && key <= 105) //數字鍵盤
            || 8 == key || 46 == key || 37 == key || 39 == key //8:backspace 46:delete 37:左 39:右 (倒退鍵、刪除鍵、左、右鍵也允許作用)
            || 190 == key || 110 == key //小數點
            || (189 == key) //減號
            ) {
            return true;
        } else {
            return false;
        }
    });
}

function formatDouble() {

    $('.doubleFormat').blur(function () { formatNumber(this); });

    $('.doubleFormat').keydown(function (e) {

        var key;
        if (window.event) {
            key = e.keyCode;
        } else if (e.which) {
            key = e.which;
        } else {
            return true;
        }

        // 只能有一個小數點
        if ($(this).val().indexOf(".") > 0 && 110 == key) {
            return false;
        }

        var intlen = $(this).attr("intLength");
        if (intlen != "" && intlen != undefined) {

            // 整數部份的最大長度限制
            var nums = $(this).val().split(".");
            if (nums[0].length == intlen
            && $(this).val().indexOf(".") < 0
            && ((key >= 48 && key <= 57) || (key >= 96 && key <= 105))) {
                return false;
            }

            if (nums.length == 2) {

                // 整數及小數部份都為最大長度後，不允許輸入
                var dotlen = $(this).attr("dotLength");

                if (nums[0].length == intlen && $(this).val().indexOf(".") > 0 && nums[1].length == dotlen
                && ((key >= 48 && key <= 57) || (key >= 96 && key <= 105))) {
                    return false;
                }
            }
        }

        // 開始數字為0且第二位不為小數點時，清除開始的0
        if ($(this).val().charAt(0) == '0' && $(this).val().charAt(1) != '.') {
            $(this).val(parseFloat($(this).val()));
        }

        if ((key >= 48 && key <= 57)
            || (key >= 96 && key <= 105) //數字鍵盤
            || 8 == key || 46 == key || 37 == key || 39 == key //8:backspace 46:delete 37:左 39:右 (倒退鍵、刪除鍵、左、右鍵也允許作用)
            || 190 == key || 110 == key //小數點
            || (189 == key) //減號
            ) {
            return true;
        } else {
            return false;
        }
    });
}

function formatNumber(obj, dot) {
    var num = $(obj).val();

    var dot = $(obj).attr("dotLength");

    if (dot == "" || dot == undefined) {
        dot = 2;
    }

    if (num != "") {

        $(obj).val(parseFloat(num).toFixed(dot));
    }
}

// 民過年日曆插件初始化
function showPicker(obj, param) {

	var settings = {
		el: obj,
		dateFmt: 'yyy/MM/dd',
		skin: 'twoer',
		errDealMode: 1,
		onpicked: function (dp) {
			jQuery(dp.el).trigger('change');
		} 
	};

	if (param && !jQuery.isEmptyObject(param)) {
		jQuery.extend(settings, param);
	};

	WdatePicker(settings);
}

//可在Javascript中使用如同C#中的string.format
//使用方式 : var fullName = String.format('Hello. My name is {0} {1}.', 'FirstName', 'LastName');
String.format = function () {
	var s = arguments[0];
	if (s == null) return "";
	for (var i = 0; i < arguments.length - 1; i++) {
		var reg = getStringFormatPlaceHolderRegEx(i);
		s = s.replace(reg, (arguments[i + 1] == null ? "" : arguments[i + 1]));
	}
	return cleanStringFormatResult(s);
}
//可在Javascript中使用如同C#中的string.format (對jQuery String的擴充方法)
//使用方式 : var fullName = 'Hello. My name is {0} {1}.'.format('FirstName', 'LastName');
String.prototype.format = function () {
	var txt = this.toString();
	for (var i = 0; i < arguments.length; i++) {
		var exp = getStringFormatPlaceHolderRegEx(i);
		txt = txt.replace(exp, (arguments[i] == null ? "" : arguments[i]));
	}
	return cleanStringFormatResult(txt);
}
//讓輸入的字串可以包含{}
function getStringFormatPlaceHolderRegEx(placeHolderIndex) {
	return new RegExp('({)?\\{' + placeHolderIndex + '\\}(?!})', 'gm')
}

//當format格式有多餘的position時，就不會將多餘的position輸出
//ex:
// var fullName = 'Hello. My name is {0} {1} {2}.'.format('firstName', 'lastName');
// 輸出的 fullName 為 'firstName lastName', 而不會是 'firstName lastName {2}'
function cleanStringFormatResult(txt) {
	if (txt == null) return "";
	return txt.replace(getStringFormatPlaceHolderRegEx("\\d+"), "");
}