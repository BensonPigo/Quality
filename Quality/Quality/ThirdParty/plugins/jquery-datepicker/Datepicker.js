/* === start === 必要：jQuery UI date picker plugin 中文化 */
jQuery(function ($) {
  //  $.datepicker.regional['zh-TW'] = {
  //      closeText: '關閉',
  //      prevText: '&#x3C;上月',
  //      nextText: '下月&#x3E;',
  //      currentText: '今天',
  //      monthNames: ['一月', '二月', '三月', '四月', '五月', '六月',
		//'七月', '八月', '九月', '十月', '十一月', '十二月'],
  //      monthNamesShort: ['一月', '二月', '三月', '四月', '五月', '六月',
		//'七月', '八月', '九月', '十月', '十一月', '十二月'],
  //      dayNames: ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'],
  //      dayNamesShort: ['周日', '周一', '周二', '周三', '周四', '周五', '周六'],
  //      dayNamesMin: ['日', '一', '二', '三', '四', '五', '六'],
  //      weekHeader: '周',
  //      dateFormat: 'yy/mm/dd',
  //      firstDay: 1,
  //      isRTL: false,
  //      showMonthAfterYear: true,
  //      yearSuffix: '年'
  //  };
    $.datepicker.regional['zh-TW'] = {
        closeText: '關閉',
        prevText: '&#x3C;上月',
        nextText: '下月&#x3E;',
        currentText: '今天',
        monthNames: ['January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'],
        monthNamesShort: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
            'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
        dayNames: ['Sunday	', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
        dayNamesShort: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
        dayNamesMin: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
        weekHeader: 'Week',
        dateFormat: 'yy/mm/dd',
        firstDay: 1,
        isRTL: false,
        showMonthAfterYear: true,
        yearSuffix: 'Year'
    };
    $.datepicker.setDefaults($.datepicker.regional['zh-TW']);
});
/* === end   === 必要：jQuery UI date picker plugin 中文化 */
