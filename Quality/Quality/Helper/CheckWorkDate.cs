using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FactoryDashBoardWeb.Helper
{
    public static class CheckWorkDate
    {
        /// <summary>
        /// 避免Client傳入錯誤的WorkDate查詢條件
        /// </summary>
        /// <param name="WorkDate"></param>
        /// <returns></returns>
        public static DateTime Check(DateTime? WorkDate)
        {
            if (!WorkDate.HasValue)
            {
                return DateTime.Now;
            }
            DateTime dt = WorkDate.Value;
            DateTime defaultTime = Convert.ToDateTime("2000/01/01 00:00");

            //小於零   dt 早於 defaultTime。
            //零       dt 與 defaultTime 相同。
            //大於零   dt 晚於 defaultTime。

            //不可以 < 預設時間
            if (DateTime.Compare(defaultTime, dt) > 0)
                return DateTime.Now;
            else
                return dt;

        }
    }
}