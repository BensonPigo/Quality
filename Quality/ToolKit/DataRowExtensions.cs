using System;
using System.Data;

public static class DataRowExtensions
{
    public static object ToValue(this object value)
    {
        // 處理 DBNull 或 null 的情況
        if (value == null || value == DBNull.Value)
        {
            return null; // 回傳 null 作為預設值
        }

        // 動態判斷型別，進行適當的轉換
        var t = value.GetType();
        switch (t.Name)
        {
            case "String":
                return value.ToValue<string>();

            case "Int32":
                return value.ToValue<int>();

            case "Int64":
                return value.ToValue<long>();

            case "Double":
                return value.ToValue<double>();

            case "Decimal":
                return value.ToValue<decimal>();

            case "Boolean":
                return value.ToValue<bool>();

            case "DateTime":
                return value.ToValue<DateTime>();

            default:
                // 若型別未明確處理，返回原始值
                return value;
        }
    }

    public static T ToValue<T>(this object value)
    {
        // 處理 DBNull 或 null 的情況，返回型別預設值
        if (value == null || value == DBNull.Value)
        {
            return default(T);
        }

        // 嘗試將值轉換為指定型別
        return (T)Convert.ChangeType(value, typeof(T));
    }
}
