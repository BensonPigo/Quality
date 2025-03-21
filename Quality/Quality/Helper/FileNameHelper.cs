using System.Text.RegularExpressions;

public static class FileNameHelper
{
    /// <summary>
    /// 將檔名中可能造成 URL 或檔案系統問題的字元，替換為底線 "_"
    /// 目前只保留英數、底線、連字號與點，其餘全部轉成 "_"
    /// </summary>
    /// <param name="originalFileName">原始檔名</param>
    /// <returns>處理後的安全檔名</returns>
    public static string SanitizeFileName(string originalFileName)
    {
        if (string.IsNullOrEmpty(originalFileName))
            return string.Empty;

        // 使用正則表達式篩選：只允許 a-zA-Z0-9 . _ -
        // 其餘字元一律以 "_" 替代
        string pattern = @"[^a-zA-Z0-9\._-]+";
        string safeFileName = Regex.Replace(originalFileName, pattern, "_");

        return safeFileName;
    }
}
