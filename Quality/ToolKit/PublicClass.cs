using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolKit
{
    public static class PublicClass
    {
        /// <summary>
        /// 取得字串右邊字數
        /// </summary>
        /// <param name="s">原字串</param>
        /// <param name="length">字數</param>
        /// <returns>string</returns>
        public static string Right(this string s, int length)
        {
            length = Math.Max(length, 0);
            if (s.Length > length)
            {
                return s.Substring(s.Length - length, length);
            }
            else
            {
                return s;
            }
        }

        /// <summary>
        /// 取得字串左邊字數
        /// </summary>
        /// <param name="s">原字串</param>
        /// <param name="length">字數</param>
        /// <returns>string</returns>
        public static string Left(this string s, int length)
        {
            length = Math.Max(length, 0);
            if (s.Length > length)
            {
                return s.Substring(0, length);
            }
            else
            {
                return s;
            }
        }

        public static bool EqualString(this Object a, Object equalTo)
        {
            return EqualString(GetString(a), GetString(equalTo));
        }

        private static bool EqualString(String a, String b)
        {
            return a.TrimEnd().Equals(b.TrimEnd(), StringComparison.InvariantCultureIgnoreCase);
        }

        public static string GetString(object obj)
        {
            if (obj == null)
                return "";
            else
                return obj.ToString();
        }

        public static int ToInt(this object val)
        {
            if (val == null)
            {
                return 0;
            }

            if (!int.TryParse(val.ToString(), out int rtnValue))
            {
                rtnValue = 0;
            }

            return rtnValue;
        }

        public static decimal? ToDecimal(this object val)
        {
            if (val == null)
            {
                return 0;
            }

            if (!decimal.TryParse(val.ToString(), out decimal rtnValue))
            {
                rtnValue = 0;
            }

            return rtnValue;
        }

        public static IEnumerable<TSource> Distinct<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            HashSet<TKey> seenKeys = new HashSet<TKey>();
            foreach (TSource element in source)
            {
                var elementValue = keySelector(element);
                if (seenKeys.Add(elementValue))
                {
                    yield return element;
                }
            }
        }

        public static string JoinToString(this IEnumerable<string> source, string joinString)
        {
            if (source.Count() == 0)
            {
                return string.Empty;
            }

            string result = string.Empty;
            foreach (string item in source)
            {
                result += $"{joinString}{item}";
            }

            return result.TrimStart(joinString.ToCharArray());
        }
    }
}
