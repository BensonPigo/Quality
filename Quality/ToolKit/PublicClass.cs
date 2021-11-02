using DatabaseObject.Public;
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

        public static List<T> CompareListValue<T>(List<T> newList, List<T> oldList, string keyCols, string compareCols) where T : CompareBase
        {
            List<T> resultList = new List<T>();
            string[] keyColArray = keyCols.Split(',');

            var typeTarget = typeof(T);

            // 判斷old 不存在於 new 表示應該被刪除
            foreach (var oldItem in oldList)
            {
                bool needDelete = !newList.Any(newItem =>
                {
                    bool isKeySame = true;
                    foreach (string fieldName in keyColArray)
                    {
                        if (!typeTarget.GetProperty(fieldName).GetValue(newItem).EqualString(typeTarget.GetProperty(fieldName).GetValue(oldItem)))
                        {
                            isKeySame = false;
                            break;
                        }
                    }
                    return isKeySame;
                });

                if (needDelete)
                {
                    oldItem.StateType = CompareStateType.Delete;
                    resultList.Add(oldItem);
                }
            }

            // 判斷需要新增或修改
            foreach (var newItem in newList)
            {
                var matchResult = oldList.Where(
                        oldItem =>
                        {
                            bool isKeySame = true;
                            foreach (string fieldName in keyColArray)
                            {
                                if (!typeTarget.GetProperty(fieldName).GetValue(newItem).EqualString(typeTarget.GetProperty(fieldName).GetValue(oldItem)))
                                {
                                    isKeySame = false;
                                    break;
                                }
                            }
                            return isKeySame;

                        }
                    );

                if (!matchResult.Any())
                {
                    newItem.StateType = CompareStateType.Add;
                    resultList.Add(newItem);
                    continue;
                }

                var matchItem = matchResult.First();

                foreach (string compareField in compareCols.Split(','))
                {
                    if (!typeTarget.GetProperty(compareField).GetValue(newItem).EqualString(typeTarget.GetProperty(compareField).GetValue(matchItem)))
                    {
                        newItem.StateType = CompareStateType.Edit;
                        resultList.Add(newItem);
                        break;
                    }
                }
            }

            return resultList;
        }
    }
}
