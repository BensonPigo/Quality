using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolKit
{
    public static class PublicClass
    {
        ///<summary>
        ///字串轉半形
        ///</summary>
        ///<paramname="input">任一字元串</param>
        ///<returns>半形字元串</returns>
        public static string ToNarrow(this string input)
        {
            char[] c = input.ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] == 12288)
                {
                    c[i] = (char)32;
                    continue;
                }
                if (c[i] > 65280 && c[i] < 65375)
                    c[i] = (char)(c[i] - 65248);
            }
            return new string(c);
        }

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

        public static decimal ToDecimal(this object val)
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

        public enum SingLocation : int
        {
            /// <summary>
            /// 左上
            /// </summary>
            TopLeft,

            /// <summary>
            /// 中上
            /// </summary>
            TopMiddle,

            /// <summary>
            /// 右上
            /// </summary>
            TopRight,

            /// <summary>
            /// 左中
            /// </summary>
            MiddleLeft,

            /// <summary>
            /// 正中
            /// </summary>
            MiddleMiddle,

            /// <summary>
            /// 右中
            /// </summary>
            MiddleRight,

            /// <summary>
            /// 左下
            /// </summary>
            LowerLeft,

            /// <summary>
            /// 中下
            /// </summary>
            LowerMiddle,

            /// <summary>
            /// 右下
            /// </summary>
            LowerRight,

            /// <summary>
            /// 正中斜體
            /// </summary>
            MiddleItalic,
        }

        public static string AddImageSignWord(byte[] ImageBytes, string signWord, SingLocation signLocation, int fontsize = 48, string fontname = "Arial Black", bool test = false)
        {
            Image img = Image.FromStream(new MemoryStream(ImageBytes));

            string imageName = $"{Guid.NewGuid()}.jpg";
            string imgPath;
            if (test)
            {
                imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
            }
            else
            {
                imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
            }

            if (string.IsNullOrEmpty(signWord))
            {
                img.Save(imgPath);
                img.Dispose();
                return imgPath;
            }

            Graphics g = Graphics.FromImage(img);
            Font drawFont = new Font(fontname, fontsize, FontStyle.Italic, GraphicsUnit.Pixel);
            SizeF crSize = g.MeasureString(signWord, drawFont);

            while (crSize.Width > img.Width)
            {
                fontsize -= 3;
                drawFont = new Font(fontname, fontsize, FontStyle.Italic, GraphicsUnit.Pixel);
                crSize = g.MeasureString(signWord, drawFont);
            }

            while (crSize.Width < img.Width * 0.5)
            {
                fontsize += 3;
                drawFont = new Font(fontname, fontsize, FontStyle.Italic, GraphicsUnit.Pixel);
                crSize = g.MeasureString(signWord, drawFont);
            }

            float xpos = 0;
            float ypos = 0;

            switch (signLocation)
            {
                case SingLocation.TopLeft:
                    xpos = (float)img.Width * (float).01;
                    ypos = (float)img.Height * (float).01;
                    break;
                case SingLocation.TopMiddle:
                    xpos = ((float)img.Width * (float).50) - (crSize.Width / 2);
                    ypos = (float)img.Height * (float).01;
                    break;
                case SingLocation.TopRight:
                    xpos = ((float)img.Width * (float).99) - crSize.Width;
                    ypos = (float)img.Height * (float).01;
                    break;
                case SingLocation.MiddleLeft:
                    xpos = (float)img.Width * (float).01;
                    ypos = ((float)img.Height * (float).50) - (crSize.Height / 2);
                    break;
                case SingLocation.MiddleMiddle:
                    xpos = ((float)img.Width * (float).50) - (crSize.Width / 2);
                    ypos = ((float)img.Height * (float).50) - (crSize.Height / 2);
                    break;
                case SingLocation.MiddleRight:
                    xpos = ((float)img.Width * (float).99) - crSize.Width;
                    ypos = ((float)img.Height * (float).50) - (crSize.Height / 2);
                    break;
                case SingLocation.LowerLeft:
                    xpos = (float)img.Width * (float).01;
                    ypos = ((float)img.Height * (float).99) - crSize.Height;
                    break;
                case SingLocation.LowerMiddle:
                    xpos = ((float)img.Width * (float).50) - (crSize.Width / 2);
                    ypos = ((float)img.Height * (float).99) - crSize.Height;
                    break;
                case SingLocation.LowerRight:
                    xpos = ((float)img.Width * (float).99) - crSize.Width;
                    ypos = ((float)img.Height * (float).99) - crSize.Height;
                    break;
                case SingLocation.MiddleItalic:
                    // 高大於寬的圖片，在偏移0.5
                    float diffWidth = img.Height > img.Width ? (float).5 : 1;
                    float transHeight = img.Height > img.Width ? img.Height * (float).7 : img.Height * (float).8;
                    float transRotate = img.Height > img.Width ? 30 : 20;
                    xpos = (float)img.Width * (float).2 * diffWidth;
                    ypos = 0;
                    g.TranslateTransform(0, transHeight);
                    g.RotateTransform(-transRotate);
                    break;
            }

            g.DrawString(signWord, drawFont, new SolidBrush(Color.FromArgb(75, 255, 255, 255)), xpos, ypos);
            img.Save(imgPath);

            g.Dispose();
            img.Dispose();

            return imgPath;
        }
    }
}
