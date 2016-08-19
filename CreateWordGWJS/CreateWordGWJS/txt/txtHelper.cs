using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordGWJS.txt
{
    public class txtHelper
    {
        /// <summary>
        /// 读取txt，按行读取
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<string> txtLines(string path)
        {
            List<string> lstStr = new List<string>();
            string[] lines = System.IO.File.ReadAllLines(path,Encoding.Default);
            foreach (string line in lines)
            {
                lstStr.Add(line);
            }
            return lstStr;
        }
        /// <summary>
        /// 读取txt文件，并以string形式返回里面所有内容
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string readtxt(string path)
        {
            string text = System.IO.File.ReadAllText(path,Encoding.Default);
            return text;
        }
    }
}
