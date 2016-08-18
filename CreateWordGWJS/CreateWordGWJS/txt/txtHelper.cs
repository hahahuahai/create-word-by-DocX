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
        /// 读取txt
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
    }
}
