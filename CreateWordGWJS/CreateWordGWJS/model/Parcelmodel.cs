using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordGWJS.model
{
    public class Parcelmodel
    {
        #region 成员变量
        public string name { get; set; } //宗地名称
        public int num { get; set; }//宗地里面建筑的个数。（表格里面合并单元格时需要）
        #endregion
    }
}
