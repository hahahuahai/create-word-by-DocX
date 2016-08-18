using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test
{
    public class FDXXmodel
    {
        //地块名称
        public string ZDXX_MC { get; set; }
        //地块占地面积
        public decimal ZDXX_ZDMJ { get; set; }
        //建筑名称
        public string FCXX_JZMC { get; set; }
        //建筑层数地上
        public decimal FCXX_DSCS { get; set; }
        //建筑层数地下
        public decimal FCXX_DXCS { get; set; }
        //建筑结构
        public string FCXX_JZJG { get; set; }
        //建筑年代
        public decimal FCXX_JSND { get; set; }
        //建筑面积总体
        public decimal ZMJ { get; set; }
        //建筑面积地上
        public decimal FCXX_DSMJ { get; set; }
        //建筑面积地下
        public decimal FCXX_DXMJ { get; set; }
        //具体使用功能
        public string FCZK_SYGN { get; set; }
        //备注
        public string FCXX_BZ { get; set; }
    }
}
