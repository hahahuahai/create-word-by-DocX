using CreateWordGWJS.model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CreateWordGWJS
{
    public class FDXXmodel
    {
        #region 成员变量
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
        #endregion

        #region 相关操作

        /// <summary>
        /// 从数据库中获取一行行表格数据，并存入FDXXmodel中。
        /// </summary>
        /// <returns></returns>
        public static List<FDXXmodel> GetInfo()
        {
            List<FDXXmodel> lstFDXX = new List<FDXXmodel>();
            FDXXmodel fm;
            OleDbConnection mycon = null;
            OleDbDataReader myReader = null;
            try
            {
                string connstr = "Provider=OraOLEDB.Oracle;User ID=jsdlUser;Password=jsdlKey;Data Source=orcl;";
                mycon = new OleDbConnection(connstr);
                if (mycon.State == ConnectionState.Closed)
                    mycon.Open();
                string sql = "SELECT ZDXX_MC,ZDXX_ZDMJ,FCXX_JZMC,FCXX_DSCS,FCXX_DXCS,FCXX_JZJG,FCZK_JSND,(FCXX_DSMJ+FCXX_DXMJ) AS ZMJ,FCXX_DSMJ,FCXX_DXMJ,FCZK_SYGN,FCXX_BZ FROM ZDXX,FCXX,FCZK WHERE ZDXX_ID = FCXX_DKID AND FCXX_ID = FCZK_JZID AND ZDXX_SSDW = 321084";
                OleDbCommand mycom = new OleDbCommand(sql, mycon);
                myReader = mycom.ExecuteReader();//执行command并得到相应的DataReader
                while (myReader.Read())//把得到的值赋给fm对象
                {
                    fm = new FDXXmodel();
                    fm.ZDXX_MC = (string)myReader["ZDXX_MC"];
                    fm.ZDXX_ZDMJ = (decimal)myReader["ZDXX_ZDMJ"];
                    fm.FCXX_JZMC = (string)myReader["FCXX_JZMC"];
                    fm.FCXX_DSCS = (decimal)myReader["FCXX_DSCS"];
                    fm.FCXX_DXCS = (decimal)myReader["FCXX_DXCS"];
                    fm.FCXX_JZJG = (string)myReader["FCXX_JZJG"];
                    fm.FCXX_JSND = (decimal)myReader["FCZK_JSND"];
                    fm.ZMJ = (decimal)myReader["ZMJ"];
                    fm.FCXX_DSMJ = (decimal)myReader["FCXX_DSMJ"];
                    fm.FCXX_DXMJ = (decimal)myReader["FCXX_DXMJ"];
                    fm.FCZK_SYGN = (string)myReader["FCZK_SYGN"];
                    if (!DBNull.Value.Equals(myReader["FCXX_BZ"])) fm.FCXX_BZ = (string)myReader["FCXX_BZ"];//判断FCXX_BZ是否为空值（DBNull）
                    else fm.FCXX_BZ = "";

                    lstFDXX.Add(fm);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                myReader.Close();
                mycon.Close();
            }
            return lstFDXX;
        }

        /// <summary>
        /// 得到ZDXX_MC这个字段名称重复的个数，用于合并单元格。
        /// </summary>
        /// <param name="lstFDXX"></param>
        /// <returns>返回无重复的宗地名称和宗地出现的次数</returns>
        public static List<Parcelmodel> Parcels(List<FDXXmodel> lstFDXX)
        {
            int flag;
            int num = 0;
            string temp = "";
            Parcelmodel pm;
            List<Parcelmodel> lstParcel = new List<Parcelmodel>();
            foreach (FDXXmodel fm in lstFDXX)
            {
                num = 0;
                flag = 1;
                temp = fm.ZDXX_MC;
                foreach (Parcelmodel pm1 in lstParcel)
                {
                    if (pm1.name == temp) { flag = 0; break; }//说明p[]里面已经存储了该宗地信息，所以把flag值设置为0。
                }
                if (flag == 1)
                {
                    foreach (FDXXmodel fm1 in lstFDXX)
                    {
                        if (fm1.ZDXX_MC == temp) num++;
                    }
                    pm = new Parcelmodel();
                    pm.name = temp;
                    pm.num = num;
                    lstParcel.Add(pm);
                }
            }
            return lstParcel;
        }

        #endregion
    }
}
