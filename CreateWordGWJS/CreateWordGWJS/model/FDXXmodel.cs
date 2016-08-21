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
    //房地信息统计模块里面需要用到的对象

    #region  房产信息统计模块的表格数据类
    public class FDXXtbl
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
        public static List<FDXXtbl> GetInfo()
        {
            List<FDXXtbl> lstFDXX = new List<FDXXtbl>();
            FDXXtbl fm;
            OleDbConnection mycon = null;
            OleDbDataReader myReader = null;
            try
            {
                string connstr = "Provider=OraOLEDB.Oracle;User ID=jsdlUser;Password=jsdlKey;Data Source=orcl;";
                mycon = new OleDbConnection(connstr);
                if (mycon.State == ConnectionState.Closed)
                    mycon.Open();
                string sql = "SELECT ZDXX_MC,ZDXX_ZDMJ,FCXX_JZMC,FCXX_DSCS,FCXX_DXCS,FCXX_JZJG,FCZK_JSND,(FCXX_DSMJ+FCXX_DXMJ) AS ZMJ,FCXX_DSMJ,FCXX_DXMJ,FCZK_SYGN,FCXX_BZ FROM ZDXX,FCXX,FCZK WHERE ZDXX_ID = FCXX_DKID AND FCXX_ID = FCZK_JZID AND ZDXX_SSDW = 1122007";
                OleDbCommand mycom = new OleDbCommand(sql, mycon);
                myReader = mycom.ExecuteReader();//执行command并得到相应的DataReader
                while (myReader.Read())//把得到的值赋给fm对象
                {
                    fm = new FDXXtbl();
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
        public static List<Parcelmodel> Parcels(List<FDXXtbl> lstFDXX)
        {
            int flag;
            int num = 0;
            string temp = "";
            Parcelmodel pm;
            List<Parcelmodel> lstParcel = new List<Parcelmodel>();
            foreach (FDXXtbl fm in lstFDXX)
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
                    foreach (FDXXtbl fm1 in lstFDXX)
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
    #endregion     
    
    #region 房地信息统计模块的文字描述类
    public class FDXXsentence1
    {
        #region 成员变量
        //各类住房栋数
        public decimal count { get; set; }
        //占地总面积
        public decimal ZDZMJ { get; set; }
        //总建筑面积
        public decimal ZJZMJ { get; set; }
        #endregion

        public static FDXXsentence1 GetInfo()
        {
            FDXXsentence1 fs = new FDXXsentence1();
            OleDbConnection mycon = null;
            OleDbDataReader myReader = null;
            try
            {
                string connstr = "Provider=OraOLEDB.Oracle;User ID=jsdlUser;Password=jsdlKey;Data Source=orcl;";
                mycon = new OleDbConnection(connstr);
                if (mycon.State == ConnectionState.Closed)
                    mycon.Open();
                string sql = "select count(*), sum(ZDXX_ZDMJ), sum(B_FCMJ) from VIEWJSS, GSXX where ZDXX_SSDW = GSXX_ID and GSXX_ID = 1122007";//todo:路径中的高邮市供电公司要改活
                OleDbCommand mycom = new OleDbCommand(sql, mycon);
                myReader = mycom.ExecuteReader();//执行command并得到相应的DataReader
                myReader.Read();
                fs.count = (decimal)myReader["COUNT(*)"];
                fs.ZDZMJ = (decimal)myReader["SUM(ZDXX_ZDMJ)"];
                fs.ZJZMJ = (decimal)myReader["SUM(B_FCMJ)"];                                
            }
            catch (System.Exception ex)
            {

            }
            finally
            {
                myReader.Close();
                mycon.Close();
            }
            return fs;

        }
    }

    public class FDXXsentence2
    {
        #region 成员变量
        public string GNGL { get; set; }
        public decimal GNGL_MJ { get; set; }
        #endregion

        public static List<FDXXsentence2> GetInfo()
        {
            List<FDXXsentence2> lstFS = new List<FDXXsentence2>();
            FDXXsentence2 fs;
            OleDbConnection mycon = null;
            OleDbDataReader myReader = null;
            try
            {
                string connstr = "Provider=OraOLEDB.Oracle;User ID=jsdlUser;Password=jsdlKey;Data Source=orcl;";
                mycon = new OleDbConnection(connstr);
                if (mycon.State == ConnectionState.Closed)
                    mycon.Open();
                string sql = "select ZDXX_SSDW, FCZK_GNGL, sum(FCMJ) from ZDXX,(select FCXX_DKID, FCZK_GNGL, sum(FCZK_SYMJ) as FCMJ from FCZK, FCXX where FCZK_JZID = FCXX_ID group by FCXX_DKID, FCZK_GNGL) a where ZDXX_ID = FCXX_DKID and ZDXX_SSDW = 1122007 group by ZDXX_SSDW, FCZK_GNGL";//todo:路径中的高邮市供电公司要改活
                OleDbCommand mycom = new OleDbCommand(sql, mycon);
                myReader = mycom.ExecuteReader();//执行command并得到相应的DataReader
                while (myReader.Read())//把得到的值赋给fm对象
                {
                    fs = new FDXXsentence2();
                    fs.GNGL = (string)myReader["FCZK_GNGL"];
                    fs.GNGL_MJ = (decimal)myReader["SUM(FCMJ)"];

                    lstFS.Add(fs);
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                myReader.Close();
                mycon.Close();
            }
            return lstFS;
        }
    

    }

    public class FDXXsentence3
    {
        #region 成员变量
        //建成投运10年内的房屋面积
        public decimal FWMJ_10 { get; set; }
        //建成投运10-20年内的房屋面积
        public decimal FWMJ_20 { get; set; }
        //建成投运30年以上的房屋面积
        public decimal FWMJ_30 { get; set; }
        #endregion

        public static FDXXsentence3 GetInfo()
        {
            FDXXsentence3 fs = new FDXXsentence3();
            string temp = "";
            OleDbConnection mycon = null;
            OleDbDataReader myReader = null;
            try
            {
                string connstr = "Provider=OraOLEDB.Oracle;User ID=jsdlUser;Password=jsdlKey;Data Source=orcl;";
                mycon = new OleDbConnection(connstr);
                if (mycon.State == ConnectionState.Closed)
                    mycon.Open();
                string sql = "select ZDXX_SSDW, TJXM, sum(FCMJ) as FCMJ from ZDXX,(select FCXX_DKID, (case when to_char(sysdate, 'yyyy' )-FCZK_JSND <= 10 then cast('10年以内（含10年）' as nvarchar2(20)) when to_char(sysdate, 'yyyy' )-FCZK_JSND > 10 and to_char(sysdate, 'yyyy' )-FCZK_JSND <= 20 then cast('10-20年（含20年）' as nvarchar2(20)) when to_char(sysdate, 'yyyy' )-FCZK_JSND > 20 and to_char(sysdate, 'yyyy' )-FCZK_JSND <= 30 then cast('20-30年（含30年）' as nvarchar2(20)) when to_char(sysdate, 'yyyy' )-FCZK_JSND > 30 then cast('30年以上' as nvarchar2(20)) end) as TJXM, sum(FCZK_SYMJ) as FCMJ from FCZK, FCXX where FCZK_JZID = FCXX_ID group by FCXX_DKID, FCZK_JSND ) a where ZDXX_ID = a.FCXX_DKID and ZDXX_SSDW = 1122007 group by ZDXX_SSDW, TJXM";//todo:路径中的高邮市供电公司要改活
                OleDbCommand mycom = new OleDbCommand(sql, mycon);
                myReader = mycom.ExecuteReader();//执行command并得到相应的DataReader
                while (myReader.Read())//把得到的值赋给fm对象
                {
                    temp = (string)myReader["TJXM"];
                    if (temp == "10年以内（含10年）") fs.FWMJ_10 = (decimal)myReader["FCMJ"];
                    else if (temp == "30年以上") fs.FWMJ_30 = (decimal)myReader["FCMJ"];
                    else fs.FWMJ_20 = (decimal)myReader["FCMJ"];
                }
            }
            catch (Exception e)
            {
                
            }
            finally
            {
                myReader.Close();
                mycon.Close();
            }
            return fs;
        }
    }
#endregion
}
