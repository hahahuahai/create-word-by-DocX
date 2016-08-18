using Novacode;
using RealEstate.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test
{
    public partial class Form1 : Form
    {
        string WordPath = Environment.CurrentDirectory + "\\test.docx";    //文档路径
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<FDXXmodel> lstFM = GetInfo();

            //List<Company> companys = GetCompanys();
            //int i = 0;int temp = 0;
            //foreach (Company c in companys) i++;
            using (DocX document = DocX.Create(WordPath))
            {
                // Add a Table to this document.
                //Table t = document.AddTable(3, 13);
                // Specify some properties for this Table.
                //t.Alignment = Alignment.center;
                //t.Design = TableDesign.MediumGrid1Accent2;
                //foreach (Company c in companys)
                //{
                //    t.Rows[temp++].Cells[0].Paragraphs.First().Append(c.Name);
                //}
                tableHelper th = new tableHelper();
                Table t = th.Template(document);
                t = th.inserttable(t, lstFM);
                document.InsertTable(t);
                document.Save();
            }
            MessageBox.Show("创建文档成功！");
        }

        //从数据库中获取一行行表格数据，并存入FDXXmodel中。
        public List<FDXXmodel> GetInfo()
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
                System.Console.WriteLine("传递过来的异常值为：{0}", e);
            }
            finally
            {
                myReader.Close();
                mycon.Close();
            }
            return lstFDXX;
        }

        public List<Company> GetCompanys()
        {
            List<Company> lstCompany = new List<Company>();
            Company pi;

            OleDbConnection mycon = null;
            OleDbDataReader myReader = null;
            try
            {
                string connstr = "Provider=OraOLEDB.Oracle;User ID=jsdlUser;Password=jsdlKey;Data Source=orcl;";
                mycon = new OleDbConnection(connstr);
                if (mycon.State == ConnectionState.Closed)
                    mycon.Open();
                string sql = "select * from GSXX";
                OleDbCommand mycom = new OleDbCommand(sql, mycon);
                myReader = mycom.ExecuteReader();//执行command并得到相应的DataReader
                while (myReader.Read())//把得到的值赋给pi对象
                {
                    pi = new Company();
                    pi.ID = (int)myReader["GSXX_ID"];
                    pi.Name = (string)myReader["GSXX_MC"];
                    pi.X = (decimal)myReader["GSXX_X"];
                    pi.Y = (decimal)myReader["GSXX_Y"];
                    pi.BelongTo = (int)myReader["GSXX_LSDW"];
                    lstCompany.Add(pi);
                }
                foreach (Company parent in lstCompany)
                {
                    foreach (Company child in lstCompany)
                    {
                        if (child.BelongTo == parent.ID && child != parent)
                        {
                            parent.Children.Add(child);
                            child.Parent = parent;
                        }
                    }
                }

            }
            finally
            {
                myReader.Close();
                mycon.Close();
            }
            return lstCompany;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
