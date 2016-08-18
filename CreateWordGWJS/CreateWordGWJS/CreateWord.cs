using CreateWordGWJS.model;
using CreateWordGWJS.parcels;
using Novacode;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CreateWordGWJS
{

    public partial class CreateWord : Form
    {
        string WordPath = Environment.CurrentDirectory + "\\test.docx";    //文档路径
        public CreateWord()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 生成文档的点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void create_Click(object sender, EventArgs e)
        {
            List<FDXXmodel> lstFM = FDXXmodel.GetInfo();
            List<Parcelmodel> p = FDXXmodel.Parcels(lstFM);
            using (DocX document = DocX.Create(WordPath))
            {
                //tableHelper thelper = new tableHelper();
                Table t = tableHelper.Template(document);
                t = tableHelper.inserttable(t, lstFM);
                t = tableHelper.combineCells(t,p);
                document.InsertTable(t);

                parcelHelper phelper = new parcelHelper("通湖路699号地块");
                phelper.insertInfo(document);

                document.Save();
            }
            MessageBox.Show("创建文档成功！");
        }
    }
}
