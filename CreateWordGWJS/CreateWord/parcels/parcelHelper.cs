using CreateWord.log;
using CreateWord.model;
using CreateWord.picture;
using CreateWord.table;
using Novacode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.parcels
{

    public class parcelHelper
    {
        private string Name = "";//宗地名称

        public parcelHelper(string name)
        {
            Name = name;
        }
        /// <summary>
        /// 插入某个宗地的所有信息（宗地图、平面分布图、鸟瞰图、分层分户平面图、场地涉外管线布置图）
        /// </summary>
        /// <param name="document"></param>
        public void insertInfo(DocX document, List<FDXXtbl> lstFM, List<Parcelmodel> p)
        {
            try
            {
                document.InsertSectionPageBreak();  //分页符

                var h1 = document.InsertParagraph(Name);
                h1.StyleName = "Heading1";
                insertZDT(document);//宗地图
                insertZPMFBT(document);//总平面分布图
                insertNKT(document);//鸟瞰图
                insertFCFH(document, lstFM, p);//分层分户平面图
                insertGXT(document);//场地涉外管线布置图
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }
        }

        /// <summary>
        /// 插入宗地图模块
        /// </summary>
        /// <param name="document"></param>
        public void insertZDT(DocX document)
        {
            try
            {
                var h1_1 = document.InsertParagraph("宗地图");
                h1_1.StyleName = "Heading2";
                h1_1.AppendLine();
                var parcelPic = document.InsertParagraph();
                picHelper.insert(document, parcelPic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\宗地图.jpg");//todo:图片路径要改活
                parcelPic.AppendLine();

                var landusePic = document.InsertParagraph();
                landusePic.AppendLine("国有土地使用证");
                picHelper.insert(document, landusePic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\土地证.jpg");//todo:图片路径要改活
                landusePic.AppendLine();
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }
        }

        /// <summary>
        /// 插入总平面分布图模块
        /// </summary>
        /// <param name="document"></param>
        public void insertZPMFBT(DocX document)
        {


            try
            {
                var h1_2 = document.InsertParagraph("总平面分布图");
                h1_2.StyleName = "Heading2";
                h1_2.AppendLine();
                var Pic = document.InsertParagraph();
                picHelper.insert(document, Pic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\总平面分布图.jpg");//todo:图片路径要改活
                Pic.AppendLine();

                //var ownershipPic = document.InsertParagraph();
                //ownershipPic.AppendLine("房屋所有权证");
                //picHelper.insert(document, landusePic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\房屋所有权证.jpg");//todo:图片路径要改活

            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }
        }

        /// <summary>
        /// 插入鸟瞰图模块
        /// </summary>
        /// <param name="document"></param>
        public void insertNKT(DocX document)
        {


            try
            {
                var h1_3 = document.InsertParagraph("鸟瞰图（航拍）");
                h1_3.StyleName = "Heading2";
                h1_3.AppendLine();
                var AerialViewPic = document.InsertParagraph();
                picHelper.insert(document, AerialViewPic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\鸟瞰图\\正射.JPG");
                picHelper.insert(document, AerialViewPic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\鸟瞰图\\前.JPG");
                picHelper.insert(document, AerialViewPic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\鸟瞰图\\后.JPG");
                picHelper.insert(document, AerialViewPic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\鸟瞰图\\左.JPG");
                picHelper.insert(document, AerialViewPic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\鸟瞰图\\右.JPG");
                AerialViewPic.AppendLine();
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }
        }

        /// <summary>
        /// 插入分层分户平面图
        /// </summary>
        /// <param name="document"></param>
        /// <param name="lstFM"></param>
        /// <param name="p"></param>
        public void insertFCFH(DocX document, List<FDXXtbl> lstFM, List<Parcelmodel> p)
        {
            Paragraph h1_4_1;
            string path = "";

            try
            {
                var h1_4 = document.InsertParagraph("分层分户平面图");
                h1_4.StyleName = "Heading2";
                foreach (FDXXtbl fm in lstFM)
                {
                    if (fm.ZDXX_MC + "地块" == Name)
                    {
                        h1_4_1 = document.InsertParagraph(fm.FCXX_JZMC + "分层分户平面图");
                        h1_4_1.StyleName = "Heading3";
                        path = Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\" + fm.FCXX_JZMC + "\\";
                        //实景图
                        var title1 = document.InsertParagraph();//实景图标题
                        title1.Append("房产外墙面实景图");
                        title1.Alignment = Alignment.center;
                        var VirtualMap = document.InsertParagraph();//实景图图片
                        picHelper.insert(document, VirtualMap, path + "外墙实景图.JPG");
                        //平面图
                        var title2 = document.InsertParagraph();//生产综合楼分层分户平面图标题
                        title2.Append("分层分户平面图");
                        title2.Alignment = Alignment.center;
                        List<string> lstStr = txt.txtHelper.txtLines(path + "列表.txt");
                        Table t = tableHelper.PlanTable(document, lstStr, path);
                        document.InsertTable(t);
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }

        }

        /// <summary>
        /// 插入场地涉外管线布置图
        /// </summary>
        /// <param name="document"></param>
        public void insertGXT(DocX document)
        {

            try
            {
                var h1_5 = document.InsertParagraph("场地涉外管线布置图");
                h1_5.StyleName = "Heading2";
                var Pic = document.InsertParagraph();
                picHelper.insert(document, Pic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\管线图.jpg");//todo:图片路径要改活
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }

        }

    }
}
