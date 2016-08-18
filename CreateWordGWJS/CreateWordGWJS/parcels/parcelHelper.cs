using CreateWordGWJS.picture;
using Novacode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordGWJS.parcels
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
        public void insertInfo(DocX document)
        {
            document.InsertSectionPageBreak();  //分页符

            var h1 = document.InsertParagraph(Name);
            h1.StyleName = "Heading1";
            insertZDT(document);
            insertZPMFBT(document);
        }

        /// <summary>
        /// 插入宗地图模块
        /// </summary>
        /// <param name="document"></param>
        public void insertZDT(DocX document)
        {
            var h1_1 = document.InsertParagraph("宗地图");
            h1_1.StyleName = "Heading2";
            h1_1.AppendLine();
            var parcelPic = document.InsertParagraph();
            picHelper.insert(document,parcelPic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\宗地图.jpg");//todo:图片路径要改活
            parcelPic.AppendLine();

            var landusePic = document.InsertParagraph();
            landusePic.AppendLine("国有土地使用证");
            picHelper.insert(document,landusePic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\土地证.jpg");//todo:图片路径要改活
            landusePic.AppendLine();
        }

        /// <summary>
        /// 插入总平面分布图模块
        /// </summary>
        /// <param name="document"></param>
        public void insertZPMFBT(DocX document)
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

        public void insertNKT(DocX document)
        {

        }

    }
}
