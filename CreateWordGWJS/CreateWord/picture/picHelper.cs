using Novacode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.picture
{

    class picHelper
    {
        ///// <summary>
        ///// 将图片插入到指定的书签位置
        ///// </summary>
        ///// <param name="document">操作的文档</param>
        ///// <param name="BMname">书签的名字</param>
        ///// <param name="picPath">图片的路径</param>
        //public static void insertBybookmark(DocX document,string BMname,string picPath)
        //{
        //    //todo:
        //}

        /// <summary>
        /// 把图片插入到段落
        /// </summary>
        /// <param name="p"></param>
        /// <param name="picPath"></param>
        public static void insert(DocX document, Paragraph p, string picPath)
        {
            Image image = document.AddImage(picPath);

            Picture picture = image.CreatePicture();
            picture.Height = 600;
            picture.Width = 600;
            p.AppendPicture(picture);
        }

        /// <summary>
        /// 根据路径生成一个图片，返回图片
        /// </summary>
        /// <param name="document"></param>
        /// <param name="picPath"></param>
        /// <param name="height">1厘米约等于28像素</param>
        /// <param name="width">1厘米约等于28像素</param>
        /// <returns></returns>
        public static Picture getPic(DocX document, string picPath, int height, int width)
        {
            Image image = document.AddImage(picPath);

            Picture picture = image.CreatePicture();
            picture.Height = height;
            picture.Width = width;
            return picture;
        }

    }
}
