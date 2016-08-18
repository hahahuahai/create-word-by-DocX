using CreateWordGWJS.model;
using Novacode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordGWJS
{
    /// <summary>
    /// 表相关的操作
    /// </summary>
    class tableHelper
    {

        /// <summary>
        /// 创建房地信息汇总表模板（前两行，字段；第三行，总计）
        /// </summary>
        /// <returns></returns>
        public static Table Template(DocX document)
        {
            Table t = document.AddTable(3, 13);

            #region "前三行"
            t.Rows[0].Cells[0].Paragraphs.First().Append("序号").Bold();
            t.Rows[0].Cells[1].Paragraphs.First().Append("地块名称").Bold();
            t.Rows[0].Cells[2].Paragraphs.First().Append("地块占地面积（m2）").Bold();
            t.Rows[0].Cells[3].Paragraphs.First().Append("建筑名称").Bold();
            t.Rows[0].Cells[4].Paragraphs.First().Append("建筑层数").Bold();
            t.Rows[0].Cells[6].Paragraphs.First().Append("建筑结构").Bold();
            t.Rows[0].Cells[7].Paragraphs.First().Append("建筑年代").Bold();
            t.Rows[0].Cells[8].Paragraphs.First().Append("建筑面积（m2）").Bold();
            t.Rows[0].Cells[11].Paragraphs.First().Append("具体使用功能").Bold();
            t.Rows[0].Cells[12].Paragraphs.First().Append("备注").Bold();
            t.Rows[1].Cells[4].Paragraphs.First().Append("地上").Bold();
            t.Rows[1].Cells[5].Paragraphs.First().Append("地下").Bold();
            t.Rows[1].Cells[8].Paragraphs.First().Append("总体").Bold();
            t.Rows[1].Cells[9].Paragraphs.First().Append("地上").Bold();
            t.Rows[1].Cells[10].Paragraphs.First().Append("地下").Bold();
            t.Rows[2].Cells[0].Paragraphs.First().Append("合计").Bold();

            //单元格合并操作（先竖向合并，再横向合并，以免报错，因为横向合并会改变列数）
            t.MergeCellsInColumn(0, 0, 1);
            t.MergeCellsInColumn(1, 0, 1);
            t.MergeCellsInColumn(2, 0, 1);
            t.MergeCellsInColumn(3, 0, 1);
            t.MergeCellsInColumn(6, 0, 1);
            t.MergeCellsInColumn(7, 0, 1);
            t.MergeCellsInColumn(11, 0, 1);
            t.MergeCellsInColumn(12, 0, 1);
            t.Rows[0].MergeCells(4, 5);
            t.Rows[0].MergeCells(7, 9);
            #endregion


            //document.InsertTable(t);

            return t;
        }

        /// <summary>
        /// 把数据一行行插入
        /// </summary>
        /// <param name="t"></param>
        /// <param name="lsmFM"></param>
        /// <returns></returns>
        public static Table inserttable(Table t, List<FDXXmodel> lsmFM)
        {
            int temp = 3;

            foreach (FDXXmodel fm in lsmFM)
            {
                t.InsertRow();
                t.Rows[temp].Cells[1].Paragraphs.First().Append(fm.ZDXX_MC);
                t.Rows[temp].Cells[2].Paragraphs.First().Append("" + fm.ZDXX_ZDMJ);
                t.Rows[temp].Cells[3].Paragraphs.First().Append(fm.FCXX_JZMC);
                t.Rows[temp].Cells[4].Paragraphs.First().Append("" + fm.FCXX_DSCS);
                t.Rows[temp].Cells[5].Paragraphs.First().Append("" + fm.FCXX_DXCS);
                t.Rows[temp].Cells[6].Paragraphs.First().Append(fm.FCXX_JZJG);
                t.Rows[temp].Cells[7].Paragraphs.First().Append("" + fm.FCXX_JSND);
                t.Rows[temp].Cells[8].Paragraphs.First().Append("" + fm.ZMJ);
                t.Rows[temp].Cells[9].Paragraphs.First().Append("" + fm.FCXX_DSMJ);
                t.Rows[temp].Cells[10].Paragraphs.First().Append("" + fm.FCXX_DXMJ);
                t.Rows[temp].Cells[11].Paragraphs.First().Append(fm.FCZK_SYGN);
                t.Rows[temp].Cells[12].Paragraphs.First().Append(fm.FCXX_BZ);
                temp++;
            }
            return t;
        }

        /// <summary>
        /// 合并重复的地块名称
        /// </summary>
        /// <param name="t"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        public static Table combineCells(Table t, List<Parcelmodel> p)
        {
            int startCell = 3;
            int number = 1;//表中第一列的序号
            foreach (Parcelmodel pm in p)
            {
                if (pm.num != 1)
                {
                    t.MergeCellsInColumn(0, startCell, startCell + pm.num - 1); //合并序号                    
                    t.MergeCellsInColumn(1, startCell, startCell + pm.num - 1); //合并地块名称
                    t.MergeCellsInColumn(2, startCell, startCell + pm.num - 1); //合并地块占地面积
                }
                t.Rows[startCell].Cells[0].Paragraphs.First().Append("" + number++);    //添加序号
                startCell = startCell + pm.num;
            }
            return t;
        }

   
    }
}
