using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ImageDetectionBeta.ExportUtil
{
    /// <summary>
    /// Excel 表格数据操作类
    /// 
    /// 摘要：
    ///      2016/3/23添加数据表格创建的方法
    /// </summary>
    public static class NpoiUtils
    {

        public static bool CreateDataSheet(DataTable dt, string path)
        {
            if (dt.Rows.Count <= 0 || dt.Rows.Count > 255 || dt.Columns.Count > 65535)
            {
                //excel表格不能超过0xff列和0xffff行
                return false;
            }

            //创建工作薄
            HSSFWorkbook wk = new HSSFWorkbook();
            //创建一个名称为mySheet的表
            ISheet tb = wk.CreateSheet("mySheet");
            IRow row;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                row = tb.CreateRow(i);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            Random r = new Random();
            try
            {
                using (FileStream fs = File.OpenWrite(path)) //打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建是不要打开该文件！
                {
                    wk.Write(fs);   //向打开的这个xls文件中写入mySheet表并保存。
                }
            }
            catch
            {
                return false;
            }
            return true;
        }
    }
}
