using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using SyntacticSugar;
using AsposeSugar;
using System.Data;

namespace Test
{
    public partial class _Default : System.Web.UI.Page
    {
        AsCellSugar asr = new AsCellSugar();
        protected void Page_Load(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            var dt = DataTableSugar.CreateEasyTable("姓名", "姓别")
              .AddRow("张三", "女")
              .AddRow("李娜", "女");
            var dt2 = DataTableSugar.CreateEasyTable("姓名", "姓别")
             .AddRow("张三", "女")
             .AddRow("李娜", "女");



            dt.TableName = "sheetone";
            dt2.TableName = "sheettwo";
            ds.Tables.Add(dt);
            ds.Tables.Add(dt2);

            //导出dt
            //  asr.Export("dt.xls", dt);
            //导出ds多sheet
            // asr.Export("ds.xls", ds);



            var d2t = asr.ReadDataTableExcel(@"C:\Users\jailall.sun\Desktop\月星\最新应聘者导入模板2015102205.xlsx");

        }

        protected void btnDt_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            var dt = DataTableSugar.CreateEasyTable("姓名", "姓别")
              .AddRow("张三", "女")
              .AddRow("李娜", "女");

            //导出dt
            asr.Export("dt.xls", dt);
            //导出ds多sheet
            // asr.Export("ds.xls", ds);

        }

        protected void btnDs_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            var dt = DataTableSugar.CreateEasyTable("姓名", "姓别")
              .AddRow("张三", "女")
              .AddRow("李娜", "女");
            var dt2 = DataTableSugar.CreateEasyTable("姓名", "姓别")
             .AddRow("张三", "女")
             .AddRow("李娜", "女");



            dt.TableName = "sheetone";
            dt2.TableName = "sheettwo";
            ds.Tables.Add(dt);
            ds.Tables.Add(dt2);

            //导出ds多sheet
            asr.Export("ds.xls", ds);
        }

        protected void btnMany_Click(object sender, EventArgs e)
        {

            //设置列
            List<ExcelColumns> columns = new List<ExcelColumns>();
            columns.Add(new ExcelColumns() { text = "id" });
            columns.Add(new ExcelColumns() { text = "name", columngroup = "namesex" });
            columns.Add(new ExcelColumns() { text = "sex", columngroup = "namesex" });
            columns.Add(new ExcelColumns() { text = "id2" });
            columns.Add(new ExcelColumns() { text = "cat", columngroup = "Animal" });
            columns.Add(new ExcelColumns() { text = "dog", columngroup = "Animal" });
            columns.Add(new ExcelColumns() { text = "rabbit", columngroup = "Animal" });
            columns.Add(new ExcelColumns() { text = "id3" });

            //设置分组
            List<ExcelColumnsGroup> group = new List<ExcelColumnsGroup>();
            group.Add(new ExcelColumnsGroup() { name = "Animal", text = "动物" });
            group.Add(new ExcelColumnsGroup() { name = "namesex", text = "名字性别" });

            //设置数据
            DataTable dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");
            dt.Columns.Add("sex");
            dt.Columns.Add("id2");
            dt.Columns.Add("cat");
            dt.Columns.Add("dog");
            dt.Columns.Add("rabbit");
            dt.Columns.Add("id3");
            var dr = dt.NewRow();
            dr[0] = 0;
            dr[1] = 1;
            dr[2] = 2;
            dr[3] = 3;
            dr[4] = 4;
            dr[5] = 5;
            dr[6] = 6;
            dr[7] = 7;
            dt.Rows.Add(dr);
            var dr2 = dt.NewRow();
            dr2[0] = 10;
            dr2[1] = 11;
            dr2[2] = 12;
            dr2[3] = 13;
            dr2[4] = 14;
            dr2[5] = 15;
            dr2[6] = 16;
            dr2[7] = 17;
            dt.Rows.Add(dr2);

            //导出复杂表头
            asr.ExportColumnsHierarchy("1.xls", columns, group, dt);

            
        }
    }
}
