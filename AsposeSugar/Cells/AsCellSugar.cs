using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.Data;
using System.Drawing;
using System.Web;

namespace AsposeSugar
{
    public partial class AsCellSugar
    {


        public AsCellSugar()
        {

        }
        public AsCellSugar(Style thStyle, Style tdStyle)
        {
            _thStyle = thStyle;
            _tdStyle = tdStyle;
        }


        /// <summary>
        /// 导出EXCEL并且动态生成多级表头
        /// </summary>
        /// <param name="columns">列</param>
        /// <param name="group">分组</param>
        /// <param name="ds">DataSet</param>
        /// <param name="path">保存路径</param>
        public void Export(string fileName, DataSet ds)
        {
            Workbook workbook = new Workbook(); //工作簿
            if (ds != null && ds.Tables.Count > 0)
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    var dt = ds.Tables[i];
                    if (dt == null)
                        dt = new DataTable();
                    List<ExcelColumns> columns = new List<ExcelColumns>();
                    foreach (DataColumn it in dt.Columns)
                    {
                        ExcelColumns colItem = new ExcelColumns()
                        {
                            text = it.ColumnName,
                            field = it.ColumnName
                        };
                        columns.Add(colItem);
                    }
                    if (i > 0)
                        workbook.Worksheets.Add("Sheet"+(i+1));
                    Worksheet sheet = workbook.Worksheets[i]; //工作表
                   
                    SetSheet(columns, new List<ExcelColumnsGroup>(), dt, sheet);
                }
            }
            var response = HttpContext.Current.Response;
            response.Clear();
            response.Buffer = true;
            response.Charset = "utf-8";
            response.AppendHeader("Content-Disposition", "attachment;filename=" + fileName);
            response.ContentEncoding = System.Text.Encoding.UTF8;
            response.ContentType = "application/ms-excel";
            response.BinaryWrite(workbook.SaveToStream().ToArray());
            response.End();
        }

        /// <summary>
        /// 导出excel
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="columns">列</param>
        /// <param name="dt">数据源</param>
        public void Export(string fileName, DataTable dt)
        {
            if (dt == null)
                dt = new DataTable();
            List<ExcelColumns> columns = new List<ExcelColumns>();
            foreach (DataColumn it in dt.Columns)
            {
                ExcelColumns colItem = new ExcelColumns()
                {
                    text = it.ColumnName,
                    field = it.ColumnName
                };
                columns.Add(colItem);
            }
            ExportColumnsHierarchy(fileName, columns, new List<ExcelColumnsGroup>(), dt);
        }

        /// <summary>
        /// 保存excel
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="filePath">保存路径</param>
        /// <param name="dt">数据源</param>
        public void Save(string filePath, DataTable dt)
        {
            if (dt == null)
                dt = new DataTable();
            List<ExcelColumns> columns = new List<ExcelColumns>();
            foreach (DataColumn it in dt.Columns)
            {
                ExcelColumns colItem = new ExcelColumns()
                {
                    text = it.ColumnName,
                    field = it.ColumnName
                };
                columns.Add(colItem);
            }
            SaveColumnsHierarchy(columns, new List<ExcelColumnsGroup>(), dt, filePath);
        }

        /// <summary>
        /// 导出EXCEL并且动态生成多级表头
        /// </summary>
        /// <param name="columns">列</param>
        /// <param name="group">分组</param>
        /// <param name="ds">dataTable</param>
        /// <param name="path">保存路径</param>
        public void SaveColumnsHierarchy(DataSet ds, string path)
        {
            Workbook workbook = new Workbook(); //工作簿
            if (ds != null && ds.Tables.Count > 0)
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    var dt = ds.Tables[i];
                    if (dt == null)
                        dt = new DataTable();
                    List<ExcelColumns> columns = new List<ExcelColumns>();
                    foreach (DataColumn it in dt.Columns)
                    {
                        ExcelColumns colItem = new ExcelColumns()
                        {
                            text = it.ColumnName,
                            field = it.ColumnName
                        };
                        columns.Add(colItem);
                    }
                    if (i > 0)
                        workbook.Worksheets.Add("sheet" + i);
                    Worksheet sheet = workbook.Worksheets[i]; //工作表
                    sheet.Name = ds.Tables[i].TableName;
                    SetSheet(columns, new List<ExcelColumnsGroup>(), ds.Tables[i], sheet);
                }
            }
            workbook.Save(path);
        }

        /// <summary>
        /// 导出EXCEL并且动态生成多级表头
        /// </summary>
        /// <param name="columns">列</param>
        /// <param name="group">分组</param>
        /// <param name="dt">dataTable</param>
        /// <param name="path">保存路径</param>
        public void SaveColumnsHierarchy(List<ExcelColumns> columns, List<ExcelColumnsGroup> group, DataTable dt, string path)
        {

            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            SetSheet(columns, group, dt, sheet);
            workbook.Save(path);
        }



        /// <summary>
        /// 导出EXCEL并且动态生成多级表头
        /// </summary>
        /// <param name="columns">列</param>
        /// <param name="group">分组</param>
        /// <param name="dt">dataTable</param>
        /// <param name="path">保存路径</param>
        public void ExportColumnsHierarchy(string fileName, List<ExcelColumns> columns, List<ExcelColumnsGroup> group, DataTable dt)
        {

            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            SetSheet(columns, group, dt, sheet);
            var response = HttpContext.Current.Response;
            response.Clear();
            response.Buffer = true;
            response.Charset = "utf-8";
            response.AppendHeader("Content-Disposition", "attachment;filename=" + fileName);
            response.ContentEncoding = System.Text.Encoding.UTF8;
            response.ContentType = "application/ms-excel";
            response.BinaryWrite(workbook.SaveToStream().ToArray());
            response.End();
        }



        private Style _thStyle
        {
            get
            {
                Style s = new Style();
                s.Font.IsBold = true;
                s.Font.Name = "宋体";
                s.Font.Color = Color.Black;
                s.HorizontalAlignment = TextAlignmentType.Center;  //标题居中对齐
                return s;
            }
            set
            {
                _thStyle = value;
            }
        }

        private Style _tdStyle
        {
            get
            {
                Style s = new Style();
                return s;
            }
            set
            {
                _tdStyle = value;
            }
        }

        private void SetSheet(List<ExcelColumns> columns, List<ExcelColumnsGroup> group, DataTable dt, Worksheet sheet)
        {
            if (dt.TableName != "Table1") {
                sheet.Name = dt.TableName;
            }
            Cells cells = sheet.Cells;//单元格
            for (int i = 0; i <= dt.Rows.Count + 1; i++)
            {
                sheet.Cells.SetRowHeight(i, 30);
            }
            List<AsposeCellInfo> acList = new List<AsposeCellInfo>();
            List<string> acColumngroupHistoryList = new List<string>();
            int currentX = 0;
            foreach (var it in columns)
            {
                AsposeCellInfo ac = new AsposeCellInfo();
                ac.y = 0;
                if (it.columngroup == null)
                {
                    ac.text = it.text;
                    ac.x = currentX;
                    ac.xCount = 1;
                    acList.Add(ac);
                    currentX++;
                    ac.yCount = 2;
                }
                else if (!acColumngroupHistoryList.Contains(it.columngroup))//防止重复
                {
                    var sameCount = columns.Where(itit => itit.columngroup == it.columngroup).Count();
                    ac.text = group.First(itit => itit.name == it.columngroup).text;
                    ac.x = currentX;
                    ac.xCount = sameCount;
                    acList.Add(ac);
                    currentX = currentX + sameCount;
                    acColumngroupHistoryList.Add(it.columngroup);
                    ac.yCount = 1;
                    ac.groupName = it.columngroup;
                }
                else
                {
                    //暂无逻辑
                }
            }
            //表头
            foreach (var it in acList)
            {
                cells.Merge(it.y, it.x, it.yCount, it.xCount);//合并单元格
                cells[it.y, it.x].PutValue(it.text);//填写内容
                cells[it.y, it.x].SetStyle(_thStyle);
                if (!string.IsNullOrEmpty(it.groupName))
                {
                    var cols = columns.Where(itit => itit.columngroup == it.groupName).ToList();
                    foreach (var itit in cols)
                    {
                        var colsIndex = cols.IndexOf(itit);
                        cells[it.y + 1, it.x + colsIndex].PutValue(itit.text);//填写内容
                        cells[it.y + 1, it.x + colsIndex].SetStyle(_thStyle);
                    }
                }
            }
            //表格
            if (dt != null && dt.Rows.Count > 0)
            {
                var rowList = dt.AsEnumerable().ToList();
                foreach (var it in rowList)
                {
                    int dtIndex = rowList.IndexOf(it);
                    var dtColumns = dt.Columns.Cast<DataColumn>().ToList();
                    foreach (var itit in dtColumns)
                    {
                        var dtColumnsIndex = dtColumns.IndexOf(itit);
                        cells[2 + dtIndex, dtColumnsIndex].PutValue(it[dtColumnsIndex]);
                        cells[2 + dtIndex, dtColumnsIndex].SetStyle(_tdStyle);

                    }
                }
            }
        }


    }


}
