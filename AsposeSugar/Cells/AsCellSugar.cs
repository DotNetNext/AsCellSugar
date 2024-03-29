﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.Data;
using System.Drawing;
using System.Web;
using System.IO;

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
                        workbook.Worksheets.Add("Sheet" + (i + 1));
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

        /// <summary>
        /// 读取DataTable
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public DataTable ReadDataTableExcel(string filePath)
        {
            var ds = ReadDataSetExcel(filePath);
            if (ds != null && ds.Tables.Count > 0)
                return ds.Tables[0];
            return null;
        }
        /// <summary>
        /// 读取DataDataSet
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public DataSet ReadDataSetExcel(string filePath)
        {
            //返回的Excel数据
            DataSet dsExcel = new DataSet();

            //创建一个Workbook和Worksheet对象
            Worksheet wkSheet = null;
            Workbook wkBook = new Workbook(filePath);

            //遍历读取sheet
            for (int i = 0; i < wkBook.Worksheets.Count; i++)
            {
                wkSheet = wkBook.Worksheets[i];

                //声明DataTable存放sheet
                DataTable dtTemp = new DataTable();
                //设置Table名为sheet的名称
                dtTemp.TableName = wkSheet.Name;

                //遍历行
                for (int x = 0; x < wkSheet.Cells.MaxDataRow + 1; x++)
                {
                    //声明DataRow存放sheet的数据行
                    DataRow dRow = null;

                    //遍历列
                    for (int y = 0; y < wkSheet.Cells.MaxDataColumn + 1; y++)
                    {
                        //获取单元格的值
                        string value = wkSheet.Cells[x, y].StringValue.Trim();

                        //如果是第一行，则当作表头
                        if (x == 0)
                        {
                            //设置表头
                            DataColumn dCol = new DataColumn(value);
                            dtTemp.Columns.Add(dCol);
                        }

                        //非第一行，则为数据行
                        else
                        {
                            //每次循环到第一列时，实例DataRow
                            if (y == 0)
                            {
                                dRow = dtTemp.NewRow();
                            }
                            //给第Y列赋值
                            dRow[y] = value;
                        }
                    }

                    if (dRow != null)
                    {
                        dtTemp.Rows.Add(dRow);
                    }
                }

                dsExcel.Tables.Add(dtTemp);
            }

            //释放对象
            wkSheet = null;
            wkBook = null;

            //返回数据
            return dsExcel;

        }

        /// <summary>
        /// 读取DataTable
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public DataTable ReadDataTableExcel(Stream stream)
        {
            var ds = ReadDataSetExcel(stream);
            if (ds != null && ds.Tables.Count > 0)
                return ds.Tables[0];
            return null;
        }


        /// <summary>
        /// 读取DataDataSet
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public DataSet ReadDataSetExcel(Stream stream)
        {
            //返回的Excel数据
            DataSet dsExcel = new DataSet();

            //创建一个Workbook和Worksheet对象
            Worksheet wkSheet = null;
            Workbook wkBook = new Workbook(stream);

            //遍历读取sheet
            for (int i = 0; i < wkBook.Worksheets.Count; i++)
            {
                wkSheet = wkBook.Worksheets[i];

                //声明DataTable存放sheet
                DataTable dtTemp = new DataTable();
                //设置Table名为sheet的名称
                dtTemp.TableName = wkSheet.Name;

                //遍历行
                for (int x = 0; x < wkSheet.Cells.MaxDataRow + 1; x++)
                {
                    //声明DataRow存放sheet的数据行
                    DataRow dRow = null;

                    //遍历列
                    for (int y = 0; y < wkSheet.Cells.MaxDataColumn + 1; y++)
                    {
                        //获取单元格的值
                        string value = wkSheet.Cells[x, y].StringValue.Trim();

                        //如果是第一行，则当作表头
                        if (x == 0)
                        {
                            //设置表头
                            DataColumn dCol = new DataColumn(value);
                            dtTemp.Columns.Add(dCol);
                        }

                        //非第一行，则为数据行
                        else
                        {
                            //每次循环到第一列时，实例DataRow
                            if (y == 0)
                            {
                                dRow = dtTemp.NewRow();
                            }
                            //给第Y列赋值
                            dRow[y] = value;
                        }
                    }

                    if (dRow != null)
                    {
                        dtTemp.Rows.Add(dRow);
                    }
                }

                dsExcel.Tables.Add(dtTemp);
            }

            //释放对象
            wkSheet = null;
            wkBook = null;

            //返回数据
            return dsExcel;

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
            if (dt.TableName != "Table1")
            {
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
