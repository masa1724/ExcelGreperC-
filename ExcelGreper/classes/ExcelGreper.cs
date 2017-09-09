using NPOI.DDF;
using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;

namespace ExcelGreper.classes
{
    class ExcelGreper
    {
        private IList<string> m_Keywords = new List<string>();
        private IList<string> m_FilePaths = new List<string>();

        public ExcelGreper(string keyword, string filePath)
        {
            m_Keywords.Add(keyword);
            m_FilePaths.Add(filePath);
        }

        public ExcelGreper(IList<string> Keywords, IList<string> FilePaths)
        {
            m_Keywords = Keywords;
            m_FilePaths = FilePaths;
        }

        public IList<GrepResult> Execute()
        {
            IList<GrepResult> resultList = new List<GrepResult>();
            IWorkbook book = null;

            try
            {
                foreach (string filePath in m_FilePaths)
                {
                    book = WorkbookFactory.Create(filePath);

                    foreach (ISheet sheet in book)
                    {
                        SearchCells(sheet, resultList, filePath);
                        SearchShapes(sheet, resultList, filePath);
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                if (book != null) book.Close();
            }

            return resultList;
        }

        private void SearchCells(ISheet sheet, IList<GrepResult> resultList, string filePath)
        {
            for (int r = 0; r <= sheet.LastRowNum; r++)
            {
                IRow row = sheet.GetRow(r);

                for (int c = 0; c <= row.LastCellNum; c++)
                {
                    ICell cell = row?.GetCell(c);
                    string value = GetCellValue(cell);

                    // キーワードが含まれているか判定
                    if (!IsHit(value))
                    {
                        continue;
                    }

                    GrepResult result = new GrepResult(
                        GrepResult.Types.Cell,
                        filePath,
                        new CellReference(sheet.SheetName, r, c, false, false).FormatAsString(),
                        "",
                        value
                    );

                    resultList.Add(result);
                }
            }
        }

        private void SearchShapes(ISheet sheet, IList<GrepResult> resultList, string filePath)
        {
            // シート上にあるすべてのシェイプを取得
            HSSFPatriarch part = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            IList<HSSFShape> shapes = part.Children;

            foreach (HSSFShape shape in shapes)
            {
                if (!(shape is HSSFSimpleShape))
                {
                    continue;
                }

                HSSFSimpleShape sshape = (HSSFSimpleShape)shape;
                string value = sshape.String.String;

                // キーワードが含まれているか判定
                if (!IsHit(value))
                {
                    continue;
                }

                // シェイプが配置されている絶対座標(左上)からセル座標を取得
                EscherClientAnchorRecord record =
                    (EscherClientAnchorRecord)ConvertAnchor.CreateAnchor(sshape.Anchor);
                int r = record.Row1;
                int c = record.Col1;

                GrepResult result = new GrepResult(
                    GrepResult.Types.Shape,
                    filePath,
                    new CellReference(sheet.SheetName, r, c, false, false).FormatAsString(),
                    "",
                    value
                );

                resultList.Add(result);
            }
        }

        private string GetCellValue(ICell cell)
        {
            if (cell == null) return null;

            string value = null;

            switch (cell.CellType)
            {
                // 数値, 日付
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        value = cell.DateCellValue.ToString();
                    }
                    else
                    {
                        value = cell.NumericCellValue.ToString();
                    }
                    break;
                // 計算式
                case CellType.Formula:
                    value = cell.CellFormula;
                    break;
                // 真偽
                case CellType.Boolean:
                    value = cell.BooleanCellValue.ToString();
                    break;
                // 文字列
                case CellType.String:
                    value = cell.StringCellValue;
                    break;
                // 空
                case CellType.Blank:
                    value = "";
                    break;
                default:
                    break;
            }

            return value;
        }

        private bool IsHit(string text)
        {
            if (text == null || text == "")
            {
                return false;
            }

            foreach (string keyword in m_Keywords)
            {
                if (text.IndexOf(keyword) >= 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
