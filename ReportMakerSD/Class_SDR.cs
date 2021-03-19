using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportMakerSD
{
    public class Class_SDR
    {
        public string ReportMonth(DateTime dateTimeTo, string ReportMonthName)
        {
            List<string> ColumnsReport1 = new List<string>()
            {
                "Дата окончания регистрации",
                "Номер",
                "Организация заявителя",
                "Заявитель",
                "ФГП",
                "Исполнитель",
                "Услуга",
                "Тема",
                "Описание (для списков)",
                "Дата и время решения",
                "Решение (для списков)",
                "Оценка пользователя" };
            List<string> ColumnsReport3 = new List<string>()
            {
                "АСДОУ",
                "АСДУ - заявки и ремонты",
                "АСДУ - локальные системы",
                "АСДУ - ОИК",
                "АСДУ - телемеханика",
                "АСДУ - технологические системы",
                "АСДУ - централизованные системы",
                "Базовая поддержка пользователей",
                "Инженерно-хозяйственное обеспечение",
                "Информационная безопасность",
                "Каналы связи",
                "Обеспечение деятельности ИТ",
                "ОИК НП",
                "Поддержка аппаратных платформ",
                "Порталы и информационно - справочные системы",
                "Проекты",
                "Сети передачи данных",
                "Системное администрирование",
                "Телефония, конференцсвязь",
                "Техническое обслуживание",
                "Финансово - экономическая деятельность",
                "Разное"
            };

            // Сортируем по оперзоне и организации заявителя 
            var SDR0Data0Sorted = Form0.SDR0Data0.Where(x => x.OAreaApplicant != "").OrderBy(x => x.OAreaApplicant).ThenBy(x => x.Org).ToList();

            //Формируем отчетный файл
            using (var package = new ExcelPackage())
            {
                #region Заполняем 1 лист
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets.Add("Низкая оценка");
                int count1 = 0;
                foreach (var row in SDR0Data0Sorted)
                {
                    if ((row.Rating == "Удовлетворительное" || row.Rating == "Низкое") && row.OAreaFSG == "ОЗ ОДУ Северо-Запада")
                    {
                        worksheet1.Cells[2 + count1, 1].Value = row.DateEndReg;
                        worksheet1.Cells[2 + count1, 1].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                        worksheet1.Cells[2 + count1, 2].Value = row.Number;
                        worksheet1.Cells[2 + count1, 3].Value = row.Org;
                        worksheet1.Cells[2 + count1, 4].Value = row.Applicant;
                        worksheet1.Cells[2 + count1, 5].Value = row.FSG;
                        worksheet1.Cells[2 + count1, 6].Value = row.Executor;
                        worksheet1.Cells[2 + count1, 7].Value = row.Service;
                        worksheet1.Cells[2 + count1, 8].Value = row.Subject;
                        worksheet1.Cells[2 + count1, 9].Value = row.Description;
                        worksheet1.Cells[2 + count1, 10].Value = row.DateAnswer;
                        worksheet1.Cells[2 + count1, 11].Value = row.AnswerText;
                        worksheet1.Cells[2 + count1, 12].Value = row.Rating;
                        count1++;
                    }
                }

                // Красим 1 лист
                ExcelRange rg1 = worksheet1.Cells[1, 1, count1 + 1, 12];
                ExcelTable tab1 = worksheet1.Tables.Add(rg1, "Table1");
                tab1.TableStyle = TableStyles.Medium10;

                int countColumns1 = 0;
                foreach (string ColumnsReport in ColumnsReport1)
                {
                    worksheet1.Cells[1, countColumns1 + 1].Value = ColumnsReport;
                    countColumns1++;
                }
                worksheet1.Calculate();
                worksheet1.View.FreezePanes(2, 1);
                worksheet1.View.ZoomScale = 100;
                worksheet1.Row(1).Height = 70;
                worksheet1.Row(1).Style.WrapText = true;
                worksheet1.Row(1).Style.Font.Bold = true;
                worksheet1.Cells.AutoFitColumns();  //Autofit columns for all cells
                worksheet1.Column(1).Width = 15;
                worksheet1.Column(2).Width = 7;
                worksheet1.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet1.Column(3).Width = 20;
                worksheet1.Column(4).Width = 32;
                worksheet1.Column(5).Width = 25;
                worksheet1.Column(6).Width = 30;
                worksheet1.Column(7).Width = 32;
                worksheet1.Column(8).Width = 20;
                worksheet1.Column(9).Width = 20;
                worksheet1.Column(10).Width = 15;
                worksheet1.Column(11).Width = 20;
                worksheet1.Column(12).Width = 20;
                //worksheet.Column(15).Width = 20;
                //worksheet.Column(15).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet1.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet1.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet1.HeaderFooter.OddFooter.RightAlignedText = string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                worksheet1.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                worksheet1.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FileName;
                worksheet1.PrinterSettings.Orientation = eOrientation.Landscape;
                worksheet1.PrinterSettings.FitToPage = true;
                worksheet1.PrinterSettings.FitToWidth = 1;
                worksheet1.PrinterSettings.FitToHeight = 0;
                worksheet1.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\" Обращения с низкой оценкой";
                #endregion

                #region Заполняем 2 лист
                //Utils.DebugInfoWrite("Формируем 2 лист отчета за месяц.", true, true);

                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("Количество обращений по СО");

                // добавить сортировку по результату
                var SDR0DataGroupApplicant = from a in Form0.SDR0Data0
                                             group a by a.OAreaApplicant into b
                                             orderby b.Key
                                             where b.Key != ""
                                             select new { b.Key, RequestCount = b.Count() };
                //SDR0Data0.GroupBy(x => x.OAreaApplicant).Select(g => new { g.Key, RequestCount = g.Count() });

                int RowCount = 0;
                foreach (var row in SDR0DataGroupApplicant)
                {
                    worksheet2.Cells[2 + RowCount, 1].Value = row.Key;
                    worksheet2.Cells[2 + RowCount, 2].Value = row.RequestCount;
                    RowCount++;
                }
                string OAreaPrev = "ОЗ Исполнительный аппарат";
                {
                    //int RowCount = 1;
                    //foreach (var row in SDR0Data0Sorted)
                    //{
                    //    if (OAreaPrev == row.OAreaApplicant)
                    //    { count2_2++; }
                    //    else
                    //    {
                    //        count2_2 = 1;
                    //        count2++;
                    //        OAreaPrev = row.OAreaApplicant;
                    //    }
                    //    worksheet2.Cells[1 + count2, 1].Value = OAreaPrev;
                    //    worksheet2.Cells[1 + count2, 2].Value = count2_2;
                }
                //worksheet2.Cells[2, 1, RowCount + 1, 2].Sort(1);

                // Красим 2 лист
                ExcelRange rg2 = worksheet2.Cells[1, 1, RowCount + 2, 4];
                ExcelTable tab2 = worksheet2.Tables.Add(rg2, "Table2");
                tab2.TableStyle = TableStyles.Medium13;

                worksheet2.Cells[1, 1].Value = "Операционная зона";
                worksheet2.Cells[1, 2].Value = "Количество за отчетный месяц";
                worksheet2.Cells[1, 3].Value = "Количество за отчетный месяц - 1";
                worksheet2.Cells[1, 4].Value = "Количество за отчетный месяц - 2";
                worksheet2.Cells[RowCount + 2, 1].Value = "Общий итог";
                worksheet2.Cells[RowCount + 2, 2].Formula = "=SUM(" + worksheet2.Cells[2, 2].Address + ":" + worksheet2.Cells[RowCount + 1, 2].Address + ")";
                worksheet2.Cells[RowCount + 2, 3].Formula = "=SUM(" + worksheet2.Cells[2, 3].Address + ":" + worksheet2.Cells[RowCount + 1, 3].Address + ")";
                worksheet2.Cells[RowCount + 2, 4].Formula = "=SUM(" + worksheet2.Cells[2, 4].Address + ":" + worksheet2.Cells[RowCount + 1, 4].Address + ")";

                //worksheet2.Calculate();
                worksheet2.View.FreezePanes(2, 1);
                worksheet2.View.ZoomScale = 100;
                worksheet2.Row(1).Height = 30;
                worksheet2.Row(1).Style.WrapText = true;
                worksheet2.Row(1).Style.Font.Bold = true;
                worksheet2.Cells.AutoFitColumns();  //Autofit columns for all cells
                worksheet2.Column(1).Width = 30;
                worksheet2.Column(2).Width = 20;
                worksheet2.Column(3).Width = 23;
                worksheet2.Column(4).Width = 23;
                worksheet2.Row(RowCount + 2).Style.Font.Bold = true;
                worksheet2.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet2.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet2.HeaderFooter.OddFooter.RightAlignedText = string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                worksheet2.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                worksheet2.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FileName;
                worksheet2.PrinterSettings.Orientation = eOrientation.Landscape;
                worksheet2.PrinterSettings.FitToPage = true;
                worksheet2.PrinterSettings.FitToWidth = 1;
                worksheet2.PrinterSettings.FitToHeight = 0;
                worksheet2.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Количество обращений по СО";
                #endregion

                #region Заполняем 3 лист
                //Utils.DebugInfoWrite("Формируем 3 лист отчета за месяц.", true, true);
                ExcelWorksheet worksheet3 = package.Workbook.Worksheets.Add("Количество обращений по ОЗ С-З");
                int count3 = 1, count3_2 = 0;
                OAreaPrev = "Архангельское РДУ";
                foreach (var row in SDR0Data0Sorted)
                {
                    if (row.OAreaApplicant.IndexOf("ОЗ ОДУ Северо-Запада") >= 0)
                    {
                        if (OAreaPrev == row.Org)
                        { count3_2++; }
                        else
                        {
                            count3_2 = 1;
                            count3++;
                            OAreaPrev = row.Org;
                        }
                        worksheet3.Cells[1 + count3, 1].Value = OAreaPrev;
                        worksheet3.Cells[1 + count3, 2].Value = count3_2;
                    }
                }
                //worksheet3.Cells[2, 1, count3, 2].Sort(1);

                // Красим 2 лист
                ExcelRange rg3 = worksheet3.Cells[1, 1, count3 + 2, 4];
                ExcelTable tab3 = worksheet3.Tables.Add(rg3, "Table3");
                tab3.TableStyle = TableStyles.Medium13;

                worksheet3.Cells[1, 1].Value = "Филиалы";
                worksheet3.Cells[1, 2].Value = "Количество за отчетный месяц";
                worksheet3.Cells[1, 3].Value = "Количество за отчетный месяц - 1";
                worksheet3.Cells[1, 4].Value = "Количество за отчетный месяц - 2";
                count3++;
                worksheet3.Cells[count3 + 1, 1].Value = "Общий итог";
                worksheet3.Cells[count3 + 1, 2].Formula = "=SUM(" + worksheet3.Cells[2, 2].Address + ":" + worksheet3.Cells[count3, 2].Address + ")";
                worksheet3.Cells[count3 + 1, 3].Formula = "=SUM(" + worksheet3.Cells[2, 3].Address + ":" + worksheet3.Cells[count3, 3].Address + ")";
                worksheet3.Cells[count3 + 1, 4].Formula = "=SUM(" + worksheet3.Cells[2, 4].Address + ":" + worksheet3.Cells[count3, 4].Address + ")";

                //worksheet3.Calculate();
                worksheet3.View.FreezePanes(2, 1);
                worksheet3.View.ZoomScale = 100;
                worksheet3.Row(1).Height = 30;
                worksheet3.Row(1).Style.WrapText = true;
                worksheet3.Row(1).Style.Font.Bold = true;
                worksheet3.Cells.AutoFitColumns();  //Autofit columns for all cells
                worksheet3.Column(1).Width = 30;
                worksheet3.Column(2).Width = 20;
                worksheet3.Column(3).Width = 23;
                worksheet3.Column(4).Width = 23;
                worksheet3.Row(count3 + 1).Style.Font.Bold = true;
                worksheet3.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet3.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet3.HeaderFooter.OddFooter.RightAlignedText = string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                worksheet3.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                worksheet3.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FileName;
                worksheet3.PrinterSettings.Orientation = eOrientation.Landscape;
                worksheet3.PrinterSettings.FitToPage = true;
                worksheet3.PrinterSettings.FitToWidth = 1;
                worksheet3.PrinterSettings.FitToHeight = 0;
                worksheet3.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Количество обращений по ОЗ";
                #endregion

                #region Заполняем 4 лист
                //Utils.DebugInfoWrite("Формируем 4 лист отчета за месяц.", true, true);
                ExcelWorksheet worksheet4 = package.Workbook.Worksheets.Add("СО по группам услуг");

                int count4 = 0;
                OAreaPrev = "ОЗ Исполнительный аппарат";
                foreach (var row in SDR0Data0Sorted)
                {
                    if (OAreaPrev != row.OAreaApplicant && row.OAreaApplicant != "")
                    {
                        count4++;
                        OAreaPrev = row.OAreaApplicant;
                    }
                    worksheet4.Cells[count4 + 2, 1].Value = OAreaPrev;
                    for (int i = 0; i < ColumnsReport3.Count; i++)
                    {
                        if (row.ServiceGroup == ColumnsReport3[i])
                        {
                            worksheet4.Cells[count4 + 2, i + 2].Value = worksheet4.Cells[count4 + 2, i + 2].GetValue<int>() + 1;
                        }
                    }
                }
                worksheet4.Cells[count4 + 3, 1].Value = "ИТОГО";
                for (int i = 0; i <= count4 + 1; i++)
                {
                    worksheet4.Cells[i + 1, ColumnsReport3.Count + 2].Formula = "=SUM(" + worksheet4.Cells[i + 1, 2].Address + ":" + worksheet4.Cells[i + 1, ColumnsReport3.Count + 1].Address + ")";
                }

                //// Красим 4 лист
                ExcelRange rg4 = worksheet4.Cells[1, 1, count4 + 3, ColumnsReport3.Count + 2];
                ExcelTable tab4 = worksheet4.Tables.Add(rg4, "Table4");
                tab4.TableStyle = TableStyles.Medium13;

                int countColumns = 2;
                worksheet4.Cells[1, countColumns - 1].Value = "Операционная зона";
                foreach (string ColumnsReport in ColumnsReport3)
                {
                    worksheet4.Cells[1, countColumns].Value = ColumnsReport;
                    worksheet4.Cells[1, countColumns].Style.TextRotation = 90;
                    worksheet4.Cells[count4 + 3, countColumns].Formula = "=SUM(" + worksheet4.Cells[2, countColumns].Address + ":" + worksheet4.Cells[count4 + 2, countColumns].Address + ")";
                    countColumns++;
                }
                worksheet4.Cells[count4 + 3, countColumns].Formula = "=SUM(" + worksheet4.Cells[2, countColumns].Address + ":" + worksheet4.Cells[count4 + 2, countColumns].Address + ")";
                worksheet4.Cells[1, countColumns].Value = "ИТОГО";


                //worksheet4.Calculate();
                worksheet4.View.FreezePanes(2, 1);
                worksheet4.View.ZoomScale = 100;
                worksheet4.Row(1).Height = 160;
                worksheet4.Row(1).Style.WrapText = true;
                worksheet4.Row(1).Style.Font.Bold = true;
                worksheet4.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet4.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet4.Row(count4 + 3).Style.Font.Bold = true;
                worksheet4.Cells.AutoFitColumns();  //Autofit columns for all cells
                worksheet4.Column(1).Width = 30;
                worksheet4.Column(countColumns).Style.Font.Bold = true;
                worksheet4.HeaderFooter.OddFooter.RightAlignedText = string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                worksheet4.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                worksheet4.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FileName;
                worksheet4.PrinterSettings.Orientation = eOrientation.Landscape;
                worksheet4.PrinterSettings.FitToPage = true;
                worksheet4.PrinterSettings.FitToWidth = 1;
                worksheet4.PrinterSettings.FitToHeight = 0;
                worksheet4.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Количество обращений по ОЗ и группам услуг";
                #endregion

                #region Заполняем 5 лист
                //Utils.DebugInfoWrite("Формируем 5 лист отчета за месяц.", true, true);
                ExcelWorksheet worksheet5 = package.Workbook.Worksheets.Add("ОЗ С-З по группам услуг");

                int count5 = 0;
                string OrgPrev = "Архангельское РДУ";
                foreach (var row in SDR0Data0Sorted)
                {
                    if (row.OAreaApplicant.IndexOf("ОЗ ОДУ Северо-Запада") >= 0)
                    {
                        if (OrgPrev != row.Org)
                        {
                            count5++;
                            OrgPrev = row.Org;
                        }
                        worksheet5.Cells[count5 + 2, 1].Value = OrgPrev;
                        for (int i = 0; i < ColumnsReport3.Count; i++)
                        {
                            if (row.ServiceGroup == ColumnsReport3[i])
                            {
                                worksheet5.Cells[count5 + 2, i + 2].Value = worksheet5.Cells[count5 + 2, i + 2].GetValue<int>() + 1;
                            }
                        }
                    }
                }
                worksheet5.Cells[count5 + 3, 1].Value = "ИТОГО";
                for (int i = 0; i <= count5 + 1; i++)
                {
                    worksheet5.Cells[i + 1, ColumnsReport3.Count + 2].Formula = "=SUM(" + worksheet5.Cells[i + 1, 2].Address + ":" + worksheet5.Cells[i + 1, ColumnsReport3.Count + 1].Address + ")";
                }

                //// Красим 5 лист
                ExcelRange rg5 = worksheet5.Cells[1, 1, count5 + 3, ColumnsReport3.Count + 2];
                ExcelTable tab5 = worksheet5.Tables.Add(rg5, "Table5");
                tab5.TableStyle = TableStyles.Medium13;

                countColumns = 2;
                worksheet5.Cells[1, countColumns - 1].Value = "Филиалы";
                foreach (string ColumnsReport in ColumnsReport3)
                {
                    worksheet5.Cells[1, countColumns].Value = ColumnsReport;
                    worksheet5.Cells[1, countColumns].Style.TextRotation = 90;
                    worksheet5.Cells[count5 + 3, countColumns].Formula = "=SUM(" + worksheet5.Cells[2, countColumns].Address + ":" + worksheet5.Cells[count5 + 2, countColumns].Address + ")";
                    countColumns++;
                }
                worksheet5.Cells[count5 + 3, countColumns].Formula = "=SUM(" + worksheet5.Cells[2, countColumns].Address + ":" + worksheet5.Cells[count5 + 2, countColumns].Address + ")";
                worksheet5.Cells[1, countColumns].Value = "ИТОГО";


                //worksheet5.Calculate();
                worksheet5.View.FreezePanes(2, 1);
                worksheet5.View.ZoomScale = 100;
                worksheet5.Row(1).Height = 160;
                worksheet5.Row(1).Style.WrapText = true;
                worksheet5.Row(1).Style.Font.Bold = true;
                worksheet5.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet5.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet5.Row(count5 + 3).Style.Font.Bold = true;
                worksheet5.Cells.AutoFitColumns();  //Autofit columns for all cells
                worksheet5.Column(1).Width = 40;
                worksheet5.Column(countColumns).Style.Font.Bold = true;
                worksheet5.HeaderFooter.OddFooter.RightAlignedText = string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                worksheet5.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                worksheet5.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FileName;
                worksheet5.PrinterSettings.Orientation = eOrientation.Landscape;
                worksheet5.PrinterSettings.FitToPage = true;
                worksheet5.PrinterSettings.FitToWidth = 1;
                worksheet5.PrinterSettings.FitToHeight = 0;
                worksheet5.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Количество обращений по филиалам и группам услуг";
                #endregion

                // set some document properties
                package.Workbook.Properties.Title = "Отчеты за месяц";
                package.Workbook.Properties.Author = "Романов С.П.";
                package.Workbook.Properties.Comments = "Пример заполнения отчета в Excel 2007 используя EPPlus";

                // set some extended property values
                package.Workbook.Properties.Company = "ОДУ СЗ СО ЕЭС";

                // set some custom property values
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");
                package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");

                var xlFile = Utils.GetFileInfo(ReportMonthName + ".xlsx");
                // save our new workbook in the output directory and we are done!
                package.SaveAs(xlFile);
                return xlFile.FullName;
            }
        }
    }
}

