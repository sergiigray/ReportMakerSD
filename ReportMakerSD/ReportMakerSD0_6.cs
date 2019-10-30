// SD. Отчет по ОДУ
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportMakerSD
{
    public static partial class SDR0
    {
        public static string RunReport6(DateTime dateTimeTo)
        {
            List<string> ColumnsReport0 = new List<string>() { "Дата окончания регистрации", "Номер", "Статус", "Организация заявителя", "Заявитель", "ФГП", "Исполнитель", "Услуга", "Тема", "Описание (для списков)", "Код ожидания", "Причина ожидания", "Срок ожидания", "Расчетная дата решения обращения", "Просрочен на (выполнение)" };

            // Сортируем по организации заявителя

            //var SDR0Data0Sorted = MainForm.SDR0Data0.OrderByDescending(x => x.DateEndReg).ToList();
            //var SDR0Data0Sorted = MainForm.SDR0Data0.Where(p => p.OAreaFSG == "ОЗ ОДУ Северо-Запада").OrderByDescending(x => x.DateEndReg).ToList();
            var SDR0Data0Sorted = from UserRequest in MainForm.SDR0Data0
                                  where UserRequest.OAreaFSG == "ОЗ ОДУ Северо-Запада"
                                  where (UserRequest.Status == "2 Назначен" || UserRequest.Status == "3 Выполняется" || UserRequest.Status == "4 В ожидании")
                                  orderby UserRequest.DateEndReg descending
                                  select UserRequest;

            //Формируем отчетный файл
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets.Add("Архангельское РДУ");
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("Балтийское РДУ");
                ExcelWorksheet worksheet3 = package.Workbook.Worksheets.Add("Карельское РДУ");
                ExcelWorksheet worksheet4 = package.Workbook.Worksheets.Add("Кольское РДУ");
                ExcelWorksheet worksheet5 = package.Workbook.Worksheets.Add("Коми РДУ");
                ExcelWorksheet worksheet6 = package.Workbook.Worksheets.Add("Ленинградское РДУ");
                ExcelWorksheet worksheet7 = package.Workbook.Worksheets.Add("Новгородское РДУ");

                // Добавляем шапки
                int countColumns = 0;
                foreach (ExcelWorksheet wrksheet in package.Workbook.Worksheets)
                {
                    foreach (string ColumnsReport in ColumnsReport0)
                    {
                        wrksheet.Cells[1, countColumns + 1].Value = ColumnsReport;
                        if (ColumnsReport.IndexOf("Просро") >= 0)
                        {
                            wrksheet.Cells[1, countColumns + 1].Value = wrksheet.Cells[1, countColumns + 1].Value + " по состоянию на " + MainForm.SDR0Data;
                        }
                        countColumns++;

                        wrksheet.Row(1).Height = 70;
                        wrksheet.Row(1).Style.WrapText = true;
                        wrksheet.Row(1).Style.Font.Bold = true;
                        wrksheet.Cells.AutoFitColumns();
                        wrksheet.Column(1).Width = 16;
                        wrksheet.Column(2).Width = 10;
                        wrksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wrksheet.Column(3).Width = 15;
                        wrksheet.Column(4).Width = 15;
                        wrksheet.Column(5).Width = 25;
                        wrksheet.Column(6).Width = 35;
                        wrksheet.Column(7).Width = 25;
                        wrksheet.Column(8).Width = 20;
                        wrksheet.Column(9).Width = 20;
                        wrksheet.Column(10).Width = 20;
                        wrksheet.Column(11).Width = 20;
                        wrksheet.Column(12).Width = 20;
                        wrksheet.Column(13).Width = 16;
                        wrksheet.Column(14).Width = 16;
                        wrksheet.Column(15).Width = 20;
                        wrksheet.Column(15).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        wrksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wrksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        wrksheet.View.FreezePanes(2, 1);
                        wrksheet.View.ZoomScale = 90;

                        wrksheet.HeaderFooter.OddFooter.RightAlignedText = string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                        wrksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                        wrksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FileName;
                        wrksheet.PrinterSettings.Orientation = eOrientation.Landscape;
                        wrksheet.PrinterSettings.FitToPage = true;
                        wrksheet.PrinterSettings.FitToWidth = 1;
                        wrksheet.PrinterSettings.FitToHeight = 0;
                    }
                    countColumns = 0;
                }

                // Добавляем данные
                int count = 0, count1 = 0, count2 = 0, count3 = 0, count4 = 0, count5 = 0, count6 = 0, count7 = 0;
                DateTime dtDateTimeWeek = dateTimeTo.AddDays(7);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                foreach (var row in SDR0Data0Sorted)
                {
                    //if (row.OAreaFSG != "ОЗ ОДУ Северо-Запада") continue; // Если ФГП не ОЗ, то берем следующую запись.

                    //    if (row.Status == "2 Назначен" || row.Status == "3 Выполняется" || row.Status == "4 В ожидании")
                    //    {
                    string[] fgp1 = row.FSG.Split('\\');
                    string fgp = fgp1[0].Trim();
                    switch (fgp)
                    {
                        case "Архангельское РДУ":
                            worksheet = package.Workbook.Worksheets[1];
                            count = count1;
                            count1++;
                            break;
                        case "Балтийское РДУ":
                            worksheet = package.Workbook.Worksheets[2];
                            count = count2;
                            count2++;
                            break;
                        case "Карельское РДУ":
                            worksheet = package.Workbook.Worksheets[3];
                            count = count3;
                            count3++;
                            break;
                        case "Кольское РДУ":
                            worksheet = package.Workbook.Worksheets[4];
                            count = count4;
                            count4++;
                            break;
                        case "Коми РДУ":
                            worksheet = package.Workbook.Worksheets[5];
                            count = count5;
                            count5++;
                            break;
                        case "Ленинградское РДУ":
                            worksheet = package.Workbook.Worksheets[6];
                            count = count6;
                            count6++;
                            break;
                        case "Новгородское РДУ":
                            worksheet = package.Workbook.Worksheets[7];
                            count = count7;
                            count7++;
                            break;
                        default:
                            continue;
                    }
                    worksheet.Cells[2 + count, 1].Value = row.DateEndReg;
                    worksheet.Cells[2 + count, 1].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                    worksheet.Cells[2 + count, 2].Value = row.Number;
                    worksheet.Cells[2 + count, 3].Value = row.Status;
                    worksheet.Cells[2 + count, 4].Value = row.Org;
                    worksheet.Cells[2 + count, 5].Value = row.Applicant;
                    worksheet.Cells[2 + count, 6].Value = row.FSG;
                    worksheet.Cells[2 + count, 7].Value = row.Executor;
                    worksheet.Cells[2 + count, 8].Value = row.Service;
                    worksheet.Cells[2 + count, 9].Value = row.Subject;
                    worksheet.Cells[2 + count, 10].Value = row.Description;
                    worksheet.Cells[2 + count, 11].Value = row.WaitingCode;
                    worksheet.Cells[2 + count, 12].Value = row.WaitingReason;
                    worksheet.Cells[2 + count, 14].Value = row.DateCalc;
                    worksheet.Cells[2 + count, 14].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                    worksheet.Cells[2 + count, 13].Value = row.DateWait;
                    worksheet.Cells[2 + count, 13].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                    worksheet.Cells[2 + count, 15].Value = row.Expired;
                    //}
                }

                // Красим 1 лист
                ExcelRange rg1 = worksheet1.Cells[1, 1, count1 + 1, 15];
                ExcelTable tab1 = worksheet1.Tables.Add(rg1, "Table1");
                tab1.TableStyle = TableStyles.Medium13;

                // Красим 2 лист
                ExcelRange rg2 = worksheet2.Cells[1, 1, count2 + 1, 15];
                ExcelTable tab2 = worksheet2.Tables.Add(rg2, "Table2");
                tab2.TableStyle = TableStyles.Medium13;

                // Красим 3 лист
                ExcelRange rg3 = worksheet3.Cells[1, 1, count3 + 1, 15];
                ExcelTable tab3 = worksheet3.Tables.Add(rg3, "Table3");
                tab3.TableStyle = TableStyles.Medium13;

                // Красим 4 лист
                ExcelRange rg4 = worksheet4.Cells[1, 1, count4 + 1, 15];
                ExcelTable tab4 = worksheet4.Tables.Add(rg4, "Table4");
                tab4.TableStyle = TableStyles.Medium13;

                // Красим 5 лист
                ExcelRange rg5 = worksheet5.Cells[1, 1, count5 + 1, 15];
                ExcelTable tab5 = worksheet5.Tables.Add(rg5, "Table5");
                tab5.TableStyle = TableStyles.Medium13;

                // Красим 6 лист
                ExcelRange rg6 = worksheet6.Cells[1, 1, count6 + 1, 15];
                ExcelTable tab6 = worksheet6.Tables.Add(rg6, "Table6");
                tab6.TableStyle = TableStyles.Medium13;

                // Красим 7 лист
                ExcelRange rg7 = worksheet7.Cells[1, 1, count7 + 1, 15];
                ExcelTable tab7 = worksheet7.Tables.Add(rg7, "Table7");
                tab7.TableStyle = TableStyles.Medium13;

                //worksheet2.Calculate();

                // Подготовка отчета для печати
                // Надо альбомную, все столбцы на страницу,
                worksheet1.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Нерешенные обращения Архангельского РДУ";
                worksheet2.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Нерешенные обращения Балтийского РДУ";
                worksheet3.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Нерешенные обращения Карельское РДУ";
                worksheet4.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Нерешенные обращения Кольское РДУ";
                worksheet5.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Нерешенные обращения Коми РДУ";
                worksheet6.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Нерешенные обращения Ленинградского РДУ";
                worksheet7.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Нерешенные обращения Новгородского РДУ";

                //worksheet.PrinterSettings.RepeatRows = worksheet.Cells["3:3"];
                //worksheet.PrinterSettings.RepeatColumns = worksheet.Cells["A:G"];

                // Change the sheet view to show it in page layout mode
                //worksheet.View.PageLayoutView = true;

                // set some document properties
                package.Workbook.Properties.Title = "Нерешенные обращения по РДУ";
                package.Workbook.Properties.Author = "Романов С.П.";
                package.Workbook.Properties.Comments = "Пример заполнения отчета в Excel 2007 используя EPPlus";

                // set some extended property values
                package.Workbook.Properties.Company = "ОДУ СЗ СО ЕЭС";

                // set some custom property values
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");
                package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");

                var xlFile = Utils.GetFileInfo("SD. Нерешенные обращения по РДУ на " + dateTimeTo.ToString("dd.MM.yyyy") + ".xlsx");
                // save our new workbook in the output directory and we are done!
                package.SaveAs(xlFile);
                return xlFile.FullName;
            }
        }
    }
}
