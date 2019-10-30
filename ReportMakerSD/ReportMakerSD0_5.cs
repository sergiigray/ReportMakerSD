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
        public static string RunReportFive(DateTime dateTimeTo)
        {
            List<string> ColumnsReport0 = new List<string>() { "Дата окончания регистрации", "Номер", "Статус", "Организация заявителя", "Заявитель", "ФГП", "Исполнитель", "Услуга", "Тема", "Описание (для списков)", "Код ожидания", "Причина ожидания", "Срок ожидания", "Расчетная дата решения обращения", "Просрочен на (выполнение)" };

            // Сортируем по организации заявителя
            var SDR0Data0Sorted = MainForm.SDR0Data0.OrderByDescending(x => x.DateEndReg).ToList();
            //Формируем отчетный файл
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets.Add("ОДУ Северо-Запада");
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("Вне ОЗ Северо-Запада");

                // Добавляем шапку на лист1
                int countColumns = 0;
                foreach (string ColumnsReport in ColumnsReport0)
                {
                    worksheet1.Cells[1, countColumns + 1].Value = ColumnsReport;
                    if (ColumnsReport.IndexOf("Просро") >= 0)
                    {
                        worksheet1.Cells[1, countColumns + 1].Value = worksheet1.Cells[1, countColumns + 1].Value + " по состоянию на " + MainForm.SDR0Data;
                    }
                    countColumns++;
                }
                // Добавляем шапку на 2 лист
                countColumns = 0;
                foreach (string ColumnsReport in ColumnsReport0)
                {
                    worksheet2.Cells[1, countColumns + 1].Value = ColumnsReport;
                    if (ColumnsReport.IndexOf("Просро") >= 0)
                    {
                        worksheet2.Cells[1, countColumns + 1].Value = worksheet2.Cells[1, countColumns + 1].Value + " по состоянию на " + MainForm.SDR0Data;
                    }
                    countColumns++;
                }
                // Добавляем данные
                int count1 = 0, count2 = 0;
                DateTime dtDateTimeWeek = dateTimeTo.AddDays(7);

                foreach (var row in SDR0Data0Sorted)
                {
                    if (row.Status == "2 Назначен" || row.Status == "3 Выполняется" || row.Status == "4 В ожидании")
                    {
                        if (row.FSG.IndexOf("ОДУ Северо-Запада \\") >= 0)
                        {
                            worksheet1.Cells[2 + count1, 1].Value = row.DateEndReg;
                            worksheet1.Cells[2 + count1, 1].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                            worksheet1.Cells[2 + count1, 2].Value = row.Number;
                            worksheet1.Cells[2 + count1, 3].Value = row.Status;
                            worksheet1.Cells[2 + count1, 4].Value = row.Org;
                            worksheet1.Cells[2 + count1, 5].Value = row.Applicant;
                            worksheet1.Cells[2 + count1, 6].Value = row.FSG;
                            worksheet1.Cells[2 + count1, 7].Value = row.Executor;
                            worksheet1.Cells[2 + count1, 8].Value = row.Service;
                            worksheet1.Cells[2 + count1, 9].Value = row.Subject;
                            worksheet1.Cells[2 + count1, 10].Value = row.Description;
                            worksheet1.Cells[2 + count1, 11].Value = row.WaitingCode;
                            worksheet1.Cells[2 + count1, 12].Value = row.WaitingReason;
                            worksheet1.Cells[2 + count1, 14].Value = row.DateCalc;
                            worksheet1.Cells[2 + count1, 14].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                            worksheet1.Cells[2 + count1, 13].Value = row.DateWait;
                            worksheet1.Cells[2 + count1, 13].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                            worksheet1.Cells[2 + count1, 15].Value = row.Expired;
                            count1++;
                        }
                        else if ((row.FSG.IndexOf("ИА \\") >= 0) || (row.FSG.IndexOf("ПР \\") >= 0) || (row.FSG.IndexOf("Администраторы Naumen") >= 0))
                        {
                            worksheet2.Cells[2 + count2, 1].Value = row.DateEndReg;
                            worksheet2.Cells[2 + count2, 1].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                            worksheet2.Cells[2 + count2, 2].Value = row.Number;
                            worksheet2.Cells[2 + count2, 3].Value = row.Status;
                            worksheet2.Cells[2 + count2, 4].Value = row.Org;
                            worksheet2.Cells[2 + count2, 5].Value = row.Applicant;
                            worksheet2.Cells[2 + count2, 6].Value = row.FSG;
                            worksheet2.Cells[2 + count2, 7].Value = row.Executor;
                            worksheet2.Cells[2 + count2, 8].Value = row.Service;
                            worksheet2.Cells[2 + count2, 9].Value = row.Subject;
                            worksheet2.Cells[2 + count2, 10].Value = row.Description;
                            worksheet2.Cells[2 + count2, 11].Value = row.WaitingCode;
                            worksheet2.Cells[2 + count2, 12].Value = row.WaitingReason;
                            worksheet2.Cells[2 + count2, 14].Value = row.DateCalc;
                            worksheet2.Cells[2 + count2, 14].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                            worksheet2.Cells[2 + count2, 13].Value = row.DateWait;
                            worksheet2.Cells[2 + count2, 13].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";
                            worksheet2.Cells[2 + count2, 15].Value = row.Expired;
                            count2++;
                        }
                    }
                }

                // Красим 1 лист
                ExcelRange rg1 = worksheet1.Cells[1, 1, count1 + 1, 15];
                ExcelTable tab1 = worksheet1.Tables.Add(rg1, "Table1");
                tab1.TableStyle = TableStyles.Medium13;

                // Красим 2 лист
                ExcelRange rg2 = worksheet2.Cells[1, 1, count2 + 1, 15];
                ExcelTable tab2 = worksheet2.Tables.Add(rg2, "Table2");
                tab2.TableStyle = TableStyles.Medium13;

                worksheet1.Calculate();
                worksheet2.Calculate();

                worksheet1.View.FreezePanes(2, 1);
                worksheet1.View.ZoomScale = 90;
                worksheet2.View.FreezePanes(2, 1);
                worksheet2.View.ZoomScale = 90;

                worksheet1.Row(1).Height = 70;
                worksheet1.Row(1).Style.WrapText = true;
                worksheet1.Row(1).Style.Font.Bold = true;
                worksheet1.Cells.AutoFitColumns();  //Autofit columns for all cells
                worksheet1.Column(1).Width = 16;
                worksheet1.Column(2).Width = 10;
                worksheet1.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet1.Column(3).Width = 15;
                //worksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet1.Column(4).Width = 15;
                worksheet1.Column(5).Width = 25;
                worksheet1.Column(6).Width = 35;
                worksheet1.Column(7).Width = 25;
                worksheet1.Column(8).Width = 20;
                worksheet1.Column(9).Width = 20;
                worksheet1.Column(10).Width = 20;
                worksheet1.Column(11).Width = 20;
                worksheet1.Column(12).Width = 20;
                worksheet1.Column(15).Width = 20;
                worksheet1.Column(15).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet1.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet1.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet2.Row(1).Height = 70;
                worksheet2.Row(1).Style.WrapText = true;
                worksheet2.Row(1).Style.Font.Bold = true;
                worksheet2.Cells.AutoFitColumns();  //Autofit columns for all cells
                worksheet2.Column(1).Width = 16;
                worksheet2.Column(2).Width = 10;
                worksheet2.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet2.Column(3).Width = 15;
                //worksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet2.Column(4).Width = 15;
                worksheet2.Column(5).Width = 25;
                worksheet2.Column(6).Width = 35;
                worksheet2.Column(7).Width = 25;
                worksheet2.Column(8).Width = 20;
                worksheet2.Column(9).Width = 20;
                worksheet2.Column(10).Width = 20;
                worksheet2.Column(11).Width = 20;
                worksheet2.Column(12).Width = 20;
                worksheet2.Column(15).Width = 20;
                worksheet2.Column(15).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet2.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet2.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // Подготовка отчета для печати
                // Надо альбомную, все столбцы на страницу,
                worksheet1.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Нерешенные обращения по ФГП ОДУ";
                worksheet1.HeaderFooter.OddFooter.RightAlignedText = string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                worksheet1.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                worksheet1.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FileName;
                worksheet1.PrinterSettings.Orientation = eOrientation.Landscape;
                worksheet1.PrinterSettings.FitToPage = true;
                worksheet1.PrinterSettings.FitToWidth = 1;
                worksheet1.PrinterSettings.FitToHeight = 0;

                worksheet2.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"Нерешенные обращения сотрудников ОЗ вне ОЗ";
                worksheet2.HeaderFooter.OddFooter.RightAlignedText = string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                worksheet2.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                worksheet2.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FileName;
                worksheet2.PrinterSettings.Orientation = eOrientation.Landscape;
                worksheet2.PrinterSettings.FitToPage = true;

                //worksheet.PrinterSettings.RepeatRows = worksheet.Cells["3:3"];
                //worksheet.PrinterSettings.RepeatColumns = worksheet.Cells["A:G"];

                // Change the sheet view to show it in page layout mode
                //worksheet.View.PageLayoutView = true;

                // set some document properties
                package.Workbook.Properties.Title = "Нерешенные обращения по ОДУ";
                package.Workbook.Properties.Author = "Романов С.П.";
                package.Workbook.Properties.Comments = "Пример заполнения отчета в Excel 2007 используя EPPlus";

                // set some extended property values
                package.Workbook.Properties.Company = "ОДУ СЗ СО ЕЭС";

                // set some custom property values
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");
                package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");

                var xlFile = Utils.GetFileInfo("SD. Нерешенные обращения по ОДУ на " + dateTimeTo.ToString("dd.MM.yyyy") + ".xlsx");
                // save our new workbook in the output directory and we are done!
                package.SaveAs(xlFile);
                return xlFile.FullName;
            }
        }
    }
}
