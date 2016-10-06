using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

namespace ReportMakerSD
{
    public partial class MainForm : Form
    {
        private object[,] NameColumnsONLY;
        private int CountNameColumnsONLY;
        private string[] StatusesNames;
        private int CountStatusesNames;
        private object[,] NameFilialBranch;
        private int CountNameFilialBranch;
        private string[] FGPNames;
        private int CountFGPNames;
        private string[] WayOfGivingTreatments;
        private int CountWayOfGivingTreatments;
        private string stringTO = Properties.Settings.Default.TO;

        public MainForm()
        //Конструктор
        {
            InitializeComponent();
            Text = " " + Properties.Settings.Default.ProjectName; //изменяем наименование формы.

            //Определяем папку для хранения отчетов экспорта из СД и место хранения EXE.
            string PathToFileExportSD = @"D:\";
            string PathToEXE = Environment.CurrentDirectory;
            if (Properties.Settings.Default.FolderWithExportSdDefault.ToString().Contains(@":\")) //Если указан прямой путь, то хорошо.
            {
                PathToFileExportSD = Properties.Settings.Default.FolderWithExportSdDefault.ToString();
            }
            else //Если прямой путь к отчетам не указан, то шагаем до папки из параметров программы.
            {
                PathToFileExportSD = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString();
                PathToFileExportSD = PathToFileExportSD + @"\" + Properties.Settings.Default.FolderWithExportSdDefault.ToString();
            }
            if (Directory.Exists(PathToFileExportSD) == false) //если вдруг путь до отчетов не действителен, то пальцем в небо.
            {
                PathToFileExportSD = @"C:\";
                Properties.Settings.Default.FolderWithExportSdDefault = PathToFileExportSD;
            }

            textBox_FileExportSDName.Text = PathToFileExportSD;
            openFileDialog_ExportSDSelect.InitialDirectory = PathToFileExportSD;
            
            //Определяем свежий отчет экспорта из СД.
            DirectoryInfo fileSystemInfo = new DirectoryInfo(PathToFileExportSD);
            DateTime dt = DateTime.Now.AddDays(-7);
            foreach (FileSystemInfo fileSI in fileSystemInfo.GetFiles("exportSD*.xl*"))
            {
                //if (fileSI.Extension == ".xls" | fileSI.Extension == ".xlsx")//добавить нужные форматы
                //{
                    if (dt < Convert.ToDateTime(fileSI.CreationTime))
                    {
                        dt = Convert.ToDateTime(fileSI.CreationTime);
                        textBox_FileExportSDName.Text = fileSI.Name;
                    }
                //}
            }
            textBox_FileExportSDName.Text = PathToFileExportSD +@"\"+ textBox_FileExportSDName.Text;
            openFileDialog_ExportSDSelect.FileName = textBox_FileExportSDName.Text;

            dateTimePicker_TimeFrom.Format = DateTimePickerFormat.Custom;
            dateTimePicker_TimeFrom.CustomFormat = "dd.MM.yyyy HH:mm";
            dateTimePicker_TimeTo.Format = DateTimePickerFormat.Custom;
            dateTimePicker_TimeTo.CustomFormat = "dd.MM.yyyy HH:mm";
            switch (DateTime.Now.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    dateTimePicker_TimeFrom.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day - 7, Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    dateTimePicker_TimeTo.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, Properties.Settings.Default.TimeFromHour - 1, Properties.Settings.Default.TimeFromMinute - 1, 0);
                    break;
                default:
                    dateTimePicker_TimeFrom.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day - Convert.ToInt16(DateTime.Now.DayOfWeek) + 1, Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    dateTimePicker_TimeTo.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, Properties.Settings.Default.TimeToHour, Properties.Settings.Default.TimeToMinute, 0);
                    break;
            }

            //Определяем соответствуют ли даты отчету.
            TestFileExportSDandDateTime();

            //Подготовительная часть завершена.
            wStr("Часть 1 - завершена.");

            wStr("\r\nЧасть 2 - Проверка и формирование справочников.");

            wStr("\r\nОбязательные поля:");
            //Указатели обязательных полей на №столбцов в исходных данных
            //Остальные колонки:Дата окончания регистрации:Номер:Статус:Организация заявителя:Заявитель:ФГП:Исполнитель:Услуга:Тема:Дата и время решения:Решение:Способ подачи обращения
            //Наименование,№столбца в исходных данных (для первого=кол-во необязательных),
            NameColumnsONLY = new object[Properties.Settings.Default.ColumnsFilterNames.Split(':').Length, 3];
            int count = 0;
            foreach (string NameColumnONLY in Properties.Settings.Default.ColumnsFilterNames.Split(':'))
            {
                NameColumnsONLY[count, 0] = NameColumnONLY;
                NameColumnsONLY[count, 1] = 0;
                NameColumnsONLY[count, 2] = 0;
                wStr("   " + count.ToString() + " - '" + NameColumnsONLY[count, 0].ToString() + "';");
                count++;
            }
            CountNameColumnsONLY = count;
            wStr("Итого полей: " + CountNameColumnsONLY);

            wStr("\r\nСправочник статусов:");
            StatusesNames = Properties.Settings.Default.StatusesNames.Split(':');
            count = 0;
            foreach (string StatusNames in StatusesNames)
            { wStr("   '" + StatusNames + ";"); count++; }
            CountStatusesNames = count;
            wStr("Итого статусов: " + CountStatusesNames);

            wStr("\r\nСправочник филиалов:");
            //Наименование филиала (замена nameODU &etc), строки (замена rowNumberODU&etc), счетчик (замена countODU &etc)
            //NameFilialBranch[0,0].ToString(), Convert.ToInt32(NameFilialBranch[0, 1]), Convert.ToInt32(NameFilialBranch[0, 2])
            NameFilialBranch = new object[Properties.Settings.Default.FilialBranchName.Split(':').Length, 3];
            count = 0;
            foreach (string NameFilialBranch_ in Properties.Settings.Default.FilialBranchName.Split(':'))
            {
                NameFilialBranch[count, 0] = NameFilialBranch_;
                NameFilialBranch[count, 1] = "";
                NameFilialBranch[count, 2] = 0;
                wStr("   " + count.ToString() + " - '" + NameFilialBranch[count, 0].ToString() + "';");
                count++;
            }
            CountNameFilialBranch = count;
            wStr("Итого филиалов: " + CountNameFilialBranch);

            wStr("\r\nСправочник ФГП:");
            FGPNames = Properties.Settings.Default.FGPNames.Split(':');
            count = 0;
            foreach (string FGPName in FGPNames)
            { wStr("   '" + FGPName + ";"); count++; }
            CountFGPNames = count;
            wStr("Итого ФГП: " + CountFGPNames);

            wStr("\r\nCпособы подачи обращения:");
            WayOfGivingTreatments = Properties.Settings.Default.WayOfGivingTreatments.Split(':');
            count = 0;
            foreach (string WayOfGivingTreatment in WayOfGivingTreatments)
            { wStr("   '" + WayOfGivingTreatment + ";"); count++; }
            CountWayOfGivingTreatments = count;
            wStr("Итого способов: " + CountWayOfGivingTreatments);

            //Определяем добавочные отчеты.
            wStr("\r\nОпределяем дополнительные отчеты:");
            checkBox_ReportTO.Checked = Properties.Settings.Default.ReportTO;
            wStr(" - Отчет по подрядным организациям;");
            checkBox_ReportRDU.Checked = Properties.Settings.Default.ReportRDU;
            wStr(" - Отчет по нерешенным обращениям;");
            wStr("\r\nЖелательно закрыть EXCEL!");
        }

        void wStr(string text)
        //метод для заполнения текстового окна отладочной информацией.
        {
            textBox_DebugInfo.AppendText(text + "\n");
            //textBox_DebugInfo.SelectionStart = textBox_DebugInfo.Text.Length;
            //textBox_DebugInfo.SelectedText = textBox_DebugInfo.Text + " ";
            //textBox_DebugInfo.ScrollToCaret();
            textBox_DebugInfo.Refresh();
        }

        private void TestFileExportSDandDateTime( )
        //Тестирование даты в наименовании файла экспорта из СД и указанных дат в форме.
        {

        }

        private object[,] LoadDataFromFileExportSD(string FileNameExportSD)
        //Загрузка данных из файла.
        //Передаем полный путь к файлу с данными
        //Получаем массив с данными
        {
            Microsoft.Office.Interop.Excel.Application oExcelApp = null;
            Workbooks openbooks = null;
            Workbook openbook = null;
            Sheets openSheets = null;
            Worksheet openSheet = null;
            Range openCellsFirst = null;
            object[,] openData = null;
            try
            {
                oExcelApp = new Microsoft.Office.Interop.Excel.Application();
                oExcelApp.DisplayAlerts = false;
                oExcelApp.Visible = false;
                //FileName = "D:\\0\\export.xlsx";
                //Excel.Workbook openBook = excelApp.Workbooks.Open(openFileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                openbooks = oExcelApp.Workbooks;
                openbook = openbooks.Open(FileNameExportSD);
                openSheets = openbook.Sheets;
                //openSheet = openbook.Sheets[1];
                openSheet = openSheets[1];
                openCellsFirst = openSheet.Cells[1, 1];
                //Ranges openRange = null;
                //openRange = openSheet.Range[openSheet.Cells[1, 1], openSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell)].Value;
                openData = openSheet.Range[openCellsFirst, openSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell)].Value;
            }
                finally
            {
                openbook.Close(false, false, false);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openCellsFirst);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openSheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openSheets);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openbooks);
                ////Подготовка к убийству процесса Excel
                //int ExcelPID = 0;
                //int Hwnd = 0;
                //Hwnd = application.Hwnd;
                //System.Diagnostics.Process ExcelProcess;
                //GetWindowThreadProcessId((IntPtr)Hwnd, out ExcelPID);
                //ExcelProcess = System.Diagnostics.Process.GetProcessById(ExcelPID);
                ////Конец подготовки к убийству процесса Excel
                oExcelApp.DisplayAlerts = true;
                oExcelApp.Visible = true;
                oExcelApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oExcelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                openCellsFirst = null;
                openSheet = null;
                openSheets = null;
                openbook = null;
                openbooks = null;
                oExcelApp = null;
				GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.GetTotalMemory(true);
                ////Убийство процесса Excel
                //ExcelProcess.Kill();
                //ExcelProcess = null;
            }
            return openData;
        }
        
        private void SaveDataToReportFile(string ReportFileName, string sheetNames, object[,,] ReportSaveData)
        //Формирование XLS-файла и выгрузка данных
        {
            Microsoft.Office.Interop.Excel.Application oExcelApp = null;
            Workbooks openbooks = null;
            Workbook openbook = null;
            Sheets openSheets = null;
            Range openRange = null;

            int NumberSheets = ReportSaveData.GetLength(0);
            int NumberRow = ReportSaveData.GetLength(1);
            int NumberCol = ReportSaveData.GetLength(2);
            object[,] DataForSheet = new object[NumberRow, NumberCol];
            string[] sheetsNames = sheetNames.Split(',');

            try
            {
                oExcelApp = new Microsoft.Office.Interop.Excel.Application();
                // true => false
                oExcelApp.DisplayAlerts = false;
                oExcelApp.Visible = false;
                //Кол-во листов в новой книге
                oExcelApp.SheetsInNewWorkbook = NumberSheets;
                //Ширина колонок
                //oExcelApp.Columns.ColumnWidth = 15;

                openbooks = oExcelApp.Workbooks;
                //создаем новую книгу
                openbook = openbooks.Add();
                openSheets = openbook.Sheets;
                for (int i=1; i<=NumberSheets; i++)
                {
                    //Задаем наименование листа
                    openbook.Sheets[i].Name = sheetsNames[i-1];
                    //Заполняем шапку
                    for(int n=1;n<NameColumnsONLY.GetLength(0);n++)
                    {
                        openbook.Sheets[i].cells[1, n].Value = NameColumnsONLY[n, 0];
                        openbook.Sheets[i].cells[1, n].WrapText = true;
                        openbook.Sheets[i].cells[1, n].Font.Bold = true;
                        openbook.Sheets[i].cells[1, n].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        openbook.Sheets[i].cells[1, n].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        openbook.Sheets[i].cells[1, n].Borders.LineStyle = XlLineStyle.xlContinuous;
                        openbook.Sheets[i].cells[1, n].Interior.Color = ColorTranslator.ToOle(Color.CornflowerBlue);
                        openbook.Sheets[i].Columns[n].ColumnWidth = 20;
                    }
                    //Формируем необходимый массив с данными для листа
                    for (int r=0;r<NumberRow;r++) //перебираем все строки
                    {
                        for (int c=0;c<NumberCol-1;c++)
                        {
                            DataForSheet[r, c] = ReportSaveData[i-1, r, c+1];
                        }
                    }
                    openbook.Sheets[i].Range[openbook.Sheets[i].cells[2,1], openbook.Sheets[i].cells[NumberRow, NumberCol]].Value = DataForSheet;
                    //Форматирование
                    //по строчное. Долго выполняется.
                    for (int r = 2; r <= NumberRow; r++)
                    {
                        openRange = openbook.Sheets[i].Range[openbook.Sheets[i].cells[r, 1], openbook.Sheets[i].cells[r, NumberCol - 1]];
                        var t = openbook.Sheets[i].cells[r, 1].value;
                        if (t == null) break;
                        openRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        openRange.WrapText = null;
                        if(r % 2 != 0)
                        {
                            openRange.Interior.Color = ColorTranslator.ToOle(Color.LightSkyBlue);
                        }
                        else
                        {
                            openRange.Interior.Color = ColorTranslator.ToOle(Color.MintCream);
                        }
                    }
                    //openbook.Sheets[i].Range[openbook.Sheets[i].cells[r, 1], openbook.Sheets[i].cells[r, NumberCol]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    //openbook.Sheets[i].Range[openbook.Sheets[i].cells[2, 1], openbook.Sheets[i].cells[NumberRow, NumberCol]].
                }
                //Сохранение файла
                openbook.SaveAs((Environment.CurrentDirectory + "\\" + ReportFileName + ".xlsx"));//Возникает вопрос о перезаписи?
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                openbook.Close(false, false, false);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openRange);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openSheets);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openbooks);
                oExcelApp.DisplayAlerts = true;
                oExcelApp.Visible = true;
                //oExcelApp.Interactive = true;
                //oExcelApp.ScreenUpdating = true;
                //oExcelApp.UserControl = true;
                oExcelApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oExcelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                openRange = null;
                openSheets = null;
                openbook = null;
                openbooks = null;
                oExcelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.GetTotalMemory(true);
            }
        }

        private void button_ExportSDSelect_Click(object sender, EventArgs e)
        //Действия по выбору файла экспорта при нажатии кнопки
        {
            openFileDialog_ExportSDSelect.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog_ExportSDSelect.ShowDialog() == DialogResult.OK)
            {
                textBox_FileExportSDName.Text = openFileDialog_ExportSDSelect.FileName;
                //? Properties.Settings.Default.FolderWithExportSdDefault = openFile.FileName;
                //textBox_FileExportSDName.Text = Path.GetFileName(openFileName);

                //Если пришлось выбрать файл в ручную, то поправляем дату согласно дате отчета
                TestFileExportSDandDateTime();
            }
        }

        private void button_SendMail_Click(object sender, EventArgs e)
        //Отправка писем с отчетами.
        //!!!сделаем потом
        {
            //string usersODU = "kava@odusz.so-ups.ru; adv@odusz.so-ups.ru; nikolaev@odusz.so-ups.ru; smirnov@odusz.so-ups.ru; makarov@odusz.so-ups.ru; hicom@odusz.so-ups.ru; bsp@odusz.so-ups.ru; maksim@odusz.so-ups.ru; chulkov@odusz.so-ups.ru; maltsev@odusz.so-ups.ru";
            //string usCopyODU = "info_sd@odusz.so-ups.ru; blinov-al@odusz.so-ups.ru";

            //string usersRDU = "vif@arhrdu.so-ups.ru; glebov@balticrdu.so-ups.ru; shlyakov@karelia.so-ups.ru; iklouzay@kola.so-ups.ru; potoskuev@komirdu.so-ups.ru; msa@lenrdu.so-ups.ru; KuzminVA@novrdu.so-ups.ru";
            //string usCopyRDU = "maksim@odusz.so-ups.ru; info_sd@odusz.so-ups.ru";

            //string usersTO = "kava@odusz.so-ups.ru; adv@odusz.so-ups.ru; nikolaev@odusz.so-ups.ru; smirnov@odusz.so-ups.ru; makarov@odusz.so-ups.ru; hicom@odusz.so-ups.ru; bsp@odusz.so-ups.ru; maksim@odusz.so-ups.ru; chulkov@odusz.so-ups.ru; maltsev@odusz.so-ups.ru; vif@arhrdu.so-ups.ru; glebov@balticrdu.so-ups.ru; shlyakov@karelia.so-ups.ru; iklouzay@kola.so-ups.ru; potoskuev@komirdu.so-ups.ru; msa@lenrdu.so-ups.ru; KuzminVA@novrdu.so-ups.ru";
            //string usCopyTO = "lopintsev@odusz.so-ups.ru; pyatnitskiy-ev@odusz.so-ups.ru; blinov-al@odusz.so-ups.ru; info_sd@odusz.so-ups.ru";

            //string signature = "С уважением,\nСпециалист ТРСОО\nГруппы обработки обращений СОЭ АСУ,\nт. 39-99";

            //Outlook.Application outApp = new Outlook.Application();

            //Outlook.MailItem newMailODU = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
            //newMailODU.To = usersODU;
            //newMailODU.CC = usCopyODU;
            //newMailODU.Subject = "Нерешенные обращения по ОДУ Северо-Запада с " + DateFromPicker.Value.ToString("dd.MM.yyyy") + " по " + DateToPicker.Value.ToString("dd.MM.yyyy");
            //newMailODU.Body = "Нерешенные обращения по ОДУ Северо-Запада с " + DateFromPicker.Value.ToString("dd.MM.yyyy") + " по " + DateToPicker.Value.ToString("dd.MM.yyyy") + ".\n\n" + signature;
            //newMailODU.Attachments.Add(newBookODUPath);
            //newMailODU.Display(false);

            //Outlook.MailItem newMailRDU = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
            //newMailRDU.To = usersRDU;
            //newMailRDU.CC = usCopyRDU;
            //newMailRDU.Subject = "Нерешенные обращения по РДУ с " + DateFromPicker.Value.ToString("dd.MM.yyyy") + " по " + DateToPicker.Value.ToString("dd.MM.yyyy");
            //newMailRDU.Body = "Нерешенные обращения по РДУ с " + DateFromPicker.Value.ToString("dd.MM.yyyy") + " по " + DateToPicker.Value.ToString("dd.MM.yyyy") + ".\n\n" + signature;
            //newMailRDU.Attachments.Add(newBookRDUPath);
            //newMailRDU.Display(false);

            //Outlook.MailItem newMailTO = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
            //newMailTO.To = usersTO;
            //newMailTO.CC = usCopyTO;
            //newMailTO.Subject = "Нерешенные обращения по подрядным организациям и все поданные обращения на подрядные организации с " + DateFromPicker.Value.ToString("dd.MM.yyyy") + " по " + DateToPicker.Value.ToString("dd.MM.yyyy");
            //newMailTO.Body = "Нерешенные обращения по подрядным организациям и все поданные обращения на подрядные организации с " + DateFromPicker.Value.ToString("dd.MM.yyyy") + " по " + DateToPicker.Value.ToString("dd.MM.yyyy") + ".\n\n" + signature;
            //newMailTO.Attachments.Add(newBookTOPath);
            //newMailTO.Display(false);

            //wStr("E-mails created success!");

        }

        private void button_parsingReport_Click(object sender, EventArgs e)
        //Основная часть
        {
            button_parsingReport.Enabled = false;
            wStr("\r\nПолучаем массив данных");
            object[,] exportSDData = LoadDataFromFileExportSD(textBox_FileExportSDName.Text);
            wStr("Получили.");

            wStr("\r\nОбработка данных из отчета...");
            //Определение размерности массива исходных данных
            int lastRowIndex = exportSDData.GetLength(0);
            int lastCellIndex = exportSDData.GetLength(1);
            wStr("\r\nКоличество строк   = " + lastRowIndex);
            wStr("Количество колонок = " + lastCellIndex);

            // Определение местонахождения для обязательных полей по количеству обязательных столбцов
            for (int j = 1; j < CountNameColumnsONLY; j++)
            {
                //перебираем все столбцы
                for (int i = 1; i <= lastCellIndex; i++)
                {
                    if (exportSDData[1, i].ToString() == NameColumnsONLY[j, 0].ToString())
                    {
                        NameColumnsONLY[j, 1] = i;
                        wStr("№ Колонки '" + NameColumnsONLY[j, 0].ToString() + "' = " + i);
                    }
                }
                if (Convert.ToInt32(NameColumnsONLY[j, 1]) == 0)
                {
                    wStr("№ Колонки '" + NameColumnsONLY[j, 0].ToString() + "' НЕ ОПРЕДЕЛЕН!!! ");
                    return;
                }
            }
            NameColumnsONLY[0, 1] = lastCellIndex - CountNameColumnsONLY;
            wStr("Количество остальных колонок = " + NameColumnsONLY[0, 1]);



            if (checkBox_ReportTO.Checked)
            //Формируем отчет по подрядным организациям

            //one sheet
            //Условие: (6=ФГП начинается с "ТО \" ) и ( (3=Статус = {4 В ожидании:3 Выполняется:2 Назначен:1 Новый}) или (1=Дата окончания регистрации в необходимых пределах)
            //4=Филиал заявителя должен быть нашей ОЗ

            //two sheet
            //
            {
                //Наименование отчета
                string NameReport = "Отчет по подрядным организациям с " + dateTimePicker_TimeFrom.Value.ToString("dd.MM.yyyy") + " по " + dateTimePicker_TimeTo.Value.ToString("dd.MM.yyyy");
                wStr("\r\nФормируем отчет: " + NameReport);
                string countTO1 = "";
                string countTO2 = "";
                for (int i = 2; i <= lastRowIndex; i++) //Перебираем все строки с данными
                {
                    for (int f=0; f< CountNameFilialBranch; f++) //Перебираем филиалы
                    {
                        if (exportSDData[i, Convert.ToInt32(NameColumnsONLY[4, 1])].ToString().Contains(NameFilialBranch[f,0].ToString()))
                        {
                            for (int j = 0; j < CountStatusesNames; j++) //Перебираем все статусы
                            {
                                if ( exportSDData[i, Convert.ToInt32(NameColumnsONLY[3, 1])].ToString().Contains(StatusesNames[j]) )
                                {
                                    countTO1 = countTO1 + i.ToString() + ",";
                                }

                            }
                            if (Convert.ToDateTime(exportSDData[i, Convert.ToInt32(NameColumnsONLY[1, 1])]) >= dateTimePicker_TimeFrom.Value && Convert.ToDateTime(exportSDData[i, Convert.ToInt32(NameColumnsONLY[1, 1])]) <= dateTimePicker_TimeTo.Value)
                            {
                                countTO2 = countTO2 + i.ToString() + ",";
                            }
                        }
                    }
                }
                //Заполняем массив для отчета
                string[] numberRowTO;
                int countRowTO = 0;
                object[,,] ReportTO;

                if ( (countTO1.Split(',')).Length > (countTO2.Split(',')).Length )
                    { countRowTO = (countTO1.Split(',')).Length; };
                if ( (countTO1.Split(',')).Length < (countTO2.Split(',')).Length)
                    { countRowTO = (countTO2.Split(',')).Length; };
                ReportTO = new object[2, countRowTO, CountNameColumnsONLY];

                if (countRowTO != 0)
                {
                    numberRowTO = countTO1.Split(',');
                    countRowTO = numberRowTO.Length;
                    for (int i = 0; i < countRowTO - 1; i++) //Перебираем все необходимые строки с данными
                    {
                        for (int j = 0; j < CountNameColumnsONLY; j++)    //Перебираем необходимые поля
                        {
                            ReportTO[0,i,j] = exportSDData[Convert.ToInt32(numberRowTO[i]), Convert.ToInt32(NameColumnsONLY[j, 1])];
                        }
                    }
                    numberRowTO = countTO2.Split(',');
                    countRowTO = numberRowTO.Length;
                    for (int i = 0; i < countRowTO - 1; i++) //Перебираем все необходимые строки с данными
                    {
                        for (int j = 0; j < CountNameColumnsONLY; j++)    //Перебираем необходимые поля
                        {
                            ReportTO[1, i, j] = exportSDData[Convert.ToInt32(numberRowTO[i]), Convert.ToInt32(NameColumnsONLY[j, 1])];
                        }
                    }

                }
                SaveDataToReportFile(NameReport, "Нерешенные по ТО"+","+ "ТО за неделю", ReportTO);

                wStr("\r\nОтчет по подрядным организациям сформирован.");
            }

            //временно
            //button_parsingReport.Enabled = true;
            //return;

            if (checkBox_ReportRDU.Checked)
            //Формируем отчет: "Нерешенные обращения по РДУ за период".
            //Кол-во листов = кол-во филиалов
            //Условие: (6=ФГП начинается с "ТО \" ) и ( (3=Статус = {4 В ожидании:3 Выполняется:2 Назначен:1 Новый}) или (1=Дата окончания регистрации в необходимых пределах)
            //4=Филиал заявителя должен быть нашей ОЗ
            {
                //Наименование отчета
                string NameReport = "Нерешенные обращения по РДУ за период с " + dateTimePicker_TimeFrom.Value.ToString("dd.MM.yyyy") + " по " + dateTimePicker_TimeTo.Value.ToString("dd.MM.yyyy");
                wStr("\r\nФормируем отчет: " + NameReport);
                string ReportRDUSheets = "";//Наименования листов в книге

                for (int i = 2; i <= lastRowIndex; i++) //Перебираем все строки с данными
                {
                    for (int f = 0; f < CountNameFilialBranch; f++) //Перебираем филиалы
                    {
                        if (exportSDData[i, Convert.ToInt32(NameColumnsONLY[4, 1])].ToString().Contains(NameFilialBranch[f, 0].ToString()))
                        {
                            for (int j = 0; j < CountStatusesNames; j++) //Перебираем все статусы
                            {
                                if (exportSDData[i, Convert.ToInt32(NameColumnsONLY[3, 1])].ToString().Contains(StatusesNames[j]))
                                {
                                    NameFilialBranch[f, 1] = NameFilialBranch[f, 1] + i.ToString() + ","; //сохраняем №строки с данными
                                }

                            }
                            
                        }
                    }
                }
                //Заполняем массив для отчета
                object[,,] ReportRDU;
                int ReportRDU_max = 0;
                
                //размерность массива
                for (int n = 1; n < CountNameFilialBranch; n++)
                {
                    int t = NameFilialBranch[n, 1].ToString().Split(',').Length;
                    if (ReportRDU_max < t) ReportRDU_max = t;
                    ReportRDUSheets = ReportRDUSheets + NameFilialBranch[n, 0].ToString() + ",";
                }
                ReportRDU = new object[CountNameFilialBranch-1, ReportRDU_max, CountNameColumnsONLY];
                
                if (ReportRDU_max != 0) //Если выбрана хоть одна запись, то работаем
                {
                    for (int f = 1; f < CountNameFilialBranch; f++)
                    {
                        string[] numberRow = NameFilialBranch[f, 1].ToString().Split(',');//массив строк №записей
                        int countRow = numberRow.Length;//кол-во записей
                        for(int r = 0; r < countRow; r++)//перебираем все записи
                        {
                            for (int c = 0; c < CountNameColumnsONLY; c++)    //Перебираем необходимые поля
                            {
                                if(numberRow[r] != "") ReportRDU[f-1, r, c] = exportSDData[Convert.ToInt32(numberRow[r]), Convert.ToInt32(NameColumnsONLY[c, 1])];
                            }
                        }
                    }
                }
                SaveDataToReportFile(NameReport, ReportRDUSheets, ReportRDU);
                wStr("\r\nОтчет 'Нерешенные обращения по РДУ за период' сформирован.");

                ReportRDU_max = NameFilialBranch[0, 1].ToString().Split(',').Length;
                ReportRDU = new object[1, ReportRDU_max, CountNameColumnsONLY];
                {
                    string[] numberRow = NameFilialBranch[0, 1].ToString().Split(',');//массив строк №записей
                    int countRow = numberRow.Length;//кол-во записей
                    for (int r = 0; r < countRow; r++)//перебираем все записи
                    {
                        for (int c = 0; c < CountNameColumnsONLY; c++)    //Перебираем необходимые поля
                        {
                            if (numberRow[r] != "") ReportRDU[0, r, c] = exportSDData[Convert.ToInt32(numberRow[r]), Convert.ToInt32(NameColumnsONLY[c, 1])];
                        }
                    }
                }
                NameReport = "Нерешенные обращения по ОДУ за период с " + dateTimePicker_TimeFrom.Value.ToString("dd.MM.yyyy") + " по " + dateTimePicker_TimeTo.Value.ToString("dd.MM.yyyy");
                ReportRDUSheets = NameFilialBranch[0, 0].ToString();
                SaveDataToReportFile(NameReport, ReportRDUSheets, ReportRDU);
                wStr("\r\nОтчет 'Нерешенные обращения по ОДУ за период' сформирован.");
            }

            button_parsingReport.Enabled = true;
        }
    }
}
