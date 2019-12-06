using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace ReportMakerSD
{
    public partial class MainForm : Form
    {
        #region Инициируем переменные.
        /// <summary>
        /// Наименования обязательных полей в отчете
        /// </summary>
        private object[,] NameColumnsONLY;
        /// <summary>
        /// Наименования обязательных полей в отчете для РДУ.
        /// Берем из параметра ColumnsFilterNamesRDU
        /// </summary>
        private object[,] NameColumnsRDUONLY;
        /// <summary>
        /// Количество обязательных полей в отчете
        /// </summary>
        private int CountNameColumnsONLY;
        /// <summary>
        /// Количество обязательных полей в отчете РДУ
        /// </summary>
        private int CountNameColumnsRDUONLY;
        /// <summary>
        /// Наименования статусов
        /// </summary>
        private string[] StatusesNames;
        /// <summary>
        /// Количество статусов
        /// </summary>
        private int CountStatusesNames;

        private string[] FilialBranchName;
        /// <summary>
        /// Наименование филиалов ОЗ
        /// </summary>
        private object[,] NameFilialBranch;
        /// <summary>
        /// Количество филиалов ОЗ
        /// </summary>
        private int CountNameFilialBranch;
        /// <summary>
        /// Наименования ФГП
        /// </summary>
        private string[] FGPNames;
        /// <summary>
        /// Количество ФГП
        /// </summary>
        private int CountFGPNames;
        /// <summary>
        /// Источники оповещения
        /// </summary>
        private string[] WayOfGivingTreatments;
        /// <summary>
        /// Количество источников оповещения.
        /// </summary>
        private int CountWayOfGivingTreatments;
        /// <summary>
        /// Начинаются с этого ФГП-подрядчиков
        /// </summary>
        private string stringTO = Properties.Settings.Default.TO;
        /// <summary>
        /// Массив с исходными данными
        /// </summary>
        public object[,] exportSDData;
        /// <summary>
        /// Количество строк в исходных данных
        /// </summary>
        private int lastRowIndex;
        /// <summary>
        /// Количество колонок в исходных данных.
        /// </summary>
        private int lastCellIndex;
        /// <summary>
        ///Активация Excel
        /// </summary>
        Excel.Application oExcelApp = null;
        Workbooks oExcelBooks;
        _Workbook oExcelBook;
        Sheets oExcelSheets;
        _Worksheet oExcelSheet;
        Range oExcelRange;
        Workbooks openbooks = null;
        Workbook openbook = null;
        //Workbook openbook_TO = null;
        Sheets openSheets = null;
        //Range openRange = null;
        Worksheet openSheet = null;
        Range openCellsFirst = null;

        //Делаем количество массивов равных количеству отчетов/книг (всего задуманных).
        /// <summary>
        /// Массив данных по общей книге
        /// </summary>
        private object[][,] DataForReport_based;
        /// <summary>
        /// Массив данных по подрядникам
        /// </summary>
        private object[][,] DataForReport_TO; //по подрядникам
        /// <summary>
        /// Массив данных по ОДУ
        /// </summary>
        private object[][,] DataForReport_ODU; //по ОДУ
        /// <summary>
        /// Количество записей для отчета по ОДУ
        /// </summary>
        private int ReportODU = 0;
        private int ReportODU2 = 0;
        /// <summary>
        /// Массив данных по РДУ[Филиал][обращения,колонка]
        /// </summary>
        private object[][,] DataForReport_RDU;//по РДУ
        /// <summary>
        /// Количество записей для отчета по РДУ
        /// </summary>
        private int ReportRDU = 0;

        string PathToFileExportSD = @"D:\";
        /// <summary>
        /// Наименование отчета за период
        /// </summary>
        string NameReport_based;
        /// <summary>
        /// Наименования листов в книге отчета за период
        /// </summary>
        public string[] Report_based_Sheets = Properties.Settings.Default.Report_based_SheetsName.Split(':');
        string NameReport_TO;   //Наименование отчета по подрядникам;
        string NameReport_RDU;   //Наименование отчета по РДУ;
        string[] ReportRDU_Sheets; //Наименования листов в книге отчета по РДУ
        string NameReport_ODU;   //Наименование отчета по ОДУ;
        string ReportODU_Sheets; //Наименования листов в книге отчета по ОДУ
        /// <summary>
        /// Версия сборки в название формы
        /// </summary>
        String strVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(); // Версия сборки в название формы
        /// <summary>
        /// Путь до обновления EXE-файла
        /// </summary>
        string strInstallPath = Properties.Settings.Default.InstallPath;    // Путь до обновления EXE-файла
        /// <summary>
        /// Дата создания файла экспорта
        /// </summary>
        DateTime FileExportSD_CreationTime = DateTime.Now;
        /// <summary>
        /// Временны переменные
        /// </summary>
        string Report1Path;
        string Report2Path;
        string Report3Path;
        string Report4Path;
        string Report5Path;
        string Report6Path;
        public static string SDR0Data; //Дата отчета
        public static List<SDRData> SDR0Data0 = new List<SDRData>();
        #endregion

        public MainForm()
        //Конструктор
        {
            InitializeComponent();

            Text = " " + Properties.Settings.Default.ProjectName + " (v." + strVersion + ")"; //изменяем наименование формы.

            #region Определяем рабочую папку и файл экспорта.
            // Если нет файла экспорта в папке запуска программы, то переходим в папку загрузки профиля пользователя.
            //PathToFileExportSD = Environment.SpecialFolder.UserProfile.ToString() + @"\" + Properties.Settings.Default.FolderWithExportSdDefault.ToString() + "\\";
            //PathToFileExportSD = Environment.CurrentDirectory.ToString() + @"\Downloads\";
            PathToFileExportSD = Environment.CurrentDirectory.ToString();
            //DirectoryInfo fileSystemInfo = new DirectoryInfo(PathToFileExportSD);
            //var fileMore = fileSystemInfo.GetFiles("exportSD*.xl*");
            DirectoryInfo fileSystemInfo = new DirectoryInfo(PathToFileExportSD);
            var fileMore = fileSystemInfo.GetFiles("exportSD*.xl*");
            if (fileMore.Length == 0)
            {
                PathToFileExportSD = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\" + Properties.Settings.Default.FolderWithExportSdDefault.ToString();
                fileSystemInfo = new DirectoryInfo(PathToFileExportSD);
                fileMore = fileSystemInfo.GetFiles("exportSD*.xl*");
            }
            textBox_FileExportSDName.Text = PathToFileExportSD;
            openFileDialog_ExportSDSelect.InitialDirectory = PathToFileExportSD;

            #region Определяем свежий отчет экспорта из СД
            if ((fileMore is null) || fileMore.Length == 0)
            {
                wStr("Укажите файл отчета вручную.", true);
                button_parsingReport.Enabled = false;
            }
            else
            {
                DateTime dt = Convert.ToDateTime(fileMore[0].CreationTime);
                textBox_FileExportSDName.Text = fileMore[0].Name;
                foreach (FileSystemInfo fileSI in fileMore)
                {
                    if (dt < Convert.ToDateTime(fileSI.CreationTime))
                    {
                        dt = Convert.ToDateTime(fileSI.CreationTime);
                        textBox_FileExportSDName.Text = fileSI.Name;
                    }
                }
                textBox_FileExportSDName.Text = PathToFileExportSD + @"\" + textBox_FileExportSDName.Text;
                FileExportSD_CreationTime = dt;
                //wStr("Получили данные из выгрузки по новому...", true);
                //SDR0Data0 = SDRData.SDRLoadData(PathToFileExportSD, textBox_FileExportSDName.Text);
                //SDR0Data = FileExportSD_CreationTime.ToString();
            }
            #endregion

            //textBox_FileExportSDName.Text = PathToFileExportSD +@"\"+ textBox_FileExportSDName.Text; // Дает удвоение пути
            openFileDialog_ExportSDSelect.FileName = textBox_FileExportSDName.Text;
            #endregion
            #region Определяем даты отчета.
            //Определяем необходимые даты для отчета.
            //1. Если сегодня понедельник и дата создания файла сегодня, то период выставляем "За прошлую неделю"
            //2. Если сегодня пятница и дата создания файла сегодня, то период выставляем "За текущую неделю"
            //3. Если дата создания файла = понедельник и этот понедельник уже прошел, то период выставляем "За неделю"
            switch (DateTime.Now.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    if (FileExportSD_CreationTime.Date == DateTime.Now.Date) comboBox_SelectPeriod.SelectedItem = "За прошлую неделю";
                    else comboBox_SelectPeriod.SelectedItem = "За неделю";
                    break;
                case DayOfWeek.Friday:
                    if (FileExportSD_CreationTime.Date == DateTime.Now.Date) comboBox_SelectPeriod.SelectedItem = "За текущую неделю";
                    break;
                default:
                    comboBox_SelectPeriod.SelectedItem = "За неделю";
                    break;
            }
            #endregion

            //Подготовительная часть завершена.
            wStr("Часть 1 - завершена.", true, true);
            button_parsingReport.Focus();
            button_parsingReport.Select();
        }
        ///<summary>
        ///Метод для заполнения текстового окна отладочной информацией.
        ///text - Текст для вывода;
        ///NewLine - Признак новой строки(Необязательный параметр);
        ///DebugLine - Признак отладочной информации(Необязательный параметр).
        ///</summary>
        /// <param name="text">Текст для вывода</param>
        void wStr(string text, bool NewLine = false, bool DebugLine = false)
        {
            if (DebugLine & !checkBox_DebugInfo.Checked)
            //Если инфа для вывода относится к Дебаг, а флаг вывода не установлен - то пропускаем.
            {
                return;
            }
            if (checkBox_DebugInfo.Checked)
            {
                text = DateTime.Now.Minute.ToString("D2") + "." + DateTime.Now.Second.ToString("D2") + "." + DateTime.Now.Millisecond.ToString("D3") + "   " + text;
            }
            if (NewLine)
            {
                text = "\r\n" + text;
            }
            text = text + "\n";
            textBox_DebugInfo.AppendText(text);
            textBox_DebugInfo.Refresh();
        }

        /// <summary>
        /// Метод выбора файла экспорта при нажатии кнопки
        /// </summary>
        void button_ExportSDSelect_Click(object sender, EventArgs e)
        {
            openFileDialog_ExportSDSelect.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog_ExportSDSelect.ShowDialog() == DialogResult.OK)
            {
                textBox_FileExportSDName.Text = openFileDialog_ExportSDSelect.FileName;
                button_parsingReport.Enabled = true;
            }
        }

        /// <summary>
        /// Отправка писем с отчетами
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void button_SendMail_Click(object sender, EventArgs e)
        {
            string signature = "С уважением,\nСпециалист ТРСОО\nГруппы обработки обращений СОЭ АСУ,\nт. 39-99";
            string body = "";
            Outlook.Application outApp = new Outlook.Application();
            #region Почта по филиалам.
            if (Report5Path != null)
            {
                Outlook.MailItem newMailODU = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
                newMailODU.To = Properties.Settings.Default.ReportODU_MailTO;
                newMailODU.CC = Properties.Settings.Default.ReportODU_MailCopy;
                newMailODU.Subject = "SD. Отчет по ОДУ";
                newMailODU.Attachments.Add(Report5Path);
                newMailODU.Body = "Нерешенные обращения специалистами ОДУ.\n\n" + signature;
                newMailODU.Display(false);

                Outlook.MailItem newMailRDU = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
                newMailRDU.To = Properties.Settings.Default.ReportRDU_MailTO;
                newMailRDU.CC = Properties.Settings.Default.ReportRDU_MailCopy;
                newMailRDU.Subject = "SD. Отчет по РДУ" + NameReport_RDU;
                newMailRDU.Attachments.Add(Report6Path);
                newMailRDU.Body = "Нерешенные обращения специалистами РДУ.\n\n" + signature;
                newMailRDU.Display(false);
            }
            if (checkBox_ReportOne.Checked)
            {
                Outlook.MailItem newMailODU2 = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
                newMailODU2.To = Properties.Settings.Default.ReportODU_MailTO + ";" + Properties.Settings.Default.ReportRDU_MailTO;
                newMailODU2.CC = Properties.Settings.Default.ReportODU_MailCopy + ";" + Properties.Settings.Default.ReportRDU_MailCopy;
                newMailODU2.Subject = "SD. Отчет по филиалам ";
                newMailODU2.Body = "SD. Отчет по филиалам.\n\n" + signature;
                newMailODU2.Attachments.Add(Report1Path);
                newMailODU2.Display(false);
            }
            #endregion
            #region Почта по подрядникам.
            if (checkBox_Report4.Checked || checkBox_Report2.Checked)
            {
                Outlook.MailItem newMailTO = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
                newMailTO.To = Properties.Settings.Default.ReportTO_MailTO;
                newMailTO.CC = Properties.Settings.Default.ReportTO_MailCopy;
                newMailTO.Subject = "SD. Отчет по подрядным организациям";
                newMailTO.Body = signature;
                if (checkBox_ReportTO.Checked) { newMailTO.Body.Insert(0, NameReport_TO + ".\n\n"); }
                if (checkBox_ReportTO.Checked) { newMailTO.Attachments.Add(PathToFileExportSD + "\\" + NameReport_TO + ".xlsx"); }
                if (checkBox_Report2.Checked) { newMailTO.Attachments.Add(Report2Path); }
                if (checkBox_Report4.Checked) { newMailTO.Attachments.Add(Report4Path); }
                newMailTO.Display(false);
            }
            #endregion
            #region Почта:Отчет за период.
            if (checkBox_ReportBase.Checked || checkBox_Report3.Checked)
            {
                Outlook.MailItem newMailTO = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
                newMailTO.To = Properties.Settings.Default.Report_SD;
                newMailTO.Subject = "SD. Отчет за период";
                if (checkBox_ReportBase.Checked)
                {
                    body += NameReport_based + ".\n\n";
                    newMailTO.Attachments.Add(PathToFileExportSD + "\\" + NameReport_based + ".xlsx");
                }
                if (checkBox_Report3.Checked)
                {
                    newMailTO.Attachments.Add(Report3Path);
                }
                body += signature;
                newMailTO.Body = body;
                newMailTO.Display(false);
            }
            #endregion
            wStr("Письма с отчетами сформированы.", true, false);
        }

        /// <summary>
        /// Действия по первой кнопке "Обработка данных"
        /// </summary>
        void button_parsingReport_Click(object sender, EventArgs e)
        {
            //Подготовка:
            button_parsingReport.Enabled = false;
            button_OpenFolder.Enabled = false;
            button_SendMail.Enabled = false;
            //Запуск тестовых отчета
            if (checkBox_ReportOne.Checked || checkBox_Report2.Checked || checkBox_Report3.Checked || checkBox_Report4.Checked || checkBox_Report5.Checked)
            {
                wStr("Получили данные из выгрузки по новому...", true);
                SDR0Data0 = SDRData.SDRLoadData(PathToFileExportSD, textBox_FileExportSDName.Text);
                SDR0Data = FileExportSD_CreationTime.ToString();
            }
            if (checkBox_ReportOne.Checked) Report1Path = SDR0.RunReportOne(dateTimePicker_TimeTo.Value);
            if (checkBox_Report2.Checked) Report2Path = SDR0.RunReportTwo(dateTimePicker_TimeTo.Value);
            if (checkBox_Report3.Checked) Report3Path = SDR0.RunReportThree(dateTimePicker_TimeTo.Value);
            if (checkBox_Report4.Checked) Report4Path = SDR0.RunReportFour(dateTimePicker_TimeFrom.Value, dateTimePicker_TimeTo.Value);
            if (checkBox_Report5.Checked)
            {
                Report5Path = SDR0.RunReportFive(dateTimePicker_TimeTo.Value);
                Report6Path = SDR0.RunReport6(dateTimePicker_TimeTo.Value);
            }
            // Отчеты по старому
            #region Отчеты по старому
            if (checkBox_ReportBase.Checked || checkBox_ReportTO.Checked)
            {

                wStr("Часть 2 - Проверка и формирование справочников.", true, true);
                #region Справочники.
                #region Справочники. Обязательные поля.
                wStr("Обязательные поля:", true, true);
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
                    wStr("   " + count.ToString() + " - '" + NameColumnsONLY[count, 0].ToString() + "';", false, true);
                    count++;
                }
                CountNameColumnsONLY = count;
                wStr("Итого полей: " + CountNameColumnsONLY, true, true);
                #endregion
                #region Справочники. Обязательные поля для РДУ.
                wStr("Обязательные поля отчета РДУ:", true, true);
                //Указатели обязательных полей на №столбцов в исходных данных
                //Остальные колонки:Дата окончания регистрации:Номер:Статус:Организация заявителя:Заявитель:ФГП:Исполнитель:Услуга:Тема:Код ожидания:Причина ожидания:Дата и время решения:Решение:Способ подачи обращения
                //Наименование,№столбца в исходных данных (для первого=кол-во необязательных),
                NameColumnsRDUONLY = new object[Properties.Settings.Default.ColumnsFilterNamesRDU.Split(':').Length, 3];
                count = 0;
                foreach (string NameColumnRDUONLY in Properties.Settings.Default.ColumnsFilterNamesRDU.Split(':'))
                {
                    NameColumnsRDUONLY[count, 0] = NameColumnRDUONLY;
                    NameColumnsRDUONLY[count, 1] = 0;
                    NameColumnsRDUONLY[count, 2] = 0;
                    wStr("   " + count.ToString() + " - '" + NameColumnsRDUONLY[count, 0].ToString() + "'; ", false, true);
                    count++;
                }
                CountNameColumnsRDUONLY = count;
                wStr("Итого полей отчета РДУ: " + CountNameColumnsONLY, true, true);
                #endregion

                wStr("Справочник статусов:", true, true);
                StatusesNames = Properties.Settings.Default.StatusesNames.Split(':');
                count = 0;
                foreach (string StatusNames in StatusesNames)
                { wStr("   '" + StatusNames + ";", false, true); count++; }
                CountStatusesNames = count;
                wStr("Итого статусов: " + CountStatusesNames, false, true);

                wStr("Справочник филиалов:", true, true);
                //Наименование филиала (замена nameODU &etc), строки (замена rowNumberODU&etc), счетчик (замена countODU &etc)
                FilialBranchName = Properties.Settings.Default.FilialBranchName.Split(':');    //20170529-Список филиалов.
                                                                                               //NameFilialBranch[0,0].ToString(), Convert.ToInt32(NameFilialBranch[0, 1]), Convert.ToInt32(NameFilialBranch[0, 2])
                                                                                               //NameFilialBranch = new object[Properties.Settings.Default.FilialBranchName.Split(':').Length, 3]; //20170529.Доп.филиал="Вне ОЗ"
                NameFilialBranch = new object[Properties.Settings.Default.FilialBranchName.Split(':').Length + 1, 3];
                count = 0;
                foreach (string NameFilialBranch_ in Properties.Settings.Default.FilialBranchName.Split(':'))
                {
                    NameFilialBranch[count, 0] = NameFilialBranch_;
                    NameFilialBranch[count, 1] = "";
                    NameFilialBranch[count, 2] = 0;
                    wStr("   " + count.ToString() + " - '" + NameFilialBranch[count, 0].ToString() + "';", false, true);
                    count++;
                }
                CountNameFilialBranch = count;
                NameFilialBranch[CountNameFilialBranch, 0] = "Вне ОЗ Северо-Запада"; //20170529.Доп.филиал="Вне ОЗ"
                NameFilialBranch[CountNameFilialBranch, 1] = ""; //20170529.Доп.филиал="Вне ОЗ"
                NameFilialBranch[CountNameFilialBranch, 2] = 0; //20170529.Доп.филиал="Вне ОЗ"

                wStr("Итого филиалов: " + CountNameFilialBranch, false, true);

                wStr("Справочник ФГП:", true, true);
                FGPNames = Properties.Settings.Default.FGPNames.Split(':');
                count = 0;
                foreach (string FGPName in FGPNames)
                { wStr("   '" + FGPName + ";", false, true); count++; }
                CountFGPNames = count;
                wStr("Итого ФГП: " + CountFGPNames, false, true);

                wStr("Cпособы подачи обращения:", true, true);
                WayOfGivingTreatments = Properties.Settings.Default.WayOfGivingTreatments.Split(':');
                count = 0;
                foreach (string WayOfGivingTreatment in WayOfGivingTreatments)
                { wStr("   '" + WayOfGivingTreatment + ";", false, true); count++; }
                CountWayOfGivingTreatments = count;
                wStr("Итого способов: " + CountWayOfGivingTreatments, false, true);
                #endregion
                wStr("Желательно закрыть EXCEL!\r\n", true);
                wStr("Получаем массив данных. => ", true);
                wStr(" По старому...");
                exportSDData = LoadDataFromFileExportSD(textBox_FileExportSDName.Text);
                wStr(" Получили.\r\n");


                wStr("Обработка данных из отчета... \r\n", true);
                //Определение размерности массива исходных данных
                lastRowIndex = exportSDData.GetLength(0);
                wStr("Количество строк   = " + lastRowIndex + ". ", true, true);
                lastCellIndex = exportSDData.GetLength(1);
                wStr("Количество колонок = " + lastCellIndex + ". ", true, true);

                #region Обязательные поля - общие
                // Определение местонахождения для обязательных полей по количеству обязательных столбцов
                for (int j = 1; j < CountNameColumnsONLY; j++)
                {
                    //перебираем все столбцы
                    for (int i = 1; i <= lastCellIndex; i++)
                    {
                        if (exportSDData[1, i].ToString() == NameColumnsONLY[j, 0].ToString())
                        {
                            NameColumnsONLY[j, 1] = i;
                            wStr("№ Колонки '" + NameColumnsONLY[j, 0].ToString() + "' = " + i, false, true);
                        }
                    }
                    if (Convert.ToInt32(NameColumnsONLY[j, 1]) == 0)
                    {
                        wStr("№ Колонки '" + NameColumnsONLY[j, 0].ToString() + "' НЕ ОПРЕДЕЛЕН!!! ", false, true);
                        button_parsingReport.Enabled = true;
                        button_OpenFolder.Enabled = true;
                        button_SendMail.Enabled = false;
                        button_OpenFolder.Focus();
                        return; //Выход по ошибке.
                    }
                }
                NameColumnsONLY[0, 1] = lastCellIndex - CountNameColumnsONLY;
                wStr("Количество остальных колонок = " + NameColumnsONLY[0, 1], false, true);
                #endregion

                #region Обязательные поля - РДУ
                // Определение местонахождения для обязательных полей РДУ по количеству обязательных столбцов
                for (int j = 1; j < CountNameColumnsRDUONLY; j++)
                {
                    //перебираем все столбцы
                    for (int i = 1; i <= lastCellIndex; i++)
                    {
                        if (exportSDData[1, i].ToString() == NameColumnsRDUONLY[j, 0].ToString())
                        {
                            NameColumnsRDUONLY[j, 1] = i;
                            wStr("№ Колонки '" + NameColumnsRDUONLY[j, 0].ToString() + "' = " + i, false, true);
                        }
                    }
                    if (Convert.ToInt32(NameColumnsRDUONLY[j, 1]) == 0)
                    {
                        wStr("№ Колонки '" + NameColumnsRDUONLY[j, 0].ToString() + "' НЕ ОПРЕДЕЛЕН!!! ", false, true);
                        button_parsingReport.Enabled = true;
                        button_OpenFolder.Enabled = true;
                        button_SendMail.Enabled = false;
                        button_OpenFolder.Focus();
                        return; //Выход по ошибке.
                    }
                }
                NameColumnsRDUONLY[0, 1] = lastCellIndex - CountNameColumnsRDUONLY;
                wStr("Количество остальных колонок = " + NameColumnsRDUONLY[0, 1], false, true);
                #endregion
            }
            if (checkBox_ReportBase.Checked) ParsingDataForReport_Based();
            if (checkBox_ReportTO.Checked) ParsingDataForReport_TO();
            //Создание Excel-файлов.
            if (checkBox_ReportBase.Checked || checkBox_ReportTO.Checked)
            {
                CreateFiles();
            }
            #endregion

            wStr(" ", true);
            wStr("Все готово. Спасибо за внимание. \r\n", true);
            button_parsingReport.Enabled = true;
            button_OpenFolder.Enabled = true;
            button_SendMail.Enabled = true;
            button_SendMail.Focus();
        }

        void button_OpenFolder_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", PathToFileExportSD);
        }

        /// <summary>
        /// Настройка дат по результату списка
        /// </summary>
        private void comboBox_SelectPeriod_SelectedIndexChanged(object sender, EventArgs e)
        {
            dateTimePicker_TimeTo.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, Properties.Settings.Default.TimeFromHour - 1, Properties.Settings.Default.TimeFromMinute - 1, 0);
            switch (comboBox_SelectPeriod.SelectedItem.ToString())
            {
                case "За текущую неделю":
                    dateTimePicker_TimeFrom.Value = new DateTime(
                        DateTime.Now.AddDays(-(Convert.ToInt16(DateTime.Now.DayOfWeek)) + 1).Year,
                        DateTime.Now.AddDays(-(Convert.ToInt16(DateTime.Now.DayOfWeek)) + 1).Month,
                        DateTime.Now.AddDays(-(Convert.ToInt16(DateTime.Now.DayOfWeek)) + 1).Day,
                        //Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                        0, 0, 0);
                    switch (DateTime.Now.DayOfWeek)
                    {
                        case DayOfWeek.Monday:
                            dateTimePicker_TimeFrom.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 
                                Properties.Settings.Default.TimeFromHour - 1, Properties.Settings.Default.TimeFromMinute - 1, 0);
                            break;
                        case DayOfWeek.Friday:
                            //dateTimePicker_TimeFrom.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, Properties.Settings.Default.TimeFromHour - 1, Properties.Settings.Default.TimeFromMinute - 1, 0);
                            dateTimePicker_TimeTo.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 13, 59, 0);
                            break;
                        default:
                            dateTimePicker_TimeFrom.Value = new DateTime(
                                DateTime.Now.AddDays(-(Convert.ToInt16(DateTime.Now.DayOfWeek)) + 1).Year,
                                DateTime.Now.AddDays(-(Convert.ToInt16(DateTime.Now.DayOfWeek)) + 1).Month,
                                DateTime.Now.AddDays(-(Convert.ToInt16(DateTime.Now.DayOfWeek)) + 1).Day,
                                //Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                                0, 0, 0);
                            break;
                    }
                    break;
                case "За 7 дней":
                    dateTimePicker_TimeFrom.Value = new DateTime(
                        DateTime.Now.AddDays(-7).Year,
                        DateTime.Now.AddDays(-7).Month,
                        DateTime.Now.AddDays(-7).Day,
                        Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    dateTimePicker_TimeTo.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 59, 0);
                    break;
                case "За прошлую неделю":
                    dateTimePicker_TimeFrom.Value = new DateTime(
                        DateTime.Now.AddDays(-(6 + Convert.ToInt32(DateTime.Now.DayOfWeek))).Year,
                        DateTime.Now.AddDays(-(6 + Convert.ToInt32(DateTime.Now.DayOfWeek))).Month,
                        DateTime.Now.AddDays(-(6 + Convert.ToInt32(DateTime.Now.DayOfWeek))).Day,
                        Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    dateTimePicker_TimeTo.Value = new DateTime(
                        DateTime.Now.Year,
                        DateTime.Now.Month,
                        DateTime.Now.AddDays(-1).Day,
                        23, 59, 0);
                    break;
                case "За неделю":
                    if (FileExportSD_CreationTime.DayOfWeek == DayOfWeek.Monday)
                    {
                        dateTimePicker_TimeFrom.Value = FileExportSD_CreationTime.AddDays(-7);
                        dateTimePicker_TimeFrom.Value = new DateTime(dateTimePicker_TimeFrom.Value.Year, dateTimePicker_TimeFrom.Value.Month, dateTimePicker_TimeFrom.Value.Day,
                            Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                        dateTimePicker_TimeTo.Value = FileExportSD_CreationTime.AddDays(-3);
                        dateTimePicker_TimeTo.Value = new DateTime(dateTimePicker_TimeTo.Value.Year, dateTimePicker_TimeTo.Value.Month, dateTimePicker_TimeTo.Value.Day, 16, 15, 0);
                    }
                    else if (FileExportSD_CreationTime.DayOfWeek == DayOfWeek.Friday)
                    {
                        dateTimePicker_TimeFrom.Value = FileExportSD_CreationTime.AddDays(-5);
                        dateTimePicker_TimeFrom.Value = new DateTime(dateTimePicker_TimeFrom.Value.Year, dateTimePicker_TimeFrom.Value.Month, dateTimePicker_TimeFrom.Value.Day,
                            Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                        dateTimePicker_TimeTo.Value = FileExportSD_CreationTime;
                        dateTimePicker_TimeTo.Value = new DateTime(dateTimePicker_TimeTo.Value.Year, dateTimePicker_TimeTo.Value.Month, dateTimePicker_TimeTo.Value.Day, 16, 15, 0);
                    }
                    else
                    {
                        dateTimePicker_TimeFrom.Value = new DateTime(
                        DateTime.Now.AddDays(-(Convert.ToInt16(DateTime.Now.DayOfWeek)) + 1).Year,
                        DateTime.Now.AddDays(-(Convert.ToInt16(DateTime.Now.DayOfWeek)) + 1).Month,
                        DateTime.Now.AddDays(-(Convert.ToInt16(DateTime.Now.DayOfWeek)) + 1).Day,
                        Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    }
                    break;
                case "За прошлый месяц":
                    dateTimePicker_TimeFrom.Value = FileExportSD_CreationTime.AddMonths(-1);
                    dateTimePicker_TimeFrom.Value = new DateTime(dateTimePicker_TimeFrom.Value.Year, dateTimePicker_TimeFrom.Value.Month, 1, 0, 0, 0);
                    dateTimePicker_TimeTo.Value = new DateTime(FileExportSD_CreationTime.Year, FileExportSD_CreationTime.Month, 1).AddDays(-1);
                    break;
                case "За текущий квартал":
                    dateTimePicker_TimeFrom.Value = new DateTime(DateTime.Now.Year,
                        ((DateTime.Now.Month / 4) * 3 + 1),
                        1, Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    break;
                case "За текущие полгода":
                    if (DateTime.Now.Month > 6)
                    {
                        dateTimePicker_TimeFrom.Value = new DateTime(DateTime.Now.Year, 6, 1, Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    }
                    else
                    {
                        dateTimePicker_TimeFrom.Value = new DateTime(DateTime.Now.Year, 1, 1, Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    }
                    break;
                case "За текущий год":
                    dateTimePicker_TimeFrom.Value = new DateTime(DateTime.Now.Year, 1, 1, Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    break;
                default:
                    break;
            }
        }

    }
}
