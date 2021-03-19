using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace ReportMakerSD
{
    public partial class Form0 : Form
    {
        #region Инициируем переменные.
        private string stringTO = Properties.Settings.Default.TO;
        /// <summary>
        /// Массив с исходными данными
        /// </summary>
        public object[,] exportSDData;

        readonly string PathToFileExportSD = @"D:\";
        //string NameReport_based;
        /// <summary>
        /// Наименования листов в книге отчета за период
        /// </summary>
        public string[] Report_based_Sheets = Properties.Settings.Default.Report_based_SheetsName.Split(':');
        //string NameReport_TO;   //Наименование отчета по подрядникам;
        //string NameReport_RDU;   //Наименование отчета по РДУ;
        //string[] ReportRDU_Sheets; //Наименования листов в книге отчета по РДУ
        //string NameReport_ODU;   //Наименование отчета по ОДУ;
        //string ReportODU_Sheets; //Наименования листов в книге отчета по ОДУ
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
        string Report3Name;

        public static DateTime DateTimeNowWeek; // Дата окончания текущей недели (для отчета)
        public static string SDR0Data; //Дата отчета
        public static List<SDRData> SDR0Data0 = new List<SDRData>();

        public static bool DebugInfoWriteenabler = false;
        #endregion

        public Form0()
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
                DebugInfoWrite("Укажите файл отчета вручную.", true);
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
                    else comboBox_SelectPeriod.SelectedItem = "За текущую неделю";
                    break;
                case DayOfWeek.Friday:
                    if (FileExportSD_CreationTime.Date == DateTime.Now.Date) comboBox_SelectPeriod.SelectedItem = "За текущую неделю";
                    break;
                default:
                    comboBox_SelectPeriod.SelectedItem = "За текущую неделю";
                    break;
            }
            #endregion

            //Подготовительная часть завершена.
            button_parsingReport.Focus();
            button_parsingReport.Select();
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
        void button_SendMail_Click(object sender, EventArgs e)
        {
            DebugInfoWrite("Формируем письма с отчетами...", true, true);
            string signature = "С уважением,\nСпециалист ТРСОО\nГруппы обработки обращений СОЭ АСУ,\nт. 39-99";
            string body = "SD. Еженедельный отчет по " + dateTimePicker_TimeTo.Value.ToString("dd/MM/yyyy") + ".";
            Outlook.Application outApp = new Outlook.Application();
            #region Еженедельный отчет
            if (checkBox_Report1.Checked || checkBox_Report2.Checked)
            {
                Outlook.MailItem newMailODU = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
                newMailODU.To = textBox_WeekMailTo.Text;
                newMailODU.CC = "info_sd@odusz.so-ups.ru";
                newMailODU.Subject = body;
                newMailODU.Body = body + "\n\n" + signature;
                if (checkBox_Report1.Checked) newMailODU.Attachments.Add(Report1Path);
                if (checkBox_Report2.Checked) newMailODU.Attachments.Add(Report2Path);
                newMailODU.Display(false);
            }
            #endregion
            #region Отчет за месяц
            if (checkBox_Report3.Checked)
            {
                Outlook.MailItem newMailTO = (Outlook.MailItem)outApp.CreateItem(Outlook.OlItemType.olMailItem);
                newMailTO.To = textBox_WeekMailTo.Text;
                newMailTO.CC = "info_sd@odusz.so-ups.ru";
                newMailTO.Subject = Report3Name + ".";
                if (checkBox_Report3.Checked)
                {
                    newMailTO.Attachments.Add(Report3Path);
                }
                newMailTO.Body = Report3Name + ".\n\n" + signature;
                newMailTO.Display(false);
            }
            #endregion
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
            if (checkBox_Report1.Checked || checkBox_Report2.Checked || checkBox_Report3.Checked)
            {
                SDR0Data0 = SDRData.SDRLoadData(PathToFileExportSD, textBox_FileExportSDName.Text);
                SDR0Data = FileExportSD_CreationTime.ToString();
            }
            if (checkBox_Report1.Checked) Report1Path = SDR0.RunReportOne(dateTimePicker_TimeTo.Value, DateTimeNowWeek);
            if (checkBox_Report2.Checked) Report2Path = SDR0.RunReportTwo(dateTimePicker_TimeTo.Value, DateTimeNowWeek);
            if (checkBox_Report3.Checked)
            {
                switch(dateTimePicker_TimeTo.Value.ToString("MM"))
                {
                    case "01":
                        Report3Name = "SD. Отчет за январь " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "02":
                        Report3Name = "SD. Отчет за февраль " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "03":
                        Report3Name = "SD. Отчет за март " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "04":
                        Report3Name = "SD. Отчет за апрель " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "05":
                        Report3Name = "SD. Отчет за май " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "06":
                        Report3Name = "SD. Отчет за июнь " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "07":
                        Report3Name = "SD. Отчет за июль " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "08":
                        Report3Name = "SD. Отчет за август " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "09":
                        Report3Name = "SD. Отчет за сентябрь " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "10":
                        Report3Name = "SD. Отчет за октябрь " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "11":
                        Report3Name = "SD. Отчет за ноябрь " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    case "12":
                        Report3Name = "SD. Отчет за декабрь " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                    default:
                        Report3Name = "SD. Отчет за " + dateTimePicker_TimeTo.Value.ToString("MM") + " " + dateTimePicker_TimeTo.Value.ToString("yyyy");
                        break;
                }
                //Report3Name += ".";
                Report3Path = SDR0.ReportMonth(dateTimePicker_TimeTo.Value, Report3Name);
            }
            //if (checkBox_Report4.Checked) Report4Path = SDR0.RunReportFour(dateTimePicker_TimeFrom.Value, dateTimePicker_TimeTo.Value);
            //if (checkBox_Report5.Checked)
            //{
            //    Report5Path = SDR0.RunReportFive(dateTimePicker_TimeTo.Value);
            //    Report6Path = SDR0.RunReport6(dateTimePicker_TimeTo.Value);
            //}
            
            DebugInfoWrite(" ", true);
            DebugInfoWrite("Все готово. Спасибо за внимание. \r\n", true);
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
                        DateTime.Now.AddDays(1- (Convert.ToInt16(DateTime.Now.DayOfWeek))).Year,
                        DateTime.Now.AddDays(1- (Convert.ToInt16(DateTime.Now.DayOfWeek))).Month,
                        DateTime.Now.AddDays(1- (Convert.ToInt16(DateTime.Now.DayOfWeek))).Day, 0, 0, 0);
                    dateTimePicker_TimeTo.Value = new DateTime(
                        DateTime.Now.AddDays(0 - 1).Year,
                        DateTime.Now.AddDays(0 - 1).Month,
                        DateTime.Now.AddDays(0 - 1).Day, 23, 59, 0);
                    if (!checkBox_Report1.Checked) checkBox_Report1.Checked = true;
                    if (!checkBox_Report2.Checked) checkBox_Report2.Checked = true;
                    if (checkBox_Report3.Checked) checkBox_Report3.Checked = false;
                    DateTimeNowWeek = dateTimePicker_TimeFrom.Value.AddDays(7);
                    break;
                case "За 7 дней":
                    dateTimePicker_TimeFrom.Value = new DateTime(
                        DateTime.Now.AddDays(-7).Year,
                        DateTime.Now.AddDays(-7).Month,
                        DateTime.Now.AddDays(-7).Day,
                        Properties.Settings.Default.TimeFromHour, Properties.Settings.Default.TimeFromMinute, 0);
                    dateTimePicker_TimeTo.Value = new DateTime(DateTime.Now.Year, DateTime.Now.AddDays(-1).Month, DateTime.Now.AddDays(-1).Day, 23, 59, 0);
                    if (!checkBox_Report1.Checked) checkBox_Report1.Checked = true;
                    if (!checkBox_Report2.Checked) checkBox_Report2.Checked = true;
                    if (checkBox_Report3.Checked) checkBox_Report3.Checked = false;
                    DateTimeNowWeek = dateTimePicker_TimeTo.Value.AddDays(7);
                    break;
                case "За прошлую неделю":
                    dateTimePicker_TimeFrom.Value = new DateTime(
                        DateTime.Now.AddDays(-(6 + Convert.ToInt32(DateTime.Now.DayOfWeek))).Year,
                        DateTime.Now.AddDays(-(6 + Convert.ToInt32(DateTime.Now.DayOfWeek))).Month,
                        DateTime.Now.AddDays(-(6 + Convert.ToInt32(DateTime.Now.DayOfWeek))).Day,
                        0, 0, 0);
                    dateTimePicker_TimeTo.Value = new DateTime(
                        DateTime.Now.AddDays(0 - DateTime.Now.DayOfWeek).Year,
                        DateTime.Now.AddDays(0 - DateTime.Now.DayOfWeek).Month,
                        DateTime.Now.AddDays(0 - DateTime.Now.DayOfWeek).Day,
                        23, 59, 0);
                    if (!checkBox_Report1.Checked) checkBox_Report1.Checked = true;
                    if (!checkBox_Report2.Checked) checkBox_Report2.Checked = true;
                    if (checkBox_Report3.Checked) checkBox_Report3.Checked = false;
                    DateTimeNowWeek = dateTimePicker_TimeTo.Value.AddDays(7);
                    break;
                case "За прошлый месяц":
                    dateTimePicker_TimeFrom.Value = new DateTime(
                        DateTime.Now.Year,
                        DateTime.Now.AddMonths(-1).Month,1, 0, 0, 0);
                    dateTimePicker_TimeTo.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1, 23, 59, 0);
                    dateTimePicker_TimeTo.Value = dateTimePicker_TimeTo.Value.AddDays(-1);
                    // Для авторегулирования дат по старому отчетному файлу выгрузки.
                    //dateTimePicker_TimeFrom.Value = FileExportSD_CreationTime.AddMonths(-1);
                    //dateTimePicker_TimeFrom.Value = new DateTime(dateTimePicker_TimeFrom.Value.Year, dateTimePicker_TimeFrom.Value.Month, 1, 0, 0, 0);
                    //dateTimePicker_TimeTo.Value = new DateTime(FileExportSD_CreationTime.Year, FileExportSD_CreationTime.Month, 1, 23, 59, 59);
                    //dateTimePicker_TimeTo.Value = dateTimePicker_TimeTo.Value.AddDays(-1);
                    if (checkBox_Report1.Checked) checkBox_Report1.Checked = false;
                    if (checkBox_Report2.Checked) checkBox_Report2.Checked = false;
                    if (!checkBox_Report3.Checked) checkBox_Report3.Checked = true;
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

        private void checkBox_DebugInfo_CheckedChanged(object sender, EventArgs e)
        {
            //if (checkBox_DebugInfo.Checked) DebugInfoWriteenabler = true;
        }

        private void checkBox_DebugInfo_CheckStateChanged(object sender, EventArgs e)
        {
            //if (checkBox_DebugInfo.Checked) DebugInfoWriteenabler = true;
                    //else DebugInfoWriteenabler = false;
        }

        ///<summary>
        ///Метод для заполнения текстового окна отладочной информацией.
        ///text - Текст для вывода;
        ///NewLine - Признак новой строки(Необязательный параметр);
        ///DebugLine - Признак отладочной информации(Необязательный параметр).
        ///</summary>
        /// <param name="text">Текст для вывода</param>
        public void DebugInfoWrite(string text, bool NewLine = false, bool DebugLine = false)
        {
            if (DebugLine & !Form0.DebugInfoWriteenabler)
            //Если инфа для вывода относится к Дебаг, а флаг вывода не установлен - то пропускаем.
            {
                return;
            }
            if (Form0.DebugInfoWriteenabler)
            {
                text = DateTime.Now.Minute.ToString("D2") + "." + DateTime.Now.Second.ToString("D2") + "." + DateTime.Now.Millisecond.ToString("D3") + "   " + text;
            }
            if (NewLine)
            {
                text = "\r\n" + text;
            }
            text += "\n";
            //textBox_DebugInfo.AppendText(text);
            //textBox_DebugInfo.Refresh();
        }

        private void textBox_WeekMailTo_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
