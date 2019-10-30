using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ReportMakerSD
{
    public partial class MainForm : Form
    {
        /// <summary>
        /// Метод заполнения массива для отчета по подрядникам
        /// </summary>
        void ParsingDataForReport_TO()
        {
            //Подготовка
            NameReport_TO = "Отчет по подрядным организациям с " + dateTimePicker_TimeFrom.Value.ToString("dd.MM.yyyy") + " по " + dateTimePicker_TimeTo.Value.ToString("dd.MM.yyyy");  //Наименование отчета
            wStr("Собираем данные для отчета: " + NameReport_TO, true);
            DataForReport_TO = new object[2][,];    //В книге 2 листа;
            string countTO1 = ""; //Кол-во обращений для 1 листа;
            string countTO2 = ""; //Кол-во обращений для 1 листа;
            string[] numberRowTO;   //Времянка для количества строк на листе.
            int countRowTO = 0; //Максимальное количество строк в отчете (по двум листам)
                                //формирование списка с необходимыми строками для отчета.
            for (int i = 2; i <= lastRowIndex; i++) //Перебираем все строки с данными.
            {
                if (exportSDData[i, Convert.ToInt32(NameColumnsONLY[6, 1])].ToString().Contains("ТО \\"))  //Берем только ФГП содержащие "ТО \\"
                {
                    for (int f = 0; f < CountNameFilialBranch; f++) //Перебираем филиалы.
                    {
                        if (exportSDData[i, Convert.ToInt32(NameColumnsONLY[4, 1])].ToString().Contains(NameFilialBranch[f, 0].ToString())) //Нашли подходящий филиал...
                        {
                            for (int j = 0; j < CountStatusesNames; j++) //Перебираем все статусы.
                            {
                                if (exportSDData[i, Convert.ToInt32(NameColumnsONLY[3, 1])].ToString().Contains(StatusesNames[j]))  //Если нашли требуемый статус, то...
                                {
                                    countTO1 = countTO1 + i.ToString() + ",";
                                }
                            }
                            if (Convert.ToDateTime(exportSDData[i, Convert.ToInt32(NameColumnsONLY[1, 1])]) >= dateTimePicker_TimeFrom.Value && Convert.ToDateTime(exportSDData[i, Convert.ToInt32(NameColumnsONLY[1, 1])]) <= dateTimePicker_TimeTo.Value)
                            //Если искомое в необходимых пределах, то...
                            {
                                countTO2 = countTO2 + i.ToString() + ",";
                            }
                        }
                    }
                }
            }
            //Заполняем массив для отчета
            if ((countTO1.Split(',')).Length > (countTO2.Split(',')).Length) { countRowTO = (countTO1.Split(',')).Length; };
            if ((countTO1.Split(',')).Length < (countTO2.Split(',')).Length) { countRowTO = (countTO2.Split(',')).Length; };
            //Определяем размерность каждого листа в отчете
            DataForReport_TO[0] = new object[(countTO1.Split(',')).Length, CountNameColumnsONLY];//Количество строк и колонок на 1 листе
            DataForReport_TO[1] = new object[(countTO2.Split(',')).Length, CountNameColumnsONLY];//Количество строк и колонок на 2 листе
            if (countRowTO != 0)
            {
                //Массив для 1-листа.
                numberRowTO = countTO1.Split(',');
                countRowTO = numberRowTO.Length;
                for (int i = 0; i < countRowTO - 1; i++) //Перебираем все необходимые строки с данными
                {
                    for (int j = 0; j < CountNameColumnsONLY - 1; j++)    //Перебираем необходимые поля
                    {
                        DataForReport_TO[0][i, j] = exportSDData[Convert.ToInt32(numberRowTO[i]), Convert.ToInt32(NameColumnsONLY[j + 1, 1])];
                    }
                }
                //Массив для 2-листа.
                numberRowTO = countTO2.Split(',');
                countRowTO = numberRowTO.Length;
                for (int i = 0; i < countRowTO - 1; i++) //Перебираем все необходимые строки с данными
                {
                    for (int j = 0; j < CountNameColumnsONLY - 1; j++)    //Перебираем необходимые поля
                    {
                        DataForReport_TO[1][i, j] = exportSDData[Convert.ToInt32(numberRowTO[i]), Convert.ToInt32(NameColumnsONLY[j + 1, 1])];
                    }
                }
            }
            //Получили массив DataForReport_TO с двумя листами и данными для отчета.
            //wStr("Данные для отчета по подрядным организациям подготовлены./r/n",true);
            wStr(" => Собрали...\r\n", true);
        }

        ///<summary>
        ///Загрузка данных из файла.
        ///Передаем полный путь к файлу с данными
        ///Получаем массив с данными
        ///</summary>
        ///<param FileNameExportSD="string">путь до файла с данными</param>
        object[,] LoadDataFromFileExportSD(string FileNameExportSD)
        {
            object[,] openData = null;
            try
            {
                oExcelApp = new Excel.Application();
                oExcelApp.DisplayAlerts = false;
                oExcelApp.Visible = false;
                //FileName = "D:\\0\\export.xlsx";
                //Excel.Workbook openBook = excelApp.Workbooks.Open(openFileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                openbooks = oExcelApp.Workbooks;
                openbook = openbooks.Open(FileNameExportSD);
                openSheets = openbook.Sheets;
                openSheet = openSheets[1];
                openCellsFirst = openSheet.Cells[1, 1];
                openData = openSheet.Range[openCellsFirst, openSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell)].Value;

                wStr("Проверка на '='. => ", true, true);
                for (int row = 1; row < openData.GetLength(0); row++)
                {
                    for (int col = 1; col < openData.GetLength(1); col++)
                    {
                        if (openData[row, col] != null)
                        {
                            openData[row, col] = openData[row, col].ToString().TrimStart('=');
                        }
                    }
                }
                wStr("Проверка на '=' произведена.", false, true);
                //!!!Сделать поиск и замену первых "=" во всех ячейках.
            }
            #region Закрываем Excel
            finally
            {
                openbook.Close(false, false, false);
                //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openCellsFirst);
                //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openSheet);
                //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openSheets);
                //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openbook);
                //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(openbooks);

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
                //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oExcelApp);
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
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
            #endregion
            return openData;
        }

        void SaveDataToReportFile(string ReportFileName, string sheetNames, object[,,] ReportSaveData)
        //Формирование XLS-файла и выгрузка данных
        //Прекращение поддержки по формированию button_CreateFiles_Click от 20161101
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
                for (int i = 1; i <= NumberSheets; i++)
                {
                    //Задаем наименование листа
                    openbook.Sheets[i].Name = sheetsNames[i - 1];
                    //Заполняем шапку
                    for (int n = 1; n < NameColumnsONLY.GetLength(0); n++)
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
                    for (int r = 0; r < NumberRow; r++) //перебираем все строки
                    {
                        for (int c = 0; c < NumberCol - 1; c++)
                        {
                            DataForSheet[r, c] = ReportSaveData[i - 1, r, c + 1];
                        }
                    }
                    openbook.Sheets[i].Range[openbook.Sheets[i].cells[2, 1], openbook.Sheets[i].cells[NumberRow, NumberCol]].Value = DataForSheet;
                    //Форматирование
                    //по строчное. Долго выполняется.
                    for (int r = 2; r <= NumberRow; r++)
                    {
                        openRange = openbook.Sheets[i].Range[openbook.Sheets[i].cells[r, 1], openbook.Sheets[i].cells[r, NumberCol - 1]];
                        var t = openbook.Sheets[i].cells[r, 1].value;
                        if (t == null) break;
                        openRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        openRange.WrapText = null;
                        if (r % 2 != 0)
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

        /// <summary>
        /// Метод заполнения массивов для отчетов по филиалам
        /// </summary>
        void ParsingDataForReport_RDU()
        {
            //Формируем данные для отчета: "Нерешенные обращения по РДУ за период" в массиве object[][,] DataForReport_RDU;
            //Кол-во листов = кол-во филиалов
            //Условие: (6=ФГП начинается с "ТО \" ) и ( (3=Статус = {4 В ожидании:3 Выполняется:2 Назначен:1 Новый})
            //4=Филиал заявителя должен быть нашей ОЗ

            //Необходимые данные:
            //              exportSDData            - исходные данные;
            //              NameColumnsONLY         - наименования колонок в отчете;
            //object[][,]   DataForReport_RDU       - массив для данных для отчета;
            //int           CountNameFilialBranch   - количество филиалов;
            //object[,]     NameFilialBranch        - справочник филиалов;
            //int           lastRowIndex            - количество строк исходных данных;
            //string        NameReport_RDU          - Наименование отчета по подрядникам;
            ReportRDU = 0;
            int nFilialBranchName = 0;

            NameReport_RDU = "Нерешенные обращения по РДУ за период с " + dateTimePicker_TimeFrom.Value.ToString("dd.MM.yyyy") + "  по " + dateTimePicker_TimeTo.Value.ToString("dd.MM.yyyy") + ". ";   //Наименование отчета
            wStr("Формируем данные для отчета: " + NameReport_RDU, true, false);

            for (int i = 2; i <= lastRowIndex; i++) //Перебираем все строки с данными
            {
                if (Array.IndexOf(StatusesNames, exportSDData[i, Convert.ToInt32(NameColumnsRDUONLY[3, 1])]) != -1)   //Выбираем только действующие статусы.
                {
                    if (exportSDData[i, Convert.ToInt32(NameColumnsRDUONLY[6, 1])].ToString().Contains("\\"))  //Отсекаем ФГП без филиала.
                    {
                        if (!exportSDData[i, Convert.ToInt32(NameColumnsRDUONLY[6, 1])].ToString().Contains("ТО \\")) //Исключаем обращения на подрядников.
                        {
                            string strFGP = exportSDData[i, Convert.ToInt32(NameColumnsRDUONLY[6, 1])].ToString();

                            nFilialBranchName = Array.IndexOf(FilialBranchName, strFGP.Substring(0, strFGP.IndexOf('\\', 0) - 1));
                            if (nFilialBranchName != -1)
                            {
                                NameFilialBranch[nFilialBranchName, 1] = NameFilialBranch[nFilialBranchName, 1] + i.ToString() + ","; //сохраняем массив №строк с данными по филиалам. Третий столбец пуст - не используем.
                                NameFilialBranch[nFilialBranchName, 2] = Convert.ToInt32(NameFilialBranch[nFilialBranchName, 2]) + 1;
                            }
                            else //заносим ФГП других ОЗ.
                            {
                                NameFilialBranch[CountNameFilialBranch, 1] = NameFilialBranch[CountNameFilialBranch, 1] + i.ToString() + ","; //сохраняем массив №строк с данными по филиалам. Третий столбец пуст - не используем.
                                NameFilialBranch[CountNameFilialBranch, 2] = Convert.ToInt32(NameFilialBranch[CountNameFilialBranch, 2]) + 1;
                            }
                        }
                        //else //заносим подрядников. только для теста
                        //{
                        //    NameFilialBranch[CountNameFilialBranch, 1] = NameFilialBranch[CountNameFilialBranch, 1] + i.ToString() + ","; //сохраняем массив №строк с данными по филиалам. Третий столбец пуст - не используем.
                        //    NameFilialBranch[CountNameFilialBranch, 2] = Convert.ToInt32(NameFilialBranch[CountNameFilialBranch, 2]) + 1;
                        //}
                    }
                    else //заносим все остальное
                    {
                        NameFilialBranch[CountNameFilialBranch, 1] = NameFilialBranch[CountNameFilialBranch, 1] + i.ToString() + ","; //сохраняем массив №строк с данными по филиалам. Третий столбец пуст - не используем.
                        NameFilialBranch[CountNameFilialBranch, 2] = Convert.ToInt32(NameFilialBranch[CountNameFilialBranch, 2]) + 1;
                    }
                }
            }
            ReportRDU_Sheets = new string[CountNameFilialBranch - 1];    //Формируем массив с именами листов для РДУ.
            //размерность массива
            for (int n = 1; n < CountNameFilialBranch; n++)
            {
                ReportRDU_Sheets[n - 1] = NameFilialBranch[n, 0].ToString();   //Собираем имена листов.
            }
            DataForReport_RDU = new object[CountNameFilialBranch - 1][,];   //Формируем массив с данными по каждому РДУ.
            for (int n = 1; n < CountNameFilialBranch; n++)
            {
                DataForReport_RDU[n - 1] = new object[Convert.ToInt32(NameFilialBranch[n, 2]), CountNameColumnsRDUONLY];//Формируем массивы под каждый лист/филиал.
                ReportRDU += Convert.ToInt32(NameFilialBranch[n, 2]);
            }
            if (ReportRDU != 0) //Если выбрана хоть одна запись, то работаем
            {
                for (int f = 1; f < CountNameFilialBranch/* - 1*/; f++) //с 1 - тк пропускаем ОДУ.
                {
                    string[] numberRow = NameFilialBranch[f, 1].ToString().Split(',');//массив строк №записей
                    int countRow = numberRow.Length - 1;//кол-во записей
                    for (int r = 0; r < countRow; r++)//перебираем все записи
                    {
                        for (int c = 1; c < CountNameColumnsRDUONLY; c++)    //Перебираем необходимые поля
                        {
                            if (numberRow[r] != "")
                            {
                                DataForReport_RDU[f - 1][r, c - 1] = exportSDData[Convert.ToInt32(numberRow[r]), Convert.ToInt32(NameColumnsRDUONLY[c, 1])];
                            }
                        }
                    }
                }
                wStr("Данные для отчета 'Нерешенные обращения по РДУ за период' сформированы.\r\n", true);
            }
            else  //Иначе отчключаем создание отчета по РДУ и выводим сообщение.
            {
                wStr("Данные для отчета 'Нерешенные обращения по РДУ за период' ОТСУТСТВУЮТ!!!", true);
                //checkBox_ReportRDU.Checked = false;
            }

            //object[][,] DataForReport_ODU; //по ОДУ
            //При этом сборку данных делали выше. Здесь только запаковка в массив для отчета.
            NameReport_ODU = "Нерешенные обращения по ОДУ Северо-Запада за период с " + dateTimePicker_TimeFrom.Value.ToString("dd.MM.yyyy") + "  по " + dateTimePicker_TimeTo.Value.ToString("dd.MM.yyyy") + ". ";
            wStr("Формируем данные для отчета: " + NameReport_ODU, true);
            //ReportRDU_max = NameFilialBranch[0, 1].ToString().Split(',').Length;
            ReportODU = Convert.ToInt32(NameFilialBranch[0, 2]);
            ReportODU2 = Convert.ToInt32(NameFilialBranch[CountNameFilialBranch, 2]);
            if (ReportODU != 0) //Если выбрана хоть одна запись, то работаем
            {
                DataForReport_ODU = new object[2][,];   //1 лист - так как ОДУ; 2 - в ответственности вне ОЗ СЗ.
                DataForReport_ODU[0] = new object[ReportODU, CountNameColumnsONLY];//Количество строк и колонок на листе
                //DataForReport_ODU[1] = new object[ReportODU2, CountNameColumnsONLY];//Считаем, что нерешенные обращения ОДУ всегда есть.
                {
                    string[] numberRow = NameFilialBranch[0, 1].ToString().Split(',');//массив строк №записей
                    for (int r = 0; r <= ReportODU; r++)//перебираем все записи
                    {
                        for (int c = 1; c < CountNameColumnsONLY; c++)    //Перебираем необходимые поля
                        {
                            if (numberRow[r] != "")
                            {
                                DataForReport_ODU[0][r, c - 1] = exportSDData[Convert.ToInt32(numberRow[r]), Convert.ToInt32(NameColumnsONLY[c, 1])];
                            }
                        }
                    }
                }
                ReportODU_Sheets = NameFilialBranch[0, 0].ToString();
                wStr("Данные для отчета 'Нерешенные обращения по ОДУ за период' подготовлены.\r\n", true);
                if (ReportODU2 != 0) //Если выбрана хоть одна запись, то работаем/ Считаем, что нерешенные обращения ОДУ всегда есть.
                {
                    DataForReport_ODU[1] = new object[ReportODU2, CountNameColumnsRDUONLY];
                    {
                        string[] numberRow = NameFilialBranch[CountNameFilialBranch, 1].ToString().Split(',');//массив строк №записей
                        for (int r = 0; r <= ReportODU2; r++)//перебираем все записи
                        {
                            for (int c = 1; c < CountNameColumnsONLY; c++)    //Перебираем необходимые поля
                            {
                                if (numberRow[r] != "")
                                {
                                    DataForReport_ODU[1][r, c - 1] = exportSDData[Convert.ToInt32(numberRow[r]), Convert.ToInt32(NameColumnsONLY[c, 1])];
                                }
                            }
                        }
                    }
                    //ReportODU_Sheets = NameFilialBranch[0, 0].ToString();
                    //wStr("Данные для отчета 'Нерешенные обращения по ОДУ за период' подготовлены.");
                }
                else
                {
                    wStr("Данные для отчета 'Нерешенные обращения ВНЕ ОЗ СЗ за период' ОТСУТСТВУЮТ!!!", true);
                    //checkBox_ReportRDU.Checked = false;
                }
            }
            else
            {
                wStr("Данные для отчета 'Нерешенные обращения по ОДУ за период' ОТСУТСТВУЮТ!!!", true);
                //checkBox_ReportRDU.Checked = false;
            }
        }

        /// <summary>
        /// Метод заполнения массивов для отчета за период
        /// </summary>
        void ParsingDataForReport_Based()
        {
            //Формируем данные для отчета: "Обращения зарегистрированные за период с  по .xlsx" в массиве object[][,] DataForReport_based;
            //Каждый лист отдельный отчет:
            //1=все обращения за период. На основе этих данных строятся остальные листы.
            //1.Условие: (1=Дата окончания регистрации в необходимых пределах)
            //Осторожно с выгрузкой!!!
            //
            //Условие: (6=ФГП начинается с "ТО \" ) и ( (3=Статус = {4 В ожидании:3 Выполняется:2 Назначен:1 Новый}) или (1=Дата окончания регистрации в необходимых пределах)
            //4=Филиал заявителя должен быть нашей ОЗ

            //Необходимые данные:
            //              exportSDData            - исходные данные;
            //              NameColumnsONLY         - наименования колонок в отчете;
            //object[][,]   DataForReport_based     - массив для данных для отчета;
            //int           CountNameFilialBranch   - количество филиалов;
            //object[,]     NameFilialBranch        - справочник филиалов;
            //int           lastRowIndex            - количество строк исходных данных;
            //int           lastCellIndex           - количество колонок исходных данных;
            //string        NameReport_based          - Наименование отчета по подрядникам;
            string[] numberRow; //Массив с №строк.
            bool testFGP;
            int Report_based_SheetsNumber = Properties.Settings.Default.Report_based_SheetsName.Split(':').Length;  //В книге 5 листов. По количеству наименований листов отчета.
            if (Report_based_SheetsNumber != 5)
            {
                wStr("Наименований листов не равно 5: " + Properties.Settings.Default.Report_based_SheetsName, true, false);
                return;
            }
            int int_count;  //Считаем количество обращений (+строк).
            //Наименование отчета
            NameReport_based = "Обращения зарегистрированные за период с " + dateTimePicker_TimeFrom.Value.ToString("dd.MM.yyyy") + " по " + dateTimePicker_TimeTo.Value.ToString("dd.MM.yyyy");
            wStr("Формируем данные для отчета: " + NameReport_based, true);
            //Формируем базовый массив.
            string count = "1,";    //Учитываем шапку.
            for (int i = 2; i <= lastRowIndex; i++) //Перебираем все строки с данными (собираем подходящие №строк исходных данных)
            {
                //Если дата в требуемых пределах считаем:
                if (Convert.ToDateTime(exportSDData[i, Convert.ToInt32(NameColumnsONLY[1, 1])]) >= dateTimePicker_TimeFrom.Value && Convert.ToDateTime(exportSDData[i, Convert.ToInt32(NameColumnsONLY[1, 1])]) <= dateTimePicker_TimeTo.Value)
                {
                    count = count + i.ToString() + ",";
                }
            }
            int_count = Convert.ToInt32(count.Split(',').Length) - 1;    //Считаем количество обращений (+строк). И минусуем пустое окончание.
            DataForReport_based = new object[Report_based_SheetsNumber][,];
            #region DataForReport_based_1
            //Формируем 1 лист. Это все прошедшие по фильтру обращения.
            wStr("Формируем данные для 1-листа: ", false, true);
            DataForReport_based[0] = new object[int_count, lastCellIndex];  //Формируем массив для 1 листа.
            numberRow = count.Split(',');
            for (int r = 0; r < int_count; r++)//перебираем все записи.
            {
                for (int c = 0; c < lastCellIndex; c++) //Перебираем все колонки.
                {
                    DataForReport_based[0][r, c] = exportSDData[Convert.ToInt32(numberRow[r]), c + 1];
                }
            }
            wStr("Сформировали данные для 1-листа: ", false, true);
            #endregion
            #region DataForReport_based_2
            wStr("Формируем данные для 2-листа: ", false, true);
            //Формируем 2 лист.
            DataForReport_based[1] = new object[2, CountNameFilialBranch];  //Формируем массив для 2 листа.
            for (int i = 1; i <= int_count - 1; i++) //Перебираем все строки с уже необходимыми данными
            {
                for (int f = 0; f < CountNameFilialBranch; f++) //Перебираем все филиалы (РДУ и ОДУ)
                {
                    DataForReport_based[1][0, f] = NameFilialBranch[f, 0]; //сохраняем наименование филиала
                    if (DataForReport_based[0][i, Convert.ToInt32(NameColumnsONLY[4, 1]) - 1].ToString().Contains(NameFilialBranch[f, 0].ToString())) //Если найден необходимый филиал, то
                    {
                        DataForReport_based[1][1, f] = Convert.ToInt32(DataForReport_based[1][1, f]) + 1;   //и количество обращений по нему.
                    }
                    if (DataForReport_based[1][1, f] == null)
                    {
                        DataForReport_based[1][1, f] = 0;
                    }
                }
            }
            wStr("Сформировали данные для 2-листа: ", false, true);
            #endregion
            #region DataForReport_based_3
            wStr("Формируем данные для 3-листа: ", false, true);
            //Формируем 3 лист.
            testFGP = true;
            DataForReport_based[2] = new object[CountFGPNames + 2, CountNameFilialBranch + 2];  //Формируем массив для 3 листа.
            for (int i = 0; i < CountFGPNames; i++)
            {
                DataForReport_based[2][i + 1, 0] = FGPNames[i]; //Заполняем первую колонку наименованиями ФГП
            }
            DataForReport_based[2][CountFGPNames + 1, 0] = "ИТОГО"; //Итого по столбцам.
            DataForReport_based[2][0, CountNameFilialBranch + 1] = "ИТОГО"; //Итого по строкам.

            for (int i = 1; i <= int_count - 1; i++) //Перебираем все строки с уже отобранными данными.
            {
                testFGP = true;
                for (int f = 0; f < CountNameFilialBranch; f++) //Перебираем все филиалы (РДУ и ОДУ).
                {
                    DataForReport_based[2][0, f + 1] = NameFilialBranch[f, 0]; //сохраняем наименование филиала, т.е. заполняем первую строку.
                    if (DataForReport_based[0][i, Convert.ToInt32(NameColumnsONLY[4, 1]) - 1].ToString().Contains(NameFilialBranch[f, 0].ToString())) //Если найден необходимый филиал, то...
                    {
                        testFGP = true;
                        for (int fgp = 0; fgp < CountFGPNames; fgp++)   //Перебираем необходимые ФГП
                        {
                            if (DataForReport_based[0][i, Convert.ToInt32(NameColumnsONLY[6, 1]) - 1].ToString().Contains(FGPNames[fgp]))   //Если ФГП найдет, то...
                            {
                                DataForReport_based[2][fgp + 1, f + 1] = Convert.ToInt32(DataForReport_based[2][fgp + 1, f + 1]) + 1;
                                testFGP = false;
                                break;
                            }
                        }
                        //Перевести прочее так же в справочную информацию!!!
                        if (testFGP)    //Иначе запихиваем все в прочие.
                        {
                            DataForReport_based[2][14, f + 1] = Convert.ToInt32(DataForReport_based[2][14, f + 1]) + 1;
                        }
                    }
                    if (!testFGP)
                    {
                        break;
                    }
                }
            }
            wStr("Сформировали данные для 3-листа: ", false, true);
            #endregion
            #region DataForReport_based_4
            wStr("Формируем данные для 4-листа: ", false, true);
            //Формируем 4 лист.
            //    private string[] FGPNames;
            //private int CountFGPNames;
            testFGP = true;
            DataForReport_based[3] = new object[CountFGPNames + 2, CountNameFilialBranch + 2];  //Формируем массив для 3 листа.
            for (int i = 0; i < CountFGPNames; i++)
            {
                DataForReport_based[3][i + 1, 0] = FGPNames[i]; //Заполняем первую колонку наименованиями ФГП
            }
            DataForReport_based[3][CountFGPNames + 1, 0] = "ИТОГО"; //Итого по столбцам.
            DataForReport_based[3][0, CountNameFilialBranch + 1] = "ИТОГО"; //Итого по строкам.

            for (int i = 1; i <= int_count - 1; i++) //Перебираем все строки с уже отобранными данными.
            {
                testFGP = true;
                for (int f = 0; f < CountNameFilialBranch; f++) //Перебираем все филиалы (РДУ и ОДУ).
                {
                    DataForReport_based[3][0, f + 1] = NameFilialBranch[f, 0]; //сохраняем наименование филиала, т.е. заполняем первую строку.
                    if (DataForReport_based[0][i, Convert.ToInt32(NameColumnsONLY[6, 1]) - 1].ToString().Contains(NameFilialBranch[f, 0].ToString())) //Если найден необходимый филиал, то...
                    {
                        testFGP = true;
                        for (int fgp = 0; fgp < CountFGPNames; fgp++)   //Перебираем необходимые ФГП
                        {
                            if (DataForReport_based[0][i, Convert.ToInt32(NameColumnsONLY[6, 1]) - 1].ToString().Contains(FGPNames[fgp]))   //Если ФГП найдет, то...
                            {
                                DataForReport_based[3][fgp + 1, f + 1] = Convert.ToInt32(DataForReport_based[3][fgp + 1, f + 1]) + 1;
                                testFGP = false;
                                break;
                            }
                        }
                    }
                    if (!testFGP)
                    {
                        break;
                    }
                }
            }
            wStr("Сформировали данные для 4-листа: ", false, true);
            #endregion
            #region DataForReport_based_5
            wStr("Формируем данные для 5-листа: ", false, true);
            //Формируем 5 лист.
            testFGP = true;
            DataForReport_based[4] = new object[CountWayOfGivingTreatments + 2, CountNameFilialBranch + 2];  //Формируем массив для 3 листа.
            for (int i = 0; i < CountWayOfGivingTreatments; i++)
            {
                DataForReport_based[4][i + 1, 0] = WayOfGivingTreatments[i]; //Заполняем первую колонку
            }
            DataForReport_based[4][CountWayOfGivingTreatments + 1, 0] = "ИТОГО"; //Итого по столбцам.
            DataForReport_based[4][0, CountNameFilialBranch + 1] = "ИТОГО"; //Итого по строкам.

            for (int i = 1; i <= int_count - 1; i++) //Перебираем все строки с уже отобранными данными.
            {
                testFGP = true;
                for (int f = 0; f < CountNameFilialBranch; f++) //Перебираем все филиалы (РДУ и ОДУ).
                {
                    DataForReport_based[4][0, f + 1] = NameFilialBranch[f, 0]; //сохраняем наименование филиала, т.е. заполняем первую строку.
                    if (DataForReport_based[0][i, Convert.ToInt32(NameColumnsONLY[4, 1]) - 1].ToString().Contains(NameFilialBranch[f, 0].ToString()))
                    {
                        testFGP = true;
                        for (int row = 0; row < CountWayOfGivingTreatments; row++) //Перебираем все источники.
                        {
                            if (DataForReport_based[0][i, Convert.ToInt32(NameColumnsONLY[12, 1]) - 1].ToString().Contains(WayOfGivingTreatments[row].ToString())) //Если найден необходимый источник:
                            {
                                DataForReport_based[4][row + 1, f + 1] = Convert.ToInt32(DataForReport_based[4][row + 1, f + 1]) + 1;
                                testFGP = false;
                                break;
                            }
                        }
                    }
                    if (!testFGP)
                    {
                        break;
                    }
                }
            }
            wStr("Сформировали данные для 5-листа: ", false, true);
            #endregion

            wStr("Данные для отчета '" + NameReport_based + "' сформированы.");
        }

        /// <summary>
        /// Создание всех файлов-отчетов
        /// </summary>
        void CreateFiles()
        {
            button_parsingReport.Enabled = false;
            button_OpenFolder.Enabled = false;
            button_SendMail.Enabled = false;

            oExcelApp = new Excel.Application();
            oExcelApp.DisplayAlerts = false;
            // true => false
            oExcelApp.ScreenUpdating = false;   //Отключение обновления
            oExcelApp.Visible = false;
            //oExcelApp.ScreenUpdating = true;
            //oExcelApp.Visible = true;
            //Кол-во листов в новой книге
            oExcelApp.SheetsInNewWorkbook = 1;
            //Ширина колонок
            //oExcelApp.Columns.ColumnWidth = 15;
            oExcelBooks = oExcelApp.Workbooks;
            openbooks = oExcelApp.Workbooks;
            //Проверка требуемых отчетов по заполненностью массивов.
            //Если нет сформированных массивов то выход.

            #region Report_based_toExcel
            if (DataForReport_based != null)
            { //Делаем базовый отчет за период
                //При изменении количества строк на листах 3 и 4 необходимо исправить формулы ИТОГО по столбцам!!!
                wStr("Формируем '" + NameReport_based + "'.", true);
                oExcelBook = oExcelBooks.Add(System.Reflection.Missing.Value); //Добавляем книгу;
                oExcelSheets = oExcelBook.Worksheets;   //Определяем набор листов.
                oExcelSheet = (_Worksheet)oExcelSheets.Item[1];   //Берем 1 лист.
                wStr("Именуем листы в книге.", true, true);
                for (int i = DataForReport_based.GetLength(0) - 1; i > 0; i--)
                {
                    oExcelBook.Sheets[1].Name = Report_based_Sheets[i]; //Переименовываем лист.
                    oExcelBook.Sheets.Add(System.Reflection.Missing.Value); //Вставляем лист.
                    oExcelBook.Sheets[1].Name = Report_based_Sheets[i - 1]; //Иначе последний лист остается без имени/
                }
                for (int i = 0; i <= DataForReport_based.GetLength(0) - 1; i++)   //ОБработка каждого листа поочередно.
                {
                    wStr("Формирование " + (i + 1).ToString() + "-листа.", true, true);
                    int NumberRow = DataForReport_based[i].GetLength(0);    //Определяю количество строк на лист.
                    int NumberCol = DataForReport_based[i].GetLength(1);    //Определяю количество колонок на лист.
                    oExcelSheet = (_Worksheet)oExcelSheets.Item[i + 1];   //Беру один лист.

                    wStr("Формирование шапки. Задаем диапазон.", true, true);
                    oExcelRange = oExcelSheet.Range["A1"]; //Берем первую ячейку.
                    oExcelRange = oExcelRange.Resize[1, DataForReport_based[i].GetLength(1)]; //И переопределяю диапазон на необходимый.
                    wStr("Формирование шапки. Перенос строк.", true, true);
                    oExcelRange.WrapText = true;
                    wStr("Формирование шапки. Жирный шрифт.", true, true);
                    oExcelRange.Font.Bold = true;
                    //wStr("Формирование шапки.", true, true);
                    //oExcelRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    wStr("Формирование шапки. Ширина колонок.", true, true);
                    if (i == 1)
                    {
                        oExcelRange.ColumnWidth = 25;
                    }
                    else
                    {
                        oExcelRange.ColumnWidth = 23;
                    }
                    wStr("Формирование шапки. Опять диапазон.", true, true);

                    oExcelRange = oExcelSheet.Range["A1"]; //Берем первую ячейку.
                    oExcelRange = oExcelRange.Resize[NumberRow, NumberCol]; //И переопределяю диапазон на необходимый.
                    //oExcelRange.Borders.LineStyle = XlLineStyle.xlContinuous;   //Устанавливаю границы ячеек.

                    wStr("Вставляем данные.", false, true);
                    oExcelRange.Value = DataForReport_based[i];
                    /*oExcelRange.Value2 = DataForReport_based[i];*/    //Поехали Даты.
                    #region Форматируем таблицу
                    wStr("Форматируем таблицу.", false, true);
                    //oExcelRange.WrapText = null;
                    oExcelRange.WrapText = false;
                    for (int j = 1; j <= NumberRow; j++)    //Перебираем каждую строку листа для форматирования.
                    {
                        oExcelRange = oExcelSheet.Range["A" + j.ToString()]; //Берем первую ячейку.
                        oExcelRange = oExcelRange.Resize[1, NumberCol]; //И переопределяю диапазон на необходимый.
                        oExcelRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        if (j == 1)
                        {
                            oExcelRange.Interior.Color = Color.CornflowerBlue;
                            oExcelRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            oExcelRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);  //Установил на первую строку автофильтр.
                        }
                        else if (j % 2 == 0)
                        {
                            oExcelRange.Interior.Color = Color.LightSkyBlue;
                            //Если это 3 лист:
                            if (i >= 2)
                            {
                                oExcelSheet.Range[oExcelSheet.Cells[j, NumberCol], oExcelSheet.Cells[j, NumberCol]].FormulaLocal = "=СУММ(RC[-8]:RC[-1])";
                            }
                        }
                        else
                        {
                            oExcelRange.Interior.Color = Color.MintCream;
                            //Если это 3 лист:
                            if (i >= 2)
                            {
                                oExcelSheet.Range[oExcelSheet.Cells[j, NumberCol], oExcelSheet.Cells[j, NumberCol]].FormulaLocal = "=СУММ(RC[-8]:RC[-1])";
                            }
                        }
                    }
                    switch (i)
                    {
                        case 0:
                            oExcelSheet.Columns[1].ColumnWidth = 11;
                            oExcelSheet.Columns[2].ColumnWidth = 20;
                            oExcelSheet.Columns[4].ColumnWidth = 15;
                            oExcelSheet.Columns[5].ColumnWidth = 15;
                            break;
                        case 1:
                            //Формирование графика.
                            ChartObjects chartsobjrcts = (ChartObjects)oExcelSheet.ChartObjects(Type.Missing);
                            ChartObject chartsobjrct = chartsobjrcts.Add(0, 30, 1080, 390); //60,870 - было. Зависит от ширины столбцов (25).
                            Chart chart = chartsobjrct.Chart;
                            chart.SetSourceData(oExcelSheet.get_Range("A1", "H2"), Type.Missing);
                            chart.ChartType = XlChartType.xlColumnClustered;
                            chart.Legend.Delete();
                            Chart excelchart = chartsobjrct.Chart.Location(XlChartLocation.xlLocationAsObject, "График");

                            break;
                        case 2:
                            oExcelSheet.Columns[1].ColumnWidth = 45;
                            oExcelSheet.Range[oExcelSheet.Cells[NumberRow + 1, 1], oExcelSheet.Cells[NumberRow + 1, 1]].Value = "Отчет отражает обращения сотрудников ОЗ.";
                            oExcelRange.Font.Bold = true; //Выделяем ИТОГО по столбцам.
                            oExcelRange.Font.Bold = true; //Выделяем ИТОГО по строкам.
                            oExcelSheet.Range[oExcelSheet.Cells[NumberRow, 2], oExcelSheet.Cells[NumberRow, NumberCol]].FormulaLocal = "=СУММ(R[-21]C:R[-1]C)";
                            break;
                        case 3:
                            oExcelSheet.Columns[1].ColumnWidth = 45;
                            oExcelSheet.Range[oExcelSheet.Cells[NumberRow + 1, 1], oExcelSheet.Cells[NumberRow + 1, 1]].Value = "Отчет отражает обращения, решенные специалистами ФГП ОЗ.";
                            oExcelRange.Font.Bold = true; //Выделяем ИТОГО по столбцам.
                            oExcelRange.Font.Bold = true; //Выделяем ИТОГО по строкам.
                            oExcelSheet.Range[oExcelSheet.Cells[NumberRow, 2], oExcelSheet.Cells[NumberRow, NumberCol]].FormulaLocal = "=СУММ(R[-21]C:R[-1]C)";
                            break;
                        case 4:
                            oExcelSheet.Columns[1].ColumnWidth = 30;
                            oExcelRange.Font.Bold = true; //Выделяем ИТОГО по столбцам.
                            oExcelRange.Font.Bold = true; //Выделяем ИТОГО по строкам.
                            oExcelSheet.Range[oExcelSheet.Cells[NumberRow, 2], oExcelSheet.Cells[NumberRow, NumberCol]].FormulaLocal = "=СУММ(R[-9]C:R[-1]C)";
                            break;
                        default:
                            break;
                    }
                    #endregion
                }
                oExcelBook.SaveAs(PathToFileExportSD + "\\" + NameReport_based + ".xlsx");  //Сохранили книгу;
                oExcelBooks.Close(); //Закрыли книгу;
                wStr("'" + NameReport_based + "' сформирован.", true);
            }
            #endregion

            #region Report_TO_toExcel
            if (DataForReport_TO != null)
            {
                wStr("Формируем отчет: '" + NameReport_TO + "'. ", true, false);
                //Формируем отчет по подрядникам;, false
                //создаем новую книгу
                openbook = openbooks.Add(); //Добавляем книгу;
                openSheets = openbook.Sheets;
                openbook.Sheets.Add();
                openbook.Sheets[1].Name = "Нерешенные по ТО";   //Наименование первого листа.
                openbook.Sheets[2].Name = "ТО за период";   //Наименование второго листа.
                for (int i = 0; i <= 1; i++)   //Положили данные на каждый i-лист
                {
                    for (int n = 1; n < NameColumnsONLY.GetLength(0); n++)  //Положили шапку
                    {
                        openbook.Sheets[i + 1].cells[1, n].WrapText = true;
                        openbook.Sheets[i + 1].cells[1, n].Font.Bold = true;
                        //openbook.Sheets[i + 1].cells[1, n].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        openbook.Sheets[i + 1].cells[1, n].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        //openbook.Sheets[i + 1].cells[1, n].Borders.LineStyle = XlLineStyle.xlContinuous;
                        //openbook.Sheets[i + 1].cells[1, n].Interior.Color = ColorTranslator.ToOle(Color.CornflowerBlue);
                        openbook.Sheets[i + 1].Columns[n].ColumnWidth = 25;
                        openbook.Sheets[i + 1].cells[1, n].Value = NameColumnsONLY[n, 0];
                    }
                    int NumberRow = DataForReport_TO[i].GetLength(0);
                    int NumberCol = DataForReport_TO[i].GetLength(1) - 1;
                    //Вставили данные:
                    openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].cells[2, 1], openbook.Sheets[i + 1].cells[NumberRow, NumberCol]].Value = DataForReport_TO[i];
                    //Установили требуемые стили;
                    openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].cells[2, 1], openbook.Sheets[i + 1].cells[NumberRow, NumberCol]].Borders.LineStyle = XlLineStyle.xlContinuous;
                    openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].cells[2, 1], openbook.Sheets[i + 1].cells[NumberRow, NumberCol]].WrapText = null;
                    for (int j = 1; j <= NumberRow; j++)
                    {
                        if (j == 1)
                        {
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Interior.Color = Color.CornflowerBlue;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Font.Bold = true;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Borders.LineStyle = XlLineStyle.xlContinuous;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].AutoFilter();
                        }
                        else if (j % 2 == 0)
                        {
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Interior.Color = Color.LightSkyBlue;
                        }
                        else
                        {
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Interior.Color = Color.MintCream;
                        }
                    }
                }
                openbook.SaveAs(PathToFileExportSD + "\\" + NameReport_TO + ".xlsx");  //Сохранили книгу;
                openbooks.Close(); //Закрыли книгу;
                //wStr("'" + NameReport_TO + "' сформирован.");
                wStr(" => отчет сформирован...\r\n", true);
            }
            #endregion

            #region Report_RDU_toExcel
            if (DataForReport_RDU != null && ReportRDU != 0)
            {
                wStr("Формируем отчет: '" + NameReport_RDU + "'. ", true, false);
                //Формируем отчет по РДУ;
                //string NameReport_RDU;   //Наименование отчета по РДУ;
                //string ReportRDUSheets = ""; //Наименования листов в книге отчета по РДУ
                //private object[][,] DataForReport_RDU;//по РДУ
                openbook = openbooks.Add(); //Добавляем книгу;
                openSheets = openbook.Sheets;
                for (int i = DataForReport_RDU.GetLength(0) - 1; i > 0; i--)
                {
                    openbook.Sheets[1].Name = ReportRDU_Sheets[i]; //Переименовываем листы.
                    openbook.Sheets.Add(); //Вставляем необходимое количество листов.
                    openbook.Sheets[1].Name = ReportRDU_Sheets[i - 1]; //Иначе последний лист остается без имени
                    //oExcelApp.ActiveWindow.Zoom = 95;  //Тест масштаба листа. Нефурычит!!!
                }
                for (int i = 0; i <= DataForReport_RDU.GetLength(0) - 1; i++)   //Положили данные на каждый i-лист
                {
                    for (int n = 1; n < NameColumnsRDUONLY.GetLength(0); n++)  //Положили шапку
                    {
                        openbook.Sheets[i + 1].cells[1, n].WrapText = true;
                        openbook.Sheets[i + 1].cells[1, n].Font.Bold = true;
                        openbook.Sheets[i + 1].cells[1, n].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        openbook.Sheets[i + 1].Columns[n].ColumnWidth = 25;
                        openbook.Sheets[i + 1].cells[1, n].Value = NameColumnsRDUONLY[n, 0];
                    }
                    int NumberRow = DataForReport_RDU[i].GetLength(0);
                    if (NumberRow <= 1) { NumberRow = 2; }; //Если нет данных по РДУ.
                    int NumberCol = DataForReport_RDU[i].GetLength(1) - 1;
                    //Вставили данные:
                    if (DataForReport_RDU[i].Length != 0) //Проверка на наличие данных по РДУ: Если нет, пропускаем с оформлением шапки.
                    {
                        openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].cells[2, 1], openbook.Sheets[i + 1].cells[NumberRow + 1, NumberCol]].Value = DataForReport_RDU[i];
                        //Установили требуемые стили;
                        openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].cells[2, 1], openbook.Sheets[i + 1].cells[NumberRow + 1, NumberCol]].Borders.LineStyle = XlLineStyle.xlContinuous;
                        openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].cells[2, 1], openbook.Sheets[i + 1].cells[NumberRow + 1, NumberCol]].WrapText = null;
                    }
                    //for ( int j = 1; j <= NumberRow ; j++)
                    for (int j = 1; j < NumberRow + 2; j++)
                    {
                        if (j == 1) //Не должно срабатывать.
                        {
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Interior.Color = Color.CornflowerBlue;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Font.Bold = true;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Borders.LineStyle = XlLineStyle.xlContinuous;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].AutoFilter();
                        }
                        else if (j % 2 == 0)
                        {
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Interior.Color = Color.LightSkyBlue;
                        }
                        else
                        {
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Interior.Color = Color.MintCream;
                        }
                    }

                }
                openbook.SaveAs(PathToFileExportSD + "\\" + NameReport_RDU + ".xlsx");  //Сохранили книгу;
                openbooks.Close(); //Закрыли книгу;
                wStr("'" + NameReport_RDU + "' сформирован. \r\n", true);
            }
            #endregion

            #region Report_ODU_toExcel
            if (DataForReport_ODU != null && ReportODU != 0)
            {
                //Формируем отчет по ОДУ;
                wStr("Формируем отчет: '" + NameReport_ODU + "'. ", true, false);
                //Формируем отчет по ОДУ;
                //string NameReport_ODU;   //Наименование отчета по ОДУ;
                //string ReportODUSheets; //Наименования листов в книге отчета по ОДУ
                //private object[][,] DataForReport_ODU; //по ОДУ
                openbook = openbooks.Add(); //Добавляем книгу;
                openSheets = openbook.Sheets;
                openbook.Sheets.Add();
                openbook.Sheets[1].Name = ReportODU_Sheets; //Переименовываем листы.
                openbook.Sheets[2].Name = "Вне ОЗ Северо-Запада";
                for (int i = 0; i <= DataForReport_ODU.GetLength(0) - 1; i++)   //Положили данные на каждый i-лист
                {
                    for (int n = 1; n < NameColumnsONLY.GetLength(0); n++)  //Положили шапку
                    {
                        openbook.Sheets[i + 1].cells[1, n].WrapText = true;
                        openbook.Sheets[i + 1].cells[1, n].Font.Bold = true;
                        openbook.Sheets[i + 1].cells[1, n].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        openbook.Sheets[i + 1].Columns[n].ColumnWidth = 25;
                        openbook.Sheets[i + 1].cells[1, n].Value = NameColumnsONLY[n, 0];
                    }
                    int NumberRow = DataForReport_ODU[i].GetLength(0);
                    int NumberCol = DataForReport_ODU[i].GetLength(1) - 1;
                    //Вставили данные:
                    openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].cells[2, 1], openbook.Sheets[i + 1].cells[NumberRow, NumberCol]].Value = DataForReport_ODU[i];
                    //Установили требуемые стили;
                    openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].cells[2, 1], openbook.Sheets[i + 1].cells[NumberRow, NumberCol]].Borders.LineStyle = XlLineStyle.xlContinuous;
                    openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].cells[2, 1], openbook.Sheets[i + 1].cells[NumberRow, NumberCol]].WrapText = null;
                    for (int j = 1; j <= NumberRow; j++)
                    {
                        if (j == 1)
                        {
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Interior.Color = Color.CornflowerBlue;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Font.Bold = true;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Borders.LineStyle = XlLineStyle.xlContinuous;
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].AutoFilter();
                        }
                        else if (j % 2 == 0)
                        {
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Interior.Color = Color.LightSkyBlue;
                        }
                        else
                        {
                            openbook.Sheets[i + 1].Range[openbook.Sheets[i + 1].Cells[j, 1], openbook.Sheets[i + 1].Cells[j, NumberCol]].Interior.Color = Color.MintCream;
                        }
                    }
                }
                openbook.SaveAs(PathToFileExportSD + "\\" + NameReport_ODU + ".xlsx");  //Сохранили книгу;
                openbooks.Close(); //Закрыли книгу;
                //wStr("'" + NameReport_ODU + "' сформирован. \r\n", true);
                wStr(" => Сформирован. \r\n", true);
            }
            #endregion

            oExcelApp.DisplayAlerts = true;
            oExcelApp.Visible = true;
            oExcelApp.Quit();

            button_parsingReport.Enabled = true;
            button_OpenFolder.Enabled = true;
            button_SendMail.Enabled = true;
            button_SendMail.Focus();
        }

    }
}
