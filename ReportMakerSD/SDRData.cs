using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ReportMakerSD
{
    public class SDRData
    {
        // 15
        // DateEndReg; Number; Status; Org; Applicant; FSG; Executor; Service; Subject; Description; WaitingCode; WaitingReason; DateCalc; DateWait; Expired
        // Дата окончания регистрации; Номер; Статус; Организация заявителя; Заявитель; ФГП; Исполнитель; Услуга; Тема; Описание (для списков); Код ожидания; Причина ожидания; Расчетная дата решения обращения; Срок ожидания; Просрочен на(выполнение)

        public DateTime DateEndReg { get; set; } // Дата окончания регистрации
        public static string Name1 { get { return "Дата окончания регистрации"; } }
        public const string DateEndReg_Name = "Дата окончания регистрации";
        public static int DateEndReg_ColumnLoad { get; set; } = 0;  // зачем хранить всегда? может убрать после рабочей версии метода загрузки данных?

        public string Number { get; set; } // Номер
        public static string Name2 { get { return "Номер"; } }
        public const string Number_Name = "Номер";
        public static int Number_ColumnLoad { get; set; } = 0;

        public string Status { get; set; } // Статус
        public static string Name3 { get { return "Статус"; } }
        public const string Status_Name = "Статус";
        public static int Status_ColumnLoad { get; set; } = 0;

        public string Org { get; set; }     //Организация заявителя
        public static string Name4 { get { return "Организация заявителя"; } }
        public const string Org_Name = "Организация заявителя";
        public static int Org_ColumnLoad { get; set; } = 0;

        public string Applicant { get; set; } //Заявитель
        public static string Name5 { get { return "Заявитель"; } }
        public const string Applicant_Name = "Заявитель";
        public static int Applicant_ColumnLoad { get; set; } = 0;

        public string FSG { get; set; } //ФГП
        public static string Name6 { get { return "ФГП"; } }
        public const string FSG_Name = "ФГП";
        public static int FSG_ColumnLoad { get; set; } = 0;

        public string Executor { get; set; } //Исполнитель
        public static string Name7 { get { return "Исполнитель"; } }
        public const string Executor_Name = "Исполнитель";
        public static int Executor_ColumnLoad { get; set; } = 0;

        public string Service { get; set; } //Услуга
        public static string Name8 { get { return "Услуга"; } }
        public const string Service_Name = "Услуга";
        public static int Service_ColumnLoad { get; set; } = 0;

        public string Subject { get; set; } //Тема
        public static string Name9 { get { return "Тема"; } }
        public const string Subject_Name = "Тема";
        public static int Subject_ColumnLoad { get; set; } = 0;

        public string Description { get; set; } //Описание (для списков)
        public static string Name10 { get { return "Описание (для списков)"; } }
        public const string Description_Name = "Описание (для списков)";
        public static int Description_ColumnLoad { get; set; } = 0;

        public string WaitingCode { get; set; } //Код ожидания
        public static string Name11 { get { return "Код ожидания"; } }
        public const string WaitingCode_Name = "Код ожидания";
        public static int WaitingCode_ColumnLoad { get; set; } = 0;

        public string WaitingReason { get; set; } //Причина ожидания
        public static string Name12 { get { return "Причина ожидания"; } }
        public const string WaitingReason_Name = "Причина ожидания";
        public static int WaitingReason_ColumnLoad { get; set; } = 0;

        public DateTime DateCalc { get; set; } //Расчетная дата решения обращения
        public static string Name13 { get { return "Расчетная дата решения обращения"; } }
        public const string DateCalc_Name = "Расчетная дата решения обращения";
        public static int DateCalc_ColumnLoad { get; set; } = 0;

        public string DateWait { get; set; } //Срок ожидания
        public static string Name14 { get { return "Срок ожидания"; } }
        public const string DateWait_Name = "Срок ожидания";
        public static int DateWait_ColumnLoad { get; set; } = 0;

        public string Expired { get; set; } //Просрочен на(выполнение)
        public static string Name15 { get { return "Просрочен на (выполнение)"; } }
        public const string Expired_Name = "Просрочен на (выполнение)";
        public static int Expired_ColumnLoad { get; set; } = 0;

        public string Rating { get; set; } //Просрочен на(выполнение)
        public static string Name16 { get { return "Оценка пользователя"; } }
        public const string Rating_Name = "Оценка пользователя";
        public static int Rating_ColumnLoad { get; set; } = 0;

        public string DateAnswer { get; set; }
        public static string Name17 { get { return "Дата и время решения"; } }
        public const string DateAnswer_Name = "Дата и время решения";
        public static int DateAnswer_ColumnLoad { get; set; } = 0;

        public string AnswerText { get; set; }
        public static string Name18 { get { return "Решение (для списков)"; } }
        public const string AnswerText_Name = "Решение (для списков)";
        public static int AnswerText_ColumnLoad { get; set; } = 0;

        public string OAreaApplicant { get; set; }
        public static string Name19 { get { return "Оперзона заявителя"; } }
        public const string OAreaApplicant_Name = "Оперзона заявителя";
        public static int OAreaApplicant_ColumnLoad { get; set; } = 0;

        public string ServiceGroup { get; set; }
        public static string Name20 { get { return "Группа услуг"; } }
        public const string ServiceGroup_Name = "Группа услуг";
        public static int ServiceGroup_ColumnLoad { get; set; } = 0;

        public string OAreaFSG { get; set; }
        public static string Name43 { get { return "Оперзона ФГП"; } }
        public const string OAreaFSG_Name = "Оперзона ФГП";
        public static int OAreaFSG_ColumnLoad { get; set; } = 0;

        // Тестовый метод
        public void GetInfo()
        {
            Console.WriteLine($"Вывод данных по обращению");
        }


        // Метод загрузки данных в список из файла
        public static List<SDRData> SDRLoadData(string PathToFileExportSD, string FilePath)
        {
            Utils.OutputDir = new DirectoryInfo(PathToFileExportSD);
            FileInfo existingFile = new FileInfo(FilePath);
            List<SDRData> SDRData0 = new List<SDRData>();

            using (ExcelPackage packageData = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheetData = packageData.Workbook.Worksheets[1];
                foreach (var firstRowCell in worksheetData.Cells[1, 1, 1, worksheetData.Dimension.End.Column])
                {
                    switch (firstRowCell.Text)
                    {
                        case DateEndReg_Name:
                            SDRData.DateEndReg_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Number_Name:
                            SDRData.Number_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Status_Name:
                            SDRData.Status_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Org_Name:
                            SDRData.Org_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Applicant_Name:
                            SDRData.Applicant_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case FSG_Name:
                            SDRData.FSG_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Executor_Name:
                            SDRData.Executor_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Service_Name:
                            SDRData.Service_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Subject_Name:
                            SDRData.Subject_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Description_Name:
                            SDRData.Description_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case WaitingCode_Name:
                            SDRData.WaitingCode_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case WaitingReason_Name:
                            SDRData.WaitingReason_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case DateCalc_Name:
                            SDRData.DateCalc_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case DateWait_Name:
                            SDRData.DateWait_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Expired_Name:
                            SDRData.Expired_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case Rating_Name:
                            SDRData.Rating_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case DateAnswer_Name:
                            SDRData.DateAnswer_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case AnswerText_Name:
                            SDRData.AnswerText_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case OAreaApplicant_Name:
                            SDRData.OAreaApplicant_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case ServiceGroup_Name:
                            SDRData.ServiceGroup_ColumnLoad = firstRowCell.Start.Column;
                            break;
                        case OAreaFSG_Name:
                            SDRData.OAreaFSG_ColumnLoad = firstRowCell.Start.Column;
                            break;

                        default:
                            // Учет пропущенных столбцов??
                            break;
                    }
                }
                // Перебираем строки исходного файла и заносим данные в список

                foreach (var firstRowCell in worksheetData.Cells[2, 1, worksheetData.Dimension.End.Row, 1])
                {
                    DateTime tDateEndReg = DateTime.Now;
                    if (DateEndReg_ColumnLoad != 0) tDateEndReg = DateTime.FromOADate(double.Parse(worksheetData.Cells[firstRowCell.Start.Row, DateEndReg_ColumnLoad].Value.ToString()));
                    string tNumber = "";
                    if (Number_ColumnLoad != 0) tNumber = worksheetData.Cells[firstRowCell.Start.Row, Number_ColumnLoad].Value.ToString();
                    string tStatus = "";
                    if (Status_ColumnLoad != 0) tStatus = worksheetData.Cells[firstRowCell.Start.Row, Status_ColumnLoad].Value.ToString();
                    string tOrg = "";
                    if (Org_ColumnLoad != 0) tOrg = worksheetData.Cells[firstRowCell.Start.Row, Org_ColumnLoad].Value.ToString();
                    string tApplicant = "";
                    if (Applicant_ColumnLoad != 0) tApplicant = worksheetData.Cells[firstRowCell.Start.Row, Applicant_ColumnLoad].Value.ToString();
                    string tFSG = "";
                    if (FSG_ColumnLoad != 0) tFSG = worksheetData.Cells[firstRowCell.Start.Row, FSG_ColumnLoad].Value.ToString();
                    string tExecutor = "";
                    if (Executor_ColumnLoad != 0) tExecutor = worksheetData.Cells[firstRowCell.Start.Row, Executor_ColumnLoad].Value.ToString();
                    string tService = "";
                    if (Service_ColumnLoad != 0) tService = worksheetData.Cells[firstRowCell.Start.Row, Service_ColumnLoad].Value.ToString();
                    string tSubject = "";
                    if (Subject_ColumnLoad != 0) tSubject = worksheetData.Cells[firstRowCell.Start.Row, Subject_ColumnLoad].Value.ToString();
                    string tDescription = "";
                    if (Description_ColumnLoad != 0) tDescription = worksheetData.Cells[firstRowCell.Start.Row, Description_ColumnLoad].Text;
                    string tWaitingCode = "";
                    if (WaitingCode_ColumnLoad != 0) tWaitingCode = worksheetData.Cells[firstRowCell.Start.Row, WaitingCode_ColumnLoad].Text;
                    string tWaitingReason = "";
                    if (WaitingReason_ColumnLoad != 0) tWaitingReason = worksheetData.Cells[firstRowCell.Start.Row, WaitingReason_ColumnLoad].Text;
                    DateTime tDateCalc = DateTime.Now;
                    if (DateCalc_ColumnLoad != 0) tDateCalc = DateTime.FromOADate(double.Parse(worksheetData.Cells[firstRowCell.Start.Row, DateCalc_ColumnLoad].Value.ToString()));
                    string tDateWait = "";
                    if (DateWait_ColumnLoad != 0) tDateWait = worksheetData.Cells[firstRowCell.Start.Row, DateWait_ColumnLoad].Text;
                    string tExpired = "";
                    if (Expired_ColumnLoad != 0) tExpired = worksheetData.Cells[firstRowCell.Start.Row, Expired_ColumnLoad].Text;
                    string tRating = "";
                    if (Expired_ColumnLoad != 0) tRating = worksheetData.Cells[firstRowCell.Start.Row, Rating_ColumnLoad].Text;
                    string tDateAnswer = "";
                    if (DateAnswer_ColumnLoad != 0) tDateAnswer = worksheetData.Cells[firstRowCell.Start.Row, DateAnswer_ColumnLoad].Text;
                    string tAnswerText = "";
                    if (AnswerText_ColumnLoad != 0) tAnswerText = worksheetData.Cells[firstRowCell.Start.Row, AnswerText_ColumnLoad].Text;
                    string tOArea = "";
                    if (OAreaApplicant_ColumnLoad != 0) tOArea = worksheetData.Cells[firstRowCell.Start.Row, OAreaApplicant_ColumnLoad].Text;
                    string tServiceGroup = "";
                    if (ServiceGroup_ColumnLoad != 0) tServiceGroup = worksheetData.Cells[firstRowCell.Start.Row, ServiceGroup_ColumnLoad].Text;
                    string tOAreaFSG = "";
                    if (OAreaFSG_ColumnLoad != 0) tOAreaFSG = worksheetData.Cells[firstRowCell.Start.Row, OAreaFSG_ColumnLoad].Text;


                    SDRData0.Add(new SDRData()
                    {
                        DateEndReg = tDateEndReg,
                        Number = tNumber,
                        Status = tStatus,
                        Org = tOrg,
                        Applicant = tApplicant,
                        FSG = tFSG,
                        Executor = tExecutor,
                        Service = tService,
                        Subject = tSubject,
                        Description = tDescription,
                        WaitingCode = tWaitingCode,
                        WaitingReason = tWaitingReason,
                        DateCalc = tDateCalc,
                        DateWait = tDateWait,
                        Expired = tExpired,
                        Rating = tRating,
                        DateAnswer = tDateAnswer,
                        AnswerText = tAnswerText,
                        OAreaApplicant = tOArea,
                        ServiceGroup = tServiceGroup,
                        OAreaFSG = tOAreaFSG
                    });
                }
            }
            return SDRData0;
        }
    }
}
