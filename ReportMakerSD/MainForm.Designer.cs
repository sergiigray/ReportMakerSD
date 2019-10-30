using System;
using System.ComponentModel;

namespace ReportMakerSD
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.openFileDialog_ExportSDSelect = new System.Windows.Forms.OpenFileDialog();
            this.dateTimePicker_TimeFrom = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker_TimeTo = new System.Windows.Forms.DateTimePicker();
            this.label_TimeFrom = new System.Windows.Forms.Label();
            this.label_TimeTo = new System.Windows.Forms.Label();
            this.textBox_FileExportSDName = new System.Windows.Forms.TextBox();
            this.button_ExportSDSelect = new System.Windows.Forms.Button();
            this.label_FileExportSDName = new System.Windows.Forms.Label();
            this.textBox_DebugInfo = new System.Windows.Forms.TextBox();
            this.button_parsingReport = new System.Windows.Forms.Button();
            this.button_SendMail = new System.Windows.Forms.Button();
            this.checkBox_ReportTO = new System.Windows.Forms.CheckBox();
            this.button_OpenFolder = new System.Windows.Forms.Button();
            this.panel_Button = new System.Windows.Forms.Panel();
            this.panel_textBox = new System.Windows.Forms.Panel();
            this.panel_DateTime = new System.Windows.Forms.Panel();
            this.comboBox_SelectPeriod = new System.Windows.Forms.ComboBox();
            this.panel_Report = new System.Windows.Forms.Panel();
            this.checkBox_Report4 = new System.Windows.Forms.CheckBox();
            this.checkBox_Report5 = new System.Windows.Forms.CheckBox();
            this.checkBox_Report3 = new System.Windows.Forms.CheckBox();
            this.checkBox_Report2 = new System.Windows.Forms.CheckBox();
            this.checkBox_ReportOne = new System.Windows.Forms.CheckBox();
            this.checkBox_ReportBase = new System.Windows.Forms.CheckBox();
            this.panel_Files = new System.Windows.Forms.Panel();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.panel_ProgressBar = new System.Windows.Forms.Panel();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.checkBox_DebugInfo = new System.Windows.Forms.CheckBox();
            this.panel_Button.SuspendLayout();
            this.panel_textBox.SuspendLayout();
            this.panel_DateTime.SuspendLayout();
            this.panel_Report.SuspendLayout();
            this.panel_Files.SuspendLayout();
            this.panel_ProgressBar.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog_ExportSDSelect
            // 
            this.openFileDialog_ExportSDSelect.FileName = "openFileDialog1";
            // 
            // dateTimePicker_TimeFrom
            // 
            this.dateTimePicker_TimeFrom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.dateTimePicker_TimeFrom.CustomFormat = "dd.MM.yyyy HH:mm";
            this.dateTimePicker_TimeFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker_TimeFrom.Location = new System.Drawing.Point(3, 19);
            this.dateTimePicker_TimeFrom.Name = "dateTimePicker_TimeFrom";
            this.dateTimePicker_TimeFrom.Size = new System.Drawing.Size(125, 20);
            this.dateTimePicker_TimeFrom.TabIndex = 2;
            this.dateTimePicker_TimeFrom.Value = new System.DateTime(2016, 7, 27, 8, 30, 0, 0);
            // 
            // dateTimePicker_TimeTo
            // 
            this.dateTimePicker_TimeTo.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.dateTimePicker_TimeTo.CustomFormat = "dd.MM.yyyy HH:mm";
            this.dateTimePicker_TimeTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker_TimeTo.Location = new System.Drawing.Point(146, 19);
            this.dateTimePicker_TimeTo.Name = "dateTimePicker_TimeTo";
            this.dateTimePicker_TimeTo.Size = new System.Drawing.Size(125, 20);
            this.dateTimePicker_TimeTo.TabIndex = 3;
            this.dateTimePicker_TimeTo.Value = new System.DateTime(2016, 7, 27, 8, 30, 0, 0);
            // 
            // label_TimeFrom
            // 
            this.label_TimeFrom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label_TimeFrom.AutoSize = true;
            this.label_TimeFrom.Location = new System.Drawing.Point(23, 3);
            this.label_TimeFrom.Name = "label_TimeFrom";
            this.label_TimeFrom.Size = new System.Drawing.Size(78, 13);
            this.label_TimeFrom.TabIndex = 1;
            this.label_TimeFrom.Text = "Время начала";
            // 
            // label_TimeTo
            // 
            this.label_TimeTo.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.label_TimeTo.AutoSize = true;
            this.label_TimeTo.Location = new System.Drawing.Point(155, 3);
            this.label_TimeTo.Name = "label_TimeTo";
            this.label_TimeTo.Size = new System.Drawing.Size(96, 13);
            this.label_TimeTo.TabIndex = 1;
            this.label_TimeTo.Text = "Время окончания";
            // 
            // textBox_FileExportSDName
            // 
            this.textBox_FileExportSDName.Location = new System.Drawing.Point(6, 19);
            this.textBox_FileExportSDName.Name = "textBox_FileExportSDName";
            this.textBox_FileExportSDName.Size = new System.Drawing.Size(582, 20);
            this.textBox_FileExportSDName.TabIndex = 0;
            this.textBox_FileExportSDName.TabStop = false;
            // 
            // button_ExportSDSelect
            // 
            this.button_ExportSDSelect.Location = new System.Drawing.Point(594, 19);
            this.button_ExportSDSelect.Name = "button_ExportSDSelect";
            this.button_ExportSDSelect.Size = new System.Drawing.Size(88, 20);
            this.button_ExportSDSelect.TabIndex = 0;
            this.button_ExportSDSelect.TabStop = false;
            this.button_ExportSDSelect.Text = "Выбрать файл";
            this.button_ExportSDSelect.UseVisualStyleBackColor = true;
            this.button_ExportSDSelect.Click += new System.EventHandler(this.button_ExportSDSelect_Click);
            // 
            // label_FileExportSDName
            // 
            this.label_FileExportSDName.AutoSize = true;
            this.label_FileExportSDName.Location = new System.Drawing.Point(3, 4);
            this.label_FileExportSDName.Name = "label_FileExportSDName";
            this.label_FileExportSDName.Size = new System.Drawing.Size(154, 13);
            this.label_FileExportSDName.TabIndex = 1;
            this.label_FileExportSDName.Text = "Файл с исходными данными";
            // 
            // textBox_DebugInfo
            // 
            this.textBox_DebugInfo.AcceptsReturn = true;
            this.textBox_DebugInfo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox_DebugInfo.Location = new System.Drawing.Point(0, 0);
            this.textBox_DebugInfo.Multiline = true;
            this.textBox_DebugInfo.Name = "textBox_DebugInfo";
            this.textBox_DebugInfo.ReadOnly = true;
            this.textBox_DebugInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox_DebugInfo.Size = new System.Drawing.Size(682, 215);
            this.textBox_DebugInfo.TabIndex = 0;
            this.textBox_DebugInfo.TabStop = false;
            this.textBox_DebugInfo.WordWrap = false;
            // 
            // button_parsingReport
            // 
            this.button_parsingReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_parsingReport.Location = new System.Drawing.Point(0, 3);
            this.button_parsingReport.Name = "button_parsingReport";
            this.button_parsingReport.Size = new System.Drawing.Size(135, 23);
            this.button_parsingReport.TabIndex = 0;
            this.button_parsingReport.Text = "Обработка данных";
            this.button_parsingReport.UseVisualStyleBackColor = true;
            this.button_parsingReport.Click += new System.EventHandler(this.button_parsingReport_Click);
            // 
            // button_SendMail
            // 
            this.button_SendMail.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button_SendMail.Enabled = false;
            this.button_SendMail.Location = new System.Drawing.Point(547, 3);
            this.button_SendMail.Name = "button_SendMail";
            this.button_SendMail.Size = new System.Drawing.Size(135, 23);
            this.button_SendMail.TabIndex = 0;
            this.button_SendMail.Text = "Отправка писем";
            this.button_SendMail.UseVisualStyleBackColor = true;
            this.button_SendMail.Click += new System.EventHandler(this.button_SendMail_Click);
            // 
            // checkBox_ReportTO
            // 
            this.checkBox_ReportTO.AutoSize = true;
            this.checkBox_ReportTO.Enabled = false;
            this.checkBox_ReportTO.ForeColor = System.Drawing.SystemColors.ActiveBorder;
            this.checkBox_ReportTO.Location = new System.Drawing.Point(134, 3);
            this.checkBox_ReportTO.Name = "checkBox_ReportTO";
            this.checkBox_ReportTO.Size = new System.Drawing.Size(207, 17);
            this.checkBox_ReportTO.TabIndex = 0;
            this.checkBox_ReportTO.TabStop = false;
            this.checkBox_ReportTO.Text = "Отчет по подрядным организациям";
            this.checkBox_ReportTO.UseVisualStyleBackColor = true;
            this.checkBox_ReportTO.Visible = false;
            // 
            // button_OpenFolder
            // 
            this.button_OpenFolder.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button_OpenFolder.Enabled = false;
            this.button_OpenFolder.Location = new System.Drawing.Point(258, 3);
            this.button_OpenFolder.Name = "button_OpenFolder";
            this.button_OpenFolder.Size = new System.Drawing.Size(162, 23);
            this.button_OpenFolder.TabIndex = 0;
            this.button_OpenFolder.TabStop = false;
            this.button_OpenFolder.Text = "Открыть папку с отчетами";
            this.button_OpenFolder.UseVisualStyleBackColor = true;
            this.button_OpenFolder.Click += new System.EventHandler(this.button_OpenFolder_Click);
            // 
            // panel_Button
            // 
            this.panel_Button.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.panel_Button.Controls.Add(this.button_parsingReport);
            this.panel_Button.Controls.Add(this.button_OpenFolder);
            this.panel_Button.Controls.Add(this.button_SendMail);
            this.panel_Button.Location = new System.Drawing.Point(0, 465);
            this.panel_Button.Name = "panel_Button";
            this.panel_Button.Size = new System.Drawing.Size(682, 29);
            this.panel_Button.TabIndex = 0;
            // 
            // panel_textBox
            // 
            this.panel_textBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel_textBox.Controls.Add(this.textBox_DebugInfo);
            this.panel_textBox.Location = new System.Drawing.Point(0, 215);
            this.panel_textBox.Name = "panel_textBox";
            this.panel_textBox.Size = new System.Drawing.Size(682, 215);
            this.panel_textBox.TabIndex = 8;
            // 
            // panel_DateTime
            // 
            this.panel_DateTime.Controls.Add(this.comboBox_SelectPeriod);
            this.panel_DateTime.Controls.Add(this.dateTimePicker_TimeFrom);
            this.panel_DateTime.Controls.Add(this.label_TimeFrom);
            this.panel_DateTime.Controls.Add(this.dateTimePicker_TimeTo);
            this.panel_DateTime.Controls.Add(this.label_TimeTo);
            this.panel_DateTime.Location = new System.Drawing.Point(0, 149);
            this.panel_DateTime.Name = "panel_DateTime";
            this.panel_DateTime.Size = new System.Drawing.Size(473, 49);
            this.panel_DateTime.TabIndex = 0;
            // 
            // comboBox_SelectPeriod
            // 
            this.comboBox_SelectPeriod.FormattingEnabled = true;
            this.comboBox_SelectPeriod.Items.AddRange(new object[] {
            "За текущую неделю",
            "За 7 дней",
            "За прошлую неделю",
            "За прошлый месяц",
            "За текущий квартал",
            "За текущие полгода",
            "За текущий год",
            "За неделю"});
            this.comboBox_SelectPeriod.Location = new System.Drawing.Point(288, 19);
            this.comboBox_SelectPeriod.Name = "comboBox_SelectPeriod";
            this.comboBox_SelectPeriod.Size = new System.Drawing.Size(185, 21);
            this.comboBox_SelectPeriod.TabIndex = 4;
            this.comboBox_SelectPeriod.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectPeriod_SelectedIndexChanged);
            // 
            // panel_Report
            // 
            this.panel_Report.Controls.Add(this.checkBox_Report4);
            this.panel_Report.Controls.Add(this.checkBox_Report5);
            this.panel_Report.Controls.Add(this.checkBox_Report3);
            this.panel_Report.Controls.Add(this.checkBox_Report2);
            this.panel_Report.Controls.Add(this.checkBox_ReportOne);
            this.panel_Report.Controls.Add(this.checkBox_ReportBase);
            this.panel_Report.Controls.Add(this.checkBox_ReportTO);
            this.panel_Report.Location = new System.Drawing.Point(0, 49);
            this.panel_Report.Name = "panel_Report";
            this.panel_Report.Size = new System.Drawing.Size(682, 94);
            this.panel_Report.TabIndex = 0;
            // 
            // checkBox_Report4
            // 
            this.checkBox_Report4.AutoSize = true;
            this.checkBox_Report4.Checked = true;
            this.checkBox_Report4.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_Report4.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.checkBox_Report4.Location = new System.Drawing.Point(3, 3);
            this.checkBox_Report4.Name = "checkBox_Report4";
            this.checkBox_Report4.Size = new System.Drawing.Size(207, 17);
            this.checkBox_Report4.TabIndex = 0;
            this.checkBox_Report4.TabStop = false;
            this.checkBox_Report4.Text = "Отчет по подрядным организациям";
            this.checkBox_Report4.UseVisualStyleBackColor = true;
            // 
            // checkBox_Report5
            // 
            this.checkBox_Report5.AutoSize = true;
            this.checkBox_Report5.Checked = true;
            this.checkBox_Report5.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_Report5.ForeColor = System.Drawing.Color.ForestGreen;
            this.checkBox_Report5.Location = new System.Drawing.Point(3, 26);
            this.checkBox_Report5.Name = "checkBox_Report5";
            this.checkBox_Report5.Size = new System.Drawing.Size(125, 17);
            this.checkBox_Report5.TabIndex = 0;
            this.checkBox_Report5.TabStop = false;
            this.checkBox_Report5.Text = "Отчет по филиалам";
            this.checkBox_Report5.UseVisualStyleBackColor = true;
            // 
            // checkBox_Report3
            // 
            this.checkBox_Report3.AutoSize = true;
            this.checkBox_Report3.ForeColor = System.Drawing.Color.Violet;
            this.checkBox_Report3.Location = new System.Drawing.Point(364, 49);
            this.checkBox_Report3.Name = "checkBox_Report3";
            this.checkBox_Report3.Size = new System.Drawing.Size(76, 17);
            this.checkBox_Report3.TabIndex = 0;
            this.checkBox_Report3.TabStop = false;
            this.checkBox_Report3.Text = "Месячник";
            this.checkBox_Report3.UseVisualStyleBackColor = true;
            // 
            // checkBox_Report2
            // 
            this.checkBox_Report2.AutoSize = true;
            this.checkBox_Report2.ForeColor = System.Drawing.Color.Maroon;
            this.checkBox_Report2.Location = new System.Drawing.Point(364, 3);
            this.checkBox_Report2.Name = "checkBox_Report2";
            this.checkBox_Report2.Size = new System.Drawing.Size(244, 17);
            this.checkBox_Report2.TabIndex = 0;
            this.checkBox_Report2.TabStop = false;
            this.checkBox_Report2.Text = "Нерешенные по подрядным организациям";
            this.checkBox_Report2.UseVisualStyleBackColor = true;
            // 
            // checkBox_ReportOne
            // 
            this.checkBox_ReportOne.AutoSize = true;
            this.checkBox_ReportOne.ForeColor = System.Drawing.Color.Crimson;
            this.checkBox_ReportOne.Location = new System.Drawing.Point(364, 26);
            this.checkBox_ReportOne.Name = "checkBox_ReportOne";
            this.checkBox_ReportOne.Size = new System.Drawing.Size(240, 17);
            this.checkBox_ReportOne.TabIndex = 0;
            this.checkBox_ReportOne.TabStop = false;
            this.checkBox_ReportOne.Text = "Нерешенные по дате решения обращения";
            this.checkBox_ReportOne.UseVisualStyleBackColor = true;
            // 
            // checkBox_ReportBase
            // 
            this.checkBox_ReportBase.AutoSize = true;
            this.checkBox_ReportBase.Location = new System.Drawing.Point(3, 72);
            this.checkBox_ReportBase.Name = "checkBox_ReportBase";
            this.checkBox_ReportBase.Size = new System.Drawing.Size(250, 17);
            this.checkBox_ReportBase.TabIndex = 0;
            this.checkBox_ReportBase.TabStop = false;
            this.checkBox_ReportBase.Text = "Обращения зарегистрированные за период";
            this.checkBox_ReportBase.UseVisualStyleBackColor = true;
            // 
            // panel_Files
            // 
            this.panel_Files.Controls.Add(this.button_ExportSDSelect);
            this.panel_Files.Controls.Add(this.label_FileExportSDName);
            this.panel_Files.Controls.Add(this.textBox_FileExportSDName);
            this.panel_Files.Location = new System.Drawing.Point(0, 1);
            this.panel_Files.Name = "panel_Files";
            this.panel_Files.Size = new System.Drawing.Size(682, 42);
            this.panel_Files.TabIndex = 0;
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.WorkerReportsProgress = true;
            this.backgroundWorker.WorkerSupportsCancellation = true;
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // panel_ProgressBar
            // 
            this.panel_ProgressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel_ProgressBar.Controls.Add(this.progressBar);
            this.panel_ProgressBar.Location = new System.Drawing.Point(0, 428);
            this.panel_ProgressBar.Name = "panel_ProgressBar";
            this.panel_ProgressBar.Size = new System.Drawing.Size(682, 34);
            this.panel_ProgressBar.TabIndex = 9;
            // 
            // progressBar
            // 
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBar.Location = new System.Drawing.Point(0, 0);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(682, 34);
            this.progressBar.TabIndex = 0;
            // 
            // checkBox_DebugInfo
            // 
            this.checkBox_DebugInfo.AutoSize = true;
            this.checkBox_DebugInfo.Location = new System.Drawing.Point(3, 194);
            this.checkBox_DebugInfo.Name = "checkBox_DebugInfo";
            this.checkBox_DebugInfo.Size = new System.Drawing.Size(212, 17);
            this.checkBox_DebugInfo.TabIndex = 10;
            this.checkBox_DebugInfo.TabStop = false;
            this.checkBox_DebugInfo.Text = "Вывод дополнительной информации";
            this.checkBox_DebugInfo.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(682, 498);
            this.Controls.Add(this.checkBox_DebugInfo);
            this.Controls.Add(this.panel_ProgressBar);
            this.Controls.Add(this.panel_Files);
            this.Controls.Add(this.panel_Report);
            this.Controls.Add(this.panel_DateTime);
            this.Controls.Add(this.panel_textBox);
            this.Controls.Add(this.panel_Button);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "ReportMakerSD";
            this.panel_Button.ResumeLayout(false);
            this.panel_textBox.ResumeLayout(false);
            this.panel_textBox.PerformLayout();
            this.panel_DateTime.ResumeLayout(false);
            this.panel_DateTime.PerformLayout();
            this.panel_Report.ResumeLayout(false);
            this.panel_Report.PerformLayout();
            this.panel_Files.ResumeLayout(false);
            this.panel_Files.PerformLayout();
            this.panel_ProgressBar.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            throw new NotImplementedException();
        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog_ExportSDSelect;
        private System.Windows.Forms.DateTimePicker dateTimePicker_TimeFrom;
        private System.Windows.Forms.DateTimePicker dateTimePicker_TimeTo;
        private System.Windows.Forms.Label label_TimeFrom;
        private System.Windows.Forms.Label label_TimeTo;
        private System.Windows.Forms.TextBox textBox_FileExportSDName;
        private System.Windows.Forms.Button button_ExportSDSelect;
        private System.Windows.Forms.Label label_FileExportSDName;
        private System.Windows.Forms.TextBox textBox_DebugInfo;
        private System.Windows.Forms.Button button_parsingReport;
        private System.Windows.Forms.Button button_SendMail;
        public System.Windows.Forms.CheckBox checkBox_ReportTO;
        private System.Windows.Forms.Button button_OpenFolder;
        private System.Windows.Forms.Panel panel_Button;
        private System.Windows.Forms.Panel panel_textBox;
        private System.Windows.Forms.Panel panel_DateTime;
        private System.Windows.Forms.Panel panel_Report;
        private System.Windows.Forms.Panel panel_Files;
        public System.Windows.Forms.CheckBox checkBox_ReportBase;
        private System.Windows.Forms.Panel panel_ProgressBar;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.ComboBox comboBox_SelectPeriod;
        public System.ComponentModel.BackgroundWorker backgroundWorker;
        public System.Windows.Forms.CheckBox checkBox_ReportOne;
        public System.Windows.Forms.CheckBox checkBox_DebugInfo;
        public System.Windows.Forms.CheckBox checkBox_Report2;
        public System.Windows.Forms.CheckBox checkBox_Report3;
        public System.Windows.Forms.CheckBox checkBox_Report4;
        public System.Windows.Forms.CheckBox checkBox_Report5;
    }
}

