using System;
using System.ComponentModel;

namespace ReportMakerSD
{
    public partial class Form0
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form0));
            this.openFileDialog_ExportSDSelect = new System.Windows.Forms.OpenFileDialog();
            this.dateTimePicker_TimeFrom = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker_TimeTo = new System.Windows.Forms.DateTimePicker();
            this.label_TimeFrom = new System.Windows.Forms.Label();
            this.label_TimeTo = new System.Windows.Forms.Label();
            this.textBox_FileExportSDName = new System.Windows.Forms.TextBox();
            this.button_ExportSDSelect = new System.Windows.Forms.Button();
            this.label_FileExportSDName = new System.Windows.Forms.Label();
            this.button_parsingReport = new System.Windows.Forms.Button();
            this.button_SendMail = new System.Windows.Forms.Button();
            this.button_OpenFolder = new System.Windows.Forms.Button();
            this.panel_Button = new System.Windows.Forms.Panel();
            this.panel_DateTime = new System.Windows.Forms.Panel();
            this.comboBox_SelectPeriod = new System.Windows.Forms.ComboBox();
            this.panel_Report = new System.Windows.Forms.Panel();
            this.checkBox_Report3 = new System.Windows.Forms.CheckBox();
            this.checkBox_Report2 = new System.Windows.Forms.CheckBox();
            this.checkBox_Report1 = new System.Windows.Forms.CheckBox();
            this.panel_Files = new System.Windows.Forms.Panel();
            this.panel_WeekMailTo = new System.Windows.Forms.Panel();
            this.textBox_WeekMailTo = new System.Windows.Forms.TextBox();
            this.label_WeekMailTo = new System.Windows.Forms.Label();
            this.panel_Button.SuspendLayout();
            this.panel_DateTime.SuspendLayout();
            this.panel_Report.SuspendLayout();
            this.panel_Files.SuspendLayout();
            this.panel_WeekMailTo.SuspendLayout();
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
            this.dateTimePicker_TimeFrom.Value = new System.DateTime(2020, 3, 10, 0, 0, 0, 0);
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
            this.dateTimePicker_TimeTo.Value = new System.DateTime(2020, 3, 10, 23, 59, 0, 0);
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
            this.label_FileExportSDName.Size = new System.Drawing.Size(157, 13);
            this.label_FileExportSDName.TabIndex = 1;
            this.label_FileExportSDName.Text = "Файл с исходными данными:";
            // 
            // button_parsingReport
            // 
            this.button_parsingReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_parsingReport.Location = new System.Drawing.Point(0, 3);
            this.button_parsingReport.Name = "button_parsingReport";
            this.button_parsingReport.Size = new System.Drawing.Size(170, 23);
            this.button_parsingReport.TabIndex = 0;
            this.button_parsingReport.Text = "Обработка данных";
            this.button_parsingReport.UseVisualStyleBackColor = true;
            this.button_parsingReport.Click += new System.EventHandler(this.button_parsingReport_Click);
            // 
            // button_SendMail
            // 
            this.button_SendMail.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button_SendMail.Enabled = false;
            this.button_SendMail.Location = new System.Drawing.Point(521, 3);
            this.button_SendMail.Name = "button_SendMail";
            this.button_SendMail.Size = new System.Drawing.Size(170, 23);
            this.button_SendMail.TabIndex = 0;
            this.button_SendMail.Text = "Отправка писем";
            this.button_SendMail.UseVisualStyleBackColor = true;
            this.button_SendMail.Click += new System.EventHandler(this.button_SendMail_Click);
            // 
            // button_OpenFolder
            // 
            this.button_OpenFolder.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button_OpenFolder.Enabled = false;
            this.button_OpenFolder.Location = new System.Drawing.Point(262, 3);
            this.button_OpenFolder.Name = "button_OpenFolder";
            this.button_OpenFolder.Size = new System.Drawing.Size(170, 23);
            this.button_OpenFolder.TabIndex = 0;
            this.button_OpenFolder.TabStop = false;
            this.button_OpenFolder.Text = "Открыть папку с отчетами";
            this.button_OpenFolder.UseVisualStyleBackColor = true;
            this.button_OpenFolder.Click += new System.EventHandler(this.button_OpenFolder_Click);
            // 
            // panel_Button
            // 
            this.panel_Button.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel_Button.Controls.Add(this.button_parsingReport);
            this.panel_Button.Controls.Add(this.button_OpenFolder);
            this.panel_Button.Controls.Add(this.button_SendMail);
            this.panel_Button.Location = new System.Drawing.Point(0, 314);
            this.panel_Button.Name = "panel_Button";
            this.panel_Button.Size = new System.Drawing.Size(691, 29);
            this.panel_Button.TabIndex = 0;
            // 
            // panel_DateTime
            // 
            this.panel_DateTime.Controls.Add(this.comboBox_SelectPeriod);
            this.panel_DateTime.Controls.Add(this.dateTimePicker_TimeFrom);
            this.panel_DateTime.Controls.Add(this.label_TimeFrom);
            this.panel_DateTime.Controls.Add(this.dateTimePicker_TimeTo);
            this.panel_DateTime.Controls.Add(this.label_TimeTo);
            this.panel_DateTime.Location = new System.Drawing.Point(0, 98);
            this.panel_DateTime.Name = "panel_DateTime";
            this.panel_DateTime.Size = new System.Drawing.Size(473, 49);
            this.panel_DateTime.TabIndex = 0;
            // 
            // comboBox_SelectPeriod
            // 
            this.comboBox_SelectPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_SelectPeriod.FormattingEnabled = true;
            this.comboBox_SelectPeriod.Items.AddRange(new object[] {
            "За прошлую неделю",
            "За текущую неделю",
            "За 7 дней",
            "За прошлый месяц"});
            this.comboBox_SelectPeriod.Location = new System.Drawing.Point(288, 19);
            this.comboBox_SelectPeriod.Name = "comboBox_SelectPeriod";
            this.comboBox_SelectPeriod.Size = new System.Drawing.Size(185, 21);
            this.comboBox_SelectPeriod.TabIndex = 4;
            this.comboBox_SelectPeriod.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectPeriod_SelectedIndexChanged);
            // 
            // panel_Report
            // 
            this.panel_Report.Controls.Add(this.checkBox_Report3);
            this.panel_Report.Controls.Add(this.checkBox_Report2);
            this.panel_Report.Controls.Add(this.checkBox_Report1);
            this.panel_Report.Location = new System.Drawing.Point(0, 49);
            this.panel_Report.Name = "panel_Report";
            this.panel_Report.Size = new System.Drawing.Size(691, 44);
            this.panel_Report.TabIndex = 0;
            // 
            // checkBox_Report3
            // 
            this.checkBox_Report3.AutoSize = true;
            this.checkBox_Report3.ForeColor = System.Drawing.Color.Violet;
            this.checkBox_Report3.Location = new System.Drawing.Point(3, 24);
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
            this.checkBox_Report2.Checked = true;
            this.checkBox_Report2.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_Report2.ForeColor = System.Drawing.Color.Maroon;
            this.checkBox_Report2.Location = new System.Drawing.Point(3, 3);
            this.checkBox_Report2.Name = "checkBox_Report2";
            this.checkBox_Report2.Size = new System.Drawing.Size(244, 17);
            this.checkBox_Report2.TabIndex = 0;
            this.checkBox_Report2.TabStop = false;
            this.checkBox_Report2.Text = "Нерешенные по подрядным организациям";
            this.checkBox_Report2.UseVisualStyleBackColor = true;
            // 
            // checkBox_Report1
            // 
            this.checkBox_Report1.AutoSize = true;
            this.checkBox_Report1.Checked = true;
            this.checkBox_Report1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_Report1.ForeColor = System.Drawing.Color.Crimson;
            this.checkBox_Report1.Location = new System.Drawing.Point(288, 3);
            this.checkBox_Report1.Name = "checkBox_Report1";
            this.checkBox_Report1.Size = new System.Drawing.Size(240, 17);
            this.checkBox_Report1.TabIndex = 0;
            this.checkBox_Report1.TabStop = false;
            this.checkBox_Report1.Text = "Нерешенные по дате решения обращения";
            this.checkBox_Report1.UseVisualStyleBackColor = true;
            // 
            // panel_Files
            // 
            this.panel_Files.Controls.Add(this.button_ExportSDSelect);
            this.panel_Files.Controls.Add(this.label_FileExportSDName);
            this.panel_Files.Controls.Add(this.textBox_FileExportSDName);
            this.panel_Files.Location = new System.Drawing.Point(0, 1);
            this.panel_Files.Name = "panel_Files";
            this.panel_Files.Size = new System.Drawing.Size(691, 42);
            this.panel_Files.TabIndex = 0;
            // 
            // panel_WeekMailTo
            // 
            this.panel_WeekMailTo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel_WeekMailTo.Controls.Add(this.textBox_WeekMailTo);
            this.panel_WeekMailTo.Location = new System.Drawing.Point(0, 166);
            this.panel_WeekMailTo.Name = "panel_WeekMailTo";
            this.panel_WeekMailTo.Size = new System.Drawing.Size(691, 145);
            this.panel_WeekMailTo.TabIndex = 11;
            // 
            // textBox_WeekMailTo
            // 
            this.textBox_WeekMailTo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox_WeekMailTo.Location = new System.Drawing.Point(0, 0);
            this.textBox_WeekMailTo.Multiline = true;
            this.textBox_WeekMailTo.Name = "textBox_WeekMailTo";
            this.textBox_WeekMailTo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox_WeekMailTo.Size = new System.Drawing.Size(691, 145);
            this.textBox_WeekMailTo.TabIndex = 0;
            this.textBox_WeekMailTo.Tag = "";
            this.textBox_WeekMailTo.Text = resources.GetString("textBox_WeekMailTo.Text");
            this.textBox_WeekMailTo.TextChanged += new System.EventHandler(this.textBox_WeekMailTo_TextChanged);
            // 
            // label_WeekMailTo
            // 
            this.label_WeekMailTo.AutoSize = true;
            this.label_WeekMailTo.Location = new System.Drawing.Point(3, 150);
            this.label_WeekMailTo.Name = "label_WeekMailTo";
            this.label_WeekMailTo.Size = new System.Drawing.Size(283, 13);
            this.label_WeekMailTo.TabIndex = 12;
            this.label_WeekMailTo.Text = "Почтовые адреса получателей еженедельного отчета:";
            // 
            // Form0
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(691, 342);
            this.Controls.Add(this.label_WeekMailTo);
            this.Controls.Add(this.panel_WeekMailTo);
            this.Controls.Add(this.panel_Files);
            this.Controls.Add(this.panel_DateTime);
            this.Controls.Add(this.panel_Report);
            this.Controls.Add(this.panel_Button);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form0";
            this.Text = "ReportMakerSD";
            this.panel_Button.ResumeLayout(false);
            this.panel_DateTime.ResumeLayout(false);
            this.panel_DateTime.PerformLayout();
            this.panel_Report.ResumeLayout(false);
            this.panel_Report.PerformLayout();
            this.panel_Files.ResumeLayout(false);
            this.panel_Files.PerformLayout();
            this.panel_WeekMailTo.ResumeLayout(false);
            this.panel_WeekMailTo.PerformLayout();
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
        private System.Windows.Forms.Button button_parsingReport;
        private System.Windows.Forms.Button button_SendMail;
        private System.Windows.Forms.Button button_OpenFolder;
        private System.Windows.Forms.Panel panel_Button;
        private System.Windows.Forms.Panel panel_DateTime;
        private System.Windows.Forms.Panel panel_Report;
        private System.Windows.Forms.Panel panel_Files;
        private System.Windows.Forms.ComboBox comboBox_SelectPeriod;
        public System.Windows.Forms.CheckBox checkBox_Report1;
        public System.Windows.Forms.CheckBox checkBox_Report2;
        public System.Windows.Forms.CheckBox checkBox_Report3;
        private System.Windows.Forms.Panel panel_WeekMailTo;
        private System.Windows.Forms.TextBox textBox_WeekMailTo;
        private System.Windows.Forms.Label label_WeekMailTo;
    }
}

