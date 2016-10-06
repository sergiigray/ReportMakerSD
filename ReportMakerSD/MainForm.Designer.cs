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
            this.checkBox_ReportRDU = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // openFileDialog_ExportSDSelect
            // 
            this.openFileDialog_ExportSDSelect.FileName = "openFileDialog1";
            // 
            // dateTimePicker_TimeFrom
            // 
            this.dateTimePicker_TimeFrom.CustomFormat = "dd.MM.yyyy HH:mm";
            this.dateTimePicker_TimeFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker_TimeFrom.Location = new System.Drawing.Point(15, 180);
            this.dateTimePicker_TimeFrom.Name = "dateTimePicker_TimeFrom";
            this.dateTimePicker_TimeFrom.Size = new System.Drawing.Size(125, 20);
            this.dateTimePicker_TimeFrom.TabIndex = 3;
            this.dateTimePicker_TimeFrom.Value = new System.DateTime(2016, 7, 27, 8, 30, 0, 0);
            // 
            // dateTimePicker_TimeTo
            // 
            this.dateTimePicker_TimeTo.CustomFormat = "dd.MM.yyyy HH:mm";
            this.dateTimePicker_TimeTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker_TimeTo.Location = new System.Drawing.Point(158, 180);
            this.dateTimePicker_TimeTo.Name = "dateTimePicker_TimeTo";
            this.dateTimePicker_TimeTo.Size = new System.Drawing.Size(125, 20);
            this.dateTimePicker_TimeTo.TabIndex = 4;
            this.dateTimePicker_TimeTo.Value = new System.DateTime(2016, 7, 27, 8, 30, 0, 0);
            // 
            // label_TimeFrom
            // 
            this.label_TimeFrom.AutoSize = true;
            this.label_TimeFrom.Location = new System.Drawing.Point(12, 164);
            this.label_TimeFrom.Name = "label_TimeFrom";
            this.label_TimeFrom.Size = new System.Drawing.Size(78, 13);
            this.label_TimeFrom.TabIndex = 1;
            this.label_TimeFrom.Text = "Время начала";
            // 
            // label_TimeTo
            // 
            this.label_TimeTo.AutoSize = true;
            this.label_TimeTo.Location = new System.Drawing.Point(155, 164);
            this.label_TimeTo.Name = "label_TimeTo";
            this.label_TimeTo.Size = new System.Drawing.Size(96, 13);
            this.label_TimeTo.TabIndex = 1;
            this.label_TimeTo.Text = "Время окончания";
            // 
            // textBox_FileExportSDName
            // 
            this.textBox_FileExportSDName.Location = new System.Drawing.Point(15, 24);
            this.textBox_FileExportSDName.Name = "textBox_FileExportSDName";
            this.textBox_FileExportSDName.Size = new System.Drawing.Size(352, 20);
            this.textBox_FileExportSDName.TabIndex = 0;
            this.textBox_FileExportSDName.TabStop = false;
            // 
            // button_ExportSDSelect
            // 
            this.button_ExportSDSelect.Location = new System.Drawing.Point(373, 24);
            this.button_ExportSDSelect.Name = "button_ExportSDSelect";
            this.button_ExportSDSelect.Size = new System.Drawing.Size(88, 20);
            this.button_ExportSDSelect.TabIndex = 5;
            this.button_ExportSDSelect.Text = "Выбрать файл";
            this.button_ExportSDSelect.UseVisualStyleBackColor = true;
            this.button_ExportSDSelect.Click += new System.EventHandler(this.button_ExportSDSelect_Click);
            // 
            // label_FileExportSDName
            // 
            this.label_FileExportSDName.AutoSize = true;
            this.label_FileExportSDName.Location = new System.Drawing.Point(12, 9);
            this.label_FileExportSDName.Name = "label_FileExportSDName";
            this.label_FileExportSDName.Size = new System.Drawing.Size(154, 13);
            this.label_FileExportSDName.TabIndex = 1;
            this.label_FileExportSDName.Text = "Файл с исходными данными";
            // 
            // textBox_DebugInfo
            // 
            this.textBox_DebugInfo.AcceptsReturn = true;
            this.textBox_DebugInfo.Location = new System.Drawing.Point(15, 206);
            this.textBox_DebugInfo.Multiline = true;
            this.textBox_DebugInfo.Name = "textBox_DebugInfo";
            this.textBox_DebugInfo.ReadOnly = true;
            this.textBox_DebugInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox_DebugInfo.Size = new System.Drawing.Size(446, 222);
            this.textBox_DebugInfo.TabIndex = 0;
            this.textBox_DebugInfo.TabStop = false;
            this.textBox_DebugInfo.WordWrap = false;
            // 
            // button_parsingReport
            // 
            this.button_parsingReport.Location = new System.Drawing.Point(15, 434);
            this.button_parsingReport.Name = "button_parsingReport";
            this.button_parsingReport.Size = new System.Drawing.Size(135, 23);
            this.button_parsingReport.TabIndex = 1;
            this.button_parsingReport.Text = "Обработка данных";
            this.button_parsingReport.UseVisualStyleBackColor = true;
            this.button_parsingReport.Click += new System.EventHandler(this.button_parsingReport_Click);
            // 
            // button_SendMail
            // 
            this.button_SendMail.Location = new System.Drawing.Point(326, 434);
            this.button_SendMail.Name = "button_SendMail";
            this.button_SendMail.Size = new System.Drawing.Size(135, 23);
            this.button_SendMail.TabIndex = 2;
            this.button_SendMail.Text = "Создание черновиков";
            this.button_SendMail.UseVisualStyleBackColor = true;
            this.button_SendMail.Click += new System.EventHandler(this.button_SendMail_Click);
            // 
            // checkBox_ReportTO
            // 
            this.checkBox_ReportTO.AutoSize = true;
            this.checkBox_ReportTO.Location = new System.Drawing.Point(15, 50);
            this.checkBox_ReportTO.Name = "checkBox_ReportTO";
            this.checkBox_ReportTO.Size = new System.Drawing.Size(207, 17);
            this.checkBox_ReportTO.TabIndex = 0;
            this.checkBox_ReportTO.TabStop = false;
            this.checkBox_ReportTO.Text = "Отчет по подрядным организациям";
            this.checkBox_ReportTO.UseVisualStyleBackColor = true;
            // 
            // checkBox_ReportRDU
            // 
            this.checkBox_ReportRDU.AutoSize = true;
            this.checkBox_ReportRDU.Checked = true;
            this.checkBox_ReportRDU.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ReportRDU.Location = new System.Drawing.Point(15, 73);
            this.checkBox_ReportRDU.Name = "checkBox_ReportRDU";
            this.checkBox_ReportRDU.Size = new System.Drawing.Size(125, 17);
            this.checkBox_ReportRDU.TabIndex = 0;
            this.checkBox_ReportRDU.TabStop = false;
            this.checkBox_ReportRDU.Text = "Отчет по филиалам";
            this.checkBox_ReportRDU.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(473, 470);
            this.Controls.Add(this.checkBox_ReportRDU);
            this.Controls.Add(this.checkBox_ReportTO);
            this.Controls.Add(this.button_SendMail);
            this.Controls.Add(this.button_parsingReport);
            this.Controls.Add(this.textBox_DebugInfo);
            this.Controls.Add(this.button_ExportSDSelect);
            this.Controls.Add(this.textBox_FileExportSDName);
            this.Controls.Add(this.label_TimeTo);
            this.Controls.Add(this.label_FileExportSDName);
            this.Controls.Add(this.label_TimeFrom);
            this.Controls.Add(this.dateTimePicker_TimeTo);
            this.Controls.Add(this.dateTimePicker_TimeFrom);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

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
        public System.Windows.Forms.CheckBox checkBox_ReportRDU;
    }
}

