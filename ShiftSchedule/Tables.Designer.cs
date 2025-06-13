namespace ShiftSchedule
{
    partial class Tables
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.менюToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.подключитьсяКБДToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выходToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.cmdChooseTable = new System.Windows.Forms.ComboBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.addRecordButton = new System.Windows.Forms.Button();
            this.editRecordButton = new System.Windows.Forms.Button();
            this.deleteRecordButton = new System.Windows.Forms.Button();
            this.Reportbutton = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.менюToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1154, 30);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // менюToolStripMenuItem
            // 
            this.менюToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.подключитьсяКБДToolStripMenuItem,
            this.выходToolStripMenuItem});
            this.менюToolStripMenuItem.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.менюToolStripMenuItem.Name = "менюToolStripMenuItem";
            this.менюToolStripMenuItem.Size = new System.Drawing.Size(59, 24);
            this.менюToolStripMenuItem.Text = "Файл";
            // 
            // подключитьсяКБДToolStripMenuItem
            // 
            this.подключитьсяКБДToolStripMenuItem.Name = "подключитьсяКБДToolStripMenuItem";
            this.подключитьсяКБДToolStripMenuItem.Size = new System.Drawing.Size(227, 26);
            this.подключитьсяКБДToolStripMenuItem.Text = "Подключиться к БД";
            this.подключитьсяКБДToolStripMenuItem.Click += new System.EventHandler(this.подключитьсяКБДToolStripMenuItem_Click);
            // 
            // выходToolStripMenuItem
            // 
            this.выходToolStripMenuItem.Name = "выходToolStripMenuItem";
            this.выходToolStripMenuItem.Size = new System.Drawing.Size(227, 26);
            this.выходToolStripMenuItem.Text = "Выход";
            this.выходToolStripMenuItem.Click += new System.EventHandler(this.выходToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(12, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(171, 22);
            this.label1.TabIndex = 1;
            this.label1.Text = "Выберите таблицу:";
            // 
            // cmdChooseTable
            // 
            this.cmdChooseTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.cmdChooseTable.FormattingEnabled = true;
            this.cmdChooseTable.Location = new System.Drawing.Point(199, 50);
            this.cmdChooseTable.Name = "cmdChooseTable";
            this.cmdChooseTable.Size = new System.Drawing.Size(255, 30);
            this.cmdChooseTable.TabIndex = 2;
            this.cmdChooseTable.SelectedIndexChanged += new System.EventHandler(this.cmdChooseTable_SelectedIndexChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(16, 94);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(927, 399);
            this.dataGridView1.TabIndex = 3;
            // 
            // addRecordButton
            // 
            this.addRecordButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.addRecordButton.Location = new System.Drawing.Point(957, 94);
            this.addRecordButton.Name = "addRecordButton";
            this.addRecordButton.Size = new System.Drawing.Size(168, 45);
            this.addRecordButton.TabIndex = 4;
            this.addRecordButton.Text = "Добавить";
            this.addRecordButton.UseVisualStyleBackColor = true;
            this.addRecordButton.Click += new System.EventHandler(this.addRecordButton_Click);
            // 
            // editRecordButton
            // 
            this.editRecordButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.editRecordButton.Location = new System.Drawing.Point(957, 145);
            this.editRecordButton.Name = "editRecordButton";
            this.editRecordButton.Size = new System.Drawing.Size(168, 45);
            this.editRecordButton.TabIndex = 5;
            this.editRecordButton.Text = "Редактировать";
            this.editRecordButton.UseVisualStyleBackColor = true;
            this.editRecordButton.Click += new System.EventHandler(this.editRecordButton_Click);
            // 
            // deleteRecordButton
            // 
            this.deleteRecordButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.deleteRecordButton.Location = new System.Drawing.Point(957, 196);
            this.deleteRecordButton.Name = "deleteRecordButton";
            this.deleteRecordButton.Size = new System.Drawing.Size(168, 45);
            this.deleteRecordButton.TabIndex = 6;
            this.deleteRecordButton.Text = "Удалить";
            this.deleteRecordButton.UseVisualStyleBackColor = true;
            this.deleteRecordButton.Click += new System.EventHandler(this.deleteRecordButton_Click);
            // 
            // Reportbutton
            // 
            this.Reportbutton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.Reportbutton.Location = new System.Drawing.Point(957, 448);
            this.Reportbutton.Name = "Reportbutton";
            this.Reportbutton.Size = new System.Drawing.Size(168, 45);
            this.Reportbutton.TabIndex = 7;
            this.Reportbutton.Text = "Отчеты";
            this.Reportbutton.UseVisualStyleBackColor = true;
            this.Reportbutton.Click += new System.EventHandler(this.Reportbutton_Click);
            // 
            // Tables
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1154, 507);
            this.Controls.Add(this.Reportbutton);
            this.Controls.Add(this.deleteRecordButton);
            this.Controls.Add(this.editRecordButton);
            this.Controls.Add(this.addRecordButton);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.cmdChooseTable);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Tables";
            this.Text = "График смены. Таблицы";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem менюToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem подключитьсяКБДToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem выходToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmdChooseTable;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button addRecordButton;
        private System.Windows.Forms.Button editRecordButton;
        private System.Windows.Forms.Button deleteRecordButton;
        private System.Windows.Forms.Button Reportbutton;
    }
}

