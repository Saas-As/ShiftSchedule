using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace ShiftSchedule
{
    /// <summary>
    /// Класс формы для формирования и отображения отчётов о сменах.
    /// </summary>
    public partial class Reports : Form
    {
        private BusinessLogic _businessLogic;

        /// <summary>
        /// Конструктор формы Reports.
        /// </summary>
        /// <param name="businessLogic">Объект бизнес-логики для выполнения запросов к базе данных.</param>
        public Reports(BusinessLogic businessLogic)
        {
            InitializeComponent();
            _businessLogic = businessLogic;
        }

        /// <summary>
        /// Обработчик события изменения выбранного типа отчёта.
        /// Очищает панель критериев и добавляет соответствующие элементы управления в зависимости от выбранного типа отчёта.
        /// </summary>
        private void cmbReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnGenerate.Enabled = true;
            pnlCriteria.Controls.Clear();
            var label = new System.Windows.Forms.Label { Text = "Критерии:", Location = new System.Drawing.Point(10, 10) };
            pnlCriteria.Controls.Add(label);

            switch (cmbReportType.SelectedItem.ToString())
            {
                case "Смены по дате":
                    AddDateCriteriaControls();
                    break;
                case "Смены по подразделению":
                    AddDepartmentCriteriaControls();
                    break;
                case "Смены по начальнику смены":
                    AddShiftManagerCriteriaControls();
                    break;
                case "Сводный отчет за период":
                    AddSummaryCriteriaControls();
                    break;
            }
        }

        /// <summary>
        /// Добавляет элементы управления для выбора даты начала и конца для отчёта "Смены по дате".
        /// </summary>
        private void AddDateCriteriaControls()
        {
            var lblFrom = new System.Windows.Forms.Label
            {
                Text = "С:",
                Location = new System.Drawing.Point(10, 40),
                AutoSize = true
            };

            var dtpFrom = new DateTimePicker
            {
                Name = "dtpFrom",
                Location = new System.Drawing.Point(40, 40),
                Format = DateTimePickerFormat.Short,
                Width = 120,
                Value = DateTime.Today
            };

            var lblTo = new System.Windows.Forms.Label
            {
                Text = "По:",
                Location = new System.Drawing.Point(10, 80),
                AutoSize = true
            };

            var dtpTo = new DateTimePicker
            {
                Name = "dtpTo",
                Location = new System.Drawing.Point(40, 80),
                Format = DateTimePickerFormat.Short,
                Width = 120,
                Value = DateTime.Today.AddDays(7)
            };

            pnlCriteria.Controls.AddRange(new Control[] { lblFrom, dtpFrom, lblTo, dtpTo });
        }

        /// <summary>
        /// Добавляет элемент управления для выбора подразделения для отчёта "Смены по подразделению".
        /// </summary>
        private void AddDepartmentCriteriaControls()
        {
            var lblDept = new System.Windows.Forms.Label { Text = "Подразделение:", Location = new System.Drawing.Point(10, 40) };
            var cmbDept = new ComboBox { Name = "cmbDept", Location = new System.Drawing.Point(120, 40), Width = 200 };

            // Заполнение из таблицы Подразделения
            var depts = _businessLogic.GetTableData("Подразделения");
            foreach (DataRow row in depts.Rows)
            {
                cmbDept.Items.Add(row["Подразделение"].ToString());
            }

            pnlCriteria.Controls.AddRange(new Control[] { lblDept, cmbDept });
        }

        /// <summary>
        /// Добавляет элемент управления для выбора начальника смены для отчёта "Смены по начальнику смены".
        /// </summary>
        private void AddShiftManagerCriteriaControls()
        {
            var lblManager = new System.Windows.Forms.Label { Text = "Начальник:", Location = new System.Drawing.Point(10, 40) };
            var cmbManager = new ComboBox { Name = "cmbManager", Location = new System.Drawing.Point(120, 40), Width = 200 };

            // Заполнение из таблицы Начальники смен
            var managers = _businessLogic.GetTableData("Начальники смен");
            foreach (DataRow row in managers.Rows)
            {
                cmbManager.Items.Add(row["ФИО_начальника_смены"].ToString());
            }

            pnlCriteria.Controls.AddRange(new Control[] { lblManager, cmbManager });
        }

        /// <summary>
        /// Добавляет элементы управления для выбора месяца и года для сводного отчёта за период.
        /// </summary>
        private void AddSummaryCriteriaControls()
        {
            var lblMonth = new System.Windows.Forms.Label { Text = "Месяц:", Location = new System.Drawing.Point(10, 40) };
            var cmbMonth = new ComboBox
            {
                Name = "cmbMonth",
                Location = new System.Drawing.Point(120, 40),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // Заполняем месяцы
            for (int i = 1; i <= 12; i++)
            {
                cmbMonth.Items.Add(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i));
            }
            cmbMonth.SelectedIndex = DateTime.Now.Month - 1;

            var lblYear = new System.Windows.Forms.Label { Text = "Год:", Location = new System.Drawing.Point(10, 80) };
            var numYear = new NumericUpDown
            {
                Name = "numYear",
                Location = new System.Drawing.Point(120, 80),
                Width = 200,
                Minimum = 2000,
                Maximum = 2100,
                Value = DateTime.Now.Year
            };

            pnlCriteria.Controls.AddRange(new Control[] { lblMonth, cmbMonth, lblYear, numYear });
        }

        /// <summary>
        /// Обработчик события нажатия кнопки "Сформировать отчёт".
        /// Формирует отчёт в зависимости от выбранного типа и отображает его в DataGridView.
        /// </summary>
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable reportData = null;

                switch (cmbReportType.SelectedItem.ToString())
                {
                    case "Смены по дате":
                        reportData = GenerateDateReport();
                        break;
                    case "Смены по подразделению":
                        reportData = GenerateDepartmentReport();
                        break;
                    case "Смены по начальнику смены":
                        reportData = GenerateShiftManagerReport();
                        break;
                    case "Сводный отчет за период":
                        reportData = GenerateSummaryReport();
                        break;
                }
                btnExportExcel.Enabled = true;
                btnExportWord.Enabled = true;

                if (reportData != null)
                {
                    dgvReport.DataSource = reportData;
                    dgvReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка формирования отчета: {ex.Message}", "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Формирует отчёт "Смены по дате".
        /// </summary>
        private System.Data.DataTable GenerateDateReport()
        {
            var dtpFrom = pnlCriteria.Controls.Find("dtpFrom", true).FirstOrDefault() as DateTimePicker;
            var dtpTo = pnlCriteria.Controls.Find("dtpTo", true).FirstOrDefault() as DateTimePicker;

            if (dtpFrom == null || dtpTo == null)
            {
                MessageBox.Show("Не найдены элементы выбора даты");
                return null;
            }

            if (dtpFrom.Value > dtpTo.Value)
            {
                MessageBox.Show("Дата 'С' не может быть позже даты 'По'");
                return null;
            }

            string query = $@"SELECT 
                            s.[Код смены], 
                            s.[Дата],
                            p.[Подразделение],
                            r.[ФИО_руководителя],
                            n.[ФИО_начальника_смены],
                            k.[Количество рабочих],
                            d.[Длительность смены]
                            FROM (((([Смены] s
                            LEFT JOIN [Подразделения] p ON s.[ID_подразделения] = p.[ID_подразделения])
                            LEFT JOIN [Руководители] r ON s.[ID_руководителя] = r.[ID_руководителя])
                            LEFT JOIN [Начальники смен] n ON s.[ID_начальника_смены] = n.[ID_начальника_смены])
                            LEFT JOIN [Количество рабочих] k ON s.[ID_количества_рабочих] = k.[ID_количества_рабочих])
                            LEFT JOIN [Длительности смен] d ON s.[ID_длительности_смены] = d.[ID_длительности_смены]
                            WHERE FORMAT(s.[Дата], 'dd.MM.yyyy') BETWEEN '{dtpFrom.Value:dd.MM.yyyy}' AND '{dtpTo.Value:dd.MM.yyyy}'
                            ORDER BY s.[Дата]";

            return _businessLogic.ExecuteCustomQuery(query);
        }

        /// <summary>
        /// Формирует отчёт "Смены по подразделению".
        /// </summary>
        private System.Data.DataTable GenerateDepartmentReport()
        {
            var cmbDept = pnlCriteria.Controls.Find("cmbDept", true).FirstOrDefault() as ComboBox;

            if (cmbDept == null || cmbDept.SelectedItem == null)
            {
                MessageBox.Show("Выберите подразделение");
                return null;
            }

            string query = $@"SELECT 
                                s.[Код смены], 
                                s.[Дата],
                                p.[Подразделение],
                                r.[ФИО_руководителя],
                                n.[ФИО_начальника_смены],
                                k.[Количество рабочих],
                                d.[Длительность смены]
                                FROM (((([Смены] s
                                LEFT JOIN [Подразделения] p ON s.[ID_подразделения] = p.[ID_подразделения])
                                LEFT JOIN [Руководители] r ON s.[ID_руководителя] = r.[ID_руководителя])
                                LEFT JOIN [Начальники смен] n ON s.[ID_начальника_смены] = n.[ID_начальника_смены])
                                LEFT JOIN [Количество рабочих] k ON s.[ID_количества_рабочих] = k.[ID_количества_рабочих])
                                LEFT JOIN [Длительности смен] d ON s.[ID_длительности_смены] = d.[ID_длительности_смены]
                                WHERE p.[Подразделение] = '{cmbDept.SelectedItem}'
                                ORDER BY  s.[Дата]";

            return _businessLogic.ExecuteCustomQuery(query);
        }

        /// <summary>
        /// Формирует отчёт "Смены по начальнику смены".
        /// </summary>
        private System.Data.DataTable GenerateShiftManagerReport()
        {
            var cmbManager = pnlCriteria.Controls.Find("cmbManager", true).FirstOrDefault() as ComboBox;

            if (cmbManager == null || cmbManager.SelectedItem == null)
            {
                MessageBox.Show("Выберите начальника");
                return null;
            }

            string query = $@"SELECT 
                            s.[Код смены], 
                            s.[Дата],
                            p.[Подразделение],
                            r.[ФИО_руководителя],
                            n.[ФИО_начальника_смены],
                            k.[Количество рабочих],
                            d.[Длительность смены]
                            FROM (((([Смены] s
                            LEFT JOIN [Подразделения] p ON s.[ID_подразделения] = p.[ID_подразделения])
                            LEFT JOIN [Руководители] r ON s.[ID_руководителя] = r.[ID_руководителя])
                            LEFT JOIN [Начальники смен] n ON s.[ID_начальника_смены] = n.[ID_начальника_смены])
                            LEFT JOIN [Количество рабочих] k ON s.[ID_количества_рабочих] = k.[ID_количества_рабочих])
                            LEFT JOIN [Длительности смен] d ON s.[ID_длительности_смены] = d.[ID_длительности_смены]
                            WHERE n.[ФИО_начальника_смены] = '{cmbManager.SelectedItem}'
                            ORDER BY  s.[Дата]";

            return _businessLogic.ExecuteCustomQuery(query);
        }

        /// <summary>
        /// Формирует сводный отчёт за период.
        /// </summary>
        private System.Data.DataTable GenerateSummaryReport()
        {
            var cmbMonth = pnlCriteria.Controls.Find("cmbMonth", true).FirstOrDefault() as ComboBox;
            var numYear = pnlCriteria.Controls.Find("numYear", true).FirstOrDefault() as NumericUpDown;

            if (cmbMonth == null || numYear == null || cmbMonth.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите месяц и год");
                return null;
            }

            int month = cmbMonth.SelectedIndex + 1;
            int year = (int)numYear.Value;

            string query = $@"SELECT 
                            p.[Подразделение],
                            COUNT(s.[Код смены]) AS [Количество смен],
                            SUM(k.[Количество рабочих]) AS [Общее количество рабочих],
                            SUM(d.[Длительность смены]) AS [Общая длительность (часов)]
                            FROM (([Смены] s
                            LEFT JOIN [Подразделения] p ON s.[ID_подразделения] = p.[ID_подразделения])
                            LEFT JOIN [Количество рабочих] k ON s.[ID_количества_рабочих] = k.[ID_количества_рабочих])
                            LEFT JOIN [Длительности смен] d ON s.[ID_длительности_смены] = d.[ID_длительности_смены]
                            WHERE MONTH(s.[Дата]) = {month} AND YEAR(s.[Дата]) = {year}
                            GROUP BY p.[Подразделение]";

            return _businessLogic.ExecuteCustomQuery(query);
        }

        /// <summary>
        /// Обработчик события загрузки формы.
        /// Отключает кнопки "Сформировать отчёт", "Excel" и "Word" при загрузке формы.
        /// </summary>
        private void Reports_Load(object sender, EventArgs e)
        {
            btnGenerate.Enabled = false;
            btnExportExcel.Enabled = false;
            btnExportWord.Enabled = false;
        }

        /// <summary>
        /// Обработчик события нажатия кнопки "Вернуться".
        /// Закрывает текущую форму.
        /// </summary>
        private void btnReturn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Обработчик события нажатия кнопки "Excel".
        /// Экспортирует данные из DataGridView в Excel.
        /// </summary>
        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (dgvReport.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            try
            {
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;
                Workbook workbook = excelApp.Workbooks.Add();
                Worksheet worksheet = (Worksheet)workbook.ActiveSheet;

                // Заголовки столбцов
                for (int i = 0; i < dgvReport.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = dgvReport.Columns[i].HeaderText;
                }

                // Данные
                for (int i = 0; i < dgvReport.Rows.Count; i++)
                {
                    for (int j = 0; j < dgvReport.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dgvReport.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // Автоподбор ширины столбцов
                worksheet.Columns.AutoFit();

                MessageBox.Show("Данные успешно экспортированы в Excel");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта в Excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Обработчик события нажатия кнопки "Word".
        /// Экспортирует данные из DataGridView в Word.
        /// </summary>
        private void btnExportWord_Click(object sender, EventArgs e)
        {
            if (dgvReport.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }
            try
            {
                var WordApp = new Microsoft.Office.Interop.Word.Application();
                WordApp.Visible = true;
                Document document = WordApp.Documents.Add();

                // Добавляем заголовок
                Paragraph title = document.Content.Paragraphs.Add();
                title.Range.Text = $"Отчет: {cmbReportType.SelectedItem}";
                title.Range.Font.Bold = 1;
                title.Range.Font.Size = 14;
                title.Format.SpaceAfter = 24;
                title.Range.InsertParagraphAfter();

                // Создаем таблицу
                Table table = document.Tables.Add(
                    document.Content,
                    dgvReport.Rows.Count + 1,
                    dgvReport.Columns.Count
                    );
                table.Borders.Enable = 1;
                // Заголовки столбцов
                for (int i = 0; i < dgvReport.Columns.Count; i++)
                {
                    table.Cell(1, i + 1).Range.Text = dgvReport.Columns[i].HeaderText;
                    table.Cell(1, i + 1).Range.Font.Bold = 1;
                }

                // Данные
                for (int i = 0; i < dgvReport.Rows.Count; i++)
                {
                    for (int j = 0; j < dgvReport.Columns.Count; j++)
                    {
                        table.Cell(i + 2, j + 1).Range.Text = dgvReport.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // Автоподбор таблицы
                table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                MessageBox.Show("Данные успешно экспортированы в Word");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта в Word: ex.Message");
            }
        }
    }
}
