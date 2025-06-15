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
    /// Форма для генерации отчетов по данным смен.
    /// Предоставляет функционал для:
    /// - Выбора типа отчета (по дате, подразделению и т.д.)
    /// - Формирования отчетных данных
    /// - Отображения отчетов в табличном виде
    /// - Экспорта отчетов в Excel и Word
    /// </summary>
    public partial class Reports : Form
    {
        // Объект бизнес-логики для выполнения запросов к базе данных
        private BusinessLogic _businessLogic;

        /// <summary>
        /// Конструктор формы Reports.
        /// Инициализирует форму и сохраняет ссылку на объект бизнес-логики.
        /// </summary>
        /// <param name="businessLogic">Объект бизнес-логики для работы с данными</param>
        public Reports(BusinessLogic businessLogic)
        {
            InitializeComponent();
            // Сохраняем переданный объект бизнес-логики
            _businessLogic = businessLogic;
        }

        /// <summary>
        /// Обработчик изменения выбранного типа отчета.
        /// В зависимости от выбранного типа добавляет соответствующие элементы управления для ввода параметров отчета.
        /// </summary>
        private void cmbReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Активируем кнопку генерации отчета
            btnGenerate.Enabled = true;

            // Очищаем панель с критериями от предыдущих элементов
            pnlCriteria.Controls.Clear();

            // Добавляем общую надпись "Критерии:"
            var label = new System.Windows.Forms.Label { Text = "Критерии:", Location = new System.Drawing.Point(10, 10) };
            pnlCriteria.Controls.Add(label);

            // В зависимости от выбранного типа отчета добавляем соответствующие элементы управления
            switch (cmbReportType.SelectedItem.ToString())
            {
                case "Смены по дате":
                    AddDateCriteriaControls(); // Добавляем элементы для выбора дат
                    break;
                case "Смены по подразделению":
                    AddDepartmentCriteriaControls(); // Добавляем выпадающий список подразделений
                    break;
                case "Смены по начальнику смены":
                    AddShiftManagerCriteriaControls();  // Добавляем выпадающий список начальников
                    break;
                case "Сводный отчет за период":
                    AddSummaryCriteriaControls(); // Добавляем элементы для выбора месяца и года
                    break;
            }
        }

        /// <summary>
        /// Добавляет элементы управления для выбора диапазона дат (от/до).
        /// Используется для отчетов "Смены по дате".
        /// </summary>
        private void AddDateCriteriaControls()
        {
            // Метка "С:"
            var lblFrom = new System.Windows.Forms.Label
            {
                Text = "С:",
                Location = new System.Drawing.Point(10, 40),
                AutoSize = true
            };

            // Поле выбора начальной даты
            var dtpFrom = new DateTimePicker
            {
                Name = "dtpFrom", // Уникальное имя для последующего поиска
                Location = new System.Drawing.Point(40, 40),
                Format = DateTimePickerFormat.Short, // Короткий формат даты
                Width = 120,
                Value = DateTime.Today // Текущая дата по умолчанию
            };

            // Метка "По:"
            var lblTo = new System.Windows.Forms.Label
            {
                Text = "По:",
                Location = new System.Drawing.Point(10, 80),
                AutoSize = true
            };

            // Поле выбора конечной даты
            var dtpTo = new DateTimePicker
            {
                Name = "dtpTo",
                Location = new System.Drawing.Point(40, 80),
                Format = DateTimePickerFormat.Short,
                Width = 120,
                Value = DateTime.Today.AddDays(7) // Текущая дата + 7 дней по умолчанию
            };

            // Добавляем все элементы на панель
            pnlCriteria.Controls.AddRange(new Control[] { lblFrom, dtpFrom, lblTo, dtpTo });
        }

        /// <summary>
        /// Добавляет выпадающий список для выбора подразделения.
        /// Используется для отчетов "Смены по подразделению".
        /// </summary>
        private void AddDepartmentCriteriaControls()
        {
            // Метка "Подразделение:"
            var lblDept = new System.Windows.Forms.Label
            {
                Text = "Подразделение:",
                Location = new System.Drawing.Point(10, 40)
            };

            // Выпадающий список подразделений
            var cmbDept = new ComboBox
            {
                Name = "cmbDept",
                Location = new System.Drawing.Point(120, 40),
                Width = 200
            };

            // Получаем данные о подразделениях из базы данных
            var depts = _businessLogic.GetTableData("Подразделения");
            // Заполняем выпадающий список названиями подразделений
            foreach (DataRow row in depts.Rows)
            {
                cmbDept.Items.Add(row["Подразделение"].ToString());
            }

            // Добавляем элементы на панель
            pnlCriteria.Controls.AddRange(new Control[] { lblDept, cmbDept });
        }

        /// <summary>
        /// Добавляет выпадающий список для выбора начальника смены.
        /// Используется для отчетов "Смены по начальнику смены".
        /// </summary>
        private void AddShiftManagerCriteriaControls()
        {
            // Метка "Начальник:"
            var lblManager = new System.Windows.Forms.Label
            {
                Text = "Начальник:",
                Location = new System.Drawing.Point(10, 40)
            };

            // Выпадающий список начальников
            var cmbManager = new ComboBox
            {
                Name = "cmbManager",
                Location = new System.Drawing.Point(120, 40),
                Width = 200
            };

            // Получаем данные о начальниках смен из базы данных
            var managers = _businessLogic.GetTableData("Начальники смен");
            // Заполняем выпадающий список ФИО начальников
            foreach (DataRow row in managers.Rows)
            {
                cmbManager.Items.Add(row["ФИО_начальника_смены"].ToString());
            }

            // Добавляем элементы на панель
            pnlCriteria.Controls.AddRange(new Control[] { lblManager, cmbManager });
        }

        /// <summary>
        /// Добавляет элементы управления для выбора месяца и года.
        /// Используется для "Сводного отчета за период".
        /// </summary>
        private void AddSummaryCriteriaControls()
        {
            // Метка "Месяц:"
            var lblMonth = new System.Windows.Forms.Label
            {
                Text = "Месяц:",
                Location = new System.Drawing.Point(10, 40)
            };

            // Выпадающий список месяцев
            var cmbMonth = new ComboBox
            {
                Name = "cmbMonth",
                Location = new System.Drawing.Point(120, 40),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList // Запрет на ручной ввод
            };

            // Заполняем список названиями месяцев
            for (int i = 1; i <= 12; i++)
            {
                cmbMonth.Items.Add(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i));
            }
            // Устанавливаем текущий месяц по умолчанию
            cmbMonth.SelectedIndex = DateTime.Now.Month - 1;

            // Метка "Год:"
            var lblYear = new System.Windows.Forms.Label
            {
                Text = "Год:",
                Location = new System.Drawing.Point(10, 80)
            };

            // Числовое поле для ввода года
            var numYear = new NumericUpDown
            {
                Name = "numYear",
                Location = new System.Drawing.Point(120, 80),
                Width = 200,
                Minimum = 2000, // Минимальный год
                Maximum = 2100, // Максимальный год
                Value = DateTime.Now.Year // Текущий год по умолчанию
            };

            // Добавляем элементы на панель
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

                // В зависимости от выбранного типа отчета вызываем соответствующий метод генерации
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

                // Активируем кнопки экспорта после успешной генерации
                btnExportExcel.Enabled = true;
                btnExportWord.Enabled = true;

                // Если данные получены, отображаем их в DataGridView
                if (reportData != null)
                {
                    dgvReport.DataSource = reportData;
                    // Автоподбор ширины столбцов
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
        /// <returns>DataTable с данными отчета</returns>
        private System.Data.DataTable GenerateDateReport()
        {
            // Находим элементы управления для выбора дат на панели
            var dtpFrom = pnlCriteria.Controls.Find("dtpFrom", true).FirstOrDefault() as DateTimePicker;
            var dtpTo = pnlCriteria.Controls.Find("dtpTo", true).FirstOrDefault() as DateTimePicker;

            // Проверяем, что элементы найдены
            if (dtpFrom == null || dtpTo == null)
            {
                MessageBox.Show("Не найдены элементы выбора даты");
                return null;
            }

            // Проверяем, что начальная дата не позже конечной
            if (dtpFrom.Value > dtpTo.Value)
            {
                MessageBox.Show("Дата 'С' не может быть позже даты 'По'");
                return null;
            }

            // Форматируем даты в нужный формат
            string fromDate = dtpFrom.Value.ToString("dd.MM.yyyy");
            string toDate = dtpTo.Value.ToString("dd.MM.yyyy");

            // Формируем SQL-запрос с JOIN для получения данных из связанных таблиц
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
                    WHERE s.[Дата] BETWEEN CDate('{fromDate}') AND CDate('{toDate}')
                    ORDER BY s.[Дата]";

            // Выполняем запрос через бизнес-логику и возвращаем результат
            return _businessLogic.ExecuteCustomQuery(query);
        }

        /// <summary>
        /// Формирует отчёт "Смены по подразделению".
        /// </summary>
        /// <returns>DataTable с данными отчета</returns>
        private System.Data.DataTable GenerateDepartmentReport()
        {
            // Находим выпадающий список подразделений на панели
            var cmbDept = pnlCriteria.Controls.Find("cmbDept", true).FirstOrDefault() as ComboBox;

            // Проверяем, что элемент найден и выбран элемент
            if (cmbDept == null || cmbDept.SelectedItem == null)
            {
                MessageBox.Show("Выберите подразделение");
                return null;
            }

            // Формируем SQL-запрос для выбранного подразделения
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

            // Выполняем запрос через бизнес-логику и возвращаем результат
            return _businessLogic.ExecuteCustomQuery(query);
        }

        /// <summary>
        /// Формирует отчёт "Смены по начальнику смены".
        /// </summary>
        /// <returns>DataTable с данными отчета</returns>
        private System.Data.DataTable GenerateShiftManagerReport()
        {
            // Находим выпадающий список начальников на панели
            var cmbManager = pnlCriteria.Controls.Find("cmbManager", true).FirstOrDefault() as ComboBox;

            // Проверяем, что элемент найден и выбран элемент
            if (cmbManager == null || cmbManager.SelectedItem == null)
            {
                MessageBox.Show("Выберите начальника");
                return null;
            }

            // Формируем SQL-запрос для выбранного начальника
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

            // Выполняем запрос через бизнес-логику и возвращаем результат
            return _businessLogic.ExecuteCustomQuery(query);
        }

        /// <summary>
        /// Формирует сводный отчёт за период.
        /// </summary>
        /// <returns>DataTable с данными отчета</returns>
        private System.Data.DataTable GenerateSummaryReport()
        {
            // Находим элементы выбора месяца и года на панели
            var cmbMonth = pnlCriteria.Controls.Find("cmbMonth", true).FirstOrDefault() as ComboBox;
            var numYear = pnlCriteria.Controls.Find("numYear", true).FirstOrDefault() as NumericUpDown;

            // Проверяем, что элементы найдены и месяц выбран
            if (cmbMonth == null || numYear == null || cmbMonth.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите месяц и год");
                return null;
            }

            // Получаем выбранные месяц и год
            int month = cmbMonth.SelectedIndex + 1; // Индекс месяца + 1 (т.к. индексация с 0)
            int year = (int)numYear.Value;

            // Формируем SQL-запрос с группировкой по подразделениям
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

            // Выполняем запрос через бизнес-логику и возвращаем результат
            return _businessLogic.ExecuteCustomQuery(query);
        }

        /// <summary>
        /// Обработчик события загрузки формы.
        /// Отключает кнопки "Сформировать отчёт", "Excel" и "Word" при загрузке формы.
        /// </summary>
        private void Reports_Load(object sender, EventArgs e)
        {
            // Отключаем кнопки при загрузке формы (пока не выбран тип отчета)
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
            // Проверяем, что есть данные для экспорта
            if (dgvReport.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            try
            {
                // Создаем новое приложение Excel
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true; // Делаем Excel видимым
                // Создаем новую книгу
                Workbook workbook = excelApp.Workbooks.Add();
                // Получаем активный лист
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
            // Проверяем, что есть данные для экспорта
            if (dgvReport.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }
            try
            {
                // Создаем новое приложение Word
                var WordApp = new Microsoft.Office.Interop.Word.Application();
                WordApp.Visible = true; // Делаем Word видимым

                // Создаем новый документ
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
                    dgvReport.Rows.Count + 1, // Количество строк (+1 для заголовков)
                    dgvReport.Columns.Count   // Количество столбцов
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
                MessageBox.Show($"Ошибка экспорта в Word: {ex.Message}");
            }
        }
    }
}
