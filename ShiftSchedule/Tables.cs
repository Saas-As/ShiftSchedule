using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShiftSchedule
{
    /// <summary>
    /// Главная форма приложения для работы с расписанием смен.
    /// Предоставляет интерфейс для подключения к базе данных, просмотра, 
    /// добавления, редактирования и удаления записей, а также формирования отчетов.
    /// </summary>
    public partial class Tables : Form
    {
        // Поля класса для хранения состояния приложения
        private string selectedDatabasePath;  // Путь к выбранной базе данных
        private OleDbConnection connection;   // Подключение к базе данных
        private BusinessLogic _businessLogic; // Объект бизнес-логики для работы с данными

        /// <summary>
        /// Конструктор главной формы. Инициализирует компоненты и настраивает начальное состояние.
        /// </summary>
        public Tables()
        {
            InitializeComponent(); // Инициализация компонентов формы
            _businessLogic = null; // Пока не выбрана БД, бизнес-логика не инициализирована

            // Добавляем обработчик закрытия формы
            this.FormClosing += Tables_FormClosing;
        }
        /// <summary>
        /// Обработчик события закрытия формы. Закрывает соединение с БД.
        /// </summary>
        private void Tables_FormClosing(object sender, FormClosingEventArgs e)
        {
            connection?.Close(); // Если соединение открыто, закрываем его
        }
        /// <summary>
        /// Обработчик нажатия пункта меню "Подключиться к БД"
        /// Открывает диалог выбора файла базы данных и инициализирует подключение.
        /// </summary>
        private void подключитьсяКБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Путь к папке с базами данных (относительно текущей директории)
            string databaseFolder = Path.Combine(Directory.GetCurrentDirectory(), "DataBase");

            // Проверка существования папки
            if (!Directory.Exists(databaseFolder))
            {
                MessageBox.Show("Папка с базой данных не найдена: " + databaseFolder,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Настройка диалога выбора файла
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = databaseFolder, // Начальная директория
                Title = "Выберите файл базы данных Access", // Заголовок окна
                Filter = "Базы данных Access (*.accdb; *.mdb)|*.accdb;*.mdb|Все файлы (*.*)|*.*", // Фильтры файлов
                RestoreDirectory = true // Восстанавливаем предыдущую директорию при следующем открытии
            };

            // Показываем диалог и обрабатываем результат
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Сохраняем путь к БД
                selectedDatabasePath = openFileDialog.FileName;

                // Инициализируем бизнес-логику с выбранной БД
                _businessLogic = new BusinessLogic(openFileDialog.FileName);
                
                // Вызов метода для установки соединения
                подключитьсяКБДToolStripMenuItem_Click_AfterSelection();
            }
        }
        /// <summary>
        /// Метод для установки соединения с базой данных после выбора файла.
        /// Вызывает форму аутентификации и загружает список таблиц.
        /// </summary>
        private void подключитьсяКБДToolStripMenuItem_Click_AfterSelection()
        {
            // Проверка, что выбран путь к БД
            if (string.IsNullOrEmpty(selectedDatabasePath))
            {
                MessageBox.Show("Пожалуйста, выберите файл базы данных.", "Выберите файл",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //Расширение файла для определения типа базы данных
            string extension = Path.GetExtension(selectedDatabasePath).ToLower();
            // Провайдер
            string provider = "";
            // Строка подключения
            string connectionString = "";

            // Выбираем провайдера в зависимости от расширения файла
            switch (extension)
            {
                case ".accdb":
                case ".mdb":
                    provider = "Microsoft.ACE.OLEDB.16.0";
                    connectionString = @"Provider=" + provider + ";Data Source=" + selectedDatabasePath +
                        ";Persist Security Info=False;";
                    break;
                default:
                    MessageBox.Show("Неподдерживаемый тип базы данных.", 
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
            }
            // создаем подключение к базе данных
            connection = new OleDbConnection(connectionString);

            // Показываем форму аутентификации
            var loginForm = new LoginForm(selectedDatabasePath);
            loginForm.ShowDialog();

            // Вызов метода загрузки списка таблиц
            LoadTableNames();
        }
        /// <summary>
        /// Метод загрузки списка таблиц из базы данных.
        /// Заполняет выпадающий список доступными таблицами.
        /// </summary>
        private void LoadTableNames()
        {
            try
            {
                // открываем соединение с БД
                connection.Open();

                // Сообщение об успешном подключении
                MessageBox.Show("Подключение к базе данных установлено успешно!", "Успешно", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Получаем список видимых таблиц (исключая системные и таблицу Users)
                var visibleTables = _businessLogic.GetVisibleTables();

                // Очищаем выпадающий список перед заполнением
                cmdChooseTable.Items.Clear();

                // Заполняем выпадающий список именами таблиц
                foreach (DataRow row in visibleTables.Rows)
                {
                    string tableName = row["TABLE_NAME"].ToString();
                    cmdChooseTable.Items.Add(tableName);
                }

                // активируем ComboBox и кнопку отчетов
                cmdChooseTable.Enabled = true;
                Reportbutton.Enabled = true;
            }
            // Обработка ошибки подключения
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения: " + ex.Message,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем соединение
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }
        /// <summary>
        /// Обработчик события нажатия пункта меню "Выход"
        /// Закрывает соединение с БД и завершает работу приложения
        /// </summary>

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Закрытие соединения, если оно было открыто
            connection?.Close();
            // Выход из приложения
            Application.Exit();
        }
        /// <summary>
        /// Обработчик события изменения выбранного элемента в ComboBox таблиц
        /// </summary>
        private void cmdChooseTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Вызываем метод загрузки данных при изменении выбора таблицы
            LoadData();
        }
        /// <summary>
        /// Метод загрузки данных
        /// Загружает данные из выбранной таблицы в DataGridView
        /// </summary>
        private void LoadData()
        {
            // Активируем кнопки работы с записями
            addRecordButton.Enabled = true;
            editRecordButton.Enabled = true;
            deleteRecordButton.Enabled = true;

            // Если не выбрана таблица - выходим из метода
            if (cmdChooseTable.SelectedItem == null) return;

            // Получаем имя выбранной таблицы
            string selectedTable = cmdChooseTable.SelectedItem.ToString();

            try
            {
                // Для таблицы "Смены" используем специальный метод с подстановкой значений
                if (selectedTable.Equals("Смены", StringComparison.OrdinalIgnoreCase))
                {
                    LoadShiftsDataWithNames();
                }
                else
                {
                    // Для остальных таблиц стандартная загрузка
                    connection.Open();

                    // Создаем SQL-запрос
                    OleDbCommand command = new OleDbCommand($"SELECT * FROM [{selectedTable}]", connection);

                    // Используем адаптер для заполнения DataTable
                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    // Привязываем DataTable к DataGridView
                    dataGridView1.DataSource = dataTable;
                }
                    
            }
            // Обработчик ошибок
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки данных: " + ex.Message, "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // В любом случае закрываем соединение
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }
        /// <summary>
        /// Специальный метод для загрузки данных таблицы "Смены" с подстановкой значений из связанных таблиц.
        /// </summary>
        private void LoadShiftsDataWithNames()
        {
            // SQL-запрос с JOIN для получения данных из связанных таблиц
            string query = @"SELECT 
                            s.[Код смены], 
                            s.[Дата],
                            s.[ID_подразделения],
                            p.[Подразделение] AS [Подразделение_текст],
                            s.[ID_руководителя],
                            r.[ФИО_руководителя] AS [Руководитель_текст],
                            s.[ID_начальника_смены],
                            n.[ФИО_начальника_смены] AS [Начальник_текст],
                            s.[ID_количества_рабочих],
                            k.[Количество рабочих] AS [Количество_текст],
                            s.[ID_длительности_смены],
                            d.[Длительность смены] AS [Длительность_текст]
                            FROM (((([Смены] s
                            LEFT JOIN [Подразделения] p ON s.[ID_подразделения] = p.[ID_подразделения])
                            LEFT JOIN [Руководители] r ON s.[ID_руководителя] = r.[ID_руководителя])
                            LEFT JOIN [Начальники смен] n ON s.[ID_начальника_смены] = n.[ID_начальника_смены])
                            LEFT JOIN [Количество рабочих] k ON s.[ID_количества_рабочих] = k.[ID_количества_рабочих])
                            LEFT JOIN [Длительности смен] d ON s.[ID_длительности_смены] = d.[ID_длительности_смены]";

            try
            {
                // Открытие соединения с базой данных
                connection.Open();
                // Создание команды с SQL-запросом и привязкой к открытому соединению
                OleDbCommand command = new OleDbCommand(query, connection);
                // Инициализация адаптера данных для выполнения запроса и заполнения таблицы
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                // Создание структуры DataTable для хранения результатов запроса
                DataTable dataTable = new DataTable();
                // Заполнение DataTable данными из результата выполнения SQL-запроса
                adapter.Fill(dataTable);
                // Привязываем данные к DataGridView
                dataGridView1.DataSource = dataTable;

                // Скрываем технические ID-столбцы
                dataGridView1.Columns["ID_подразделения"].Visible = false;
                dataGridView1.Columns["ID_руководителя"].Visible = false;
                dataGridView1.Columns["ID_начальника_смены"].Visible = false;
                dataGridView1.Columns["ID_количества_рабочих"].Visible = false;
                dataGridView1.Columns["ID_длительности_смены"].Visible = false;

                // Переименовываем текстовые столбцы для красивого отображения
                dataGridView1.Columns["Подразделение_текст"].HeaderText = "Подразделение";
                dataGridView1.Columns["Руководитель_текст"].HeaderText = "Руководитель";
                dataGridView1.Columns["Начальник_текст"].HeaderText = "Начальник смены";
                dataGridView1.Columns["Количество_текст"].HeaderText = "Количество рабочих";
                dataGridView1.Columns["Длительность_текст"].HeaderText = "Длительность смены";
            }
            finally
            {
                // В любом случае закрываем соединение
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }
        /// <summary>
        /// Обработчик события загрузки формы
        /// Выполняет начальную настройку элементов управления
        /// </summary>
        private void Form1_Load(object sender, EventArgs e)
        {
            // отключение кнопок
            cmdChooseTable.Enabled = false;
            addRecordButton.Enabled = false;
            editRecordButton.Enabled = false;
            deleteRecordButton.Enabled = false;
            Reportbutton.Enabled = false;

            // Настройка DataGridView
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.ReadOnly = true;
        }
        /// <summary>
        /// Обработчик события нажатия кнопки "Добавить"
        /// </summary>
        private void addRecordButton_Click(object sender, EventArgs e)
        {
            // Проверяем, выбрана ли таблица в выпадающем списке
            if (cmdChooseTable.SelectedItem == null)
            {
                MessageBox.Show("Выберите таблицу", "Выберите таблицу",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            // Создаем форму для добавления записи, передавая выбранную таблицу и экземпляр BusinessLogic
            var addForm = new AddRecord(cmdChooseTable.SelectedItem.ToString(), _businessLogic);
            if (addForm.ShowDialog() == DialogResult.OK)
            {
                RefreshData(); // Обновляем DataGridView
            }
        }
        /// <summary>
        /// Метод для обновления данных в DataGridView
        /// </summary>
        private void RefreshData()
        {
            if (cmdChooseTable.SelectedItem == null) return;

            try
            {
                // Получаем имя текущей таблицы
                string tableName = cmdChooseTable.SelectedItem.ToString();
                // Для таблицы "Смены" используем специальный метод с JOIN
                if (tableName.Equals("Смены", StringComparison.OrdinalIgnoreCase))
                {
                    LoadShiftsDataWithNames();
                }
                else
                {
                    // Для остальных таблиц стандартная загрузка
                    dataGridView1.DataSource = _businessLogic.GetTableData(tableName);
                }
                ;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка обновления: {ex.Message}");
            }
        }
        /// <summary>
        /// Обработчик события нажатия кнопки "Редактировать"
        /// Открывает форму для редактирования выбранной записи.
        /// </summary>
        private void editRecordButton_Click(object sender, EventArgs e)
        {
            // Проверяем, выбрана ли строка в DataGridView
            if (dataGridView1.SelectedRows.Count == 0) return;
            // Получаем выбранную строку
            var selectedRow = dataGridView1.SelectedRows[0];
            string tableName = cmdChooseTable.SelectedItem.ToString();

            // Получаем данные выбранной строки
            var values = new Dictionary<string, object>();
            foreach (DataGridViewCell cell in selectedRow.Cells)
            {
                string columnName = dataGridView1.Columns[cell.ColumnIndex].Name;

                // Если это текстовый столбец, пропускаем - нас интересуют только ID
                if (columnName.EndsWith("_текст")) continue;

                values[columnName] = cell.Value;
            }

            // Создаем форму редактирования
            var editForm = new EditRecord(tableName, _businessLogic, values);
            // Если форма закрыта с OK - обновляем данные
            if (editForm.ShowDialog() == DialogResult.OK)
            {
                RefreshData();
            }
        }
        /// <summary>
        /// Обработчик события нажатия кнопки "Удалить"
        /// Удаляет выбранную запись с подтверждением.
        /// </summary>
        private void deleteRecordButton_Click(object sender, EventArgs e)
        {
            // Проверяем, выбрана ли строка в DataGridView
            if (dataGridView1.SelectedRows.Count == 0) return;
            // Получаем выбранную строку
            var selectedRow = dataGridView1.SelectedRows[0];
            string tableName = cmdChooseTable.SelectedItem.ToString();
            // Получаем имя столбца идентификатора для текущей таблицы
            string idColumnName = _businessLogic.GetIdColumnName(tableName);
            // Получаем значение идентификатора из выбранной строки
            object idValue = selectedRow.Cells[idColumnName].Value;

            // Запрашиваем подтверждение удаления
            if (MessageBox.Show($"Вы уверены, что хотите удалить эту запись?",
                               "Подтверждение удаления",
                               MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                try
                {
                    // Вызываем удаление через бизнес-логику
                    _businessLogic.DeleteRecord(tableName, idColumnName, idValue);
                    // Обновляем данные
                    RefreshData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении: {ex.Message}");
                }
            }
        }
        /// <summary>
        /// Обработчик нажатия кнопки "Отчеты".
        /// Открывает форму для формирования отчетов.
        /// </summary>

        private void Reportbutton_Click(object sender, EventArgs e)
        {
            Reports reportsForm = new Reports(_businessLogic);
            reportsForm.Show();
        }
    }
}
