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
    // Класс главной формы приложения
    public partial class Form1 : Form
    {
        // Поля класса для хранения состояния приложения
        private string selectedDatabasePath;  // Путь к выбранной базе данных
        private OleDbConnection connection;   // Подключение к базе данных
        private DataTable tablesSchema;      // Схема таблиц базы данных
        public Form1()
        {
            InitializeComponent(); // Инициализация компонентов формы
        }
        /// <summary>
        /// Обработчик нажатия пункта меню "Подключиться к БД"
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Аргументы события</param>
        private void подключитьсяКБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Путь к папке с БД
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
                InitialDirectory = databaseFolder,
                Title = "Выберите файл базы данных Access",
                Filter = "Базы данных Access (*.accdb; *.mdb)|*.accdb;*.mdb|Все файлы (*.*)|*.*",
                RestoreDirectory = true
            };

            // Показываем диалог и обрабатываем результат
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Сохраняем путь к БД
                selectedDatabasePath = openFileDialog.FileName;

                // Вызов метода для установки соединения
                подключитьсяКБДToolStripMenuItem_Click_AfterSelection();
            }
        }
        /// <summary>
        /// Метод для установки соединения
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

            //Выбираем провайдера
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

            // Вызов метода загрузки списка таблиц
            LoadTableNames();
        }
        /// <summary>
        /// Метод загрузки списка таблиц
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

                // Схема таблиц из БД
                tablesSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });

                // Заполнение ComboBox именами таблиц
                foreach (DataRow row in tablesSchema.Rows)
                {
                    string tableName = row["TABLE_NAME"].ToString();
                    cmdChooseTable.Items.Add(tableName);
                }

                // активируем ComboBox
                cmdChooseTable.Enabled = true;
            }
            // Обработка ошибки подключения
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            // Если не выбрана таблица - выходим из метода
            if (cmdChooseTable.SelectedItem == null) return;

            // Получаем имя выбранной таблицы
            string selectedTable = cmdChooseTable.SelectedItem.ToString();

            try
            {
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
            // Обработчик ошибок
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрытие соединения
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
            // отключение ComboBox
            cmdChooseTable.Enabled = false;

            // Настройка DataGridView
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.ReadOnly = true;
        }
    }
}
