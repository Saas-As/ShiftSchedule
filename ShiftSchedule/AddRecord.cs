using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShiftSchedule
{
    /// <summary>
    /// Форма для добавления новых записей в базу данных.
    /// Обеспечивает:
    /// - Динамическое создание полей ввода на основе структуры таблицы
    /// - Автоматическую генерацию ID для новой записи
    /// - Валидацию вводимых данных
    /// - Специальную обработку для связанных таблиц (combobox)
    /// - Визуальное выделение обязательных полей
    /// </summary>
    public partial class AddRecord : Form
    {
        // Название таблицы, в которую добавляется запись
        private readonly string _tableName;

        // Объект бизнес-логики для работы с данными
        private readonly BusinessLogic _businessLogic;

        // Список всех элементов управления для ввода данных
        private readonly List<Control> _inputControls = new List<Control>();

        // Кнопка сохранения записи
        private Button _btnSave;

        // Кнопка отмены
        private Button _btnCancel;

        // Элемент управления для отображения ID
        private NumericUpDown _idControl;

        // Название столбца с первичным ключом
        private readonly string _idColumnName;

        /// <summary>
        /// Конструктор формы добавления записи.
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="businessLogic">Объект бизнес-логики</param>
        /// <exception cref="ArgumentNullException">Если параметры не указаны</exception>
        public AddRecord(string tableName, BusinessLogic businessLogic)
        {
            // Проверка входных параметров
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentNullException(nameof(tableName));
            if (businessLogic == null)
                throw new ArgumentNullException(nameof(businessLogic));

            // Сохраняем переданные параметры
            _tableName = tableName;
            _businessLogic = businessLogic;

            // Получаем название столбца с ID для этой таблицы
            _idColumnName = _businessLogic.GetIdColumnName(_tableName);

            // Инициализация компонентов формы
            InitializeComponent();

            // Настройка формы
            InitializeForm();
        }

        /// <summary>
        /// Настраивает основные параметры формы.
        /// </summary>
        private void InitializeForm()
        {
            // Устанавливаем заголовок формы с названием таблицы
            this.Text = $"Добавление записи в {_tableName}";

            // Позиционируем форму по центру родительского окна
            this.StartPosition = FormStartPosition.CenterParent;

            // Фиксируем размер формы
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // Запрещаем разворачивание формы на весь экран
            this.MaximizeBox = false;

            // Добавляем отступы от краев формы
            this.Padding = new Padding(20);

            // Создаем кнопки формы
            CreateButtons();

            // Загружаем поля таблицы в форму
            LoadTableFields();

            // Устанавливаем минимальный размер формы
            this.MinimumSize = new Size(450, 300);
        }

        /// <summary>
        /// Создает кнопки "Сохранить" и "Отмена".
        /// </summary>
        private void CreateButtons()
        {
            // Кнопка сохранения
            _btnSave = new Button
            {
                Text = "Сохранить",
                DialogResult = DialogResult.OK,
                Size = new Size(100, 40),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(this.ClientSize.Width - 220, this.ClientSize.Height - 70)
            };
            _btnSave.Click += BtnSave_Click;
            this.Controls.Add(_btnSave); // Добавляем кнопку на форму

            // Кнопка отмены
            _btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Size = new Size(100, 40),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(this.ClientSize.Width - 110, this.ClientSize.Height - 70)
            };
            _btnCancel.Click += (s, e) => this.Close(); // Закрываем форму при клике
            this.Controls.Add(_btnCancel); // Добавляем кнопку на форму
        }

        /// <summary>
        /// Загружает поля таблицы и создает соответствующие элементы управления.
        /// </summary>
        private void LoadTableFields()
        {
            int yPos = 20; // Начальная позиция по вертикали

            // Получаем схему таблицы из бизнес-логики
            var schema = _businessLogic.GetTableSchema(_tableName);

            // Создаем панель для группировки элементов с возможностью прокрутки
            var panel = new Panel
            {
                Location = new Point(20, 20),
                Width = this.ClientSize.Width - 40,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoScroll = true
            };
            this.Controls.Add(panel); // Добавляем панель на форму

            // Создаем элементы управления для каждого столбца таблицы
            foreach (DataRow column in schema.Rows)
            {
                string colName = column["COLUMN_NAME"].ToString();

                // Проверяем, является ли поле обязательным
                bool isRequired = _businessLogic.IsFieldRequired(_tableName, colName);

                // Создаем группу для поля (GroupBox)
                var groupBox = new GroupBox
                {
                    Text = isRequired ? $"{colName}*" : colName, // Добавляем * для обязательных полей
                    Location = new Point(0, yPos),
                    Width = panel.Width - 20,
                    Height = 60,
                    Font = new Font(this.Font, FontStyle.Regular)
                };

                Control inputControl;

                // Особый случай для поля ID
                if (colName.Equals(_idColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    // Создаем NumericUpDown для отображения ID
                    _idControl = new NumericUpDown
                    {
                        Tag = colName, // Сохраняем название столбца
                        Name = colName,
                        Location = new Point(10, 25),
                        Width = groupBox.Width - 20,
                        Minimum = 0,
                        Maximum = int.MaxValue,
                        Value = _businessLogic.GetNextId(_tableName), // Получаем следующий ID
                        ReadOnly = true, // Запрещаем редактирование ID
                        BackColor = SystemColors.Control, // Серый фон для наглядности
                        Enabled = false // Отключаем возможность изменения
                    };
                    inputControl = _idControl;
                }
                else
                {
                    // Создаем элемент управления в зависимости от типа данных
                    inputControl = CreateInputControl(column, groupBox);
                }

                // Добавляем элемент управления в группу
                groupBox.Controls.Add(inputControl);

                // Добавляем группу на панель
                panel.Controls.Add(groupBox);

                // Сохраняем элемент управления в общий список
                _inputControls.Add(inputControl);

                // Увеличиваем позицию для следующего элемента
                yPos += groupBox.Height + 10;
            }

            // Рассчитываем общую высоту всех элементов
            int totalHeight = yPos + 40; // Добавляем место для кнопок

            // Устанавливаем высоту панели
            panel.Height = totalHeight;

            // Настраиваем размер формы
            this.ClientSize = new Size(this.ClientSize.Width, totalHeight + 100);

            // Позиционируем кнопки внизу формы
            _btnSave.Top = this.ClientSize.Height - 60;
            _btnCancel.Top = this.ClientSize.Height - 60;
        }

        /// <summary>
        /// Создает элемент управления для ввода значения поля.
        /// </summary>
        /// <param name="column">Информация о столбце</param>
        /// <param name="parent">Родительский элемент</param>
        /// <returns>Созданный элемент управления</returns>
        private Control CreateInputControl(DataRow column, GroupBox parent)
        {
            string colName = column["COLUMN_NAME"].ToString();
            var dataType = (OleDbType)column["DATA_TYPE"];

            // Специальная обработка для таблицы "Смены"
            if (_tableName.Equals("Смены", StringComparison.OrdinalIgnoreCase))
            {
                // Для связанных полей создаем выпадающие списки
                switch (colName)
                {
                    case "ID_подразделения":
                        return CreateComboBoxControl(colName, parent,
                            "Подразделения", "ID_подразделения", "Подразделение");
                    case "ID_руководителя":
                        return CreateComboBoxControl(colName, parent,
                            "Руководители", "ID_руководителя", "ФИО_руководителя");
                    case "ID_количества_рабочих":
                        return CreateComboBoxControl(colName, parent,
                            "Количество рабочих", "ID_количества_рабочих", "Количество рабочих");
                    case "ID_длительности_смены":
                        return CreateComboBoxControl(colName, parent,
                            "Длительности смен", "ID_длительности_смены", "Длительность смены");
                    case "ID_начальника_смены":
                        return CreateComboBoxControl(colName, parent,
                            "Начальники смен", "ID_начальника_смены", "ФИО_начальника_смены");
                }
            }

            // Специальная обработка для таблицы "Длительности смен"
            if (_tableName.Equals("Длительности смен", StringComparison.OrdinalIgnoreCase))
            {
                switch (colName)
                {
                    case "Начало смены":
                    case "Окончание смены":
                        // Для времени используем DateTimePicker с форматом времени
                        return new DateTimePicker
                        {
                            Tag = colName,
                            Name = colName,
                            Location = new Point(10, 25),
                            Width = parent.Width - 20,
                            Format = DateTimePickerFormat.Custom,
                            CustomFormat = "HH:mm", // Формат часов:минут
                            ShowUpDown = true, // Используем счетчик вместо календаря
                            Value = DateTime.Today.Date.AddHours(8) // Значение по умолчанию - 8:00
                        };
                    case "Длительность смены":
                        // Для длительности используем NumericUpDown
                        return new NumericUpDown
                        {
                            Tag = colName,
                            Name = colName,
                            Location = new Point(10, 25),
                            Width = parent.Width - 20,
                            Minimum = 0,
                            Maximum = 24 // Максимальная длительность смены - 24 часа
                        };
                }
            }

            // Создаем элементы управления в зависимости от типа данных
            return dataType switch
            {
                OleDbType.Integer => new NumericUpDown
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Minimum = int.MinValue,
                    Maximum = int.MaxValue
                },
                OleDbType.Decimal => new NumericUpDown
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    DecimalPlaces = 2, // Два знака после запятой
                    Minimum = decimal.MinValue,
                    Maximum = decimal.MaxValue
                },
                OleDbType.Date => new DateTimePicker
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Format = DateTimePickerFormat.Short, // Короткий формат даты
                    Value = DateTime.Today, // Текущая дата по умолчанию
                    MinDate = new DateTime(2000, 1, 1), // Минимальная дата
                    MaxDate = new DateTime(2100, 1, 1)  // Максимальная дата
                },
                OleDbType.Boolean => new CheckBox
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Text = ""
                },
                _ => new TextBox // Для всех остальных типов - обычное текстовое поле
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20
                }
            };
        }

        /// <summary>
        /// Создает выпадающий список (ComboBox) для связанных таблиц.
        /// </summary>
        /// <param name="colName">Название столбца</param>
        /// <param name="parent">Родительский элемент</param>
        /// <param name="lookupTable">Таблица-справочник</param>
        /// <param name="idColumn">Столбец с ID в справочнике</param>
        /// <param name="nameColumn">Столбец с названием в справочнике</param>
        /// <returns>Настроенный ComboBox</returns>
        private ComboBox CreateComboBoxControl(string colName, GroupBox parent,
                                             string lookupTable, string idColumn, string nameColumn)
        {
            // Создаем ComboBox
            var combo = new ComboBox
            {
                Tag = colName,
                Name = colName,
                Location = new Point(10, 25),
                Width = parent.Width - 20,
                DropDownStyle = ComboBoxStyle.DropDownList // Запрещаем ручной ввод
            };

            try
            {
                // Получаем данные для заполнения списка
                var lookupData = _businessLogic.GetLookupData(lookupTable, idColumn, nameColumn);

                // Заполняем ComboBox данными
                foreach (var item in lookupData)
                {
                    // Используем KeyValuePair для хранения ID и названия
                    combo.Items.Add(new KeyValuePair<int, string>(item.Key, item.Value));
                }

                // Настраиваем отображение
                combo.DisplayMember = "Value"; // Показываем название
                combo.ValueMember = "Key";     // Но используем ID как значение

                // Выбираем первый элемент, если список не пустой
                if (combo.Items.Count > 0)
                {
                    combo.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                // В случае ошибки выводим сообщение
                MessageBox.Show($"Ошибка загрузки данных для {colName}: {ex.Message}",
                              "Ошибка",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning);
            }

            return combo;
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Сохранить".
        /// Выполняет валидацию данных и сохраняет новую запись.
        /// </summary>
        private void BtnSave_Click(object sender, EventArgs e)
        {
            // Словарь для значений полей
            var values = new Dictionary<string, object>();

            // Список незаполненных обязательных полей
            var missingFields = new List<string>();

            // Фильтруем только элементы управления для ввода данных
            var inputFields = _inputControls.Where(c =>
                c is TextBox ||
                c is NumericUpDown ||
                c is DateTimePicker ||
                c is CheckBox ||
                c is ComboBox).ToList();

            // Проверяем все элементы управления
            foreach (var control in inputFields)
            {
                string fieldName = control.Tag?.ToString();
                if (string.IsNullOrEmpty(fieldName)) continue;

                // Проверяем, является ли поле обязательным
                bool isRequired = _businessLogic.IsFieldRequired(_tableName, fieldName);

                // Получаем значение из элемента управления
                object value = GetControlValue(control);

                // Проверка обязательных полей
                if (isRequired && IsValueEmpty(value))
                {
                    missingFields.Add(fieldName);
                    control.BackColor = Color.LightPink; // Подсвечиваем поле с ошибкой
                }
                else
                {
                    // Сохраняем значение
                    values.Add(fieldName, value);

                    // Сбрасываем подсветку
                    control.BackColor = SystemColors.Window;
                }
            }

            // Если есть незаполненные обязательные поля - показываем сообщение
            if (missingFields.Count > 0)
            {
                MessageBox.Show($"Заполните обязательные поля:\n{string.Join("\n", missingFields)}",
                              "Ошибка",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Пытаемся добавить запись
                _businessLogic.InsertRecord(_tableName, values);

                // Закрываем форму с результатом OK
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                // В случае ошибки выводим сообщение
                MessageBox.Show($"Ошибка сохранения: {ex.Message}",
                              "Ошибка",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Получает значение из элемента управления.
        /// </summary>
        /// <param name="control">Элемент управления</param>
        /// <returns>Значение в нужном формате</returns>
        private object GetControlValue(Control control)
        {
            switch (control)
            {
                case ComboBox cb:
                    if (cb.SelectedItem == null) return null;

                    // Для ComboBox с KeyValuePair возвращаем ключ (ID)
                    if (cb.SelectedItem is KeyValuePair<int, string> kvp)
                        return kvp.Key;

                    // Для обычных ComboBox возвращаем выбранное значение
                    return cb.SelectedValue ?? cb.SelectedItem;

                case TextBox tb:
                    return tb.Text;

                case NumericUpDown num:
                    return Convert.ToInt32(num.Value);

                case DateTimePicker dt:
                    // Специальная обработка для времени в "Длительности смен"
                    if (_tableName.Equals("Длительности смен", StringComparison.OrdinalIgnoreCase) &&
                        (dt.Tag.ToString() == "Начало смены" || dt.Tag.ToString() == "Окончание смены"))
                    {
                        // Возвращаем DateTime с фиксированной датой и выбранным временем
                        return new DateTime(1999, 12, 30, dt.Value.Hour, dt.Value.Minute, 0);
                    }
                    return dt.Value;

                case CheckBox cb:
                    return cb.Checked;

                default:
                    return null;
            }
        }

        /// <summary>
        /// Проверяет, является ли значение пустым.
        /// </summary>
        /// <param name="value">Значение для проверки</param>
        /// <returns>True, если значение пустое</returns>
        private bool IsValueEmpty(object value)
        {
            if (value == null || value == DBNull.Value)
                return true;

            if (value is TimeSpan timeSpan)
                return timeSpan == TimeSpan.Zero;

            // Проверка для разных типов данных
            return value switch
            {
                string s => string.IsNullOrWhiteSpace(s), // Пустая строка
                decimal d => d == 0,                     // Ноль
                DateTime dt => dt == DateTime.MinValue,  // Минимальная дата
                _ => false
            };
        }
    }
}