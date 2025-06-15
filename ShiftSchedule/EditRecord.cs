using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ShiftSchedule
{
    /// <summary>
    /// Форма для редактирования существующих записей в базе данных.
    /// Обеспечивает:
    /// - Загрузку текущих значений записи
    /// - Валидацию вводимых данных
    /// - Специальную обработку для связанных таблиц (combobox)
    /// - Отображение обязательных полей
    /// - Сохранение изменений в базу данных
    /// </summary>
    public partial class EditRecord : Form
    {
        // Название таблицы, в которой редактируется запись
        private readonly string _tableName;

        // Объект бизнес-логики для работы с данными
        private readonly BusinessLogic _businessLogic;

        // Исходные значения редактируемой записи
        private readonly Dictionary<string, object> _originalValues;

        // Список всех элементов управления для ввода данных
        private readonly List<Control> _inputControls = new List<Control>();

        // Кнопка сохранения изменений
        private Button _btnSave;

        // Кнопка отмены изменений
        private Button _btnCancel;

        // Элемент управления для отображения ID (если есть)
        private NumericUpDown _idControl;

        // Название столбца с первичным ключом
        private readonly string _idColumnName;

        /// <summary>
        /// Конструктор формы редактирования.
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="businessLogic">Объект бизнес-логики</param>
        /// <param name="originalValues">Текущие значения записи</param>
        public EditRecord(string tableName, BusinessLogic businessLogic, Dictionary<string, object> originalValues)
        {
            // Проверка входных параметров
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentNullException(nameof(tableName));
            if (businessLogic == null)
                throw new ArgumentNullException(nameof(businessLogic));
            if (originalValues == null)
                throw new ArgumentNullException(nameof(originalValues));

            // Сохраняем переданные параметры
            _tableName = tableName;
            _businessLogic = businessLogic;
            _originalValues = originalValues;

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
            this.Text = $"Редактирование записи в {_tableName}";

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
        /// Создает кнопки "Сохранить изменения" и "Отмена".
        /// </summary>
        private void CreateButtons()
        {
            // Кнопка сохранения изменений
            _btnSave = new Button
            {
                Text = "Сохранить изменения",
                DialogResult = DialogResult.OK,
                Size = new Size(150, 40),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(this.ClientSize.Width - 270, this.ClientSize.Height - 70)
            };
            _btnSave.Click += BtnSave_Click;
            this.Controls.Add(_btnSave); // Добавляем кнопку на форму

            // Кнопка отмены изменений
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
                        Value = _originalValues.ContainsKey(colName) ? Convert.ToDecimal(_originalValues[colName]) : 0,
                        ReadOnly = true, // Запрещаем редактирование ID
                        BackColor = SystemColors.Control, // Серый фон для наглядности
                        Enabled = false // Отключаем возможность изменения
                    };
                    inputControl = _idControl;
                }
                else
                {
                    // Получаем текущее значение поля (если есть)
                    object initialValue = _originalValues.ContainsKey(colName) ? _originalValues[colName] : null;

                    // Создаем элемент управления в зависимости от типа данных
                    inputControl = CreateInputControl(column, groupBox, initialValue);
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
        /// Создает элемент управления для редактирования значения поля.
        /// </summary>
        /// <param name="column">Информация о столбце</param>
        /// <param name="parent">Родительский элемент</param>
        /// <param name="initialValue">Текущее значение</param>
        /// <returns>Созданный элемент управления</returns>
        private Control CreateInputControl(DataRow column, GroupBox parent, object initialValue)
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
                        return CreateComboBoxControl(colName, parent, initialValue,
                            "Подразделения", "ID_подразделения", "Подразделение");
                    case "ID_руководителя":
                        return CreateComboBoxControl(colName, parent, initialValue,
                            "Руководители", "ID_руководителя", "ФИО_руководителя");
                    case "ID_количества_рабочих":
                        return CreateComboBoxControl(colName, parent, initialValue,
                            "Количество рабочих", "ID_количества_рабочих", "Количество рабочих");
                    case "ID_длительности_смены":
                        return CreateComboBoxControl(colName, parent, initialValue,
                            "Длительности смен", "ID_длительности_смены", "Длительность смены");
                    case "ID_начальника_смены":
                        return CreateComboBoxControl(colName, parent, initialValue,
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
                        var dtPicker = new DateTimePicker
                        {
                            Tag = colName,
                            Name = colName,
                            Location = new Point(10, 25),
                            Width = parent.Width - 20,
                            Format = DateTimePickerFormat.Custom,
                            CustomFormat = "HH:mm",
                            ShowUpDown = true
                        };

                        // Устанавливаем начальное значение из originalValues
                        if (initialValue != null && initialValue != DBNull.Value)
                        {
                            dtPicker.Value = Convert.ToDateTime(initialValue);
                        }
                        else
                        {
                            dtPicker.Value = DateTime.Today.Date.AddHours(8); // Значение по умолчанию
                        }

                        return dtPicker;
                    case "Длительность смены":
                        var numericUpDown = new NumericUpDown
                        {
                            Tag = colName,
                            Name = colName,
                            Location = new Point(10, 25),
                            Width = parent.Width - 20,
                            Minimum = 0,
                            Maximum = 24,
                            DecimalPlaces = 0
                        };

                        // Устанавливаем начальное значение из originalValues
                        if (initialValue != null && initialValue != DBNull.Value)
                        {
                            numericUpDown.Value = Convert.ToDecimal(initialValue);
                        }
                        else
                        {
                            numericUpDown.Value = 8; // Значение по умолчанию - 8 часов
                        }

                        return numericUpDown;
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
                    Maximum = int.MaxValue,
                    Value = initialValue != null && initialValue != DBNull.Value ?
                        Convert.ToDecimal(initialValue) : 0
                },
                OleDbType.Decimal => new NumericUpDown
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    DecimalPlaces = 2, // Два знака после запятой
                    Minimum = decimal.MinValue,
                    Maximum = decimal.MaxValue,
                    Value = initialValue != null && initialValue != DBNull.Value ?
                        Convert.ToDecimal(initialValue) : 0
                },
                OleDbType.Date => new DateTimePicker
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Format = DateTimePickerFormat.Short, // Короткий формат даты
                    Value = initialValue != null && initialValue != DBNull.Value ?
                        Convert.ToDateTime(initialValue) : DateTime.Today,
                    MinDate = new DateTime(2000, 1, 1), // Минимальная дата
                    MaxDate = new DateTime(2100, 1, 1)  // Максимальная дата
                },
                OleDbType.Boolean => new CheckBox
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Checked = initialValue != null && initialValue != DBNull.Value &&
                        Convert.ToBoolean(initialValue)
                },
                _ => new TextBox // Для всех остальных типов - обычное текстовое поле
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Text = initialValue != null && initialValue != DBNull.Value ?
                        initialValue.ToString() : ""
                }
            };
        }

        /// <summary>
        /// Создает выпадающий список (ComboBox) для связанных таблиц.
        /// </summary>
        /// <param name="colName">Название столбца</param>
        /// <param name="parent">Родительский элемент</param>
        /// <param name="initialValue">Текущее значение</param>
        /// <param name="lookupTable">Таблица-справочник</param>
        /// <param name="idColumn">Столбец с ID в справочнике</param>
        /// <param name="nameColumn">Столбец с названием в справочнике</param>
        /// <returns>Настроенный ComboBox</returns>
        private ComboBox CreateComboBoxControl(string colName, GroupBox parent, object initialValue,
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

                // Устанавливаем выбранное значение (если оно есть)
                if (initialValue != null && initialValue != DBNull.Value)
                {
                    int initialId = Convert.ToInt32(initialValue);

                    // Ищем элемент с нужным ID
                    for (int i = 0; i < combo.Items.Count; i++)
                    {
                        var item = (KeyValuePair<int, string>)combo.Items[i];
                        if (item.Key == initialId)
                        {
                            combo.SelectedIndex = i;
                            break;
                        }
                    }
                }
                else if (combo.Items.Count > 0)
                {
                    // Если значение не задано - выбираем первый элемент
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
        /// Обработчик нажатия кнопки "Сохранить изменения".
        /// </summary>
        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // Словарь для новых значений
                var values = new Dictionary<string, object>();

                // Списки для ошибок
                var missingFields = new List<string>();
                var invalidFields = new List<string>();

                // Проверяем все элементы управления
                foreach (var control in _inputControls)
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
                        continue;
                    }

                    // Дополнительная проверка для числовых полей
                    if (!IsValueEmpty(value))
                    {
                        if (value is int intValue && intValue < 0)
                        {
                            invalidFields.Add($"{fieldName} (не может быть отрицательным)");
                            control.BackColor = Color.LightPink;
                            continue;
                        }
                    }

                    // Сохраняем значение (или DBNull для пустых значений)
                    values[fieldName] = value ?? DBNull.Value;

                    // Сбрасываем подсветку
                    control.BackColor = SystemColors.Window;
                }

                // Если есть ошибки - показываем сообщение
                if (missingFields.Count > 0 || invalidFields.Count > 0)
                {
                    var errorMessage = "";
                    if (missingFields.Count > 0)
                        errorMessage += $"Не заполнены обязательные поля:\n{string.Join(", ", missingFields)}\n\n";
                    if (invalidFields.Count > 0)
                        errorMessage += $"Некорректные значения в полях:\n{string.Join(", ", invalidFields)}";

                    MessageBox.Show(errorMessage, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Добавляем ID записи для обновления
                values[_idColumnName] = _originalValues[_idColumnName];

                // Пытаемся обновить запись
                bool success = _businessLogic.UpdateRecord(_tableName, values, _idColumnName);

                if (success)
                {
                    // Если успешно - показываем сообщение и закрываем форму
                    MessageBox.Show("Изменения успешно сохранены!", "Успех",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Не удалось сохранить изменения", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                // В случае ошибки выводим подробную информацию
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}\n\n{ex.StackTrace}",
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
                int i => i == 0,                         // Ноль
                decimal d => d == 0,                     // Ноль
                DateTime dt => dt == DateTime.MinValue,  // Минимальная дата
                bool _ => false,                         // Для CheckBox всегда заполнено
                _ => false                               // Для остальных типов считаем заполненным
            };
        }
    }
}