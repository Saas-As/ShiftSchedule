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
    public partial class EditRecord : Form
    {
        // Поле для хранения имени таблицы, в которой редактируется запись
        private readonly string _tableName;
        // Поле для хранения экземпляра класса BusinessLogic
        private readonly BusinessLogic _businessLogic;
        // Словарь с исходными значениями редактируемой записи
        private readonly Dictionary<string, object> _originalValues;
        // Список элементов управления для ввода данных
        private readonly List<Control> _inputControls = new List<Control>();
        // Кнопки для сохранения и отмены изменений
        private Button _btnSave;
        private Button _btnCancel;
        // Элемент управления для ввода ID (если необходимо)
        private NumericUpDown _idControl;
        // Имя столбца идентификатора
        private readonly string _idColumnName;

        /// <summary>
        /// Конструктор класса EditRecord.
        /// Инициализирует форму для редактирования записи в указанной таблице.
        /// </summary>
        /// <param name="tableName">Название таблицы, в которой редактируется запись.</param>
        /// <param name="businessLogic">Экземпляр класса BusinessLogic для взаимодействия с данными.</param>
        /// <param name="originalValues">Словарь с исходными значениями редактируемой записи.</param>
        /// <exception cref="ArgumentNullException">Выбрасывается, если tableName, businessLogic или originalValues равны null.</exception>
        public EditRecord(string tableName, BusinessLogic businessLogic, Dictionary<string, object> originalValues)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentNullException(nameof(tableName));
            if (businessLogic == null)
                throw new ArgumentNullException(nameof(businessLogic));
            if (originalValues == null)
                throw new ArgumentNullException(nameof(originalValues));

            _tableName = tableName;
            _businessLogic = businessLogic;
            _originalValues = originalValues;
            _idColumnName = _businessLogic.GetIdColumnName(_tableName);

            InitializeComponent();
            InitializeForm();
        }

        /// <summary>
        /// Метод для инициализации формы.
        /// Устанавливает основные настройки формы, создает элементы управления и загружает поля таблицы.
        /// </summary>
        private void InitializeForm()
        {
            // Основные настройки формы
            this.Text = $"Редактирование записи в {_tableName}";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Padding = new Padding(20);

            // Создание кнопок
            CreateButtons();

            // Загрузка полей таблицы
            LoadTableFields();

            // Устанавливаем минимальный размер формы
            this.MinimumSize = new Size(450, 300);
        }

        /// <summary>
        /// Метод для создания кнопок "Сохранить изменения" и "Отмена".
        /// </summary>
        private void CreateButtons()
        {
            // Кнопка сохранения
            _btnSave = new Button
            {
                Text = "Сохранить изменения",
                DialogResult = DialogResult.OK,
                Size = new Size(150, 40),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(this.ClientSize.Width - 270, this.ClientSize.Height - 70)
            };
            _btnSave.Click += BtnSave_Click;
            this.Controls.Add(_btnSave);

            // Кнопка отмены
            _btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Size = new Size(100, 40),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(this.ClientSize.Width - 110, this.ClientSize.Height - 70)
            };
            _btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(_btnCancel);
        }

        /// <summary>
        /// Метод для загрузки полей таблицы и создания соответствующих элементов управления для ввода данных.
        /// </summary>
        private void LoadTableFields()
        {
            int yPos = 20;
            var schema = _businessLogic.GetTableSchema(_tableName);

            // Создаем Panel для группировки элементов
            var panel = new Panel
            {
                Location = new Point(20, 20),
                Width = this.ClientSize.Width - 40,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoScroll = true
            };
            this.Controls.Add(panel);

            foreach (DataRow column in schema.Rows)
            {
                string colName = column["COLUMN_NAME"].ToString();
                bool isRequired = _businessLogic.IsFieldRequired(_tableName, colName);

                // Создаем группу для каждого поля
                var groupBox = new GroupBox
                {
                    Text = isRequired ? $"{colName}*" : colName,
                    Location = new Point(0, yPos),
                    Width = panel.Width - 20,
                    Height = 60,
                    Font = new Font(this.Font, FontStyle.Regular)
                };

                Control inputControl;

                // Для ID-поля
                if (colName.Equals(_idColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    _idControl = new NumericUpDown
                    {
                        Tag = colName,
                        Name = colName,
                        Location = new Point(10, 25),
                        Width = groupBox.Width - 20,
                        Minimum = 0,
                        Maximum = int.MaxValue,
                        Value = _originalValues.ContainsKey(colName) ? Convert.ToDecimal(_originalValues[colName]) : 0,
                        ReadOnly = true,
                        BackColor = SystemColors.Control,
                        Enabled = false
                    };
                    inputControl = _idControl;
                }
                else
                {
                    // Обычные поля
                    object initialValue = _originalValues.ContainsKey(colName) ? _originalValues[colName] : null;
                    inputControl = CreateInputControl(column, groupBox, initialValue);
                }

                groupBox.Controls.Add(inputControl);
                panel.Controls.Add(groupBox);
                _inputControls.Add(inputControl);

                yPos += groupBox.Height + 10;
            }

            // Вычисляем общую высоту панели
            int totalHeight = yPos + 40;

            // Устанавливаем высоту панели
            panel.Height = totalHeight;

            // Устанавливаем высоту формы
            this.ClientSize = new Size(this.ClientSize.Width, totalHeight + 100);

            // Позиционирование кнопок
            _btnSave.Top = this.ClientSize.Height - 60;
            _btnCancel.Top = this.ClientSize.Height - 60;
        }

        /// <summary>
        /// Метод для создания элементов управления для ввода данных в зависимости от типа данных поля.
        /// </summary>
        /// <param name="column">Строка с информацией о поле таблицы.</param>
        /// <param name="parent">Родительский элемент, в который добавляется элемент управления.</param>
        /// <param name="initialValue">Исходное значение поля.</param>
        /// <returns>Элемент управления для ввода данных.</returns>
        private Control CreateInputControl(DataRow column, GroupBox parent, object initialValue)
        {
            string colName = column["COLUMN_NAME"].ToString();
            var dataType = (OleDbType)column["DATA_TYPE"];

            if (_tableName.Equals("Смены", StringComparison.OrdinalIgnoreCase))
            {
                switch (colName)
                {
                    case "ID_подразделения":
                        return CreateComboBoxControl(colName, parent, initialValue, "Подразделения", "ID_подразделения", "Подразделение");
                    case "ID_руководителя":
                        return CreateComboBoxControl(colName, parent, initialValue, "Руководители", "ID_руководителя", "ФИО_руководителя");
                    case "ID_количества_рабочих":
                        return CreateComboBoxControl(colName, parent, initialValue, "Количество рабочих", "ID_количества_рабочих", "Количество рабочих");
                    case "ID_длительности_смены":
                        return CreateComboBoxControl(colName, parent, initialValue, "Длительности смен", "ID_длительности_смены", "Длительность смены");
                    case "ID_начальника_смены":
                        return CreateComboBoxControl(colName, parent, initialValue, "Начальники смен", "ID_начальника_смены", "ФИО_начальника_смены");
                }
            }
            // Специальная обработка для таблицы "Длительности смен"
            if (_tableName.Equals("Длительности смен", StringComparison.OrdinalIgnoreCase))
            {
                switch (colName)
                {
                    case "Начало смены":
                    case "Окончание смены":
                        return new DateTimePicker
                        {
                            Tag = colName,
                            Name = colName,
                            Location = new Point(10, 25),
                            Width = parent.Width - 20,
                            Format = DateTimePickerFormat.Custom,
                            CustomFormat = "HH:mm",
                            ShowUpDown = true,
                            Value = DateTime.Today.Date.AddHours(8) // Устанавливаем 8:00 по умолчанию
                        };
                    case "Длительность смены":
                        return new NumericUpDown
                        {
                            Tag = colName,
                            Name = colName,
                            Location = new Point(10, 25),
                            Width = parent.Width - 20,
                            Minimum = 0,
                            Maximum = 24
                        };
                }
            }

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
                    Value = initialValue != null && initialValue != DBNull.Value ? Convert.ToDecimal(initialValue) : 0
                },
                OleDbType.Decimal => new NumericUpDown
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    DecimalPlaces = 2,
                    Minimum = decimal.MinValue,
                    Maximum = decimal.MaxValue,
                    Value = initialValue != null && initialValue != DBNull.Value ? Convert.ToDecimal(initialValue) : 0
                },
                OleDbType.Date => new DateTimePicker
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Format = DateTimePickerFormat.Short,
                    Value = initialValue != null && initialValue != DBNull.Value ? Convert.ToDateTime(initialValue) : DateTime.Today,
                    MinDate = new DateTime(2000, 1, 1),
                    MaxDate = new DateTime(2100, 1, 1)
                },
                OleDbType.Boolean => new CheckBox
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Checked = initialValue != null && initialValue != DBNull.Value && Convert.ToBoolean(initialValue)
                },
                _ => new TextBox
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Text = initialValue != null && initialValue != DBNull.Value ? initialValue.ToString() : ""
                }
            };
        }
        private ComboBox CreateComboBoxControl(string colName, GroupBox parent, object initialValue,
                                     string lookupTable, string idColumn, string nameColumn)
        {
            var combo = new ComboBox
            {
                Tag = colName,
                Name = colName,
                Location = new Point(10, 25),
                Width = parent.Width - 20,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            try
            {
                // Получаем данные для ComboBox
                var lookupData = _businessLogic.GetLookupData(lookupTable, idColumn, nameColumn);

                // Заполняем ComboBox
                foreach (var item in lookupData)
                {
                    combo.Items.Add(new KeyValuePair<int, string>(item.Key, item.Value));
                }

                // Настраиваем отображение
                combo.DisplayMember = "Value";
                combo.ValueMember = "Key";

                // Устанавливаем выбранное значение
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
                    combo.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных для {colName}: {ex.Message}",
                              "Ошибка",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning);
            }

            return combo;
        }

        /// <summary>
        /// Обработчик события нажатия кнопки "Сохранить изменения".
        /// Собирает данные из элементов управления, выполняет валидацию и сохраняет изменения в базе данных.
        /// </summary>
        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                var values = new Dictionary<string, object>();
                var missingFields = new List<string>();
                var invalidFields = new List<string>();

                foreach (var control in _inputControls)
                {
                    string fieldName = control.Tag?.ToString();
                    if (string.IsNullOrEmpty(fieldName)) continue;

                    bool isRequired = _businessLogic.IsFieldRequired(_tableName, fieldName);
                    object value = GetControlValue(control);

                    if (isRequired && IsValueEmpty(value))
                    {
                        missingFields.Add(fieldName);
                        control.BackColor = Color.LightPink;
                        continue;
                    }

                    if (!IsValueEmpty(value))
                    {
                        if (value is int intValue && intValue < 0)
                        {
                            invalidFields.Add($"{fieldName} (не может быть отрицательным)");
                            control.BackColor = Color.LightPink;
                            continue;
                        }
                    }

                    values[fieldName] = value ?? DBNull.Value;
                    control.BackColor = SystemColors.Window;
                }

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

                // Вызываем обновление
                bool success = _businessLogic.UpdateRecord(_tableName, values, _idColumnName);

                if (success)
                {
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
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}\n\n{ex.StackTrace}",
                              "Ошибка",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Метод для получения значения из элемента управления.
        /// </summary>
        /// <param name="control">Элемент управления, из которого получается значение.</param>
        /// <returns>Значение, полученное из элемента управления.</returns>
        private object GetControlValue(Control control)
        {
            switch (control)
            {
                case ComboBox cb:
                    if (cb.SelectedItem == null) return null;

                    // Для ComboBox с KeyValuePair
                    if (cb.SelectedItem is KeyValuePair<int, string> kvp)
                        return kvp.Key;

                    // Для обычных ComboBox
                    return cb.SelectedValue ?? cb.SelectedItem;

                case TextBox tb:
                    return tb.Text;

                case NumericUpDown num:
                    return Convert.ToInt32(num.Value);

                case DateTimePicker dt:
                    if (_tableName.Equals("Длительности смен", StringComparison.OrdinalIgnoreCase) &&
                        (dt.Tag.ToString() == "Начало смены" || dt.Tag.ToString() == "Окончание смены"))
                    {
                        // Для времени возвращаем DateTime с фиксированной датой (например, 30.12.1999) и выбранным временем
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
        /// Метод для проверки, является ли значение пустым.
        /// </summary>
        /// <param name="value">Значение для проверки.</param>
        /// <returns>True, если значение пустое, иначе False.</returns>
        private bool IsValueEmpty(object value)
        {
            if (value == null || value == DBNull.Value)
                return true;

            if (value is TimeSpan timeSpan)
                return timeSpan == TimeSpan.Zero;

            return value switch
            {
                string s => string.IsNullOrWhiteSpace(s),
                int i => i == 0,
                decimal d => d == 0,
                DateTime dt => dt == DateTime.MinValue,
                bool _ => false, // Для CheckBox - всегда считается заполненным
                _ => false // Для других типов по умолчанию считаем заполненным
            };
        }
    }
}