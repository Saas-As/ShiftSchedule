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
    public partial class AddRecord : Form
    {
        // Поле для хранения имени таблицы, в которую добавляется запись
        private readonly string _tableName;
        // Поле для хранения экземпляра класса BusinessLogic
        private readonly BusinessLogic _businessLogic;
        // Список элементов управления для ввода данных
        private readonly List<Control> _inputControls = new List<Control>();
        // Кнопки для сохранения и отмены
        private Button _btnSave;
        private Button _btnCancel;
        // Элемент управления для ввода ID
        private NumericUpDown _idControl;
        private readonly string _idColumnName;

        /// <summary>
        /// Конструктор класса AddRecord.
        /// Инициализирует форму для добавления записи в указанную таблицу.
        /// </summary>
        /// <param name="tableName">Название таблицы, в которую добавляется запись.</param>
        /// <param name="businessLogic">Экземпляр класса BusinessLogic для взаимодействия с данными.</param>
        /// <exception cref="ArgumentNullException">Выбрасывается, если tableName или businessLogic равны null.</exception>
        public AddRecord(string tableName, BusinessLogic businessLogic)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentNullException(nameof(tableName));
            if (businessLogic == null)
                throw new ArgumentNullException(nameof(businessLogic));

            _tableName = tableName;
            _businessLogic = businessLogic;
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
            this.Text = $"Добавление записи в {_tableName}";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Padding = new Padding(20); // Добавляем отступы от краев формы

            // Создание кнопок
            CreateButtons();

            // Загрузка полей таблицы
            LoadTableFields();

            // Устанавливаем минимальный размер формы
            this.MinimumSize = new Size(450, 300);
        }

        /// <summary>
        /// Метод для создания кнопок "Сохранить" и "Отмена".
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
                        Value = _businessLogic.GetNextId(_tableName),
                        ReadOnly = true,
                        BackColor = SystemColors.Control,
                        Enabled = false
                    };
                    inputControl = _idControl;
                }
                else
                {
                    // Обычные поля
                    inputControl = CreateInputControl(column, groupBox);
                }

                groupBox.Controls.Add(inputControl);
                panel.Controls.Add(groupBox);
                _inputControls.Add(inputControl);

                yPos += groupBox.Height + 10;
            }

            // Вычисляем общую высоту панели
            int totalHeight = yPos + 40; // 40 - дополнительное пространство для кнопок

            // Устанавливаем высоту панели
            panel.Height = totalHeight;

            // Устанавливаем высоту формы
            this.ClientSize = new Size(this.ClientSize.Width, totalHeight + 100); // 100 - пространство для кнопок и отступов

            // Позиционирование кнопок
            _btnSave.Top = this.ClientSize.Height - 60;
            _btnCancel.Top = this.ClientSize.Height - 60;
        }

        /// <summary>
        /// Метод для создания элементов управления для ввода данных в зависимости от типа данных поля.
        /// </summary>
        /// <param name="column">Строка с информацией о поле таблицы.</param>
        /// <param name="parent">Родительский элемент, в который добавляется элемент управления.</param>
        /// <returns>Элемент управления для ввода данных.</returns>
        private Control CreateInputControl(DataRow column, GroupBox parent)
        {
            string colName = column["COLUMN_NAME"].ToString();
            var dataType = (OleDbType)column["DATA_TYPE"];

            if (_tableName.Equals("Смены", StringComparison.OrdinalIgnoreCase))
            {
                switch (colName)
                {
                    case "ID_подразделения":
                        return CreateComboBoxControl(colName, parent, "Подразделения", "ID_подразделения", "Подразделение");
                    case "ID_руководителя":
                        return CreateComboBoxControl(colName, parent, "Руководители", "ID_руководителя", "ФИО_руководителя");
                    case "ID_количества_рабочих":
                        return CreateComboBoxControl(colName, parent, "Количество рабочих", "ID_количества_рабочих", "Количество рабочих");
                    case "ID_длительности_смены":
                        return CreateComboBoxControl(colName, parent, "Длительности смен", "ID_длительности_смены", "Длительность смены");
                    case "ID_начальника_смены":
                        return CreateComboBoxControl(colName, parent, "Начальники смен", "ID_начальника_смены", "ФИО_начальника_смены");
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
                    Maximum = int.MaxValue
                },
                OleDbType.Decimal => new NumericUpDown
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    DecimalPlaces = 2,
                    Minimum = decimal.MinValue,
                    Maximum = decimal.MaxValue
                },
                OleDbType.Date => new DateTimePicker
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Format = DateTimePickerFormat.Short,
                    Value = DateTime.Today, // Устанавливаем сегодняшнюю дату по умолчанию
                    MinDate = new DateTime(2000, 1, 1), // Минимальная допустимая дата
                    MaxDate = new DateTime(2100, 1, 1)  // Максимальная допустимая дата
                },
                OleDbType.Boolean => new CheckBox
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20,
                    Text = ""
                },
                _ => new TextBox
                {
                    Tag = colName,
                    Name = colName,
                    Location = new Point(10, 25),
                    Width = parent.Width - 20
                }
            };
        }
        private ComboBox CreateComboBoxControl(string colName, GroupBox parent,
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
                var lookupData = _businessLogic.GetLookupData(lookupTable, idColumn, nameColumn);

                foreach (var item in lookupData)
                {
                    combo.Items.Add(new KeyValuePair<int, string>(item.Key, item.Value));
                }

                combo.DisplayMember = "Value";
                combo.ValueMember = "Key";

                if (combo.Items.Count > 0)
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
        /// Обработчик события нажатия кнопки "Сохранить".
        /// Собирает данные из элементов управления, выполняет валидацию и сохраняет их в базе данных.
        /// </summary>
        private void BtnSave_Click(object sender, EventArgs e)
        {
            var values = new Dictionary<string, object>();
            var missingFields = new List<string>();

            // Фильтруем только элементы управления для ввода данных
            var inputFields = _inputControls.Where(c =>
                c is TextBox ||
                c is NumericUpDown ||
                c is DateTimePicker ||
                c is CheckBox ||
                c is ComboBox).ToList();

            foreach (var control in inputFields) // Используем отфильтрованный список
            {
                string fieldName = control.Tag?.ToString();
                if (string.IsNullOrEmpty(fieldName)) continue;

                bool isRequired = _businessLogic.IsFieldRequired(_tableName, fieldName);
                object value = GetControlValue(control);

                if (isRequired && IsValueEmpty(value))
                {
                    missingFields.Add(fieldName);
                    control.BackColor = Color.LightPink;
                }
                else
                {
                    values.Add(fieldName, value);
                    control.BackColor = SystemColors.Window;
                }
            }

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
                _businessLogic.InsertRecord(_tableName, values);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}",
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
                    if (cb.SelectedItem is KeyValuePair<int, string> kvp)
                        return kvp.Key;
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
                decimal d => d == 0,
                DateTime dt => dt == DateTime.MinValue,
                _ => false
            };
        }
    }
}