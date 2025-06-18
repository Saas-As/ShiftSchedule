using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShiftSchedule
{
    /// <summary>
    /// Класс BusinessLogic реализует бизнес-логику приложения.
    /// Служит промежуточным слоем между пользовательским интерфейсом и уровнем доступа к данным (DataAccess).
    /// Обеспечивает:
    /// - Валидацию данных
    /// - Преобразование типов
    /// - Управление транзакциями
    /// - Обработку бизнес-правил
    /// - Работу с внешними ключами и связями между таблицами
    /// </summary>
    public class BusinessLogic
    {
        // Поле для хранения экземпляра класса DataAccess, который отвечает за доступ к данным.
        private readonly DataAccess _dataAccess;
        /// <summary>
        /// Конструктор класса BusinessLogic.
        /// Инициализация объекта DataAccess с указанным путем к базе данных.
        /// </summary>
        /// <param name="databasePath">путь к выбранной БД</param>
        public BusinessLogic(string databasePath)
        {
            // Создаем новый экземпляр DataAccess для работы с указанной БД
            _dataAccess = new DataAccess(databasePath);
        }
        /// <summary>
        /// Получает схему указанной таблицы из базы данных.
        /// Фильтрует системные таблицы.
        /// </summary>
        /// <param name="tableName">название таблицы</param>
        /// <returns>DataTable с информацией о схеме таблицы</returns>
        public DataTable GetTableSchema(string tableName)
        {
            // получение схемы таблицы из DataAccess
            var schema = _dataAccess.GetTableSchema(tableName);

            // Создаем копию структуры schema для фильтрации
            var filteredTable = schema.Clone();
            // Фильтруем системные таблицы (начинающиеся с MSys или ~)
            foreach (DataRow row in schema.Rows)
            {
                string table = row["TABLE_NAME"].ToString();
                if (!table.StartsWith("MSys") && !table.StartsWith("~"))
                {
                    filteredTable.ImportRow(row);
                }
            }

            return filteredTable;
        }
        /// <summary>
        /// Определяет имя столбца с первичным ключом для указанной таблицы.
        /// Использует предопределенный словарь для известных таблиц.
        /// Для остальных таблиц ищет столбец, содержащий "ID" в названии.
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <returns>имя столбца с первичным ключом</returns>
        /// <exception cref="Exception">Если не удается определить ID-столбец</exception>
        public string GetIdColumnName(string tableName)
        {   
            // Словарь соответствий таблиц и их ID-полей
            var idColumns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Смены", "Код смены" },
                { "Руководители", "ID_руководителя" },
                { "Подразделения", "ID_подразделения" },
                { "Длительности смен", "ID_длительности_смены" },
                { "Начальники смен", "ID_начальника_смены" },
                { "Количество рабочих", "ID_количества_рабочих" }
            };
            // проверка, есть ли таблица в словере
            if (idColumns.TryGetValue(tableName, out var idColumnName))
            {
                return idColumnName;
            }

            // Если таблицы нет в словаре, ищем поле с "ID" в названии
            var schema = GetTableSchema(tableName);
            foreach (DataRow row in schema.Rows)
            {
                string columnName = row["COLUMN_NAME"].ToString();
                // Проверяем, содержит ли название столбца "ID" (без учета регистра)
                if (columnName.IndexOf("ID", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return columnName;
                }
            }

            throw new Exception($"Не удалось определить поле ID для таблицы {tableName}");
        }
        /// <summary>
        /// Получает следующий доступный ID для новой записи в указанной таблице.
        /// Вычисляет как максимальный существующий ID + 1.
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <returns>Следующий доступный ID</returns>
        public int GetNextId(string tableName)
        {
            // Получаем имя столбца идентификатора
            string idColumnName = GetIdColumnName(tableName);
            // Получаем следующий идентификатор из DataAccess
            return _dataAccess.GetNextId(tableName, idColumnName);
        }
        /// <summary>
        /// Метод для проверки, является ли поле обязательным в таблице
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="fieldName">имя поля</param>
        /// <returns>True, если поле обязательное, иначе False</returns>
        public bool IsFieldRequired(string tableName, string fieldName)
        {
            // Получаем схему таблицы
            var schema = _dataAccess.GetTableSchema(tableName);
            // Проверяем, является ли поле обязательным (IS_NULLABLE = "NO")
            return schema.Rows.Cast<DataRow>()
                .Any(r => r["COLUMN_NAME"].ToString().Equals(fieldName, StringComparison.OrdinalIgnoreCase) &&
                         r["IS_NULLABLE"].ToString() == "NO");
        }
        /// <summary>
        /// метод для получения данных из таблицы
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <returns>DataTable с данными из таблицы</returns>
        public DataTable GetTableData(string tableName)
        {
            // Делегируем запрос к DataAccess
            return _dataAccess.GetTableData(tableName);
        }
        /// <summary>
        /// Метод для добавления новой записи в таблицу
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="values">словарь с данными для вставки</param>
        public void InsertRecord(string tableName, Dictionary<string, object> values)
        {
            // Валидируем данные перед вставкой
            ValidateData(tableName, values);

            // Делегируем вставку DataAccess
            _dataAccess.InsertRecord(tableName, values);
        }
        /// <summary>
        /// Валидирует данные перед вставкой или обновлением.
        /// Проверяет соответствие типов данных и бизнес-правила.
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="values">словарь с данными для валидации</param>
        private void ValidateData(string tableName, Dictionary<string, object> values)
        {
            // Получаем схему таблицы для проверки
            var schema = _dataAccess.GetTableSchema(tableName);

            // Проверяем каждое поле в словаре значений
            foreach (var field in values)
            {
                // Ищем описание столбца в схеме таблицы
                var column = schema.Rows.Cast<DataRow>()
                    .FirstOrDefault(r => r["COLUMN_NAME"].ToString().Equals(field.Key, StringComparison.OrdinalIgnoreCase));

                // Если столбец не найден - ошибка
                if (column == null)
                    throw new ArgumentException($"Поле {field.Key} не существует в таблице {tableName}");

                // Получаем ожидаемый тип данных для столбца
                var expectedType = GetDotNetType((OleDbType)column["DATA_TYPE"]);

                // Специальная обработка для полей времени в "Длительности смен"
                if (tableName.Equals("Длительности смен", StringComparison.OrdinalIgnoreCase) &&
                    (field.Key == "Начало смены" || field.Key == "Окончание смены"))
                {
                    if (field.Value is DateTime)
                        continue; // Тип правильный
                    else
                        throw new ArgumentException(
                            $"Неверный тип данных для поля {field.Key}. Ожидается: DateTime, получено: {field.Value?.GetType().Name ?? "null"}");
                }

                // Если значение не null, проверяем его тип
                if (field.Value != null)
                {
                    // Специальная обработка для decimal -> int
                    if (expectedType == typeof(int) && field.Value is decimal)
                    {
                        values[field.Key] = Convert.ToInt32(field.Value);
                        continue;
                    }

                    // Проверяем соответствие типов
                    if (field.Value.GetType() != expectedType)
                    {
                        throw new ArgumentException(
                            $"Неверный тип данных для поля {field.Key}. Ожидается: {expectedType.Name}, получено: {field.Value.GetType().Name}");
                    }
                }
            }
        }
        /// <summary>
        /// Метод для получения типа данных .NET, соответствующего типу данных OleDb
        /// </summary>
        /// <param name="oleDbType">тип данных OleDb</param>
        /// <returns>тип данных .NET</returns>
        private Type GetDotNetType(OleDbType oleDbType)
        {
            // Сопоставление типов OleDb и .NET
            return oleDbType switch
            {
                OleDbType.Integer => typeof(int),       // Целое число
                OleDbType.Decimal => typeof(decimal),   // Десятичное число
                OleDbType.Currency => typeof(decimal),  // Денежный тип
                OleDbType.Date => typeof(DateTime),     // Дата/время
                OleDbType.Boolean => typeof(bool),      // Логический тип
                _ => typeof(string)                     // По умолчанию - строка
            };
        }
        /// <summary>
        /// Метод для удаления записи из таблицы
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="idColumnName">имя столбца с ID</param>
        /// <param name="idValue">значение ID</param>
        public void DeleteRecord(string tableName, string idColumnName, object idValue)
        {
            // Делегируем удаление DataAccess
            _dataAccess.DeleteRecord(tableName, idColumnName, idValue);
        }
        /// <summary>
        /// Метод для обновления записи в таблице
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="values">Словарь с данными для обновления</param>   
        /// <param name="idColumnName">имя столбца с ID</param>
        /// <returns>True, если обновление прошло успешно</returns>
        public bool UpdateRecord(string tableName, Dictionary<string, object> values, string idColumnName)
        {
            try
            {
                // Делегируем обновление DataAccess
                return _dataAccess.UpdateRecord(tableName, values, idColumnName);
            }
            catch (Exception ex)
            {
                // Обрабатываем ошибки обновления
                MessageBox.Show($"Ошибка при обновлении записи: {ex.Message}",
                              "Ошибка",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Error);
                return false;
            }
        }
        /// <summary>
        /// Выполняет произвольный SQL-запрос к базе данных.
        /// </summary>
        /// <param name="query">SQL-запрос</param>
        /// <returns>DataTable с результатами запроса</returns>
        public DataTable ExecuteCustomQuery(string query)
        {
            // Делегируем выполнение запроса DataAccess
            return _dataAccess.ExecuteCustomQuery(query);
        }
        /// <summary>
        /// Получает список видимых таблиц (исключая системные).
        /// </summary>
        /// <returns>DataTable с информацией о видимых таблицах</returns>
        public DataTable GetVisibleTables()
        {
            // Делегируем запрос к DataAccess
            return _dataAccess.GetVisibleTables();
        }
        /// <summary>
        /// Получает данные для ComboBox (справочники) в формате ID - Name.
        /// Используется для отображения связанных данных (например, подразделений).
        /// </summary>
        /// <param name="lookupTableName">Имя таблицы-справочника</param>
        /// <param name="idColumn">Имя столбца с ID</param>
        /// <param name="nameColumn">Имя столбца с отображаемым значением</param>
        /// <returns>Словарь (ID, Name) для заполнения ComboBox</returns>
        public Dictionary<int, string> GetLookupData(string lookupTableName, string idColumn, string nameColumn)
        {
            // Создаем словарь для результатов
            var data = new Dictionary<int, string>();
            // Получаем данные таблицы-справочника
            var table = _dataAccess.GetTableData(lookupTableName);

            // Заполняем словарь значениями из таблицы
            foreach (DataRow row in table.Rows)
            {
                int id = Convert.ToInt32(row[idColumn]);
                string name = row[nameColumn].ToString();
                data.Add(id, name);
            }

            return data;
        }
    }
}
