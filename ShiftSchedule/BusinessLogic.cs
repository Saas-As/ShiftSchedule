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
    /// Класс BusinessLogic отвечает за реализацию бизнес-логики приложения.
    /// Он взаимодействует с уровнем доступа к данным (DataAccess) для выполнения операций с базой данных.
    /// Основные функции включают получение схемы таблиц, управление записями (вставка, обновление, удаление)
    /// и валидацию данных
    /// </summary>
    public class BusinessLogic
    {
        // Поле для хранения экземпляра класса DataAccess, который отвечает за доступ к данным.
        private readonly DataAccess _dataAccess;
        /// <summary>
        /// Инициализация объекта DataAccess с указанным путем к базе данных.
        /// </summary>
        /// <param name="databasePath">путь к выбранной БД</param>
        public BusinessLogic(string databasePath)
        {
            _dataAccess = new DataAccess(databasePath);
        }
        /// <summary>
        /// Метод для получения схемы таблицы из базы данных
        /// </summary>
        /// <param name="tableName">название таблицы</param>
        /// <returns>DataTable с информацией о схеме таблицы</returns>
        public DataTable GetTableSchema(string tableName)
        {
            // получение схемы таблицы из DataAccess
            var schema = _dataAccess.GetTableSchema(tableName);

            // фильтрация системных таблиц и таблицы Users
            var filteredTable = schema.Clone();
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
        ///  Метод для получения имени столбца, который является идентификатором записи в таблице
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <returns>имя таблицы с ID</returns>
        /// <exception cref="Exception">Ошибка</exception>
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
        /// Метод для получения следующего идентификатора для таблицы
        /// </summary>
        /// <param name="tableName">название таблицы</param>
        /// <returns>следующий ID таблицы</returns>
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
            return _dataAccess.GetTableData(tableName);
        }
        /// <summary>
        /// Метод для добавления новой записи в таблицу
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="values">словарь с данными для вставки</param>
        public void InsertRecord(string tableName, Dictionary<string, object> values)
        {
            // валидация
            ValidateData(tableName, values);
            // вставка
            _dataAccess.InsertRecord(tableName, values);
        }
        /// <summary>
        /// метод для валидации данных
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="values">словарь с данными для валидации</param>
        private void ValidateData(string tableName, Dictionary<string, object> values)
        {
            var schema = _dataAccess.GetTableSchema(tableName);

            foreach (var field in values)
            {
                var column = schema.Rows.Cast<DataRow>()
                    .FirstOrDefault(r => r["COLUMN_NAME"].ToString().Equals(field.Key, StringComparison.OrdinalIgnoreCase));

                if (column == null)
                    throw new ArgumentException($"Поле {field.Key} не существует в таблице {tableName}");

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

                if (field.Value != null)
                {
                    if (expectedType == typeof(int) && field.Value is decimal)
                    {
                        values[field.Key] = Convert.ToInt32(field.Value);
                        continue;
                    }

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
            return oleDbType switch
            {
                OleDbType.Integer => typeof(int),
                OleDbType.Decimal => typeof(decimal),
                OleDbType.Currency => typeof(decimal),
                OleDbType.Date => typeof(DateTime),
                OleDbType.Boolean => typeof(bool),
                _ => typeof(string)
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
            _dataAccess.DeleteRecord(tableName, idColumnName, idValue);
        }
        /// <summary>
        /// Метод для обновления записи в таблице
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="values">Словарь с данными для обновления</param>   
        /// <param name="idColumnName">имя столбца с ID</param>
        public bool UpdateRecord(string tableName, Dictionary<string, object> values, string idColumnName)
        {
            try
            {
                return _dataAccess.UpdateRecord(tableName, values, idColumnName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении записи: {ex.Message}",
                              "Ошибка",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Error);
                return false;
            }
        }
        public DataTable ExecuteCustomQuery(string query)
        {
            return _dataAccess.ExecuteCustomQuery(query);
        }
        public DataTable GetVisibleTables()
        {
            return _dataAccess.GetVisibleTables();
        }
        public Dictionary<int, string> GetLookupData(string lookupTableName, string idColumn, string nameColumn)
        {
            var data = new Dictionary<int, string>();
            var table = _dataAccess.GetTableData(lookupTableName);

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
