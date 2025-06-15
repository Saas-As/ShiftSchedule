using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShiftSchedule
{
    /// <summary>
    /// Класс DataAccess обеспечивает непосредственный доступ к базе данных MS Access.
    /// Реализует CRUD-операции (Create, Read, Update, Delete) и другие низкоуровневые операции с БД.
    /// Содержит методы для:
    /// - Управления подключением к БД
    /// - Выполнения SQL-запросов
    /// - Работы с транзакциями
    /// - Получения метаданных (схемы таблиц)
    /// </summary>
    public class DataAccess
    {
        private readonly string _connectionString; // строка подлключения к БД

        /// <summary>
        /// Конструктор класса DataAccess
        /// Инициализирует строку подключения на основе пути к файлу базы данных.
        /// </summary>
        /// <param name="databasePath">Путь к файлу базы данных</param>
        public DataAccess(string databasePath)
        {
            // Формируем строку подключения с использованием провайдера Microsoft ACE OLEDB
            _connectionString = $@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};";
        }

        /// <summary>
        /// Получает схему указанной таблицы из базы данных.
        /// Использует OleDbSchemaGuid для получения метаданных.
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <returns>DataTable с информацией о колонках таблицы</returns>
        public DataTable GetTableSchema(string tableName)
        {
            // Используем using для автоматического освобождения ресурсов подключения
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open(); // открываем соединение

                // Получаем схему таблицы через OleDbSchemaGuid.Columns
                // Параметры:
                // null - каталог (используется текущий)
                // null - схема (не используется для Access)
                // tableName - имя таблицы
                // null - все колонки

                return conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns,
                    new object[] { null, null, tableName, null });
            }
            // Соединение автоматически закрывается благодаря using
        }

        /// <summary>
        /// Получает следующий доступный ID для новой записи в указанной таблице.
        /// Вычисляет как максимальный существующий ID + 1.
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <param name="idColumnName">Имя колонки с ID</param>
        /// <returns>Следующий доступный ID (максимальный существующий + 1)</returns>
        public int GetNextId(string tableName, string idColumnName)
        {
            // Используем using для автоматического управления подключением
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();

                // Формируем SQL-запрос для поиска максимального ID
                var cmd = new OleDbCommand($"SELECT MAX([{idColumnName}]) FROM [{tableName}]", conn);

                // Выполняем запрос и получаем результат
                var result = cmd.ExecuteScalar();

                // Если результат NULL (нет записей), возвращаем 1, иначе максимальный ID + 1
                return result == DBNull.Value ? 1 : Convert.ToInt32(result) + 1;
            }
        }
        /// <summary>
        /// Получение всех данных из указанной таблицы
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <returns>DataTable со всеми данными таблицы</returns>
        public DataTable GetTableData(string tableName)
        {
            // Используем using для автоматического управления подключением
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();
                // Создание адаптера данных для выполнения SELECT запроса
                var adapter = new OleDbDataAdapter($"SELECT * FROM [{tableName}]", conn);
                // Создание DataTable для хранения результатов
                var dt = new DataTable();
                // Заполнение DataTable данными из БД
                adapter.Fill(dt);

                return dt;
            }
        }
        /// <summary>
        /// Добавляет новую запись в указанную таблицу.
        /// Использует транзакцию для обеспечения целостности данных.
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <param name="values">Словарь значений для вставки (имя столбца  - значение)</param>
        /// <exception cref="Exception">При ошибке выполнения запроса</exception>
        public void InsertRecord(string tableName, Dictionary<string, object> values)
        {
            // Используем using для автоматического управления подключением
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();

                // Начинаем транзакцию для обеспечения атомарности операции
                // Атомарность – это свойство, обозначающее, что транзакция
                // выполняется либо полностью, либо не выполняется вовсе.
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // Формируем списки столбцов и параметров для INSERT запроса
                        var columns = string.Join(", ", values.Keys.Select(k => $"[{k}]"));
                        var parameters = string.Join(", ", values.Keys.Select(k => $"@{k.Replace(" ", "_")}"));
                        
                        // команда для INSERT
                        var cmd = new OleDbCommand(
                            $"INSERT INTO [{tableName}] ({columns}) VALUES ({parameters})",
                            conn,
                            transaction);
                        // Добавление параметров в команду
                        foreach (var item in values)
                        {
                            // Заменяем пробелы в именах параметров на подчеркивания
                            cmd.Parameters.AddWithValue($"@{item.Key.Replace(" ", "_")}", item.Value ?? DBNull.Value);
                        }
                        // выполнение команды
                        cmd.ExecuteNonQuery();

                        // подтверждение транзакции
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        // в случае ошибки откатываем транзакцию
                        transaction.Rollback();
                        throw new Exception($"Ошибка при сохранении в таблицу {tableName}: {ex.Message}");
                    }
                }
            }
        }
        /// <summary>
        /// Удаление записи из таблицы
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <param name="idColumnName">Имя столбца с ID</param>
        /// <param name="idValue">Значение ID для удаления</param>
        public void DeleteRecord(string tableName, string idColumnName, object idValue)
        {
            // Используем using для автоматического управления подключением
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();
                // Создаем команду для DELETE запроса
                var cmd = new OleDbCommand(
                    $"DELETE FROM [{tableName}] WHERE [{idColumnName}] = @id", conn);
                // параметр с ID
                cmd.Parameters.AddWithValue("@id", idValue);
                // выполнение команды
                cmd.ExecuteNonQuery();
            }
        }
        /// <summary>
        /// Обновляет существующую запись в указанной таблице.
        /// Использует транзакцию для обеспечения целостности данных.
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <param name="values">Словарь значений для обновления</param>
        /// <param name="idColumnName">Имя колонки с ID</param>
        /// <exception cref="Exception">При ошибке выполнения запроса</exception>
        public bool UpdateRecord(string tableName, Dictionary<string, object> values, string idColumnName)
        {
            // Используем using для автоматического управления подключением
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();
                // Начинаем транзакцию
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // Подготавливаем части SET для UPDATE запроса
                        var setParts = new List<string>();
                        var parameters = new List<OleDbParameter>();
                        object idValue = null;

                        // Формируем параметры для каждого значения
                        foreach (var item in values)
                        {
                            // Пропускаем ID-столбец (он будет использован в WHERE)
                            if (item.Key.Equals(idColumnName, StringComparison.OrdinalIgnoreCase))
                            {
                                idValue = item.Value;
                                continue;
                            }
                            // Добавляем часть SET для текущего столбца
                            setParts.Add($"[{item.Key}] = ?");

                            // Создаем параметр для значения
                            parameters.Add(new OleDbParameter($"@{item.Key}", item.Value ?? DBNull.Value));
                        }

                        // Проверяем, что ID было найдено
                        if (idValue == null)
                            throw new Exception("Не найдено значение ID для обновления");

                        // Добавляем параметр для ID (используется в WHERE)
                        parameters.Add(new OleDbParameter($"@{idColumnName}", idValue));

                        // Формируем полный текст запроса
                        string query = $"UPDATE [{tableName}] SET {string.Join(", ", setParts)} " +
                                      $"WHERE [{idColumnName}] = ?";

                        // Создаем команду
                        var cmd = new OleDbCommand(query, conn, transaction);
                        // Добавляем параметры в команду
                        cmd.Parameters.AddRange(parameters.ToArray());
                        // Выполняем команду и получаем количество измененных строк
                        int affectedRows = cmd.ExecuteNonQuery();
                        // Подтверждаем транзакцию
                        transaction.Commit();

                        // Возвращаем true, если была обновлена хотя бы одна строка
                        return affectedRows > 0;
                    }
                    catch (Exception ex)
                    {
                        // В случае ошибки откатываем транзакцию
                        transaction.Rollback();
                        throw new Exception($"Ошибка при обновлении записи: {ex.Message}");
                    }
                }
            }
        }
        /// <summary>
        /// Выполняет произвольный SQL-запрос к базе данных.
        /// </summary>
        /// <param name="query">SQL-запрос для выполнения</param>
        /// <returns>DataTable с результатами запроса</returns>
        public DataTable ExecuteCustomQuery(string query)
        {
            // Используем using для автоматического управления подключением
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();
                // Создаем адаптер данных для выполнения запроса
                var adapter = new OleDbDataAdapter(query, conn);
                // Создаем DataTable для результатов
                var dt = new DataTable();
                // Заполняем DataTable данными
                adapter.Fill(dt);

                return dt;
            }
        }
        /// <summary>
        /// Получает список видимых таблиц базы данных (исключая системные таблицы и Users).
        /// </summary>
        /// <returns>DataTable с информацией о видимых таблицах</returns>
        public DataTable GetVisibleTables()
        {
            // Используем using для автоматического управления подключением
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();
                // Получаем все таблицы через OleDbSchemaGuid
                DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });

                // Создаем копию структуры для фильтрации
                var filtered = schema.Clone();
                // Фильтруем таблицы, исключая системные и таблицу Users
                foreach (DataRow row in schema.Rows)
                {
                    string name = row["TABLE_NAME"].ToString();
                    if (!name.StartsWith("MSys") &&
                        !name.StartsWith("~") &&
                        !name.Equals("Users", StringComparison.OrdinalIgnoreCase))
                    {
                        filtered.ImportRow(row);
                    }
                }
                return filtered;
            }
        }
    }
}

