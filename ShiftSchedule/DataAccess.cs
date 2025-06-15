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
    /// Класс DataAccess отвечает за работу с базой данных (Data Access Layer)
    /// Обеспечивает подключение и выполнение операций с базой данных MS Access
    /// </summary>
    public class DataAccess
    {
        private readonly string _connectionString; // строка подлключения к БД
        private OleDbConnection _connection; // подключение к БД

        /// <summary>
        /// Конструктор класса DataAccess
        /// </summary>
        /// <param name="databasePath">Путь к файлу базы данных</param>
        public DataAccess(string databasePath)
        {
            // строка подключения для MS Access
            _connectionString = $@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};";
        }

        /// <summary>
        /// Получение схемы таблицы
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <returns>DataTable с информацией о колонках таблицы</returns>
        public DataTable GetTableSchema(string tableName)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open(); // открываем соединение
                
                // Получаем схему таблицы через OleDbSchemaGuid.Columns
                // Параметры: null - каталог, null - схема, tableName - имя таблицы, null - все колонки

                return conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns,
                    new object[] { null, null, tableName, null });
            }
            // Соединение автоматически закрывается благодаря using
        }

        /// <summary>
        /// Получает следующий ID для новой записи
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <param name="idColumnName">Имя колонки с ID</param>
        /// <returns>Следующий доступный ID (максимальный существующий + 1)</returns>
        public int GetNextId(string tableName, string idColumnName)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                // команда для поиска максимального ID

                var cmd = new OleDbCommand($"SELECT MAX([{idColumnName}]) FROM [{tableName}]", conn);
                
                // Выполнение запроса
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
            using (var conn = new OleDbConnection(_connectionString))
            {
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
        /// Добавление новой записи в таблицу
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <param name="values">Словарь значений для вставки (имя колонки - значение)</param>
        /// <exception cref="Exception">Ошибка</exception>
        public void InsertRecord(string tableName, Dictionary<string, object> values)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                // транзакция для безопасного выполнения
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // списки колонок и параметров для INSERT
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
        /// Освобождение ресурсов подключения
        /// </summary>
        public void Dispose()
        {
            _connection?.Dispose(); // закрываем соединение, если оно было открыто
        }
        /// <summary>
        /// Удаление записи из таблицы
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <param name="idColumnName">Имя колонки с ID</param>
        /// <param name="idValue">Значение ID для удаления</param>
        public void DeleteRecord(string tableName, string idColumnName, object idValue)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                // команда для DELETE
                var cmd = new OleDbCommand(
                    $"DELETE FROM [{tableName}] WHERE [{idColumnName}] = @id", conn);
                // параметр с ID
                cmd.Parameters.AddWithValue("@id", idValue);
                // выполнение команды
                cmd.ExecuteNonQuery();
            }
        }
        /// <summary>
        /// Обновление существующей записи в таблицы
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <param name="values">Словарь значений для обновления</param>
        /// <param name="idColumnName">Имя колонки с ID</param>
        /// <exception cref="Exception">Ошибка</exception>

        public bool UpdateRecord(string tableName, Dictionary<string, object> values, string idColumnName)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        var setParts = new List<string>();
                        var parameters = new List<OleDbParameter>();
                        object idValue = null;

                        foreach (var item in values)
                        {
                            if (item.Key.Equals(idColumnName, StringComparison.OrdinalIgnoreCase))
                            {
                                idValue = item.Value;
                                continue;
                            }

                            setParts.Add($"[{item.Key}] = ?");
                            parameters.Add(new OleDbParameter($"@{item.Key}", item.Value ?? DBNull.Value));
                        }

                        if (idValue == null)
                            throw new Exception("Не найдено значение ID для обновления");

                        parameters.Add(new OleDbParameter($"@{idColumnName}", idValue));

                        string query = $"UPDATE [{tableName}] SET {string.Join(", ", setParts)} " +
                                      $"WHERE [{idColumnName}] = ?";

                        var cmd = new OleDbCommand(query, conn, transaction);
                        cmd.Parameters.AddRange(parameters.ToArray());

                        int affectedRows = cmd.ExecuteNonQuery();
                        transaction.Commit();

                        return affectedRows > 0;
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw new Exception($"Ошибка при обновлении записи: {ex.Message}");
                    }
                }
            }
        }
        public DataTable ExecuteCustomQuery(string query)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                var adapter = new OleDbDataAdapter(query, conn);
                var dt = new DataTable();
                adapter.Fill(dt);
                return dt;
            }
        }
        public DataTable GetVisibleTables()
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                // Получаем все таблицы
                DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });

                // Фильтруем
                var filtered = schema.Clone();
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

