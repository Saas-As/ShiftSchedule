using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ShiftSchedule
{
    /// <summary>
    /// Класс для работы с аутентификацией пользователей.
    /// Обеспечивает:
    /// - Хеширование паролей
    /// - Проверку существования пользователей
    /// - Регистрацию новых пользователей
    /// - Аутентификацию пользователей
    /// </summary>
    internal class Authentication
    {
        // Строка подключения к базе данных
        private readonly string _connectionString;

        /// <summary>
        /// Конструктор класса аутентификации.
        /// </summary>
        /// <param name="databasePath">Путь к файлу базы данных</param>
        public Authentication(string databasePath)
        {
            // Формируем строку подключения
            _connectionString = $@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};";
        }

        /// <summary>
        /// Хеширует пароль с использованием SHA256 и "соли".
        /// </summary>
        /// <param name="password">Пароль для хеширования</param>
        /// <returns>Хеш пароля в виде hex-строки</returns>
        private string HashPassword(string password)
        {
            // Используем SHA256 для хеширования
            using (var sha256 = SHA256.Create())
            {
                // Добавляем "соль" к паролю и вычисляем хеш
                var hashedBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password + "ShiftScheduleSalt"));

                // Преобразуем байты в hex-строку и возвращаем
                return BitConverter.ToString(hashedBytes).Replace("-", "").ToLower();
            }
        }

        /// <summary>
        /// Проверяет существование пользователя с указанным логином.
        /// </summary>
        /// <param name="username">Логин для проверки</param>
        /// <returns>True, если пользователь существует</returns>
        public bool UserExists(string username)
        {
            // Используем using для автоматического закрытия подключения
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();
                // Создаем команду для проверки существования пользователя
                var cmd = new OleDbCommand("SELECT COUNT(*) FROM [Users] WHERE [Username] = ?", conn);
                // Добавляем параметр с логином
                cmd.Parameters.AddWithValue("@username", username);

                // Выполняем запрос и возвращаем результат
                return (int)cmd.ExecuteScalar() > 0;
            }
        }

        /// <summary>
        /// Регистрирует нового пользователя в системе.
        /// </summary>
        /// <param name="username">Логин пользователя</param>
        /// <param name="password">Пароль пользователя</param>
        /// <returns>True, если регистрация прошла успешно</returns>
        public bool RegisterUser(string username, string password)
        {
            // Проверяем, не существует ли уже пользователь
            if (UserExists(username))
                return false;

            // Используем using для автоматического закрытия подключения
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();

                // Создаем команду для вставки нового пользователя
                var cmd = new OleDbCommand("INSERT INTO [Users] ([Username], [PasswordHash]) VALUES (?, ?)", conn);

                // Добавляем параметры
                cmd.Parameters.AddWithValue("@username", username);
                cmd.Parameters.AddWithValue("@password", HashPassword(password));

                // Выполняем команду и возвращаем результат
                return cmd.ExecuteNonQuery() > 0;
            }
        }

        /// <summary>
        /// Аутентифицирует пользователя
        /// </summary>
        /// <param name="username">Логин пользователя</param>
        /// <param name="password">Пароль пользователя</param>
        /// <returns>True, если аутентификация прошла успешно</returns>
        public bool Authenticate(string username, string password)
        {
            // Используем using для автоматического закрытия подключения
            using (var conn = new OleDbConnection(_connectionString))
            {
                // Открываем соединение
                conn.Open();
                // Создаем команду для получения хеша пароля
                var cmd = new OleDbCommand("SELECT [PasswordHash] FROM [Users] WHERE [Username] = ?", conn);
                // Добавляем параметр с логином
                cmd.Parameters.AddWithValue("@username", username);
                // Выполняем запрос
                var result = cmd.ExecuteScalar();

                // Если пользователь не найден - возвращаем false
                if (result == null || result == DBNull.Value)
                    return false;

                // Получаем сохраненный хеш из базы данных
                var storedHash = result.ToString();
                // Вычисляем хеш введенного пароля
                var inputHash = HashPassword(password);
                // Сравниваем хеши
                return storedHash.Equals(inputHash);
            }
        }
    }
}
