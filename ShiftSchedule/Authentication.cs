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
    internal class Authentication
    {
        private readonly string _connectionString;

        public Authentication(string databasePath)
        {
            _connectionString = $@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={databasePath};";
        }

        // Хеширование пароля с солью
        private string HashPassword(string password)
        {
            using (var sha256 = SHA256.Create())
            {
                var hashedBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password + "ShiftScheduleSalt"));
                return BitConverter.ToString(hashedBytes).Replace("-", "").ToLower();
            }
        }

        // Проверка существования пользователя
        public bool UserExists(string username)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                var cmd = new OleDbCommand("SELECT COUNT(*) FROM [Users] WHERE [Username] = ?", conn);
                cmd.Parameters.AddWithValue("@username", username);
                return (int)cmd.ExecuteScalar() > 0;
            }
        }

        // Регистрация нового пользователя
        public bool RegisterUser(string username, string password)
        {
            if (UserExists(username))
                return false;

            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                var cmd = new OleDbCommand("INSERT INTO [Users] ([Username], [PasswordHash]) VALUES (?, ?)", conn);
                cmd.Parameters.AddWithValue("@username", username);
                cmd.Parameters.AddWithValue("@password", HashPassword(password));
                return cmd.ExecuteNonQuery() > 0;
            }
        }

        // Аутентификация пользователя
        public bool Authenticate(string username, string password)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                var cmd = new OleDbCommand("SELECT [PasswordHash] FROM [Users] WHERE [Username] = ?", conn);
                cmd.Parameters.AddWithValue("@username", username);
                var result = cmd.ExecuteScalar();

                if (result == null || result == DBNull.Value)
                    return false;

                var storedHash = result.ToString();
                var inputHash = HashPassword(password);
                return storedHash.Equals(inputHash);
            }
        }
    }
}
