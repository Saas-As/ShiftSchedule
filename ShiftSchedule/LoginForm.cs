using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShiftSchedule
{
    /// <summary>
    /// Форма входа в систему.
    /// Обеспечивает:
    /// - Аутентификацию пользователей
    /// - Переход к форме регистрации
    /// </summary>
    public partial class LoginForm : Form
    {
        // Объект для работы с аутентификацией
        private Authentication _auth;

        // Путь к файлу базы данных
        private string _databasePath;

        /// <summary>
        /// Конструктор формы входа.
        /// </summary>
        /// <param name="databasePath">Путь к файлу базы данных</param>
        public LoginForm(string databasePath)
        {
            InitializeComponent();
            // Сохраняем путь к базе данных
            _databasePath = databasePath;
            // Создаем объект для работы с аутентификацией
            _auth = new Authentication(databasePath);
            // Настраиваем поле пароля - символы заменяются звездочками
            txtPassword.UseSystemPasswordChar = true;

            // Добавляем обработчик закрытия формы
            this.FormClosing += LoginForm_FormClosing;
        }
        /// <summary>
        /// Обработчик события закрытия формы.
        /// Завершает приложение, если форма закрыта не через кнопку входа.
        /// </summary>
        private void LoginForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Если закрываем форму не через кнопку "Войти"
            if (this.DialogResult != DialogResult.OK)
            {
                Application.Exit(); // Завершаем приложение полностью
            }
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Войти".
        /// Проверяет введенные учетные данные и выполняет вход.
        /// </summary>
        private void btnLogin_Click(object sender, EventArgs e)
        {
            // Проверяем, что введены логин и пароль
            if (string.IsNullOrWhiteSpace(txtUsername.Text) || 
                string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                MessageBox.Show("Введите логин и пароль", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Пытаемся аутентифицировать пользователя
            if (_auth.Authenticate(txtUsername.Text, txtPassword.Text))
            {
                // Если аутентификация успешна - закрываем форму с результатом OK
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("Неверный логин или пароль", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Регистрация".
        /// Открывает форму регистрации нового пользователя.
        /// </summary>
        private void btnRegister_Click(object sender, EventArgs e)
        {
            // Создаем и показываем форму регистрации
            var registerForm = new RegisterForm(_databasePath);

            // Если регистрация прошла успешно (форма закрыта с OK)
            if (registerForm.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Регистрация прошла успешно!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Обработчик изменения состояния чекбокса "Показать пароль".
        /// Переключает отображение символов в поле пароля.
        /// </summary>
        private void chkShowPassword_CheckedChanged(object sender, EventArgs e)
        {
            // Показываем/скрываем символы пароля в зависимости от состояния чекбокса
            txtPassword.UseSystemPasswordChar = !chkShowPassword.Checked;
        }
    }
}
