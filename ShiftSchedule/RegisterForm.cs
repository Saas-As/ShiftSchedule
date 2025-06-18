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
    /// Форма регистрации новых пользователей в системе.
    /// Обеспечивает:
    /// - Ввод логина и пароля
    /// - Подтверждение пароля
    /// - Валидацию введенных данных
    /// - Взаимодействие с классом Authentication
    /// </summary>
    public partial class RegisterForm : Form
    {
        // Объект для работы с аутентификацией
        private Authentication _auth;
        // Путь к файлу базы данных
        private string _databasePath;

        /// <summary>
        /// Конструктор формы регистрации
        /// </summary>
        /// <param name="databasePath">Путь к файлу базы данных</param>
        public RegisterForm(string databasePath)
        {
            InitializeComponent();
            // Сохраняем путь к базе данных
            _databasePath = databasePath;
            // Создаем объект для работы с аутентификацией
            _auth = new Authentication(databasePath);

            // Настраиваем поля паролей - символы заменяются звездочками
            txtPassword.UseSystemPasswordChar = true;
            txtConfirmPassword.UseSystemPasswordChar = true;
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Зарегистрироваться".
        /// Выполняет проверку введенных данных и регистрирует нового пользователя.
        /// </summary>
        private void btnRegister_Click(object sender, EventArgs e)
        {
            // Проверяем, что все поля заполнены
            if (string.IsNullOrWhiteSpace(txtUsername.Text) ||
               string.IsNullOrWhiteSpace(txtPassword.Text) ||
               string.IsNullOrWhiteSpace(txtConfirmPassword.Text))
            {
                MessageBox.Show("Заполните все поля", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Проверяем минимальную длину пароля (6 символов)
            if (txtPassword.Text.Length < 6)
            {
                MessageBox.Show("Пароль должен содержать не менее 6 символов", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // Очищаем поля паролей
                txtPassword.Clear();
                txtConfirmPassword.Clear();
                return;
            }

            // Проверяем совпадение паролей
            if (txtPassword.Text != txtConfirmPassword.Text)
            {
                MessageBox.Show("Пароли не совпадают", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // Очищаем поля паролей
                txtPassword.Clear();
                txtConfirmPassword.Clear();
                return;
            }

            // Проверяем, не существует ли уже пользователь с таким логином
            if (_auth.UserExists(txtUsername.Text))
            {
                MessageBox.Show("Пользователь с таким логином уже существует", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Пытаемся зарегистрировать пользователя
            if (_auth.RegisterUser(txtUsername.Text, txtPassword.Text))
            {
                // Если регистрация успешна - закрываем форму с результатом OK
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("Ошибка при регистрации", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Отмена".
        /// Закрывает форму без регистрации пользователя.
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            // Устанавливаем результат Cancel и закрываем форму
            DialogResult = DialogResult.Cancel;
            Close();
        }
        /// <summary>
        /// Обработчик изменения состояния чекбокса "Показать пароль".
        /// Переключает отображение символов в полях паролей.
        /// </summary>
        private void chkShowPassword_CheckedChanged(object sender, EventArgs e)
        {
            // Показываем/скрываем символы пароля в зависимости от состояния чекбокса
            txtPassword.UseSystemPasswordChar = !chkShowPassword.Checked;
            txtConfirmPassword.UseSystemPasswordChar = !chkShowPassword.Checked;
        }
    }
}