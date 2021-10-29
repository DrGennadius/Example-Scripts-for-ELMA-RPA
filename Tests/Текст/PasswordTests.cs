using ELMA.RPA.Scripts;
using NUnit.Framework;
using System.Text.RegularExpressions;

namespace Tests
{
    public class PasswordTests
    {
        private const int _generateIterationNumber = 500000;

        private const int _configGenerateIterationNumber = 100000;

        private const int _minPasswordLength = 4;

        private const int _maxPasswordLength = 100;

        [Test]
        public void SimpleTest()
        {
            PasswordGenerator passwordGenerator = new();

            string password = passwordGenerator.Generate(10);
            Assert.IsNotEmpty(password);
            Assert.AreEqual(password.Length, 10);

            password = passwordGenerator.Generate(10000);
            Assert.IsNotEmpty(password);
            Assert.AreEqual(password.Length, 10000);

            for (int i = 0; i < _generateIterationNumber; i++)
            {
                password = passwordGenerator.Generate(_minPasswordLength, _maxPasswordLength);
                Assert.IsNotEmpty(password);
            }

            Assert.Pass();
        }

        [Test]
        public void ConfigPasswordGenerateTest()
        {
            PasswordGenerator passwordGenerator = new();

            // Только цифры.
            passwordGenerator.Configuration = new()
            {
                UseNumeric = true,
                UseLowcase = false,
                UseUppercase = false,
                UseSpecial = false
            };
            for (int i = 0; i < _configGenerateIterationNumber; i++)
            {
                string password = passwordGenerator.Generate(_maxPasswordLength);
                Assert.IsNotEmpty(password);
                Assert.AreEqual(password.Length, _maxPasswordLength);
                Assert.IsFalse(Regex.IsMatch(password, @"\D"));
            }

            // Только нижний регистр.
            passwordGenerator.Configuration = new()
            {
                UseNumeric = false,
                UseLowcase = true,
                UseUppercase = false,
                UseSpecial = false
            };
            for (int i = 0; i < _configGenerateIterationNumber; i++)
            {
                string password = passwordGenerator.Generate(_maxPasswordLength);
                Assert.IsNotEmpty(password);
                Assert.AreEqual(password.Length, _maxPasswordLength);
                Assert.IsFalse(Regex.IsMatch(password, @"\d"));
                Assert.IsFalse(Regex.IsMatch(password, @"[A-Z]"));
            }

            // Только верхний регистр.
            passwordGenerator.Configuration = new()
            {
                UseNumeric = false,
                UseLowcase = false,
                UseUppercase = true,
                UseSpecial = false
            };
            for (int i = 0; i < _configGenerateIterationNumber; i++)
            {
                string password = passwordGenerator.Generate(_maxPasswordLength);
                Assert.IsNotEmpty(password);
                Assert.AreEqual(password.Length, _maxPasswordLength);
                Assert.IsFalse(Regex.IsMatch(password, @"\d"));
                Assert.IsFalse(Regex.IsMatch(password, @"[a-z]"));
            }

            // Только специальные символы.
            passwordGenerator.Configuration = new()
            {
                UseNumeric = false,
                UseLowcase = false,
                UseUppercase = false,
                UseSpecial = true
            };
            for (int i = 0; i < _configGenerateIterationNumber; i++)
            {
                string password = passwordGenerator.Generate(_maxPasswordLength);
                Assert.IsNotEmpty(password);
                Assert.AreEqual(password.Length, _maxPasswordLength);
                Assert.IsFalse(Regex.IsMatch(password, @"\d"));
                Assert.IsFalse(Regex.IsMatch(password, @"[A-Za-z]"));
            }

            Assert.Pass();
        }
    }
}
