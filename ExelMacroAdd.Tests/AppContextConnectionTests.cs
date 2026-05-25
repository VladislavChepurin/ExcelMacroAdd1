using NUnit.Framework;

namespace ExelMacroAdd.Tests
{
    [TestFixture]
    public class AppContextConnectionTests
    {
        [Test]
        public void UnitOfWorkConnection_UsesFrameworkParser_ForSqliteConnectionString()
        {
            using (var scope = new SqliteIntegrationTestScope())
            using (var unitOfWork = scope.UnitOfWorkFactory.Create())
            {
                var connection = unitOfWork.Context.Database.Connection;
                var parseViaFrameworkProperty = connection.GetType().GetProperty("ParseViaFramework");

                connection.Open();

                Assert.Multiple(() =>
                {
                    Assert.IsNotNull(connection);
                    Assert.AreEqual("System.Data.SQLite.SQLiteConnection", connection.GetType().FullName);
                    Assert.IsNotNull(parseViaFrameworkProperty);
                    Assert.IsTrue((bool)parseViaFrameworkProperty.GetValue(connection));
                    Assert.AreEqual("Open", connection.State.ToString());
                });
            }
        }
    }
}

