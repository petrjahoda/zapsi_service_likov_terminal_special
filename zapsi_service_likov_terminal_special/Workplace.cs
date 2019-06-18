using System;
using System.Data.SqlClient;
using System.Globalization;
using Microsoft.Extensions.Logging;
using MySql.Data.MySqlClient;
using static System.Console;

namespace zapsi_service_likov_terminal_special {
    public enum StateType {
        Idle,
        Running,
        PowerOff
    }

    public class Workplace {
        public int Oid { get; set; }
        public string Name { get; set; }
        public int DeviceOid { get; set; }
        public int ActualWorkshiftId { get; set; }
        public int PreviousWorkshiftId { get; set; }
        public DateTime LastStateDateTime { get; set; }
        public int DefaultOrder { get; set; }
        public int DefaultProduct { get; set; }
        public int WorkplaceDivisionId { get; set; }
        public int WorkplaceGroupId { get; set; }
        public int ProductionPortOid { get; set; }
        public int CountPortOid { get; set; }
        public int NokPortOid { get; set; }
        public DateTime OrderStartDate { get; set; }
        public int? OrderUserId { get; set; }
        public StateType ActualStateType { get; set; }
        public int WorkplaceIdleId { get; set; }

        public Workplace() {
            Oid = Oid;
            Name = Name;
            DeviceOid = DeviceOid;
            ActualWorkshiftId = ActualWorkshiftId;
            LastStateDateTime = LastStateDateTime;
            DefaultOrder = DefaultOrder;
            DefaultProduct = DefaultProduct;
            WorkplaceDivisionId = WorkplaceDivisionId;
            ProductionPortOid = ProductionPortOid;
            CountPortOid = CountPortOid;
            OrderStartDate = OrderStartDate;
            ActualStateType = ActualStateType;
        }

        private static void LogInfo(string text, ILogger logger) {
            var now = DateTime.Now;
            text = now + " " + text;
            if (Program.OsIsLinux) {
                WriteLine(Program.greenColor + text + Program.resetColor);
            } else {
                logger.LogInformation(text);
                ForegroundColor = ConsoleColor.Green;
                WriteLine(text);
                ForegroundColor = ConsoleColor.White;
            }
        }


        private static void LogError(string text, ILogger logger) {
            var now = DateTime.Now;
            text = now + " " + text;
            if (Program.OsIsLinux) {
                WriteLine(Program.redColor + text + Program.resetColor);
            } else {
                logger.LogInformation(text);
                ForegroundColor = ConsoleColor.Red;
                WriteLine(text);
                ForegroundColor = ConsoleColor.White;
            }
        }

        public void UpdateOrderData(ILogger logger) {
            if (Program.AddCyclesToOrder.Equals("1")) {
                if (Program.DatabaseType.Equals("mysql")) {
                    var connection = new MySqlConnection(
                        $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                    var count = GetCountForWorkplace(logger);
                    var nokCount = GetNokCountForWorkplace(logger);
                    var averageCycle = GetAverageCycleForWorkplace(count);
                    try {
                        connection.Open();
                        var command = connection.CreateCommand();

                        command.CommandText =
                            $"UPDATE `zapsi2`.`terminal_input_order` t SET t.Count={count}, t.Fail={nokCount}, t.AverageCycle={averageCycle} WHERE t.`DTE` is NULL and DeviceID={DeviceOid}";

                        try {
                            command.ExecuteNonQuery();
                        } catch (Exception error) {
                            LogError("[ " + Name + " ] --ERR-- problem updating count order: " + error.Message + command.CommandText, logger);
                        } finally {
                            command.Dispose();
                        }

                        connection.Close();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                    } finally {
                        connection.Dispose();
                    }
                } else if (Program.DatabaseType.Equals("sqlserver")) {
                    var connection = new SqlConnection
                        {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                    var count = GetCountForWorkplace(logger);
                    var averageCycle = GetAverageCycleForWorkplace(count);
                    var nokCount = GetNokCountForWorkplace(logger);

                    try {
                        connection.Open();
                        var command = connection.CreateCommand();

                        command.CommandText =
                            $"UPDATE [dbo].[terminal_input_order] SET [Count]={count}, [Fail]={nokCount}[AverageCycle]={averageCycle} WHERE [DTE] is NULL and DeviceID={DeviceOid}";
                        try {
                            command.ExecuteNonQuery();
                        } catch (Exception error) {
                            LogError("[ " + Name + " ] --ERR-- problem updating count order: " + error.Message + command.CommandText, logger);
                        } finally {
                            command.Dispose();
                        }

                        connection.Close();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                    } finally {
                        connection.Dispose();
                    }
                }
            }
        }

        private object GetNokCountForWorkplace(ILogger logger) {
            var nokCount = 0;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                var startDate = string.Format("{0:yyyy-MM-dd HH:mm:ss.ffff}", OrderStartDate);
                var actualDateTime = DateTime.Now;
                if (Program.TimezoneIsUtc) {
                    actualDateTime = DateTime.UtcNow;
                }

                var endDate = string.Format("{0:yyyy-MM-dd HH:mm:ss.ffff}", actualDateTime);

                try {
                    connection.Open();
                    var selectQuery =
                        $"Select count(oid) as count from zapsi2.device_input_digital where DT>='{startDate}' and DT<='{endDate}' and DevicePortId={NokPortOid} and zapsi2.device_input_digital.Data=1";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            nokCount = Convert.ToInt32(reader["count"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting nok count for workplace: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                var startDate = string.Format("{0:yyyy-MM-dd HH:mm:ss.ffff}", OrderStartDate);
                var actualDateTime = DateTime.Now;
                if (Program.TimezoneIsUtc) {
                    actualDateTime = DateTime.UtcNow;
                }

                var endDate = string.Format("{0:yyyy-MM-dd HH:mm:ss.ffff}", actualDateTime);

                try {
                    connection.Open();
                    var selectQuery =
                        $"Select count(oid) as count from dbo.device_input_digital where DT>='{startDate}' and DT<='{endDate}' and DevicePortId={NokPortOid} and dbo.device_input_digital.Data=1";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            nokCount = Convert.ToInt32(reader["count"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting nok count for workplace: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return nokCount;
        }

        public void CloseIdleForWorkplace(DateTime dateTimeToInsert, ILogger logger) {
            if (Program.CloseOnlyAutomaticIdles.Equals("1")) {
                var myDate = string.Format("{0:yyyy-MM-dd HH:mm:ss}", LastStateDateTime);

                var dateToInsert = string.Format("{0:yyyy-MM-dd HH:mm:ss}", dateTimeToInsert);
                if (LastStateDateTime.CompareTo(dateToInsert) > 0) {
                    dateToInsert = myDate;
                }
                if (Program.DatabaseType.Equals("mysql")) {
                    var connection = new MySqlConnection(
                        $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                    try {
                        connection.Open();
                        var command = connection.CreateCommand();
                        command.CommandText =
                            $"UPDATE `zapsi2`.`terminal_input_idle` t SET t.`DTE` = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)) WHERE t.`DTE` is NULL and DeviceID={DeviceOid} and Note like 'Automatic idle'";
                        try {
                            command.ExecuteNonQuery();
                        } catch (Exception error) {
                            LogError("[ MAIN ] --ERR-- problem closing idle in database: " + error.Message + command.CommandText, logger);
                        } finally {
                            command.Dispose();
                        }

                        connection.Close();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                    } finally {
                        connection.Dispose();
                    }
                } else if (Program.DatabaseType.Equals("sqlserver")) {
                    var connection = new SqlConnection
                        {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                    try {
                        connection.Open();
                        var command = connection.CreateCommand();


                        command.CommandText =
                            $"UPDATE [dbo].[terminal_input_idle] SET [DTE] = '{dateToInsert}', [Interval] = (datediff(second, DTS, '{dateToInsert}')) WHERE [DTE] is NULL and DeviceID={DeviceOid} and Note like 'Automatic idle'";
                        try {
                            command.ExecuteNonQuery();
                        } catch (Exception error) {
                            LogError("[ MAIN ] --ERR-- problem closing idle in database: " + error.Message + command.CommandText, logger);
                        } finally {
                            command.Dispose();
                        }

                        connection.Close();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                    } finally {
                        connection.Dispose();
                    }
                }
            } else {
                var myDate = string.Format("{0:yyyy-MM-dd HH:mm:ss}", LastStateDateTime);
                if (Program.DatabaseType.Equals("mysql")) {
                    var connection = new MySqlConnection(
                        $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                    try {
                        connection.Open();
                        var command = connection.CreateCommand();


                        command.CommandText =
                            $"UPDATE `zapsi2`.`terminal_input_idle` t SET t.`DTE` = '{myDate}', t.Interval = TIME_TO_SEC(timediff('{myDate}', DTS)) WHERE t.`DTE` is NULL and DeviceID={DeviceOid}";
                        try {
                            command.ExecuteNonQuery();
                        } catch (Exception error) {
                            LogError("[ MAIN ] --ERR-- problem closing idle in database: " + error.Message + command.CommandText, logger);
                        } finally {
                            command.Dispose();
                        }

                        connection.Close();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                    } finally {
                        connection.Dispose();
                    }
                } else if (Program.DatabaseType.Equals("sqlserver")) {
                    var connection = new SqlConnection
                        {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                    try {
                        connection.Open();
                        var command = connection.CreateCommand();


                        command.CommandText =
                            $"UPDATE [dbo].[terminal_input_idle] SET [DTE] = '{myDate}', [Interval] = (datediff(second, DTS, '{myDate}')) WHERE [DTE] is NULL and DeviceID={DeviceOid}";
                        try {
                            command.ExecuteNonQuery();
                        } catch (Exception error) {
                            LogError("[ MAIN ] --ERR-- problem closing idle in database: " + error.Message + command.CommandText, logger);
                        } finally {
                            command.Dispose();
                        }

                        connection.Close();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                    } finally {
                        connection.Dispose();
                    }
                }
            }
        }

        public void CreateIdleForWorkplace(ILogger logger, bool workplaceHasActiveOrder, DateTime dateTimeToInsert) {
            var myDate = string.Format("{0:yyyy-MM-dd HH:mm:ss}", dateTimeToInsert);
            var idleOidToInsert = 2;
            var userIdToInsert = "";
            if (workplaceHasActiveOrder) {
                idleOidToInsert = 1;
                userIdToInsert = OrderUserId.ToString();
            }
            if (userIdToInsert.Length == 0) {
                userIdToInsert = DownloadFromLoginTable(logger);
            }
            if (userIdToInsert.Equals("0") || userIdToInsert.Length == 0) {
                userIdToInsert = "NULL";
            }
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection($"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var command = connection.CreateCommand();
                    command.CommandText =
                        $"INSERT INTO `zapsi2`.`terminal_input_idle` (`DTS`, `DTE`, `IdleID`, `UserID`, `Interval`, `DeviceID`, `Note`) VALUES ('{myDate}', NULL , {idleOidToInsert}, {userIdToInsert}, 0, {DeviceOid}, 'Automatic idle')";

                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting idle into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var command = connection.CreateCommand();

                    command.CommandText =
                        $"INSERT INTO [dbo].[terminal_input_idle] ([DTS], [DTE], [IdleID], [UserID], [Interval], [DeviceID], [Note]) VALUES ('{myDate}', NULL , {idleOidToInsert}, {userIdToInsert}, 0, {DeviceOid}, 'Automatic idle')";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting idle into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        private string DownloadFromLoginTable(ILogger logger) {
            var userIdFromLoginTable = "0";
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from zapsi2.terminal_input_login where DTE is NULL and DeviceID={DeviceOid}";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            userIdFromLoginTable = Convert.ToString(reader["UserID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from dbo.terminal_input_login where DTE is null and DeviceID={DeviceOid}";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            userIdFromLoginTable = Convert.ToString(reader["UserID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
            return userIdFromLoginTable;
        }

        public void CloseOrderForWorkplace(DateTime closingDateForOrder, ILogger logger) {
            var myDate = string.Format("{0:yyyy-MM-dd HH:mm:ss}", LastStateDateTime);
            var dateToInsert = string.Format("{0:yyyy-MM-dd HH:mm:ss}", closingDateForOrder);
            if (LastStateDateTime.CompareTo(closingDateForOrder) > 0) {
                dateToInsert = myDate;
            }
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                var count = GetCountForWorkplace(logger);
                var nokCount = GetNokCountForWorkplace(logger);
                var averageCycle = GetAverageCycleForWorkplace(count);
                var averageCycleAsString = averageCycle.ToString(CultureInfo.InvariantCulture).Replace(",", ".");
                try {
                    connection.Open();
                    var command = connection.CreateCommand();
                    command.CommandText =
                        $"UPDATE `zapsi2`.`terminal_input_order` t SET t.`DTE` = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)), t.`Count`={count}, t.Fail={nokCount}, t.averageCycle={averageCycleAsString} WHERE t.`DTE` is NULL and DeviceID={DeviceOid};" +
                        $"UPDATE zapsi2.terminal_input_login t set t.DTE = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)) where t.DTE is null and t.DeviceId={DeviceOid};";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem closing order in database: " + error.Message + "\n" + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }
                    OrderUserId = 0;
                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                var count = GetCountForWorkplace(logger);
                var averageCycle = GetAverageCycleForWorkplace(count);
                var nokCount = GetNokCountForWorkplace(logger);

                try {
                    connection.Open();
                    var command = connection.CreateCommand();

                    command.CommandText =
                        $"UPDATE [dbo].[terminal_input_order] SET [DTE] = '{dateToInsert}', [Interval] = (datediff(second, DTS, '{dateToInsert}')), [Count]={count}, [Fail]={nokCount}, [AverageCycle]={averageCycle} WHERE [DTE] is NULL and DeviceID={DeviceOid};" +
                        $"UPDATE [dbo].[terminal_input_login] set [DTE] = '{dateToInsert}', [Interval] = (datediff(second, DTS, '{dateToInsert}')) where [DTE] is NULL and Deviceid={DeviceOid};";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem closing order in database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        private int GetCountForWorkplace(ILogger logger) {
            var count = 0;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                var startDate = string.Format("{0:yyyy-MM-dd HH:mm:ss.ffff}", OrderStartDate);
                var actualDateTime = DateTime.Now;
                if (Program.TimezoneIsUtc) {
                    actualDateTime = DateTime.UtcNow;
                }

                var endDate = string.Format("{0:yyyy-MM-dd HH:mm:ss.ffff}", actualDateTime);

                try {
                    connection.Open();
                    var selectQuery =
                        $"Select count(oid) as count from zapsi2.device_input_digital where DT>='{startDate}' and DT<='{endDate}' and DevicePortId={CountPortOid} and zapsi2.device_input_digital.Data=1";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            count = Convert.ToInt32(reader["count"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting count for workplace: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                var startDate = string.Format("{0:yyyy-MM-dd HH:mm:ss.ffff}", OrderStartDate);
                var actualDateTime = DateTime.Now;
                if (Program.TimezoneIsUtc) {
                    actualDateTime = DateTime.UtcNow;
                }

                var endDate = string.Format("{0:yyyy-MM-dd HH:mm:ss.ffff}", actualDateTime);

                try {
                    connection.Open();
                    var selectQuery =
                        $"Select count(oid) as count from dbo.device_input_digital where DT>='{startDate}' and DT<='{endDate}' and DevicePortId={CountPortOid} and dbo.device_input_digital.Data=1";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            count = Convert.ToInt32(reader["count"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting count for workplace: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return count;
        }

        private double GetAverageCycleForWorkplace(int count) {
            var averageCycle = 0.0;
            if (count != 0) {
                var difference = LastStateDateTime.Subtract(OrderStartDate).TotalSeconds;
                averageCycle = difference / count;
                if (averageCycle < 0) {
                    averageCycle = 0;
                }
            }
            return averageCycle;
        }

        public void CreateOrderForWorkplace(ILogger logger) {
            var workplaceModeId = GetWorkplaceModeId(logger);
            var myDate = string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now.Subtract(TimeSpan.FromSeconds(10)));
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var command = connection.CreateCommand();
                    command.CommandText =
                        $"INSERT INTO `zapsi2`.`terminal_input_order` (`DTS`, `DTE`, `OrderID`, `UserID`, `DeviceID`, `Interval`, `Count`, `Fail`, `AverageCycle`, `WorkerCount`, `WorkplaceModeID`, `Note`, `WorkshiftID`) VALUES ('{myDate}', NULL, {DefaultOrder}, NULL, {DeviceOid}, 0, DEFAULT, DEFAULT, DEFAULT, DEFAULT, {workplaceModeId}, 'NULL', {ActualWorkshiftId})";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting terminal input order into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var command = connection.CreateCommand();
                    command.CommandText =
                        $"INSERT INTO [dbo].[terminal_input_order] ([DTS], [DTE], [OrderID], [UserID], [DeviceID], [Interval], [Count], [Fail], [AverageCycle], [WorkerCount], [WorkplaceModeID], [Note], [WorkshiftID]) VALUES ('{myDate}', NULL, {DefaultOrder}, 1, {DeviceOid}, 0, DEFAULT, DEFAULT, DEFAULT, DEFAULT, {workplaceModeId}, 'NULL', {ActualWorkshiftId})";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting terminal input order into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        private int GetWorkplaceModeId(ILogger logger) {
            var workplaceModeId = 1;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from zapsi2.workplace_mode where WorkplaceId={Oid}";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            workplaceModeId = Convert.ToInt32(reader["OID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from dbo.workplace_mode where WorkplaceId={Oid}";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            workplaceModeId = Convert.ToInt32(reader["OID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return workplaceModeId;
        }

        public bool CheckIfWorkplaceHasActivedIdle(ILogger logger) {
            var workplaceHasActiveIdle = false;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from zapsi2.terminal_input_idle where DTE is null and  DeviceID={DeviceOid}";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            WorkplaceIdleId = Convert.ToInt32(reader["IdleID"]);
                            workplaceHasActiveIdle = true;
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active idle: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from dbo.[terminal_input_idle] where DTE is NULL and  DeviceID={DeviceOid}";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            workplaceHasActiveIdle = true;
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active idle: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return workplaceHasActiveIdle;
        }

        public bool CheckIfWorkplaceHasActiveOrder(ILogger logger) {
            var workplaceHasActiveOrder = false;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DTE is NULL and DeviceID={DeviceOid}";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            workplaceHasActiveOrder = true;
                            OrderStartDate = Convert.ToDateTime(reader["DTS"]);
                            try {
                                OrderUserId = Convert.ToInt32(reader["UserID"]);
                            } catch (Exception error) {
                                LogInfo("[ " + Name + " ] --INF-- Open order has no user", logger);
                            }
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from dbo.[terminal_input_order] where DTE is NULL and DeviceID={DeviceOid}";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            workplaceHasActiveOrder = true;
                            OrderStartDate = Convert.ToDateTime(reader["DTS"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return workplaceHasActiveOrder;
        }

        public void UpdateActualStateForWorkplace(ILogger logger) {
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    var stateNumber = 1;
                    connection.Open();
                    var selectQuery = $"SELECT * from zapsi2.workplace_state where DTE is NULL and WorkplaceID={Oid}";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            LastStateDateTime = Convert.ToDateTime(reader["DTS"].ToString());
                            stateNumber = Convert.ToInt32(reader["StateID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting actual workplace state: " + error.Message + selectQuery, logger);
                    }

                    selectQuery = $"SELECT * from zapsi2.state where OID={stateNumber}";
                    command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            var stateType = Convert.ToString(reader["Type"]);
                            if (stateType.Equals("running")) {
                                ActualStateType = StateType.Running;
                            } else if (stateType.Equals("idle")) {
                                ActualStateType = StateType.Idle;
                            } else {
                                ActualStateType = StateType.PowerOff;
                            }
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting actual workplace state: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }


                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    var stateNumber = 1;
                    connection.Open();
                    var selectQuery = $"SELECT * from dbo.[workplace_state] where DTE is NULL and WorkplaceID={Oid}";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            LastStateDateTime = Convert.ToDateTime(reader["DTS"].ToString());
                            stateNumber = Convert.ToInt32(reader["StateID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting actual workplace state: " + error.Message + selectQuery, logger);
                    }

                    selectQuery = $"SELECT * from dbo.[state] where OID={stateNumber}";
                    command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            var stateType = Convert.ToString(reader["Type"]);
                            if (stateType.Equals("running")) {
                                ActualStateType = StateType.Running;
                            } else if (stateType.Equals("idle")) {
                                ActualStateType = StateType.Idle;
                            } else {
                                ActualStateType = StateType.PowerOff;
                            }
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting actual workplace state: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }


                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        public int GetActualWorkShiftIdFor(ILogger logger) {
            PreviousWorkshiftId = ActualWorkshiftId;

            var actualWorkShiftId = 0;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                var actualTime = DateTime.Now;
                if (Program.TimezoneIsUtc) {
                    actualTime = DateTime.UtcNow;
                }

                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from zapsi2.workshift where Active=1 and WorkplaceDivisionID is null or WorkplaceDivisionID ={WorkplaceDivisionId}";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            var oid = Convert.ToInt32(reader["OID"]);
                            var shiftStartsAt = reader["WorkshiftStart"].ToString();
                            var shiftDuration = Convert.ToInt32(reader["WorkshiftLenght"]);
                            var shiftStartAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", System.Globalization.CultureInfo.CurrentCulture);
                            if (Program.TimezoneIsUtc) {
                                shiftStartAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", System.Globalization.CultureInfo.CurrentCulture)
                                    .ToUniversalTime();
                            }

                            var shiftEndsAtDateTime = shiftStartAtDateTime.AddMinutes(shiftDuration);
                            if (actualTime.Ticks < shiftStartAtDateTime.Ticks) {
                                actualTime = actualTime.AddDays(1);
                            }

                            if (actualTime.Ticks >= shiftStartAtDateTime.Ticks && actualTime.Ticks < shiftEndsAtDateTime.Ticks) {
                                actualWorkShiftId = oid;
                                LogInfo("[ " + Name + " ] --INF-- Actual workshift id: " + actualWorkShiftId, logger);
                            }
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting actual workshift: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                var actualTime = DateTime.Now;
                if (Program.TimezoneIsUtc) {
                    actualTime = DateTime.UtcNow;
                }

                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from dbo.[workshift] where Active=1 and WorkplaceDivisionID is null or WorkplaceDivisionID ={WorkplaceDivisionId}";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            var oid = Convert.ToInt32(reader["OID"]);
                            var shiftStartsAt = reader["WorkshiftStart"].ToString();
                            var shiftDuration = Convert.ToInt32(reader["WorkshiftLenght"]);
                            var shiftStartAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", System.Globalization.CultureInfo.CurrentCulture);
                            if (Program.TimezoneIsUtc) {
                                shiftStartAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", System.Globalization.CultureInfo.CurrentCulture)
                                    .ToUniversalTime();
                            }

                            var shiftEndsAtDateTime = shiftStartAtDateTime.AddMinutes(shiftDuration);
                            if (actualTime.Ticks < shiftStartAtDateTime.Ticks) {
                                actualTime = actualTime.AddDays(1);
                            }

                            if (actualTime.Ticks >= shiftStartAtDateTime.Ticks && actualTime.Ticks < shiftEndsAtDateTime.Ticks) {
                                actualWorkShiftId = oid;
                                LogInfo("[ " + Name + " ] --INF-- Actual workshift id: " + actualWorkShiftId, logger);
                            }
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting actual workshift: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
            CheckForFirstRunWhenPreviouslyWasNoShift(actualWorkShiftId);
            return actualWorkShiftId;
        }

        private void CheckForFirstRunWhenPreviouslyWasNoShift(int actualWorkShiftId) {
            if (PreviousWorkshiftId == 0) {
                PreviousWorkshiftId = actualWorkShiftId;
            }
        }

        public void AddCountPort(ILogger logger) {
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();

                    var selectQuery = $"select * from zapsi2.workplace_port where WorkplaceID = {Oid} and Type in ('cycle','running') order by Type asc limit 1;";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            CountPortOid = Convert.ToInt32(reader["DevicePortID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem reading from database: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();

                    var selectQuery = $"select top 1 * from dbo.workplace_port where WorkplaceID = {Oid} and Type in ('cycle','running') order by Type asc;";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            CountPortOid = Convert.ToInt32(reader["DevicePortID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem reading from database: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        public void AddProductionPort(ILogger logger) {
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();

                    var selectQuery = $"select * from zapsi2.workplace_port where WorkplaceID = {Oid} and Type in ('cycle','running') order by Type desc limit 1;";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            ProductionPortOid = Convert.ToInt32(reader["DevicePortID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem adding production port: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();

                    var selectQuery = $"select top 1 * from dbo.[workplace_port] where WorkplaceID = {Oid} and Type in ('cycle','running') order by Type desc;";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            ProductionPortOid = Convert.ToInt32(reader["DevicePortID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem adding production port: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        public void CreateDefaultOrder(ILogger logger) {
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var command = connection.CreateCommand();

                    command.CommandText =
                        $"INSERT INTO `zapsi2`.`order` (`Name`, `Barcode`, `ProductID`, `OrderStatusID`, `CountRequested`, `WorkplaceID`) VALUES ('Internal', '1', {DefaultProduct}, 1, 0, {Oid})";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting default order into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var command = connection.CreateCommand();

                    command.CommandText =
                        $"INSERT INTO [dbo].[order] ([Name], [Barcode], [ProductID], [OrderStatusID], [CountRequested], [WorkplaceID]) VALUES ('Internal', '1', {DefaultProduct}, 1, 0, {Oid})";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting default order into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        public bool CheckForDefaultOrder(ILogger logger) {
            var defaultOrderIsInDatabase = false;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    const string selectQuery = @"SELECT * from zapsi2.order where Name like 'Internal' limit 1";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            DefaultOrder = Convert.ToInt32(reader["OID"]);
                            defaultOrderIsInDatabase = true;
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- Problem checking default order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var selectQuery = $"SELECT TOP 1 * from dbo.[order] where Name like 'Internal';";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            DefaultOrder = Convert.ToInt32(reader["OID"]);
                            defaultOrderIsInDatabase = true;
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- Problem checking default order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return defaultOrderIsInDatabase;
        }


        public bool CheckForDefaultProduct(ILogger logger) {
            var defaultProductIsInDatabase = false;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    const string selectQuery = @"SELECT * from zapsi2.product where Name like 'Internal' limit 1";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            DefaultProduct = Convert.ToInt32(reader["OID"]);
                            defaultProductIsInDatabase = true;
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- Problem checking default product: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};
                try {
                    connection.Open();
                    var selectQuery = $"SELECT TOP 1 * from dbo.[product] where Name like 'Internal'";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            DefaultProduct = Convert.ToInt32(reader["OID"]);
                            defaultProductIsInDatabase = true;
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- Problem checking default product: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return defaultProductIsInDatabase;
        }

        public void CreateDefaultProduct(ILogger logger) {
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var command = connection.CreateCommand();

                    command.CommandText =
                        ("INSERT INTO `zapsi2`.`product` (`Name`, `Barcode`, `Cycle`, `IdleFromTime`, `ProductStatusID`, `Deleted`, `ProductGroupID`) VALUES ('Internal', '1', 1, 10, 1, 0, NULL)"
                        );
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting default product into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ MAIN ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};
                try {
                    connection.Open();
                    var command = connection.CreateCommand();

                    command.CommandText =
                        $"INSERT INTO [dbo].[product] ([Name], [Barcode], [Cycle], [IdleFromTime], [ProductStatusID], [Deleted], [ProductGroupID]) VALUES ('Internal', '1', 1, 10, 1, 0, NULL)";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting default product into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ MAIN ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        public bool CheckForOneFromLastOrderDte(ILogger logger) {
            var productionIsOne = false;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var selectQuery =
                        $"Select * from zapsi2.device_input_digital where DevicePortID={ProductionPortOid} and DT > (select DTE from terminal_input_order where DeviceID = {DeviceOid} order by DTS desc limit 1) order by DT desc limit 2";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            var data = Convert.ToInt32(reader["Data"]);
                            var lastDataDateTime = Convert.ToDateTime(reader["DT"]);
                            if (Program.TimezoneIsUtc) {
                                if (data == 1 && (DateTime.UtcNow - lastDataDateTime).TotalSeconds < 10) {
                                    productionIsOne = true;
                                }
                            } else {
                                if (data == 1 && (DateTime.Now - lastDataDateTime).TotalSeconds < 10) {
                                    productionIsOne = true;
                                }
                            }
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting count for workplace: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var selectQuery =
                        $"Select top 2 from dbo.device_input_digital where DevicePortID={ProductionPortOid} and DT > (select top1 DTE from terminal_input_order where DeviceID = {DeviceOid} order by DTS desc) order by DT desc";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            var data = Convert.ToInt32(reader["Data"]);
                            var lastDataDateTime = Convert.ToDateTime(reader["DT"]);
                            if (data == 1 && (DateTime.Now - lastDataDateTime).TotalSeconds < 10) {
                                productionIsOne = true;
                            }
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem getting count for workplace: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return productionIsOne;
        }

        public void AddFailPort(ILogger logger) {
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();

                    var selectQuery = $"select * from zapsi2.workplace_port where WorkplaceID = {Oid} and Type in ('fail') order by Type asc limit 1;";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            NokPortOid = Convert.ToInt32(reader["DevicePortID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem reading from database: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();

                    var selectQuery = $"select top 1 * from dbo.workplace_port where WorkplaceID = {Oid} and Type in ('fail') order by Type asc;";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        while (reader.Read()) {
                            NokPortOid = Convert.ToInt32(reader["DevicePortID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem reading from database: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }


        public void CloseAndStartOrderForWorkplaceAt(DateTime closeAndStartOrderDateTime, ILogger logger) {
            var orderId = GetOrderId(logger);
            var userId = OrderUserId;
            var anyOrderisOpen = orderId != 0;
            if (anyOrderisOpen) {
                CloseOrderForWorkplace(closeAndStartOrderDateTime, logger);
                LogInfo("[ " + Name + " ] --INF-- Order closed, with ID " + orderId, logger);
                CreateOrderForWorkplace(closeAndStartOrderDateTime, orderId, userId, 1, logger);
                LogInfo("[ " + Name + " ] --INF-- New order opened", logger);
            }
        }

        private void CreateOrderForWorkplace(DateTime startDateForOrder, int orderId, int? userId, int workplaceModeId, ILogger logger) {
            var userToInsert = "NULL";
            if (userId != null) {
                userToInsert = userId.ToString();
            }
            var dateToInsert = string.Format("{0:yyyy-MM-dd HH:mm:ss}", startDateForOrder);
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var command = connection.CreateCommand();
                    command.CommandText =
                        $"INSERT INTO `zapsi2`.`terminal_input_order` (`DTS`, `DTE`, `OrderID`, `UserID`, `DeviceID`, `Interval`, `Count`, `Fail`, `AverageCycle`, `WorkerCount`, `WorkplaceModeID`, `Note`, `WorkshiftID`) VALUES ('{dateToInsert}', NULL, {orderId}, {userToInsert}, {DeviceOid}, 0, DEFAULT, DEFAULT, DEFAULT, DEFAULT, {workplaceModeId}, 'NULL', {ActualWorkshiftId})";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting terminal input order into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var command = connection.CreateCommand();
                    command.CommandText =
                        $"INSERT INTO [dbo].[terminal_input_order] ([DTS], [DTE], [OrderID], [UserID], [DeviceID], [Interval], [Count], [Fail], [AverageCycle], [WorkerCount], [WorkplaceModeID], [Note], [WorkshiftID]) VALUES ('{dateToInsert}', NULL, {orderId}, {userToInsert}, {DeviceOid}, 0, DEFAULT, DEFAULT, DEFAULT, DEFAULT, {workplaceModeId}, 'NULL', {ActualWorkshiftId})";
                    try {
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting terminal input order into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        private int GetOrderId(ILogger logger) {
            var actualOrderId = 0;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DeviceID={DeviceOid} and DTE is null";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            actualOrderId = Convert.ToInt32(reader["OrderID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from dbo.terminal_input_order where DeviceID={DeviceOid} and DTE is null";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            actualOrderId = Convert.ToInt32(reader["OrderID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return actualOrderId;
        }


        private int GetIdleId(ILogger logger) {
            var actualIdleId = 0;
            if (Program.DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection(
                    $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from zapsi2.terminal_input_idle where DeviceID={DeviceOid} and DTE is null";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            actualIdleId = Convert.ToInt32(reader["IdleID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active idle: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            } else if (Program.DatabaseType.Equals("sqlserver")) {
                var connection = new SqlConnection
                    {ConnectionString = $"Data Source={Program.IpAddress}; Initial Catalog={Program.Database}; User id={Program.Login}; Password={Program.Password};"};

                try {
                    connection.Open();
                    var selectQuery = $"SELECT * from dbo.terminal_input_idle where DeviceID={DeviceOid} and DTE is null";
                    var command = new SqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            actualIdleId = Convert.ToInt32(reader["IdleID"]);
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ " + Name + " ] --ERR-- Problem checking active idle: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }

            return actualIdleId;
        }

        public int OpenOrderState(ILogger logger) {
            var workplaceModeId = 1;
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DTE is NULL and DeviceID={DeviceOid}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        workplaceModeId = Convert.ToInt32(reader["WorkplaceModeID"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }
            return workplaceModeId;
        }

        public bool HasSignalInOneInLastTenSeconds(ILogger logger) {
            bool workplaceHasSignalInOneInLasttenSeconds = false;
            var lastValueDateTime = DateTime.Now;
            var lastValue = 0;
            var connection = new MySqlConnection($"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.device_input_digital where deviceportID={ProductionPortOid} order by DT desc limit 1";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        lastValue = Convert.ToInt32(reader["Data"]);
                        lastValueDateTime = Convert.ToDateTime(reader["DT"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            if (lastValue == 1) {
                if ((DateTime.Now - lastValueDateTime).TotalSeconds < 10) {
                    workplaceHasSignalInOneInLasttenSeconds = true;
                }
            }

            return workplaceHasSignalInOneInLasttenSeconds;
        }

        public bool IsInProductionForMoreThanTenMinutes(ILogger logger) {
            bool workplaceIsInProduction = false;
            var stateId = 1;
            var stateDateTimeStart = DateTime.Now;
            var connection = new MySqlConnection($"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.workplace_state where WorkplaceID={Oid} and DTE is null";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        stateId = Convert.ToInt32(reader["StateID"]);
                        stateDateTimeStart = Convert.ToDateTime(reader["DTS"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }
            if (stateId == 1) {
                if ((DateTime.Now - stateDateTimeStart).TotalMinutes > 10) {
                    workplaceIsInProduction = true;
                }
            }
            return workplaceIsInProduction;
        }


        public bool TimeIsFifteenMinutesBeforeShiftCloses(ILogger logger) {
            var timeIsFifteenMinutesBeforeShiftCloses = false;
            var connection = new MySqlConnection($"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            DateTime shiftStartsAtDateTime = DateTime.Now;
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.workshift where OID = {ActualWorkshiftId}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    while (reader.Read()) {
                        var shiftStartsAt = reader["WorkshiftStart"].ToString();
                        shiftStartsAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", CultureInfo.CurrentCulture);
                        if (Program.TimezoneIsUtc) {
                            shiftStartsAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", System.Globalization.CultureInfo.CurrentCulture).ToUniversalTime();
                        }
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem getting actual workshift: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }
            LogInfo($"[ {Name} ] --INF-- Workshift time: {shiftStartsAtDateTime.Hour}:{shiftStartsAtDateTime.Minute}", logger);
            LogInfo($"[ {Name} ] --INF-- Actual time: {DateTime.Now.Hour}:{DateTime.Now.Minute}", logger);
            if (shiftStartsAtDateTime.Hour - DateTime.Now.Hour == 1 && DateTime.Now.Minute > 44) {
                LogInfo($"[ {Name} ] --INF-- It is less the 15 minutes before shifts end", logger);
                timeIsFifteenMinutesBeforeShiftCloses = true;
            } else {
                LogInfo($"[ {Name} ] --INF-- It is more the 15 minutes before shifts end", logger);
            }

            return timeIsFifteenMinutesBeforeShiftCloses;
        }

        public bool HasOpenOrderWithStartBeforeThoseFifteenMinutes(ILogger logger) {
            var thereIsOpenOrder = false;
            var workplaceHasOpenOrderWithStartBeforeFifteenMinutesToShiftsEnd = false;
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DTE is NULL and DeviceID={DeviceOid}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        OrderStartDate = Convert.ToDateTime(reader["DTS"]);
                        thereIsOpenOrder = true;
                    } 

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }
            LogInfo($"[ {Name} ] --INF-- Order start time: {OrderStartDate.Hour}:{OrderStartDate.Minute}", logger);
            if (thereIsOpenOrder && OrderStartDate.Minute < 45) {
                LogInfo($"[ {Name} ] --INF-- Order started before 15 minutes shift end interval", logger);
                workplaceHasOpenOrderWithStartBeforeFifteenMinutesToShiftsEnd = true;
            } else {
                LogInfo($"[ {Name} ] --INF-- Order", logger);
            }

            return workplaceHasOpenOrderWithStartBeforeFifteenMinutesToShiftsEnd;
        }

        public bool HasOpenOrderForMoreThanTenMinutes(ILogger logger) {
            var orderIsOpenForMoreThanTenMinutes = false;
            var workplaceHasOpenOrder = CheckIfWorkplaceHasActiveOrder(logger);
            if (workplaceHasOpenOrder) {
                if ((DateTime.Now - OrderStartDate).TotalMinutes > 10) {
                    orderIsOpenForMoreThanTenMinutes = true;
                }
            }
            return orderIsOpenForMoreThanTenMinutes;
        }
    }
}