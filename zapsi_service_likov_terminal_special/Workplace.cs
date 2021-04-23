using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
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
        public string Code  { get; set; }

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

        private object GetNokCountForWorkplace(ILogger logger) {
            var nokCount = 0;
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

            return nokCount;
        }

        public void CloseOrderForWorkplace(DateTime closingDateForOrder, bool closeUserLogin, ILogger logger) {
            var dateToInsert = string.Format("{0:yyyy-MM-dd HH:mm:ss}", closingDateForOrder);
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            var count = 0;
            var nokCount = 0;
            var averageCycleAsString = "0.0";
            try {
                connection.Open();
                var command = connection.CreateCommand();
                if (closeUserLogin) {
                    LogInfo("[ " + Name + " ] --INF-- Closing order and login", logger);
                    command.CommandText =
                        $"UPDATE `zapsi2`.`terminal_input_order` t SET t.`DTE` = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)), t.`Count`={count}, t.Fail={nokCount}, t.averageCycle={averageCycleAsString} WHERE t.`DTE` is NULL and DeviceID={DeviceOid};UPDATE zapsi2.terminal_input_login t set t.DTE = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)) where t.DTE is null and t.DeviceId={DeviceOid};";
                } else {
                    LogInfo("[ " + Name + " ] --INF-- Closing order", logger);
                    command.CommandText =
                        $"UPDATE `zapsi2`.`terminal_input_order` t SET t.`DTE` = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)), t.`Count`={count}, t.Fail={nokCount}, t.averageCycle={averageCycleAsString} WHERE t.`DTE` is NULL and DeviceID={DeviceOid};";
                }

                LogInfo("[ " + Name + " ] --INF-- " + command.CommandText, logger);
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
        }


        private int GetCountForWorkplace(ILogger logger) {
            var count = 0;
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

        public bool CheckIfWorkplaceHasActiveOrder(ILogger logger) {
            var workplaceHasActiveOrder = false;
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
                        } catch (Exception) {
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

            return workplaceHasActiveOrder;
        }

        public void UpdateActualStateForWorkplace(ILogger logger) {
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
        }

        public int GetActualWorkShiftIdFor(ILogger logger) {
            PreviousWorkshiftId = ActualWorkshiftId;

            var actualWorkShiftId = 0;
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
                        var shiftStartAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", CultureInfo.CurrentCulture);
                        if (Program.TimezoneIsUtc) {
                            shiftStartAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", CultureInfo.CurrentCulture)
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

            CheckForFirstRunWhenPreviouslyWasNoShift(actualWorkShiftId);
            return actualWorkShiftId;
        }

        private void CheckForFirstRunWhenPreviouslyWasNoShift(int actualWorkShiftId) {
            if (PreviousWorkshiftId == 0) {
                PreviousWorkshiftId = actualWorkShiftId;
            }
        }

        public void AddCountPort(ILogger logger) {
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
        }

        public void AddProductionPort(ILogger logger) {
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
        }


        public void AddFailPort(ILogger logger) {
            var connection = new MySqlConnection($"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
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
        }


        public void CloseAndStartOrderForWorkplaceAt(DateTime closeAndStartOrderDateTime, ILogger logger) {
            var orderId = GetOrderId(logger);
            var userId = GetUserId(logger);
            var actualOpenTerminalInputOrder = GetTerminalInputOrderId(logger);
            var users = GetUsersForTerminalInputOrderId(actualOpenTerminalInputOrder, logger);

            var anyOrderisOpen = orderId != 0;
            if (anyOrderisOpen) {
                CloseOrderForWorkplace(closeAndStartOrderDateTime, false, logger);
                LogInfo($"[ {Name} ] --INF-- Order closed, with ID {orderId} and user ID {userId} and datetime " + closeAndStartOrderDateTime.ToString(CultureInfo.InvariantCulture), logger);
                CreateOrderForWorkplace(closeAndStartOrderDateTime, orderId, userId, 1, logger);
                LogInfo("[ " + Name + " ] --INF-- New order opened, updating terminal input order user", logger);
                actualOpenTerminalInputOrder = GetTerminalInputOrderId(logger);
                UpdateTerminalInputOrderUser(actualOpenTerminalInputOrder, users, logger);
                LogInfo("[ " + Name + " ] --INF-- Terminal input order user updated", logger);
            }
        }

        private void UpdateTerminalInputOrderUser(int actualOpenTerminalInputOrder, List<int> users, ILogger logger) {
            var connection = new MySqlConnection($"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var command = connection.CreateCommand();
                foreach (var user in users) {
                    command.CommandText = $"INSERT INTO `zapsi2`.`terminal_input_order_user` (`TerminalInputOrderID`, `UserID`) VALUES ({actualOpenTerminalInputOrder}, {user})";
                    try {
                        LogInfo("[ " + Name + " ] --INF-- " + command.CommandText, logger);
                        command.ExecuteNonQuery();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- problem inserting terminal input order user into database: " + error.Message + command.CommandText, logger);
                    } finally {
                        command.Dispose();
                    }
                }
                connection.Close();
            } catch (Exception error) {
                LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }
        }

        private List<int> GetUsersForTerminalInputOrderId(int terminalInputOrderId, ILogger logger) {
            var users = new List<int>();
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * FROM terminal_input_order_user WHERE TerminalInputOrderID = {terminalInputOrderId}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    while (reader.Read()) {
                        users.Add(Convert.ToInt32(reader["UserID"]));
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem checking active terminal input order user: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            return users;
        }

        private int GetTerminalInputOrderId(ILogger logger) {
            var actualTerminalInputOrderId = 0;
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DeviceID={DeviceOid} and DTE is null limit 1";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        actualTerminalInputOrderId = Convert.ToInt32(reader["OID"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem checking active terminal input order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            return actualTerminalInputOrderId;
        }

        private int GetUserId(ILogger logger) {
            var actualUserId = 0;
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DeviceID={DeviceOid} and DTE is null";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        actualUserId = Convert.ToInt32(reader["UserID"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem checking active user: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            return actualUserId;
        }

        private void CreateOrderForWorkplace(DateTime startDateForOrder, int orderId, int? userId, int workplaceModeId, ILogger logger) {
            var userToInsert = "NULL";
            if (userId != null) {
                userToInsert = userId.ToString();
            }

            var dateToInsert = string.Format("{0:yyyy-MM-dd HH:mm:ss}", startDateForOrder);
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText =
                    $"INSERT INTO `zapsi2`.`terminal_input_order` (`DTS`, `DTE`, `OrderID`, `UserID`, `DeviceID`, `Interval`, `Count`, `Fail`, `AverageCycle`, `WorkerCount`, `WorkplaceModeID`, `Note`, `WorkshiftID`) " +
                    $"VALUES ('{dateToInsert}', NULL, {orderId}, {userToInsert}, {DeviceOid}, 0, DEFAULT, DEFAULT, DEFAULT, DEFAULT, {workplaceModeId}, 'NULL', {ActualWorkshiftId})";
                try {
                    LogInfo("[ " + Name + " ] --INF-- " + command.CommandText, logger);
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

        private int GetOrderId(ILogger logger) {
            var actualOrderId = 0;
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

            return actualOrderId;
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

            LogInfo("[ " + Name + " ] --INF-- Open order has workplacemode: " + workplaceModeId, logger);

            return workplaceModeId;
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

                    LogInfo("[ " + Name + " ] --INF-- StateId: " + stateId, logger);
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
            var shiftEndsAt = DateTime.Now;
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.workshift where OID = {ActualWorkshiftId}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    while (reader.Read()) {
                        var shiftStartsAt = reader["WorkshiftStart"].ToString();
                        var shiftLength = Convert.ToInt32(reader["WorkshiftLenght"].ToString());
                        var shiftStartsAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", CultureInfo.CurrentCulture);
                        if (Program.TimezoneIsUtc) {
                            shiftStartsAtDateTime = DateTime.ParseExact(shiftStartsAt, "HH:mm:ss", CultureInfo.CurrentCulture).ToUniversalTime();
                        }

                        shiftEndsAt = shiftStartsAtDateTime.AddMinutes(shiftLength);
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

            LogInfo($"[ {Name} ] --INF-- Workshift time: {shiftEndsAt.Hour}:{shiftEndsAt.Minute}", logger);
            LogInfo($"[ {Name} ] --INF-- Actual time: {DateTime.Now.Hour}:{DateTime.Now.Minute}", logger);
            if (shiftEndsAt.Hour - DateTime.Now.Hour == 1 && DateTime.Now.Minute > 44) {
                LogInfo($"[ {Name} ] --INF-- It is less the 15 minutes before shifts end", logger);
                timeIsFifteenMinutesBeforeShiftCloses = true;
            } else {
                LogInfo($"[ {Name} ] --INF-- It is more the 15 minutes before shifts end", logger);
            }

            return timeIsFifteenMinutesBeforeShiftCloses;
        }

        public bool HasOpenOrderWithStartBeforeThoseFifteenMinutes(ILogger logger) {
            var thereIsOpenOrder = false;
            var openOrderId = 0;
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
                        openOrderId = Convert.ToInt32(reader["OID"]);
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

            if (thereIsOpenOrder && (DateTime.Now - OrderStartDate).TotalMinutes > 15) {
                LogInfo($"[ {Name} ] --INF-- Order started before 15 minutes shift end interval, id: " + openOrderId, logger);
                workplaceHasOpenOrderWithStartBeforeFifteenMinutesToShiftsEnd = true;
            } else {
                LogInfo($"[ {Name} ] --INF-- Order starts in interval 15 minutes before shifts end, id: " + openOrderId, logger);
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

        public int GetWorkplaceMode(ILogger logger) {
            var workplaceModeId = 0;
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * FROM zapsi2.workplace_mode where WorkplaceModeTypeId=3 and WorkplaceId={Oid}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        workplaceModeId = Convert.ToInt32(reader["OID"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem checking workplace mode: " + error.Message + selectQuery, logger);
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

        public bool IsInProduction(ILogger logger) {
            bool workplaceIsInProduction = false;
            var stateId = 1;
            var connection = new MySqlConnection($"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.workplace_state where WorkplaceID={Oid} and DTE is null";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        stateId = Convert.ToInt32(reader["StateID"]);
                    }

                    LogInfo("[ " + Name + " ] --INF-- StateId: " + stateId, logger);
                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + Name + " ] --ERR-- Problem checking production state: " + error.Message + selectQuery, logger);
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
                workplaceIsInProduction = true;
            }

            return workplaceIsInProduction;
        }

        public void CloseLoginForWorkplace(DateTime closingDateForOrder, ILogger logger) {
            var dateToInsert = string.Format("{0:yyyy-MM-dd HH:mm:ss}", closingDateForOrder);
            var connection = new MySqlConnection($"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var command = connection.CreateCommand();
                LogInfo("[ " + Name + " ] --INF-- Closing only login", logger);
                command.CommandText =
                    $"UPDATE zapsi2.terminal_input_login t set t.DTE = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)) where t.DTE is null and t.DeviceId={DeviceOid};";

                LogInfo("[ " + Name + " ] --INF-- " + command.CommandText, logger);
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
        }

        public void CloseOrderForWorkplaceBeforeFifteenMinutes(DateTime closingDateForOrder, bool closeUserLogin, ILogger logger) {
            var myDate = string.Format("{0:yyyy-MM-dd HH:mm:ss}", LastStateDateTime);
            var dateToInsert = string.Format("{0:yyyy-MM-dd HH:mm:ss}", closingDateForOrder);
            if (LastStateDateTime.CompareTo(closingDateForOrder) > 0) {
                dateToInsert = myDate;
            }

            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            var count = GetCountForWorkplace(logger);
            var nokCount = GetNokCountForWorkplace(logger);
            var averageCycle = GetAverageCycleForWorkplace(count);
            var averageCycleAsString = averageCycle.ToString(CultureInfo.InvariantCulture).Replace(",", ".");
            try {
                connection.Open();
                var command = connection.CreateCommand();
                if (closeUserLogin) {
                    LogInfo("[ " + Name + " ] --INF-- Closing order and login", logger);
                    command.CommandText =
                        $"UPDATE `zapsi2`.`terminal_input_order` t SET t.`DTE` = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)), t.`Count`={count}, t.Fail={nokCount}, t.averageCycle={averageCycleAsString} WHERE t.`DTE` is NULL and DeviceID={DeviceOid};UPDATE zapsi2.terminal_input_login t set t.DTE = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)) where t.DTE is null and t.DeviceId={DeviceOid};";
                } else {
                    LogInfo("[ " + Name + " ] --INF-- Closing order", logger);
                    command.CommandText =
                        $"UPDATE `zapsi2`.`terminal_input_order` t SET t.`DTE` = '{dateToInsert}', t.Interval = TIME_TO_SEC(timediff('{dateToInsert}', DTS)), t.`Count`={count}, t.Fail={nokCount}, t.averageCycle={averageCycleAsString} WHERE t.`DTE` is NULL and DeviceID={DeviceOid};";
                }

                LogInfo("[ " + Name + " ] --INF-- " + command.CommandText, logger);
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
        }

        public void SendXml(string destinationUrl, string requestXml, ILogger logger) {
            LogInfo($"[ {Name} ] --INF-- Sending XML", logger);
            HttpWebRequest request = (HttpWebRequest) WebRequest.Create(destinationUrl);
            byte[] bytes = Encoding.UTF8.GetBytes(requestXml);
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = bytes.Length;
            request.Method = "POST";
            try {
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(bytes, 0, bytes.Length);
                HttpWebResponse response;
                response = (HttpWebResponse) request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK) {
                    LogInfo($"[ {Name} ] --INF-- XML sent OK", logger);
                    return;
                }
            
                LogInfo($"[ {Name} ] --INF-- XML not sent!!!", logger);
            } catch {
                LogInfo($"[ {Name} ] --INF-- XML not sent!!!", logger);
            }
        }
    }
}
