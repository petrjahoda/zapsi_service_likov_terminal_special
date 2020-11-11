using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using static System.Console;

namespace zapsi_service_likov_terminal_special {
    class Program {
        private const string BuildDate = "2020.4.2.11";
        private const string NavUrl = "http://localhost:8000/send";
        private const string DataFolder = "Logs";
        private const double InitialDownloadValue = 1000;

        private static bool _systemIsActivated;
        private static bool _databaseIsOnline;
        internal static bool OsIsLinux;
        private static int _numberOfRunningWorkplaces;
        private static bool _databaseOfflineEmailWasSent;
        private static bool _loopCanRun;
        internal static bool TimezoneIsUtc;
        private static bool _swConfigCreated;


        public static string redColor = "\u001b[31;1m";
        public static string greenColor = "\u001b[32;1m";
        public static string yellowColor = "\u001b[33;1m";
        private static string cyanColor = "\u001b[36;1m";
        public static string resetColor = "\u001b[0m";


        private static string _smtpClient;
        private static string _smtpPort;
        private static string _smtpUsername;
        private static string _smtpPassword;
        internal static string IpAddress;
        internal static string Port;
        internal static string Database;
        internal static string Login;
        public static string Password;
        private static string _customer;
        private static string _email;
        private static string _downloadEvery;
        private static string _deleteFilesAfterDays;
        public static string DatabaseType;
        public static string CloseOnlyAutomaticIdles;
        public static string AddCyclesToOrder;

        static void Main() {
            _systemIsActivated = false;
            PrintSoftwareLogo();
            var outputPath = CreateLogFileIfNotExists("0-main.txt");
            using (CreateLogger(outputPath, out var logger)) {
                CheckOsPlatform(logger);
                LogInfo("[ MAIN ] --INF-- Program built at: " + BuildDate, logger);
                CreateConfigFileIfNotExists(logger);
                LoadSettingsFromConfigFile(logger);
                SendEmail("Computer: " + Environment.MachineName + ", User: " + Environment.UserName + ", Program started at " + DateTime.Now + ", Version " + BuildDate, logger);
                var timer = new System.Timers.Timer(InitialDownloadValue);
                timer.Elapsed += (sender, e) => {
                    timer.Interval = Convert.ToDouble(_downloadEvery);
                    RunWorkplaces(logger);
                };
                RunTimer(timer);
            }
        }

        private static void RunWorkplaces(ILogger logger) {
            _systemIsActivated = true;
            CheckDatabaseConnection(logger);
            DeleteOldLogFiles(logger);
            if (_databaseIsOnline) {
                CheckNumberOfActiveWorkplaces(logger);
                CheckSystemTimeZone(logger);

                LogInfo("[ MAIN ] --INF-- Database available: " + _databaseIsOnline + ", active workplaces: " + _numberOfRunningWorkplaces, logger);
                if (!_swConfigCreated) {
                    var optionList = GetDataFromSwConfigTable(logger);
                    var enumerable = optionList as string[] ?? optionList.ToArray();
                    if (!enumerable.Contains("CustomerName")) {
                        CreateNewConfigRecord("CustomerName", logger);
                    }

                    if (!enumerable.Contains("ActivationKey")) {
                        CreateNewConfigRecord("ActivationKey", logger);
                    }

                    if (!enumerable.Contains("Email")) {
                        CreateNewConfigRecord("Email", logger);
                    }

                    if (!enumerable.Contains("CloseOnlyAutomaticIdles")) {
                        CreateNewConfigRecord("CloseOnlyAutomaticIdles", logger);
                    }

                    if (!enumerable.Contains("AddCyclesToOrder")) {
                        CreateNewConfigRecord("AddCyclesToOrder", logger);
                    }

                    if (!enumerable.Contains("SmtpClient")) {
                        CreateNewConfigRecord("SmtpClient", logger);
                    }

                    if (!enumerable.Contains("SmtpPort")) {
                        CreateNewConfigRecord("SmtpPort", logger);
                    }

                    if (!enumerable.Contains("SmtpUsername")) {
                        CreateNewConfigRecord("SmtpUsername", logger);
                    }

                    if (!enumerable.Contains("SmtpPassword")) {
                        CreateNewConfigRecord("SmtpPassword", logger);
                    }

                    _swConfigCreated = true;
                }
            }

            if (_databaseIsOnline && _numberOfRunningWorkplaces == 0 && _systemIsActivated) {
                LogInfo("[ MAIN ] --INF-- Running main loop ", logger);
                var listOfWorkplaces = GetListOfWorkplacesFromDatabase(logger);
                _numberOfRunningWorkplaces = listOfWorkplaces.Count;
                foreach (var workplace in listOfWorkplaces) {
                    LogInfo("[ MAIN ] --INF-- Starting workplace: " + workplace.Name, logger);
                    Task.Run(() => RunWorkplace(workplace));
                }
            }
        }

        private static void RunWorkplace(Workplace workplace) {
            var outputPath = CreateLogFileIfNotExists(workplace.Oid + "-" + workplace.Name + ".txt");
            var closeOpenOrders = true;
            using (var factory = CreateLogger(outputPath, out var logger)) {
                LogDeviceInfo("[ " + workplace.Name + " ] --INF-- Started running", logger);
                var timer = Stopwatch.StartNew();
                while (_databaseIsOnline && _loopCanRun && _systemIsActivated) {
                    LogDeviceInfo("[ " + workplace.Name + " ] --INF-- Inside loop started", logger);
                    UpdateWorkplace(workplace, logger);
                    var workplaceMode = workplace.GetWorkplaceMode(logger);
                    LogDeviceInfo("[ " + workplace.Name + " ] --INF-- Workplacemode number: " + workplaceMode, logger);
                    if (workplace.OpenOrderState(logger) == workplaceMode) {
                        LogDeviceInfo("[ " + workplace.Name + " ] --INF-- Open order has mode Serizeni", logger);
                        if (workplace.WorkplaceDivisionId == 2) {
                            LogDeviceInfo("[ " + workplace.Name + " ] --INF-- WorkplaceDivision is 2", logger);
                            if (workplace.IsInProduction(logger)) {
                                LogDeviceInfo("[ " + workplace.Name + " ] --INF-- Workplace is in production", logger);
                                var userId = GetUserIdFor(workplace, logger);
                                var actualOrderId = GetOrderIdFor(workplace, logger);
                                var orderNo = GetOrderNo(workplace, actualOrderId, logger);
                                var operationNo = GetOperationNo(workplace, actualOrderId, logger);
                                var workcenter = GetWorkcenter(workplace, userId, logger);
                                var consOfMeters = GetConsOfMetersFor(workplace, logger);
                                var motorHours = GetMotorHoursFor(workplace, logger);
                                var cuts = GetCutsFor(workplace, logger);

                                var orderData = "xml=" +
                                                "<ZAPSIoperations>" +
                                                "<ZAPSIoperation>" +
                                                "<type>AL</type>" +
                                                "<orderno>" + orderNo + "</orderno>" +
                                                "<operationno>" + operationNo + "</operationno>" +
                                                "<workcenter>" + workcenter + "</workcenter>" +
                                                "<machinecenter>" + workplace.Code + "</machinecenter>" +
                                                "<operationtype>Technology</operationtype>" +
                                                "<initiator>True</initiator>" +
                                                "<startdate>" + DateTime.Now.ToString(CultureInfo.InvariantCulture) + ".000</startdate>" +
                                                "<enddate/>" +
                                                "<consofmeters>" + consOfMeters + "</consofmeters>" +
                                                "<motorhours>" + motorHours + "</motorhours>" +
                                                "<cuts>" + cuts + "</cuts>" +
                                                "<note/>" +
                                                "</ZAPSIoperation>" +
                                                "</ZAPSIoperations>";
                                SendXml(NavUrl, orderData);
                                var listOfUsers = GetAdditionalUsersFor(workplace, logger);
                                foreach (var actualUserId in listOfUsers) {
                                    workcenter = GetWorkcenter(workplace, actualUserId, logger);
                                    var userData = "xml=" +
                                                   "<ZAPSIoperations>" +
                                                   "<ZAPSIoperation>" +
                                                   "<type>AL</type>" +
                                                   "<orderno>" + orderNo + "</orderno>" +
                                                   "<operationno>" + operationNo + "</operationno>" +
                                                   "<workcenter>" + workcenter + "</workcenter>" +
                                                   "<machinecenter>" + workplace.Code + "</machinecenter>" +
                                                   "<operationtype>Production</operationtype>" +
                                                   "<initiator>True</initiator>" +
                                                   "<startdate>"+ DateTime.Now.ToString(CultureInfo.InvariantCulture) + ".000</startdate>" +
                                                   "<enddate/>" +
                                                   "<consofmeters/>" +
                                                   "<motorhours/>" +
                                                   "<cuts/>" +
                                                   "<note/>" +
                                                   "</ZAPSIoperation>" +
                                                   "</ZAPSIoperations>";
                                    SendXml(NavUrl, userData);
                                }

                                workplace.CloseAndStartOrderForWorkplaceAt(DateTime.Now, logger);
                            }
                        } else if (workplace.WorkplaceDivisionId == 3) {
                            LogDeviceInfo("[ " + workplace.Name + " ] --INF-- WorkplaceDivision is 3", logger);
                            if (workplace.IsInProductionForMoreThanTenMinutes(logger)) {
                                LogDeviceInfo("[ " + workplace.Name + " ] --INF-- Is in production for more than 10 minutes", logger);
                                if (workplace.HasOpenOrderForMoreThanTenMinutes(logger)) {
                                    LogDeviceInfo("[ " + workplace.Name + " ] --INF-- Workplace has open order for more than 10 minutes", logger);
                                    var actualOrderId = GetOrderIdFor(workplace, logger);
                                    var userId = GetUserIdFor(workplace, logger);
                                    var orderNo = GetOrderNo(workplace, actualOrderId, logger);
                                    var operationNo = GetOperationNo(workplace, actualOrderId, logger);
                                    var workcenter = GetWorkcenter(workplace, userId, logger);
                                    var consOfMeters = GetConsOfMetersFor(workplace, logger);
                                    var motorHours = GetMotorHoursFor(workplace, logger);
                                    var cuts = GetCutsFor(workplace, logger);

                                    var orderData = "xml=" +
                                                    "<ZAPSIoperations>" +
                                                    "<ZAPSIoperation>" +
                                                    "<type>AL</type>" +
                                                    "<orderno>" + orderNo + "</orderno>" +
                                                    "<operationno>" + operationNo + "</operationno>" +
                                                    "<workcenter>" + workcenter + "</workcenter>" +
                                                    "<machinecenter>" + workplace.Code + "</machinecenter>" +
                                                    "<operationtype>Technology</operationtype>" +
                                                    "<initiator>True</initiator>" +
                                                    "<startdate>" + DateTime.Now.ToString(CultureInfo.InvariantCulture) + ".000</startdate>" +
                                                    "<enddate/>" +
                                                    "<consofmeters>" + consOfMeters + "</consofmeters>" +
                                                    "<motorhours>" + motorHours + "</motorhours>" +
                                                    "<cuts>" + cuts + "</cuts>" +
                                                    "<note/>" +
                                                    "</ZAPSIoperation>" +
                                                    "</ZAPSIoperations>";
                                    SendXml(NavUrl, orderData);
                                    var listOfUsers = GetAdditionalUsersFor(workplace, logger);
                                    foreach (var actualUserId in listOfUsers) {
                                        workcenter = GetWorkcenter(workplace, actualUserId, logger);
                                        var userData = "xml=" +
                                                       "<ZAPSIoperations>" +
                                                       "<ZAPSIoperation>" +
                                                       "<type>AL</type>" +
                                                       "<orderno>" + orderNo + "</orderno>" +
                                                       "<operationno>" + operationNo + "</operationno>" +
                                                       "<workcenter>" + workcenter + "</workcenter>" +
                                                       "<machinecenter>" + workplace.Code + "</machinecenter>" +
                                                       "<operationtype>Production</operationtype>" +
                                                       "<initiator>True</initiator>" +
                                                       "<startdate>" + DateTime.Now.ToString(CultureInfo.InvariantCulture) + ".000</startdate>" +
                                                       "<enddate/>" +
                                                       "<consofmeters/>" +
                                                       "<motorhours/>" +
                                                       "<cuts/>" +
                                                       "<note/>" +
                                                       "</ZAPSIoperation>" +
                                                       "</ZAPSIoperations>";
                                        SendXml(NavUrl, userData);
                                    }

                                    workplace.CloseAndStartOrderForWorkplaceAt(DateTime.Now, logger);
                                }
                            }
                        }
                    }

                    bool SendXml(string destinationUrl, string requestXml) {
                        LogDeviceInfo($"[ {workplace.Name} ] --INF-- Sending XML", logger);
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
                                LogDeviceInfo($"[ {workplace.Name} ] --INF-- XML sent OK", logger);
                                return true;

                            }
                            LogDeviceInfo($"[ {workplace.Name} ] --INF-- XML not sent!!!", logger);
                            return false;
                        } catch {
                            LogDeviceInfo($"[ {workplace.Name} ] --INF-- XML not sent!!!", logger);
                            return false;
                        }
                    }

                    LogDeviceInfo($"[ {workplace.Name} ] --INF-- Close open orders: " + closeOpenOrders, logger);
                    if (workplace.TimeIsFifteenMinutesBeforeShiftCloses(logger) && closeOpenOrders) {
                        if (workplace.HasOpenOrderWithStartBeforeThoseFifteenMinutes(logger)) {
                            workplace.CloseOrderForWorkplaceBeforeFifteenMinutes(DateTime.Now, true, logger);
                        } else {
                            workplace.CloseLoginForWorkplace(DateTime.Now, logger);
                        }

                        closeOpenOrders = false;
                    } else if (!workplace.TimeIsFifteenMinutesBeforeShiftCloses(logger) && !closeOpenOrders) {
                        closeOpenOrders = true;
                    }

                    var sleepTime = Convert.ToDouble(_downloadEvery);
                    var waitTime = sleepTime - timer.ElapsedMilliseconds;
                    if ((waitTime) > 0) {
                        LogDeviceInfo($"[ {workplace.Name} ] --INF-- Sleeping for {waitTime} ms", logger);
                        Thread.Sleep((int) (waitTime));
                    } else {
                        LogDeviceInfo("[ " + workplace.Name + " ] --INF-- Processing takes more than" + _downloadEvery + " ms", logger);
                    }

                    timer.Restart();
                }

                factory.Dispose();
                LogDeviceInfo("[ " + workplace.Name + " ] --INF-- Process ended.", logger);
                _numberOfRunningWorkplaces--;
            }
        }

        private static List<int> GetAdditionalUsersFor(Workplace workplace, ILogger logger) {
            var listOfUsers = new List<int>();
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * FROM zapsi2.terminal_input_order_user where TerminalInputOrderID = (SELECT OID from zapsi2.terminal_input_order where DTE is NULL and DeviceID={workplace.DeviceOid})";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    while (reader.Read()) {
                        var userId = Convert.ToInt32(reader["UserID"]);
                        listOfUsers.Add(userId);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + workplace.Name + " ] --ERR-- Problem checking terminal input order users: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + workplace.Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            LogInfo("[ " + workplace.Name + " ] --INF-- Open order has no of additional users: " + listOfUsers.Count, logger);

            return listOfUsers;
        }

        private static string GetConsOfMetersFor(Workplace workplace, ILogger logger) {
            var consOfMeters = 0;
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery =
                    $"SELECT SUM(Data) as result FROM zapsi2.device_input_analog where deviceportid=(SELECT DevicePortID FROM zapsi2.workplace_port where DevicePortId = 111 and WorkplaceID = {workplace.Oid}) and Dt > (SELECT DTS from zapsi2.terminal_input_order where DTE is NULL and DeviceID={workplace.DeviceOid})";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        consOfMeters = Convert.ToInt32(reader["result"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + workplace.Name + " ] --ERR-- Problem checking cons of meters: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + workplace.Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            LogInfo("[ " + workplace.Name + " ] --INF-- Open order cons of meters: " + consOfMeters, logger);

            return consOfMeters.ToString();
        }

        private static string GetMotorHoursFor(Workplace workplace, ILogger logger) {
            var cuts = 0;
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DTE is NULL and DeviceID={workplace.DeviceOid}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        cuts = Convert.ToInt32(reader["Interval"]);
                    }

                    cuts /= 3600;
                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + workplace.Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + workplace.Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            LogInfo("[ " + workplace.Name + " ] --INF-- Open order has interval / 3600: " + cuts, logger);

            return cuts.ToString(CultureInfo.InvariantCulture);
        }

        private static string GetCutsFor(Workplace workplace, ILogger logger) {
            var cuts = "error getting order count";
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DTE is NULL and DeviceID={workplace.DeviceOid}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        cuts = Convert.ToString(reader["Count"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + workplace.Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + workplace.Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            LogInfo("[ " + workplace.Name + " ] --INF-- Open order has count: " + cuts, logger);

            return cuts;
        }

        private static int GetUserIdFor(Workplace workplace, ILogger logger) {
            var userId = 1;
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DTE is NULL and DeviceID={workplace.DeviceOid}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        userId = Convert.ToInt32(reader["UserID"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + workplace.Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + workplace.Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            LogInfo("[ " + workplace.Name + " ] --INF-- Open order has userId: " + userId, logger);

            return userId;
        }

        private static string GetWorkcenter(Workplace workplace, int userId, ILogger logger) {
            var userLogin = "error getting user login";
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.user where OID={userId}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        userLogin = Convert.ToString(reader["Login"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + workplace.Name + " ] --ERR-- Problem checking user login: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + workplace.Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            LogInfo("[ " + workplace.Name + " ] --INF-- User has login: " + userLogin, logger);
            return userLogin;
        }

        private static string GetOperationNo(Workplace workplace, int actualOrderId, ILogger logger) {
            var operation = "error getting order operation";
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.order where OID={actualOrderId}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        operation = Convert.ToString(reader["OperationNo"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + workplace.Name + " ] --ERR-- Problem checking operationNo: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + workplace.Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            LogInfo("[ " + workplace.Name + " ] --INF-- Open order has operation: " + operation, logger);
            return operation;
        }

        private static string GetOrderNo(Workplace workplace, int actualOrderId, ILogger logger) {
            var orderBarcode = "error getting order barcode";
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.order where OID={actualOrderId}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        orderBarcode = Convert.ToString(reader["Barcode"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + workplace.Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + workplace.Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            LogInfo("[ " + workplace.Name + " ] --INF-- Open order has barcode: " + orderBarcode, logger);
            return orderBarcode;
        }

        private static int GetOrderIdFor(Workplace workplace, ILogger logger) {
            var orderId = 1;
            var connection = new MySqlConnection(
                $"server={Program.IpAddress};port={Program.Port};userid={Program.Login};password={Program.Password};database={Program.Database};");
            try {
                connection.Open();
                var selectQuery = $"SELECT * from zapsi2.terminal_input_order where DTE is NULL and DeviceID={workplace.DeviceOid}";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        orderId = Convert.ToInt32(reader["OrderID"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ " + workplace.Name + " ] --ERR-- Problem checking active order: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ " + workplace.Name + " ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            LogInfo("[ " + workplace.Name + " ] --INF-- Open order has orderid: " + orderId, logger);

            return orderId;
        }


        private static void UpdateWorkplace(Workplace workplace, ILogger logger) {
            workplace.AddProductionPort(logger);
            workplace.AddCountPort(logger);
            workplace.AddFailPort(logger);
            workplace.ActualWorkshiftId = workplace.GetActualWorkShiftIdFor(logger);
            workplace.UpdateActualStateForWorkplace(logger);
        }

        private static List<Workplace> GetListOfWorkplacesFromDatabase(ILogger logger) {
            var workplaces = new List<Workplace>();
            var connection = new MySqlConnection($"server={IpAddress};port={Port};userid={Login};password={Password};database={Database};");
            try {
                connection.Open();
                const string selectQuery = "SELECT * from zapsi2.workplace where DeviceID is not NULL";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    while (reader.Read()) {
                        var workplace = new Workplace {
                            Oid = Convert.ToInt32(reader["OID"]),
                            Name = Convert.ToString(reader["Name"]),
                            DeviceOid = Convert.ToInt32(reader["DeviceID"]),
                            WorkplaceDivisionId = Convert.ToInt32(reader["WorkplaceDivisionID"]),
                            WorkplaceGroupId = Convert.ToInt32(reader["WorkplaceGroupID"]),
                            Code = Convert.ToString(reader["Code"])
                        };
                        workplaces.Add(workplace);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ MAIN ] --ERR-- Problem getting list of workplaces " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ MAIN ] --ERR-- problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            return workplaces;
        }


        private static void CreateNewConfigRecord(string option, ILogger logger) {
            var key = "";
            if (option.Equals("Email")) {
                key = _email;
            }

            if (option.Equals("OfflineAfterSeconds")) {
                key = 60.ToString();
            }

            if (option.Equals("CustomerName")) {
                key = _customer;
            }

            if (option.Equals("CloseOnlyAutomaticIdles")) {
                key = CloseOnlyAutomaticIdles;
            }

            if (option.Equals("AddCyclesToOrder")) {
                key = AddCyclesToOrder;
            }

            if (option.Equals("SmtpClient")) {
                key = _smtpClient;
            }

            if (option.Equals("SmtpPort")) {
                key = _smtpPort;
            }

            if (option.Equals("SmtpUsername")) {
                key = _smtpUsername;
            }

            if (option.Equals("SmtpPassword")) {
                key = _smtpPassword;
            }


            var connection = new MySqlConnection($"server={IpAddress};port={Port};userid={Login};password={Password};database={Database};");
            try {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = $"INSERT INTO `zapsi2`.`sw_config` (`SoftID`, `Key`, `Value`, `Version`, `Note`) VALUES ('', '{option}', '{key}', '', '')";
                try {
                    command.ExecuteNonQuery();
                } catch (Exception error) {
                    LogError("[ MAIN ] --ERR-- problem inserting " + option + " into database: " + error.Message + command.CommandText, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ MAIN ] --ERR-- problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }
        }

        private static IEnumerable<string> GetDataFromSwConfigTable(ILogger logger) {
            var swConfig = new List<string>();
            var connection = new MySqlConnection($"server={IpAddress};port={Port};userid={Login};password={Password};database={Database};");
            try {
                connection.Open();
                const string selectQuery = "select * from zapsi2.sw_config";
                var command = new MySqlCommand(selectQuery, connection);


                try {
                    var reader = command.ExecuteReader();
                    while (reader.Read()) {
                        var keyData = reader["Key"].ToString();
                        swConfig.Add(keyData);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ MAIN ] --ERR-- Problem reading from database: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ MAIN ] --ERR-- problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            return swConfig;
        }

        private static void CheckSystemTimeZone(ILogger logger) {
            if (DatabaseType.Equals("mysql")) {
                var connection = new MySqlConnection($"server={IpAddress};port={Port};userid={Login};password={Password};database={Database};");
                try {
                    connection.Open();
                    const string selectQuery = "select @@system_time_zone as timezone;";
                    var command = new MySqlCommand(selectQuery, connection);
                    try {
                        var reader = command.ExecuteReader();
                        if (reader.Read()) {
                            var timezone = reader["timezone"].ToString();
                            if (timezone.Contains("UTC")) {
                                TimezoneIsUtc = true;
                            }
                        }

                        reader.Close();
                        reader.Dispose();
                    } catch (Exception error) {
                        LogError("[ MAIN ] --ERR-- Problem reading timezone: " + error.Message + selectQuery, logger);
                    } finally {
                        command.Dispose();
                    }

                    connection.Close();
                } catch (Exception error) {
                    LogError("[ MAIN ] --ERR-- problem with database: " + error.Message, logger);
                } finally {
                    connection.Dispose();
                }
            }
        }

        private static void CheckNumberOfActiveWorkplaces(ILogger logger) {
            var numberOfActivatedWorkplaces = 0;

            var connection = new MySqlConnection($"server={IpAddress};port={Port};userid={Login};password={Password};database={Database};");
            try {
                connection.Open();
                const string selectQuery = "SELECT count(oid) as count from zapsi2.workplace where DeviceID is not NULL limit 1";
                var command = new MySqlCommand(selectQuery, connection);
                try {
                    var reader = command.ExecuteReader();
                    if (reader.Read()) {
                        numberOfActivatedWorkplaces = Convert.ToInt32(reader["count"]);
                    }

                    reader.Close();
                    reader.Dispose();
                } catch (Exception error) {
                    LogError("[ MAIN ] --ERR-- Problem checking number of workplaces: " + error.Message + selectQuery, logger);
                } finally {
                    command.Dispose();
                }

                connection.Close();
            } catch (Exception error) {
                LogError("[ MAIN ] --ERR-- Problem with database: " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            if (_numberOfRunningWorkplaces != numberOfActivatedWorkplaces) {
                LogInfo("[ MAIN ] --INF-- Workplaces running: " + _numberOfRunningWorkplaces + ", change to: " + numberOfActivatedWorkplaces, logger);
                _loopCanRun = false;
            }

            if (_numberOfRunningWorkplaces == 0) {
                LogInfo("[ MAIN ] --INF-- Number of workplaces is zero, main loop can start.", logger);
                _loopCanRun = true;
            }
        }

        private static void DeleteOldLogFiles(ILogger logger) {
            var currentDirectory = Directory.GetCurrentDirectory();
            var outputPath = Path.Combine(currentDirectory, DataFolder);
            try {
                Directory.GetFiles(outputPath)
                    .Select(f => new FileInfo(f))
                    .Where(f => f.CreationTime < DateTime.Now.AddDays(Convert.ToDouble(_deleteFilesAfterDays)))
                    .ToList()
                    .ForEach(f => f.Delete());
                LogInfo("[ MAIN ] --INF-- Cleared old files.", logger);
            } catch (Exception error) {
                LogError("[ MAIN ] --ERR-- Problem clearing old log files: " + error.Message, logger);
            }
        }

        private static void CheckDatabaseConnection(ILogger logger) {
            var connection = new MySqlConnection($"server={IpAddress};port={Port};userid={Login};password={Password};database={Database};");
            try {
                connection.Open();
                _databaseIsOnline = true;
                LogInfo("[ MAIN ] --INF-- Database is available", logger);
                connection.Close();
            } catch (Exception error) {
                _databaseIsOnline = false;
                LogError("[ MAIN ] --ERR-- Database is unavailable " + error.Message, logger);
            } finally {
                connection.Dispose();
            }

            if (!_databaseIsOnline && !_databaseOfflineEmailWasSent) {
                LogError("[ MAIN ] --ERR-- Database became unavailable", logger);
                SendEmail("Database become unavailable.", logger);
                _databaseOfflineEmailWasSent = true;
            } else if (_databaseIsOnline && _databaseOfflineEmailWasSent) {
                LogInfo("[ MAIN ] --INF-- Database is available again", logger);
                SendEmail("Database is available again.", logger);
                _databaseOfflineEmailWasSent = false;
            }
        }

        private static void RunTimer(System.Timers.Timer timer) {
            timer.Start();
            while (timer.Enabled) {
                Thread.Sleep(Convert.ToInt32(InitialDownloadValue * 10));
                var text = "[ MAIN ] --INF-- Program still running.";
                var now = DateTime.Now;
                text = now + " " + text;
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
                    WriteLine(cyanColor + text + resetColor);
                } else {
                    ForegroundColor = ConsoleColor.Cyan;
                    WriteLine(text);
                    ForegroundColor = ConsoleColor.White;
                }
            }

            timer.Stop();
            timer.Dispose();
        }

        private static void SendEmail(string dataToSend, ILogger logger) {
            ServicePointManager.ServerCertificateValidationCallback = RemoteServerCertificateValidationCallback;
            var client = new SmtpClient(_smtpClient) {
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(_smtpUsername, _smtpPassword),
                Port = int.Parse(_smtpPort)
            };
            var mailMessage = new MailMessage {From = new MailAddress(_smtpUsername)};
            mailMessage.To.Add(_email);
            mailMessage.Subject = "SPECIAL LIKOV TERMINAL SERVER >> " + _customer;
            mailMessage.Body = dataToSend;
            client.EnableSsl = true;
            try {
                client.Send(mailMessage);
                LogInfo("[ MAIN ] --INF-- Email sent: " + dataToSend, logger);
            } catch (Exception error) {
                LogError("[ MAIN ] --ERR-- Cannot send email: " + dataToSend + ": " + error.Message, logger);
            }
        }

        private static bool RemoteServerCertificateValidationCallback(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
            System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors) {
            return true;
        }

        private static void LoadSettingsFromConfigFile(ILogger logger) {
            var currentDirectory = Directory.GetCurrentDirectory();
            const string configFile = "config.json";
            const string backupConfigFile = "config.json.backup";
            var outputPath = Path.Combine(currentDirectory, configFile);
            var backupOutputPath = Path.Combine(currentDirectory, backupConfigFile);
            var configFileLoaded = false;
            try {
                var configBuilder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("config.json");
                var configuration = configBuilder.Build();
                IpAddress = configuration["ipaddress"];
                Database = configuration["database"];
                Port = configuration["port"];
                Login = configuration["login"];
                Password = configuration["password"];
                _customer = configuration["customer"];
                _email = configuration["email"];
                _downloadEvery = configuration["downloadevery"];
                _deleteFilesAfterDays = configuration["deletefilesafterdays"];
                DatabaseType = configuration["databasetype"];
                CloseOnlyAutomaticIdles = configuration["closeonlyautomaticidles"];
                AddCyclesToOrder = configuration["addcyclestoorder"];
                _smtpClient = configuration["smtpclient"];
                _smtpPort = configuration["smtpport"];
                _smtpUsername = configuration["smtpusername"];
                _smtpPassword = configuration["smtppassword"];
                LogInfo("[ MAIN ] --INF-- Config loaded from file for customer: " + _customer, logger);
                configFileLoaded = true;
            } catch (Exception error) {
                LogError("[ MAIN ] --ERR-- Cannot load config from file: " + error.Message, logger);
            }

            if (!configFileLoaded) {
                LogInfo("[ MAIN ] --INF-- Loading backup file.", logger);
                File.Delete(outputPath);
                File.Copy(backupOutputPath, outputPath);
                LogInfo("[ MAIN ] --INF-- Config file updated from backup file.", logger);
                LoadSettingsFromConfigFile(logger);
            }
        }

        private static void CreateConfigFileIfNotExists(ILogger logger) {
            var currentDirectory = Directory.GetCurrentDirectory();
            const string configFile = "config.json";
            const string backupConfigFile = "config.json.backup";
            var outputPath = Path.Combine(currentDirectory, configFile);
            var backupOutputPath = Path.Combine(currentDirectory, backupConfigFile);
            var config = new Config();
            if (!File.Exists(outputPath)) {
                var dataToWrite = JsonConvert.SerializeObject(config);
                try {
                    File.WriteAllText(outputPath, dataToWrite);
                    LogInfo("[ MAIN ] --INF-- Config file created.", logger);
                    if (File.Exists(backupOutputPath)) {
                        File.Delete(backupOutputPath);
                    }

                    File.WriteAllText(backupOutputPath, dataToWrite);
                    LogInfo("[ MAIN ] --INF-- Backup file created.", logger);
                } catch (Exception error) {
                    LogError("[ MAIN ] --ERR-- Cannot create config or backup file: " + error.Message, logger);
                }
            } else {
                LogInfo("[ MAIN ] --INF-- Config file already exists.", logger);
            }
        }

        private static void CheckOsPlatform(ILogger logger) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
                OsIsLinux = true;
                LogInfo("[ MAIN ] --INF-- OS Linux, disable logging to file", logger);
            } else {
                OsIsLinux = false;
            }
        }

        private static void LogInfo(string text, ILogger logger) {
            var now = DateTime.Now;
            text = now + " " + text;
            if (OsIsLinux) {
                WriteLine(cyanColor + text + resetColor);
            } else {
                logger.LogInformation(text);
                ForegroundColor = ConsoleColor.Cyan;
                WriteLine(text);
                ForegroundColor = ConsoleColor.White;
            }
        }

        private static void LogDeviceInfo(string text, ILogger logger) {
            var now = DateTime.Now;
            text = now + " " + text;
            if (OsIsLinux) {
                WriteLine(greenColor + text + resetColor);
            } else {
                ForegroundColor = ConsoleColor.Green;
                logger.LogInformation(text);
                WriteLine(text);
                ForegroundColor = ConsoleColor.White;
            }
        }

        private static void LogError(string text, ILogger logger) {
            var now = DateTime.Now;
            text = now + " " + text;
            if (OsIsLinux) {
                WriteLine(yellowColor + text + resetColor);
            } else {
                logger.LogInformation(text);
                ForegroundColor = ConsoleColor.Yellow;
                WriteLine(text);
                ForegroundColor = ConsoleColor.White;
            }
        }

        private static LoggerFactory CreateLogger(string outputPath, out ILogger logger) {
            var factory = new LoggerFactory();
            logger = factory.CreateLogger("Likov Special Terminal");
            factory.AddFile(outputPath, LogLevel.Debug);
            return factory;
        }

        private static string CreateLogFileIfNotExists(string fileName) {
            var currentDirectory = Directory.GetCurrentDirectory();
            var logFilename = fileName;
            var outputPath = Path.Combine(currentDirectory, DataFolder, logFilename);
            var outputDirectory = Path.GetDirectoryName(outputPath);
            CreateLogDirectoryIfNotExists(outputDirectory);
            return outputPath;
        }

        private static void CreateLogDirectoryIfNotExists(string outputDirectory) {
            if (!Directory.Exists(outputDirectory)) {
                try {
                    Directory.CreateDirectory(outputDirectory);
                    var text = "[ MAIN ] --INF-- Log directory created.";
                    var now = DateTime.Now;
                    text = now + " " + text;
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
                        WriteLine(cyanColor + text + resetColor);
                    } else {
                        ForegroundColor = ConsoleColor.Cyan;
                        WriteLine(text);
                        ForegroundColor = ConsoleColor.White;
                    }
                } catch (Exception error) {
                    var text = "[ MAIN ] --ERR-- Log directory not created: " + error.Message;
                    var now = DateTime.Now;
                    text = now + " " + text;
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
                        WriteLine(redColor + text + resetColor);
                    } else {
                        ForegroundColor = ConsoleColor.Red;
                        WriteLine(text);
                        ForegroundColor = ConsoleColor.White;
                    }
                }
            }
        }

        private static void PrintSoftwareLogo() {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
                WriteLine(cyanColor + " LIKOV SPECIAL TERMINAL SERVICE  ");
            } else {
                ForegroundColor = ConsoleColor.Cyan;
                WriteLine(" LIKOV SPECIAL TERMINAL SERVICE  ");
            }
        }
    }
}