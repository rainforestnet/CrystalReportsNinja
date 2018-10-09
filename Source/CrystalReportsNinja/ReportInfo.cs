using CrystalDecisions.CrystalReports.Engine;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrystalReportsNinja
{
    public class ReportInfo
    {
        #region Properties

        public string ReportName { get; set; }
        public string FileName { get; set; }
        public ConnectionInfo[] Connections { get; set; }
        public DatasourceInfo[] Datasources { get; set; }

        public class ConnectionInfo
        {
            public string DatabaseName { get; set; }
            public bool UsesIntegratedSecurity { get; set; }
            public string ServerName { get; set; }
            public string Username { get; set; }
            public string Password { get; set; }
        }

        public class DatasourceInfo
        {
            public string Name { get; set; }
            public string Location { get; set; }
        }

        #endregion

        public ReportInfo(ReportDocument reportDocument)
        {
            ReportName = reportDocument.Name;
            FileName = reportDocument.FileName;

            Connections = new ConnectionInfo[reportDocument.DataSourceConnections.Count];
            for (var i = 0; i < reportDocument.DataSourceConnections.Count; i++)
            {
                var conn = reportDocument.DataSourceConnections[i];
                Connections[i] = new ConnectionInfo();
                Connections[i].DatabaseName = conn.DatabaseName;
                Connections[i].UsesIntegratedSecurity = conn.IntegratedSecurity;
                Connections[i].ServerName = conn.ServerName;
                Connections[i].Username = conn.UserID;
                Connections[i].Password = conn.Password;
            }

            Datasources = new DatasourceInfo[reportDocument.Database.Tables.Count];
            for (var i = 0; i < reportDocument.Database.Tables.Count; i++)
            {
                var table = reportDocument.Database.Tables[i];
                Datasources[i] = new DatasourceInfo();
                Datasources[i].Name = table.Name;
                Datasources[i].Location = table.Location;
            }
        }

        #region Output method & support

        private const int MaxConnections = 2;
        private const int MaxDatasources = 5;

        public void Output(string outputFormat, string outputPath)
        {
            // Generate the output
            var outBuffer = new StringBuilder();
            switch ((outputFormat ?? "print").ToUpper())
            {
                case "TAB":
                    outBuffer.AppendFormat("{0}\t", ReportName);
                    outBuffer.AppendFormat("{0}\t", FileName);

                    for (var i = 0; i < MaxConnections; i++)
                    {
                        ConnectionInfo conn = null;
                        if (i < Connections.Length)
                        {
                            conn = Connections[i];
                        }
                        outBuffer.AppendFormat("{0}\t", conn?.ServerName);
                        outBuffer.AppendFormat("{0}\t", conn?.DatabaseName);
                        outBuffer.AppendFormat("{0}\t", conn == null ? "" : (conn.UsesIntegratedSecurity ? "Yes" : "No"));
                        outBuffer.AppendFormat("{0}\t", conn?.Username);
                        outBuffer.AppendFormat("{0}\t", conn?.Password);
                    }


                    for (var i = 0; i < MaxDatasources; i++)
                    {
                        DatasourceInfo datasource = null;
                        if (i < Datasources.Length)
                        {
                            datasource = Datasources[i];
                        }
                        outBuffer.AppendFormat("{0}\t", datasource?.Name);
                        outBuffer.AppendFormat("{0}\t", datasource?.Location);
                    }

                    // Remove the trailing \t and add an endline
                    outBuffer.Length -= 1;
                    outBuffer.AppendLine();
                    break;
                default:
                    outBuffer.AppendFormat("Report Name: {0}\n", ReportName);
                    outBuffer.AppendFormat("File Name: {0}\n", FileName);

                    if (Connections.Length > 0)
                    {
                        for (var i = 0; i < Connections.Length; i++)
                        {
                            outBuffer.AppendFormat("Connection #{0}:\n", i + 1);
                            var conn = Connections[i];
                            outBuffer.AppendFormat("  Server Name: {0}\n", conn.ServerName);
                            outBuffer.AppendFormat("  Database Name: {0}\n", conn.DatabaseName);
                            outBuffer.AppendFormat("  Integrated Security: {0}\n", conn.UsesIntegratedSecurity ? "Yes" : "No");
                            outBuffer.AppendFormat("  Username: {0}\n", conn.Username);
                            outBuffer.AppendFormat("  Password: {0}\n", conn.Password);
                        }
                    }
                    else
                    {
                        outBuffer.AppendFormat("No connections found\n");
                    }

                    if (Datasources.Length > 0)
                    {
                        for (var i = 0; i < Datasources.Length; i++)
                        {
                            outBuffer.AppendFormat("Datasource #{0}:\n", i + 1);
                            var datasource = Datasources[i];
                            outBuffer.AppendFormat("  Name: {0}\n", datasource.Name);
                            outBuffer.AppendFormat("  Location: {0}\n", datasource.Location);
                        }
                    }
                    else
                    {
                        outBuffer.AppendFormat("No datasources found\n");
                    }

                    break;
            }

            // Output the, uh, output
            if (string.IsNullOrEmpty(outputPath))
            {
                Console.Write(outBuffer.ToString());
            }
            else
            {
                if (!File.Exists(outputPath))
                {
                    File.WriteAllText(outputPath, GetHeader(outputFormat));
                }

                using (var strm = File.AppendText(outputPath))
                {
                    strm.Write(outBuffer.ToString());
                }
            }
        }

        private string GetHeader(string outputFormat)
        {
            var outBuffer = new StringBuilder();
            switch ((outputFormat ?? "print").ToUpper())
            {
                case "TAB":
                    outBuffer.AppendFormat("Report Name\t", ReportName);
                    outBuffer.AppendFormat("Report File Name\t", FileName);

                    for (var i = 0; i < MaxConnections; i++)
                    {
                        outBuffer.AppendFormat("Conn {0}: Server\t", i + 1);
                        outBuffer.AppendFormat("Conn {0}: Database\t", i + 1);
                        outBuffer.AppendFormat("Conn {0}: Integrated Security\t", i + 1);
                        outBuffer.AppendFormat("Conn {0}: Username\t", i + 1);
                        outBuffer.AppendFormat("Conn {0}: Password\t", i + 1);
                    }

                    for (var i = 0; i < 5; i++)
                    {
                        outBuffer.AppendFormat("Datasource {0}: Name\t", i + 1);
                        outBuffer.AppendFormat("Datasource {0}: Location\t", i + 1);
                    }

                    // Remove the trailing \t and add an endline
                    outBuffer.Length -= 1;
                    outBuffer.AppendLine();
                    break;
            }

            return outBuffer.ToString();
        }

        #endregion
    }
}

