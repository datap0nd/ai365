using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookLocalAIChat.Utilities
{
    // User-initiated self update. This is never exposed to the model:
    // only the Settings window calls it after an explicit confirmation.
    public static class SelfUpdater
    {
        public const string InstallerUrl =
            "https://github.com/datap0nd/ai365/releases/latest/download/AI365Setup.exe";

        public const int MaxInstallerBytes = 100 * 1024 * 1024;
        public const int MinInstallerBytes = 200 * 1024;

        public static async Task<string> DownloadInstallerAsync(
            CancellationToken cancellationToken)
        {
            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls12;
            var path = Path.Combine(
                Path.GetTempPath(),
                "AI365-Update-" +
                Guid.NewGuid().ToString("N") +
                ".exe");
            using (var http = new HttpClient())
            {
                http.Timeout = TimeSpan.FromMinutes(5);
                using (var response = await http.GetAsync(
                    InstallerUrl,
                    HttpCompletionOption.ResponseHeadersRead,
                    cancellationToken).ConfigureAwait(true))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "The update download failed: " +
                            (int)response.StatusCode +
                            " " +
                            response.ReasonPhrase +
                            ".");
                    }

                    using (var source = await response.Content
                        .ReadAsStreamAsync()
                        .ConfigureAwait(true))
                    using (var target = File.Create(path))
                    {
                        var buffer = new byte[81920];
                        var total = 0;
                        while (true)
                        {
                            var read = await source.ReadAsync(
                                buffer,
                                0,
                                buffer.Length,
                                cancellationToken).ConfigureAwait(true);
                            if (read == 0)
                            {
                                break;
                            }

                            total += read;
                            if (total > MaxInstallerBytes)
                            {
                                throw new InvalidOperationException(
                                    "The downloaded installer is larger than expected.");
                            }

                            target.Write(buffer, 0, read);
                        }

                        if (total < MinInstallerBytes)
                        {
                            throw new InvalidOperationException(
                                "The downloaded installer looks incomplete.");
                        }
                    }
                }
            }

            using (var check = File.OpenRead(path))
            {
                var header = new byte[2];
                if (check.Read(header, 0, 2) != 2 ||
                    header[0] != 'M' ||
                    header[1] != 'Z')
                {
                    throw new InvalidOperationException(
                        "The downloaded file is not a Windows installer.");
                }
            }

            return path;
        }

        // The script receives the installer path as %1 so the file itself
        // stays pure ASCII regardless of the user's profile path. One
        // installer carries all four add-ins (Outlook, Excel,
        // PowerPoint, and Word share a single assembly), so the script waits for
        // every Office host to close before installing and the whole
        // suite updates together.
        public static string BuildUpdateScript()
        {
            var builder = new StringBuilder();
            builder.AppendLine("@echo off");
            builder.AppendLine("set \"installer=%~1\"");
            builder.AppendLine("set \"restart=%~2\"");
            builder.AppendLine("set tries=0");
            builder.AppendLine(":wait");
            builder.AppendLine("set running=0");
            builder.AppendLine(
                "tasklist /FI \"IMAGENAME eq OUTLOOK.EXE\" | " +
                "find /I \"OUTLOOK.EXE\" >nul");
            builder.AppendLine("if not errorlevel 1 set running=1");
            builder.AppendLine(
                "tasklist /FI \"IMAGENAME eq EXCEL.EXE\" | " +
                "find /I \"EXCEL.EXE\" >nul");
            builder.AppendLine("if not errorlevel 1 set running=1");
            builder.AppendLine(
                "tasklist /FI \"IMAGENAME eq POWERPNT.EXE\" | " +
                "find /I \"POWERPNT.EXE\" >nul");
            builder.AppendLine("if not errorlevel 1 set running=1");
            builder.AppendLine(
                "tasklist /FI \"IMAGENAME eq WINWORD.EXE\" | " +
                "find /I \"WINWORD.EXE\" >nul");
            builder.AppendLine("if not errorlevel 1 set running=1");
            builder.AppendLine("if %running%==0 goto install");
            builder.AppendLine("set /a tries+=1");
            builder.AppendLine("if %tries% GEQ 150 exit /b 1");
            builder.AppendLine("timeout /T 2 /NOBREAK >nul");
            builder.AppendLine("goto wait");
            builder.AppendLine(":install");
            builder.AppendLine(
                "\"%installer%\" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART");
            builder.AppendLine(
                "if not \"%restart%\"==\"\" start \"\" \"%restart%\"");
            builder.AppendLine("del \"%installer%\"");
            return builder.ToString();
        }

        // hostApplication may be null when the update starts from the
        // Excel/PowerPoint pane settings: nothing is quit and the
        // script simply waits for the user to close the Office apps.
        public static void LaunchUpdateAndQuitHost(
            object hostApplication,
            string installerPath,
            string restartExecutable)
        {
            var scriptPath = Path.Combine(
                Path.GetTempPath(),
                "AI365-Update-" +
                Guid.NewGuid().ToString("N") +
                ".cmd");
            File.WriteAllText(
                scriptPath,
                BuildUpdateScript(),
                Encoding.ASCII);
            var start = new System.Diagnostics.ProcessStartInfo
            {
                FileName = scriptPath,
                Arguments = "\"" + installerPath + "\" \"" +
                    (restartExecutable ?? string.Empty) + "\"",
                UseShellExecute = true,
                WindowStyle =
                    System.Diagnostics.ProcessWindowStyle.Hidden
            };
            System.Diagnostics.Process.Start(start);
            if (hostApplication != null)
            {
                dynamic application = hostApplication;
                application.Quit();
            }
        }

        public static string InstalledVersion()
        {
            try
            {
                return System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(
                        typeof(SelfUpdater).Assembly.Location)
                    .FileVersion ?? "unknown";
            }
            catch
            {
                return "unknown";
            }
        }
    }
}
