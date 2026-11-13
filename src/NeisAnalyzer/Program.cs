using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace NeisAnalyzer
{
    internal static class Program
    {
        private static Assembly _epplus;

        [STAThread]
        private static int Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += ResolveEmbeddedAssembly;
            return Run(args);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static int Run(string[] args)
        {
            if (args.Length > 0 && args[0] == "--cli")
                return Cli.Run(args);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length >= 4 && args[0] == "--ui-test")
            {
                float scale = 0;
                if (args.Length >= 5) float.TryParse(args[4], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out scale);
                return UiTest(args[1], args[2], args[3], scale);
            }
            Application.Run(new MainForm());
            return 0;
        }

        /// <summary>자동 시험용: 창을 띄워 파일을 넣고 분석까지 돌린 뒤 화면을 그림 파일로 저장한다. scale이 0이 아니면 그 배율 화면(예: 1.5)처럼 그린다.</summary>
        static int UiTest(string inputFolder, string outXlsx, string screenshotPng, float scale)
        {
            int code = 1;
            var form = new MainForm(scale);
            form.Shown += async (s, e) =>
            {
                try
                {
                    form.AddPaths(new[] { inputFolder });
                    Application.DoEvents();
                    Snapshot(form, screenshotPng.Replace(".png", "_1.png"));
                    var r = await form.RunAnalysisTo(outXlsx, form.ReadSettings());
                    Application.DoEvents();
                    Snapshot(form, screenshotPng);
                    code = r != null && File.Exists(outXlsx) ? 0 : 1;
                }
                catch (Exception ex) { File.WriteAllText(outXlsx + ".error.txt", ex.ToString()); code = 1; }
                finally { form.Close(); }
            };
            Application.Run(form);
            return code;
        }

        static void Snapshot(Form f, string path)
        {
            using (var bmp = new System.Drawing.Bitmap(f.Width, f.Height))
            {
                f.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, f.Width, f.Height));
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            }
        }

        private static Assembly ResolveEmbeddedAssembly(object sender, ResolveEventArgs args)
        {
            var name = new AssemblyName(args.Name).Name;
            if (!string.Equals(name, "EPPlus", StringComparison.OrdinalIgnoreCase))
                return null;
            if (_epplus != null)
                return _epplus;
            using (var stream = typeof(Program).Assembly.GetManifestResourceStream("EPPlus.dll"))
            {
                if (stream == null) return null;
                var bytes = new byte[stream.Length];
                int read = 0;
                while (read < bytes.Length)
                {
                    int n = stream.Read(bytes, read, bytes.Length - read);
                    if (n <= 0) break;
                    read += n;
                }
                _epplus = Assembly.Load(bytes);
                return _epplus;
            }
        }
    }
}
