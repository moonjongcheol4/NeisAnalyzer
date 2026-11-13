using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NeisAnalyzer
{
    /// <summary>프로그램 메인 창: 파일 끌어다 놓기 → 설정 확인 → 분석 시작 → 결과 열기</summary>
    public class MainForm : Form
    {
        static readonly Color Blue = Color.FromArgb(31, 78, 121);
        static readonly Color LightBlue = Color.FromArgb(221, 235, 247);
        static readonly Color Accent = Color.FromArgb(0, 120, 212);
        static readonly Font BaseFont = new Font("맑은 고딕", 11f);
        static readonly Font BoldFont = new Font("맑은 고딕", 11f, FontStyle.Bold);
        static readonly Font TitleFont = new Font("맑은 고딕", 18f, FontStyle.Bold);

        readonly List<DetectedFile> _files = new List<DetectedFile>();
        AnalysisSettings _settings;
        string _lastOutput;

        // 화면 요소
        Panel _dropZone; Label _dropLabel;
        ListView _fileList;
        TextBox _warnBox;
        DataGridView _perfGrid, _examGrid;
        NumericUpDown _numA, _numB, _numC, _numD, _numBoundary;
        ComboBox _cmbWrittenDec, _cmbFinalDec, _cmbBasis;
        Label _ratioSum, _status;
        Button _btnRun, _btnOpen, _btnOpenFolder, _btnClear, _btnPick;
        ProgressBar _progress;

        public MainForm()
        {
            Text = "나이스 성적 분석";
            Font = BaseFont;
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1100, 800);
            Size = new Size(1240, 860);
            AllowDrop = true;
            DragEnter += OnDragEnter; DragDrop += OnDragDrop;
            try { Icon = SystemIcons.Application; } catch { }
            BuildUi();
        }

        // ===================== 화면 구성 =====================
        void BuildUi()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3, Padding = new Padding(16) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 134));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
            Controls.Add(root);

            // --- 제목 ---
            var head = new Panel { Dock = DockStyle.Fill };
            var title = new Label { Text = "나이스 성적 분석", Font = TitleFont, ForeColor = Blue, AutoSize = true, Location = new Point(0, 0) };
            var sub = new Label
            {
                Text = "나이스에서 아래 세 가지를 엑셀로 저장한 뒤, 파일 이름을 바꾸지 말고 모두 아래 상자에 끌어다 놓으세요. (반 수만큼 ①·③ 파일을 함께 넣습니다)\n" +
                       "  ① 교과담임 → 정기시험조회/통계 → 정기시험조회 → [교과목별학생답정오표(학급별)]\n" +
                       "  ② 교과담임 → 정기시험조회/통계 → 정기시험조회 → [교과목별일람표조회(전체학급)]  : 필요시 검토(검산)용, 없어도 됨\n" +
                       "  ③ 교과담임 → 수행평가조회/통계 → 수행평가조회 → [강의실별일람표조회-전체영역(학급별)]",
                AutoSize = false, Location = new Point(2, 38), Size = new Size(1400, 92), ForeColor = Color.FromArgb(64, 64, 64)
            };
            head.Controls.Add(title); head.Controls.Add(sub);
            root.Controls.Add(head, 0, 0); root.SetColumnSpan(head, 2);

            // --- 왼쪽: 파일 ---
            var left = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(0, 0, 12, 0) };
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 110));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 62));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 38));
            root.Controls.Add(left, 0, 1);

            _dropZone = new Panel { Dock = DockStyle.Fill, BackColor = LightBlue, AllowDrop = true, Margin = new Padding(0, 0, 0, 6) };
            _dropZone.Paint += (s, e) =>
            {
                using (var pen = new Pen(Accent, 2) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash })
                    e.Graphics.DrawRectangle(pen, 2, 2, _dropZone.Width - 5, _dropZone.Height - 5);
            };
            _dropZone.Resize += (s, e) => _dropZone.Invalidate();
            _dropLabel = new Label { Text = "📂  여기에 나이스 엑셀 파일을 끌어다 놓으세요 (여러 개 한꺼번에 가능)\n파일이 든 폴더를 통째로 끌어다 놓아도 됩니다.", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("맑은 고딕", 12f, FontStyle.Bold), ForeColor = Blue, AllowDrop = true };
            _dropZone.Controls.Add(_dropLabel);
            foreach (Control c in new Control[] { _dropZone, _dropLabel }) { c.DragEnter += OnDragEnter; c.DragDrop += OnDragDrop; }
            _dropLabel.Click += (s, e) => PickFiles();
            left.Controls.Add(_dropZone, 0, 0);

            var btnRow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
            _btnPick = MakeButton("파일 선택...", 150, false); _btnPick.Click += (s, e) => PickFiles();
            _btnClear = MakeButton("목록 비우기", 130, false); _btnClear.Click += (s, e) => { _files.Clear(); RefreshAll(); };
            btnRow.Controls.Add(_btnPick); btnRow.Controls.Add(_btnClear);
            left.Controls.Add(btnRow, 0, 1);

            _fileList = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true, HeaderStyle = ColumnHeaderStyle.Nonclickable, Margin = new Padding(0, 6, 0, 6), AllowDrop = true };
            _fileList.Columns.Add("상태", 50); _fileList.Columns.Add("종류", 100); _fileList.Columns.Add("내용", 330); _fileList.Columns.Add("파일 이름", 200);
            _fileList.DragEnter += OnDragEnter; _fileList.DragDrop += OnDragDrop;
            _fileList.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) RemoveSelected(); };
            var menu = new ContextMenuStrip(); menu.Items.Add("선택한 파일 목록에서 빼기", null, (s, e) => RemoveSelected()); _fileList.ContextMenuStrip = menu;
            left.Controls.Add(_fileList, 0, 2);

            var warnPanel = new Panel { Dock = DockStyle.Fill };
            var warnTitle = new Label { Text = "확인이 필요한 사항", Font = BoldFont, ForeColor = Blue, Dock = DockStyle.Top, Height = 26 };
            _warnBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = Color.FromArgb(255, 251, 240), BorderStyle = BorderStyle.FixedSingle };
            warnPanel.Controls.Add(_warnBox); warnPanel.Controls.Add(warnTitle);
            left.Controls.Add(warnPanel, 0, 3);

            // --- 오른쪽: 설정 ---
            var right = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8, 0, 0, 0), BackColor = Color.FromArgb(250, 251, 253) };
            root.Controls.Add(right, 0, 1); root.SetColumn(right, 1);
            var settings = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            right.Controls.Add(settings);

            settings.Controls.Add(SectionLabel("설정  (파일에서 읽은 값이 자동으로 채워집니다. 필요하면 고치세요)"));

            settings.Controls.Add(SmallLabel("수행평가 영역별 만점과 반영 비율(%)"));
            _perfGrid = MakeGrid(new[] { "영역", "만점", "반영 비율(%)" }, 150);
            _perfGrid.Columns[0].ReadOnly = true;
            _perfGrid.CellValueChanged += (s, e) => UpdateRatioSum();
            settings.Controls.Add(_perfGrid);

            settings.Controls.Add(SmallLabel("지필평가 반영 비율(%)  — 고사가 하나면 100 − 수행 비율, 둘이면 나누어 입력"));
            _examGrid = MakeGrid(new[] { "고사", "반영 비율(%)" }, 90);
            _examGrid.Columns[0].ReadOnly = true;
            _examGrid.CellValueChanged += (s, e) => UpdateRatioSum();
            settings.Controls.Add(_examGrid);
            _ratioSum = new Label { AutoSize = true, Font = BoldFont, Margin = new Padding(4, 2, 0, 10) };
            settings.Controls.Add(_ratioSum);

            settings.Controls.Add(SmallLabel("성취도 기준 (이 점수 이상이면 해당 성취도, D 미만은 E)"));
            var gradeRow = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
            _numA = MakeNum(90); _numB = MakeNum(80); _numC = MakeNum(70); _numD = MakeNum(60);
            foreach (var p in new[] { ("A", _numA), ("B", _numB), ("C", _numC), ("D", _numD) })
            {
                gradeRow.Controls.Add(new Label { Text = p.Item1 + " :", AutoSize = true, Padding = new Padding(8, 8, 0, 0), Font = BoldFont });
                gradeRow.Controls.Add(p.Item2);
            }
            settings.Controls.Add(gradeRow);

            settings.Controls.Add(SmallLabel("반올림 방식 (학교 학업성적관리규정에 맞게 선택)"));
            var roundTbl = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, Margin = new Padding(0, 0, 0, 8) };
            _cmbWrittenDec = MakeCombo(new[] { "소수 둘째 자리까지 (나이스 기본)", "소수 첫째 자리까지", "정수" }, 0);
            _cmbFinalDec = MakeCombo(new[] { "소수 둘째 자리까지 (나이스 기본)", "소수 첫째 자리까지", "정수" }, 0);
            _cmbBasis = MakeCombo(new[] { "정수로 반올림한 점수 (생기부 방식: 89.5→90→A)", "반올림 전 점수 그대로 (89.5→B)" }, 0);
            AddRow(roundTbl, "지필 환산 점수", _cmbWrittenDec);
            AddRow(roundTbl, "최종 점수", _cmbFinalDec);
            AddRow(roundTbl, "성취도 판정 기준", _cmbBasis);
            settings.Controls.Add(roundTbl);

            var bRow = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
            bRow.Controls.Add(new Label { Text = "성취도 경계 표시 범위: ±", AutoSize = true, Padding = new Padding(4, 8, 0, 0) });
            _numBoundary = MakeNum(1); _numBoundary.DecimalPlaces = 1; _numBoundary.Increment = 0.5m; _numBoundary.Minimum = 0; _numBoundary.Maximum = 10;
            bRow.Controls.Add(_numBoundary);
            bRow.Controls.Add(new Label { Text = "점", AutoSize = true, Padding = new Padding(4, 8, 0, 0) });
            settings.Controls.Add(bRow);

            var privacy = new Label { Text = "🔒 이 프로그램은 인터넷에 연결하지 않으며, 학생 자료는 이 컴퓨터 밖으로 나가지 않습니다.", AutoSize = true, ForeColor = Color.FromArgb(0, 112, 60), Margin = new Padding(4, 8, 0, 0) };
            settings.Controls.Add(privacy);

            // --- 아래: 실행 ---
            var bottom = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 1 };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            _btnRun = MakeButton("▶  분석 시작", 210, true); _btnRun.Click += async (s, e) => await RunAnalysis();
            _btnOpen = MakeButton("결과 파일 열기", 170, false); _btnOpen.Enabled = false; _btnOpen.Click += (s, e) => { if (_lastOutput != null && File.Exists(_lastOutput)) Process.Start(new ProcessStartInfo(_lastOutput) { UseShellExecute = true }); };
            _btnOpenFolder = MakeButton("폴더 열기", 140, false); _btnOpenFolder.Enabled = false; _btnOpenFolder.Click += (s, e) => { if (_lastOutput != null && File.Exists(_lastOutput)) Process.Start("explorer.exe", "/select,\"" + _lastOutput + "\""); };
            _status = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Blue, Padding = new Padding(12, 0, 0, 0) };
            _progress = new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Marquee, Visible = false, Margin = new Padding(0, 20, 0, 20) };
            bottom.Controls.Add(_btnRun, 0, 0); bottom.Controls.Add(_btnOpen, 1, 0); bottom.Controls.Add(_btnOpenFolder, 2, 0); bottom.Controls.Add(_status, 3, 0); bottom.Controls.Add(_progress, 4, 0);
            root.Controls.Add(bottom, 0, 2); root.SetColumnSpan(bottom, 2);

            RefreshAll();
        }

        static Label SectionLabel(string t) => new Label { Text = t, Font = BoldFont, ForeColor = Blue, AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        static Label SmallLabel(string t) => new Label { Text = t, AutoSize = true, Margin = new Padding(2, 4, 0, 2), ForeColor = Color.FromArgb(64, 64, 64) };
        static Button MakeButton(string text, int width, bool primary)
        {
            var b = new Button { Text = text, Width = width, Height = 44, FlatStyle = FlatStyle.Flat, Font = BoldFont, Margin = new Padding(0, 0, 10, 0), Cursor = Cursors.Hand };
            if (primary) { b.BackColor = Accent; b.ForeColor = Color.White; b.FlatAppearance.BorderSize = 0; }
            else { b.BackColor = Color.White; b.ForeColor = Blue; b.FlatAppearance.BorderColor = Accent; }
            // 사용할 수 없을 때는 회색으로 보이게
            b.EnabledChanged += (s, e) =>
            {
                if (primary) b.BackColor = b.Enabled ? Accent : Color.FromArgb(180, 180, 180);
                else { b.ForeColor = b.Enabled ? Blue : Color.Gray; b.FlatAppearance.BorderColor = b.Enabled ? Accent : Color.LightGray; }
            };
            return b;
        }
        static NumericUpDown MakeNum(decimal v) => new NumericUpDown { Value = v, Minimum = 0, Maximum = 100, Width = 70, DecimalPlaces = 0, Font = BaseFont, Margin = new Padding(4) };
        static ComboBox MakeCombo(string[] items, int idx) { var c = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 340, Font = BaseFont }; c.Items.AddRange(items); c.SelectedIndex = idx; return c; }
        static void AddRow(TableLayoutPanel t, string label, Control c) { int r = t.RowCount++; t.Controls.Add(new Label { Text = label, AutoSize = true, Padding = new Padding(4, 8, 8, 0) }, 0, r); t.Controls.Add(c, 1, r); }
        static DataGridView MakeGrid(string[] cols, int height)
        {
            var g = new DataGridView { Height = height, Width = 470, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false, BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = BaseFont, Margin = new Padding(0, 0, 0, 6), AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize, EnableHeadersVisualStyles = false };
            g.ColumnHeadersDefaultCellStyle.BackColor = Blue; g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; g.ColumnHeadersDefaultCellStyle.Font = BoldFont;
            g.RowTemplate.Height = 30;
            foreach (var c in cols) g.Columns.Add(c, c);
            g.Columns[0].FillWeight = 200;
            g.DataError += (s, e) => { e.ThrowException = false; };
            return g;
        }

        // ===================== 파일 넣기 =====================
        void OnDragEnter(object sender, DragEventArgs e) { e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None; }
        void OnDragDrop(object sender, DragEventArgs e) { var paths = (string[])e.Data.GetData(DataFormats.FileDrop); AddPaths(paths); }

        void PickFiles()
        {
            using (var dlg = new OpenFileDialog { Multiselect = true, Filter = "엑셀 파일 (*.xlsx)|*.xlsx|모든 파일|*.*", Title = "나이스 엑셀 파일 선택 (여러 개 선택 가능)" })
                if (dlg.ShowDialog(this) == DialogResult.OK) AddPaths(dlg.FileNames);
        }

        public void AddPaths(IEnumerable<string> paths)
        {
            var files = new List<string>();
            foreach (var p in paths)
            {
                if (Directory.Exists(p)) files.AddRange(Directory.GetFiles(p, "*.xlsx"));
                else if (File.Exists(p)) files.Add(p);
            }
            files = files.Where(f => !Path.GetFileName(f).StartsWith("~$")).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (files.Count == 0) { SetStatus("엑셀(.xlsx) 파일이 없습니다."); return; }
            Cursor = Cursors.WaitCursor;
            SetStatus($"파일 {files.Count}개 읽는 중...");
            Application.DoEvents();
            foreach (var f in files)
            {
                if (_files.Any(x => string.Equals(x.Path, f, StringComparison.OrdinalIgnoreCase))) continue;
                _files.Add(NeisReader.Detect(f));
            }
            Cursor = Cursors.Default;
            RefreshAll();
        }

        void RemoveSelected()
        {
            var sel = _fileList.SelectedItems.Cast<ListViewItem>().Select(i => (DetectedFile)i.Tag).ToList();
            foreach (var d in sel) _files.Remove(d);
            RefreshAll();
        }

        // ===================== 화면 갱신 =====================
        void RefreshAll()
        {
            _fileList.BeginUpdate();
            _fileList.Items.Clear();
            foreach (var d in _files.OrderBy(f => f.Kind == NeisFileKind.AnswerSheet ? 0 : f.Kind == NeisFileKind.PerfSheet ? 1 : f.Kind == NeisFileKind.ScoreTable ? 2 : 3).ThenBy(f => f.Header?.ExamName).ThenBy(f => f.Header?.ClassNo))
            {
                string kind = d.Kind == NeisFileKind.AnswerSheet ? "정오표" : d.Kind == NeisFileKind.PerfSheet ? "수행평가" : d.Kind == NeisFileKind.ScoreTable ? "일람표(검산)" : "알 수 없음";
                string desc = d.Label; int idx = desc.IndexOf(" - "); if (idx > 0) desc = desc.Substring(idx + 3);
                var item = new ListViewItem(d.Kind == NeisFileKind.Unknown ? "⚠" : "✔") { Tag = d, UseItemStyleForSubItems = false };
                item.SubItems.Add(kind); item.SubItems.Add(desc); item.SubItems.Add(Path.GetFileName(d.Path));
                if (d.Kind == NeisFileKind.Unknown) { item.ForeColor = Color.Firebrick; item.SubItems[2].ForeColor = Color.Firebrick; }
                _fileList.Items.Add(item);
            }
            _fileList.EndUpdate();

            var warnings = _files.Count == 0 ? new List<string>() : Analyzer.CheckFiles(_files);
            _warnBox.Text = warnings.Count == 0 ? (_files.Count == 0 ? "" : "문제가 없습니다. 설정을 확인하고 [분석 시작]을 누르세요.") : string.Join(Environment.NewLine, warnings.Select(w => "⚠ " + w));

            // 설정 기본값 (파일 구성이 바뀔 때마다 다시 읽는다)
            _settings = Analyzer.DefaultSettings(_files);
            _perfGrid.Rows.Clear();
            foreach (var a in _settings.PerfAreas) _perfGrid.Rows.Add(a.Name, a.Max.ToString("0.##"), a.Ratio.ToString("0.##"));
            _examGrid.Rows.Clear();
            foreach (var e in _settings.ExamRatios) _examGrid.Rows.Add(e.Key, e.Value.ToString("0.##"));
            UpdateRatioSum();

            int answers = _files.Count(f => f.Kind == NeisFileKind.AnswerSheet), perfs = _files.Count(f => f.Kind == NeisFileKind.PerfSheet), tables = _files.Count(f => f.Kind == NeisFileKind.ScoreTable);
            _btnRun.Enabled = answers > 0 || perfs > 0;
            SetStatus(_files.Count == 0 ? "파일을 끌어다 놓으면 시작합니다." : $"정오표 {answers}개, 수행평가 {perfs}개, 일람표 {tables}개");
        }

        void UpdateRatioSum()
        {
            double perf = 0, exam = 0;
            foreach (DataGridViewRow r in _perfGrid.Rows) { double v; if (double.TryParse(Convert.ToString(r.Cells[2].Value), out v)) perf += v; }
            foreach (DataGridViewRow r in _examGrid.Rows) { double v; if (double.TryParse(Convert.ToString(r.Cells[1].Value), out v)) exam += v; }
            double total = perf + exam;
            _ratioSum.Text = $"수행 {perf:0.##}% + 지필 {exam:0.##}% = {total:0.##}%" + (Math.Abs(total - 100) > 0.001 ? "   ← 합이 100%가 아닙니다!" : "   ✔");
            _ratioSum.ForeColor = Math.Abs(total - 100) > 0.001 ? Color.Firebrick : Color.FromArgb(0, 112, 60);
        }

        void SetStatus(string t) { _status.Text = t; }

        /// <summary>화면의 설정값을 읽어 AnalysisSettings로 만든다.</summary>
        public AnalysisSettings ReadSettings()
        {
            var s = Analyzer.DefaultSettings(_files);
            for (int i = 0; i < _perfGrid.Rows.Count && i < s.PerfAreas.Count; i++)
            {
                double v;
                if (double.TryParse(Convert.ToString(_perfGrid.Rows[i].Cells[1].Value), out v)) s.PerfAreas[i].Max = v;
                if (double.TryParse(Convert.ToString(_perfGrid.Rows[i].Cells[2].Value), out v)) s.PerfAreas[i].Ratio = v;
            }
            foreach (DataGridViewRow r in _examGrid.Rows)
            {
                double v; var k = Convert.ToString(r.Cells[0].Value);
                if (k != null && double.TryParse(Convert.ToString(r.Cells[1].Value), out v)) s.ExamRatios[k] = v;
            }
            s.A = (double)_numA.Value; s.B = (double)_numB.Value; s.C = (double)_numC.Value; s.D = (double)_numD.Value;
            s.WrittenDecimals = 2 - _cmbWrittenDec.SelectedIndex;
            s.FinalDecimals = 2 - _cmbFinalDec.SelectedIndex;
            s.Basis = _cmbBasis.SelectedIndex == 0 ? AchievementBasis.RoundedInteger : AchievementBasis.FinalScore;
            s.BoundaryRange = (double)_numBoundary.Value;
            return s;
        }

        // ===================== 분석 실행 =====================
        async Task RunAnalysis()
        {
            if (_files.Count == 0) return;
            var s = ReadSettings();
            double total = s.PerfRatioSum + s.ExamRatios.Values.Sum();
            if (Math.Abs(total - 100) > 0.001 &&
                MessageBox.Show(this, $"반영 비율의 합이 {total:0.##}%입니다(100%가 아님).\n그래도 계속할까요?", "비율 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            if (!(s.A > s.B && s.B > s.C && s.C > s.D)) { MessageBox.Show(this, "성취도 기준은 A > B > C > D 순서여야 합니다.", "기준 확인", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            var first = _files.FirstOrDefault(f => f.Kind != NeisFileKind.Unknown) ?? _files[0];
            string defName = $"성적분석결과_{first.Header?.Grade}학년_{first.Header?.Subject}_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            string outPath;
            using (var dlg = new SaveFileDialog { Filter = "엑셀 파일 (*.xlsx)|*.xlsx", FileName = defName, InitialDirectory = Path.GetDirectoryName(first.Path), Title = "결과 엑셀 파일 저장 위치" })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                outPath = dlg.FileName;
            }
            await RunAnalysisTo(outPath, s);
        }

        public async Task<AnalysisResult> RunAnalysisTo(string outPath, AnalysisSettings s)
        {
            _btnRun.Enabled = false; _progress.Visible = true; SetStatus("분석 중...");
            var files = _files.ToList();
            AnalysisResult result = null; Exception error = null;
            var sw = Stopwatch.StartNew();
            await Task.Run(() =>
            {
                try { result = Analyzer.Analyze(files, s); ReportWriter.Write(result, outPath); }
                catch (Exception ex) { error = ex; }
            });
            _progress.Visible = false; _btnRun.Enabled = true;
            if (error != null)
            {
                SetStatus("분석 실패");
                string msg = error is IOException ? "결과 파일을 저장할 수 없습니다. 같은 이름의 파일이 엑셀에서 열려 있으면 닫고 다시 시도하세요.\n\n" + error.Message : "분석 중 문제가 생겼습니다.\n\n" + error.Message;
                MessageBox.Show(this, msg, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            _lastOutput = outPath; _btnOpen.Enabled = true; _btnOpenFolder.Enabled = true;
            bool? verify = result.ExamAnalyses.Any(e => e.Table != null) ? result.ExamAnalyses.Where(e => e.Table != null).All(e => e.VerifyAllMatch == true) : (bool?)null;
            string vmsg = verify == null ? "일람표가 없어 검산은 생략했습니다." : verify == true ? "검산: 나이스 일람표와 모두 일치 ✔" : "검산: 일람표와 다른 항목이 있습니다. '13.검산' 시트를 확인하세요.";
            SetStatus($"완료 ({sw.Elapsed.TotalSeconds:0.0}초) — 분석 {result.Count}명, 제외 {result.Students.Count - result.Count}명. {vmsg}");
            if (result.Warnings.Count > 0)
                _warnBox.Text = string.Join(Environment.NewLine, result.Warnings.Select(w => "⚠ " + w));
            return result;
        }
    }
}
