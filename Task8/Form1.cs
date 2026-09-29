using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;

namespace Sortix
{
    // ============================================================
    //  АЛГОРИТМЫ СОРТИРОВКИ
    // ============================================================
    public abstract class SortingAlgorithm
    {
        public string Name { get; protected set; } = string.Empty;
        public bool Ascending { get; set; } = true;
        public long Iterations { get; protected set; }

        public event Action<int, int>? OnCompare;
        public event Action<int, int>? OnSwap;

        protected void RaiseCompare(int i, int j) => OnCompare?.Invoke(i, j);
        protected void RaiseSwap(int i, int j) => OnSwap?.Invoke(i, j);
        protected void RaiseIteration() { Iterations++; }

        public abstract void Sort(double[] array);

        protected bool NeedSwap(double a, double b) => Ascending ? a > b : a < b;
    }

    public class BubbleSort : SortingAlgorithm
    {
        public BubbleSort() { Name = "Пузырьковая"; }
        public override void Sort(double[] a)
        {
            int n = a.Length;
            for (int i = 0; i < n - 1; i++)
            {
                bool swapped = false;
                for (int j = 0; j < n - 1 - i; j++)
                {
                    RaiseCompare(j, j + 1);
                    if (NeedSwap(a[j], a[j + 1]))
                    {
                        (a[j], a[j + 1]) = (a[j + 1], a[j]);
                        RaiseSwap(j, j + 1);
                        swapped = true;
                    }
                }
                // Итерацией считаем только завершённый рабочий проход.
                // Проверочный проход без обменов нужен алгоритму, но в статистику не входит.
                if (!swapped) break;
                RaiseIteration();
            }
        }
    }

    public class InsertionSort : SortingAlgorithm
    {
        public InsertionSort() { Name = "Вставками"; }
        public override void Sort(double[] a)
        {
            for (int i = 1; i < a.Length; i++)
            {
                double key = a[i];
                int j = i - 1;
                while (j >= 0)
                {
                    RaiseCompare(j, j + 1);
                    if (!NeedSwap(a[j], key)) break;
                    a[j + 1] = a[j];
                    RaiseSwap(j, j + 1);
                    j--;
                }
                a[j + 1] = key;
                RaiseIteration();
            }
        }
    }

    public class ShakerSort : SortingAlgorithm
    {
        public ShakerSort() { Name = "Шейкерная"; }
        public override void Sort(double[] a)
        {
            int left = 0, right = a.Length - 1;
            while (left < right)
            {
                bool swapped = false;
                for (int i = left; i < right; i++)
                {
                    RaiseCompare(i, i + 1);
                    if (NeedSwap(a[i], a[i + 1]))
                    {
                        (a[i], a[i + 1]) = (a[i + 1], a[i]);
                        RaiseSwap(i, i + 1);
                        swapped = true;
                    }
                }
                right--;

                for (int i = right; i > left; i--)
                {
                    RaiseCompare(i - 1, i);
                    if (NeedSwap(a[i - 1], a[i]))
                    {
                        (a[i - 1], a[i]) = (a[i], a[i - 1]);
                        RaiseSwap(i - 1, i);
                        swapped = true;
                    }
                }
                left++;

                // Один полный цикл шейкерной сортировки = проход вперёд + проход назад.
                // Финальный проверочный цикл без обменов не учитываем.
                if (!swapped) break;
                RaiseIteration();
            }
        }
    }

    public class QuickSort : SortingAlgorithm
    {
        public QuickSort() { Name = "Быстрая"; }
        public override void Sort(double[] a) => QuickSortRec(a, 0, a.Length - 1);

        private void QuickSortRec(double[] a, int low, int high)
        {
            if (low < high)
            {
                int pi = Partition(a, low, high);
                QuickSortRec(a, low, pi - 1);
                QuickSortRec(a, pi + 1, high);
            }
        }

        private int Partition(double[] a, int low, int high)
        {
            double pivot = a[high];
            int i = low - 1;
            for (int j = low; j < high; j++)
            {
                RaiseCompare(j, high);
                bool move = Ascending ? a[j] <= pivot : a[j] >= pivot;
                if (move)
                {
                    i++;
                    (a[i], a[j]) = (a[j], a[i]);
                    RaiseSwap(i, j);
                }
            }
            (a[i + 1], a[high]) = (a[high], a[i + 1]);
            RaiseSwap(i + 1, high);
            RaiseIteration();
            return i + 1;
        }
    }

    public class BogoSort : SortingAlgorithm
    {
        private readonly Random _rnd = new Random();
        public BogoSort() { Name = "BOGO"; }

        public override void Sort(double[] a)
        {
            var start = DateTime.Now;
            while (!IsSorted(a))
            {
                if ((DateTime.Now - start).TotalSeconds > 10)
                    throw new TimeoutException("BOGO превысила лимит 10 секунд.");
                Shuffle(a);
                RaiseIteration();
            }
        }

        private void Shuffle(double[] a)
        {
            for (int i = a.Length - 1; i > 0; i--)
            {
                int j = _rnd.Next(i + 1);
                (a[i], a[j]) = (a[j], a[i]);
                RaiseSwap(i, j);
            }
        }

        private bool IsSorted(double[] a)
        {
            for (int i = 0; i < a.Length - 1; i++)
            {
                RaiseCompare(i, i + 1);
                if (Ascending && a[i] > a[i + 1]) return false;
                if (!Ascending && a[i] < a[i + 1]) return false;
            }
            return true;
        }
    }

    // ============================================================
    //  СОСТОЯНИЕ АЛГОРИТМА
    // ============================================================
    public class AlgorithmState
    {
        public string Name = string.Empty;
        public Color BarColor;
        public double[] Data = Array.Empty<double>();
        public int CompareI = -1, CompareJ = -1;
        public int SwapI = -1, SwapJ = -1;
        public long Comparisons;
        public long Swaps;
        public long Iterations;
        public double AvgMs;
        public double MinMs;
        public double MaxMs;
        public int Runs;
        public bool Finished;
        public string? Error;
        public bool Visualize;

        public readonly object Sync = new object();

        public List<(int type, int i, int j)> EventLog = new();
        public readonly object LogSync = new object();

        public (double[] data, int ci, int cj, int si, int sj,
                long cmp, long swp, long iter,
                double avg, double min, double max, int runs,
                bool finished, string? err) Snapshot()
        {
            lock (Sync)
            {
                var copy = new double[Data.Length];
                Array.Copy(Data, copy, Data.Length);
                return (copy, CompareI, CompareJ, SwapI, SwapJ,
                        Comparisons, Swaps, Iterations,
                        AvgMs, MinMs, MaxMs, Runs,
                        Finished, Error);
            }
        }
    }

    // ============================================================
    //  ГЛАВНАЯ ФОРМА
    // ============================================================
    public partial class Form1 : Form
    {
        // ---------- Палитра ----------
        private static readonly Color BgDark = Color.FromArgb(16, 15, 22);
        private static readonly Color BgPanel = Color.FromArgb(24, 22, 33);
        private static readonly Color BgCard = BgPanel;
        private static readonly Color BgInput = Color.FromArgb(32, 29, 44);
        private static readonly Color Accent = Color.FromArgb(114, 82, 200);
        private static readonly Color AccentHover = Color.FromArgb(128, 96, 212);
        private static readonly Color AccentDown = Color.FromArgb(91, 62, 165);
        private static readonly Color AccentLight = Color.FromArgb(43, 34, 65);
        private static readonly Color TextMain = Color.FromArgb(234, 230, 244);
        private static readonly Color TextDim = Color.FromArgb(163, 153, 181);
        private static readonly Color GridLine = Color.FromArgb(52, 45, 68);
        private static readonly Color Canvas = Color.FromArgb(20, 18, 29);
        private static readonly Color ChartBg = Color.FromArgb(28, 25, 39);
        private static readonly Color ChartGrid = Color.FromArgb(50, 44, 65);
        private static readonly Color CanvasText = Color.FromArgb(236, 231, 246);
        private static readonly Color CanvasMuted = Color.FromArgb(168, 156, 190);
        private static readonly Color CompareClr = Color.FromArgb(153, 199, 255);
        private static readonly Color SwapClr = Color.FromArgb(229, 173, 238);

        // ---------- Константы ----------
        private const int MaxBogoElements = 100;
        private const int MaxVisualizeElements = 50;
        private const int RunsPerAlgorithm = 5;

        // ---------- Контролы ----------
        private Panel headerPanel = null!;
        private Label titleLabel = null!;
        private Label subtitleLabel = null!;

        private Panel leftPanel = null!;
        private Panel toolbarPanel = null!;
        private Button btnGenerate = null!;
        private Button btnExcel = null!;
        private Button btnGoogle = null!;
        private Button btnClear = null!;
        private DataGridView dataGrid = null!;
        private Label lblCount = null!;

        private Panel optionsPanel = null!;
        private CheckBox cbBubble = null!;
        private CheckBox cbInsertion = null!;
        private CheckBox cbShaker = null!;
        private CheckBox cbQuick = null!;
        private CheckBox cbBogo = null!;
        private Label lblDirection = null!;
        private StudioCombo cmbDirection = null!;
        private Label lblDelay = null!;
        private Label lblDelayValue = null!;
        private StudioSlider tbDelay = null!;

        // --- Знаков после запятой ---
        private Label lblDecimals = null!;
        private NumericUpDown nudDecimals = null!;

        private Panel workspacePanel = null!;
        private Label workspaceHint = null!;
        private Label lblStatus = null!;

        private Button btnCalculate = null!;
        private Button btnStop = null!;

        // ---------- Состояние ----------
        private double[] _currentData = Array.Empty<double>();
        private List<AlgorithmState> _states = new();
        private System.Windows.Forms.Timer _renderTimer = null!;
        private bool _isFullscreen = false;
        private bool _isRunning;
        private CancellationTokenSource? _cts;
        private int _decimalPlaces = 3;
        private int _hoveredDeleteRow = -1;

        // ---------- Цвета алгоритмов ----------
        private readonly Dictionary<string, Color> _algoColors = new()
        {
            ["Пузырьковая"] = Color.FromArgb(169, 144, 228),
            ["Вставками"] = Color.FromArgb(195, 173, 238),
            ["Шейкерная"] = Color.FromArgb(151, 143, 216),
            ["Быстрая"] = Color.FromArgb(187, 158, 234),
            ["BOGO"] = Color.FromArgb(190, 141, 209)
        };
        private Panel stagePanel = null!;
        private Label dataTitle = null!, dataNote = null!, algoTitle = null!, stageTitle = null!, stageNote = null!;
        private Button btnReport = null!;
        private TextBox reportBox = null!;
        private readonly ToolTip chartTip = new ToolTip { InitialDelay = 120, ReshowDelay = 80, AutoPopDelay = 10000 };
        private readonly List<(AlgorithmState state, Rectangle area)> _chartAreas = new();
        private bool _showReport;
        private const string EmptyHint = "Добавьте числа. Выберите алгоритмы.\nНажмите «Запустить» и наблюдайте за сортировкой.";

        public Form1()
        {
            InitializeComponent();

            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Sortix — студия сортировок";
            AutoScaleMode = AutoScaleMode.None;
            MinimumSize = new Size(U(1060), U(660));
            Size = new Size(U(1500), U(930));
            BackColor = BgDark;
            KeyPreview = true;
            HandleCreated += (s, e) => SetDarkTitleBar(this);
            if (IsHandleCreated) SetDarkTitleBar(this);

            BuildUi();
            SetupGrid();

            dataGrid.CellValueChanged += (s, e) => SyncDataFromGrid();
            dataGrid.RowsRemoved     += (s, e) => SyncDataFromGrid();
            dataGrid.UserDeletingRow += (s, e) => BeginInvoke((Action)SyncDataFromGrid);

            KeyDown += Form1_KeyDown;
            Resize += (s, e) => LayoutControls();
            DpiChanged += (s, e) => { MinimumSize = new Size(U(1060), U(660)); LayoutControls(); };
            FormClosed += (s, e) => { _renderTimer.Dispose(); chartTip.Dispose(); };

            _renderTimer = new System.Windows.Forms.Timer { Interval = 16 };
            _renderTimer.Tick += (s, e) =>
            {
                if (_isRunning) workspacePanel.Invalidate();
            };
            _renderTimer.Start();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            var area = Screen.FromControl(this).WorkingArea;
            Size = new Size(Math.Min(Width, area.Width - 24), Math.Min(Height, area.Height - 48));
            Location = new Point(area.Left + (area.Width - Width) / 2, area.Top + (area.Height - Height) / 2);
            LayoutControls();
        }

        // ============================================================
        //  FULLSCREEN
        // ============================================================
        private void EnterFullscreen()
        {
            var screen = Screen.FromControl(this);
            FormBorderStyle = FormBorderStyle.None;
            Bounds = screen.Bounds;
            WindowState = FormWindowState.Normal;
            _isFullscreen = true;
        }

        private void ExitFullscreen()
        {
            FormBorderStyle = FormBorderStyle.Sizable;
            WindowState = FormWindowState.Normal;

            var screen = Screen.FromControl(this);
            int w = Math.Min(1200, screen.WorkingArea.Width - 100);
            int h = Math.Min(700, screen.WorkingArea.Height - 100);
            Size = new Size(w, h);
            Location = new Point(
                screen.WorkingArea.Left + (screen.WorkingArea.Width - w) / 2,
                screen.WorkingArea.Top + (screen.WorkingArea.Height - h) / 2);

            _isFullscreen = false;
            LayoutControls();
        }

        // ============================================================
        //  UI
        // ============================================================
        // ---------- Интерфейс SORTIX / Studio ----------
        private int U(float value) => (int)Math.Round(value * DeviceDpi / 96f);

        private static GraphicsPath Rounded(RectangleF r, float radius)
        {
            var path = new GraphicsPath();
            float d = Math.Max(1, Math.Min(radius * 2, Math.Min(r.Width, r.Height)));
            path.AddArc(r.Left, r.Top, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private sealed class Surface : Panel
        {
            public Surface() { DoubleBuffered = true; ResizeRedraw = true; }
        }

        internal sealed class StudioButton : Button
        {
            [System.ComponentModel.DefaultValue(false)]
            public bool Primary { get; set; }
            private bool hovered;
            public StudioButton()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0;
                Cursor = Cursors.Hand; Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            }
            protected override void OnMouseEnter(EventArgs e) { hovered = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { hovered = false; Invalidate(); base.OnMouseLeave(e); }
            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Parent?.BackColor ?? BackColor);
                float scale = DeviceDpi / 96f;
                var r = new RectangleF(1, 1, Math.Max(1, Width - 3), Math.Max(1, Height - 3));
                using var path = Rounded(r, 8 * scale);
                Color fill = !Enabled ? BgInput : Primary ? (hovered ? AccentHover : Accent) : (hovered ? AccentLight : BgInput);
                if (Enabled && Capture && MouseButtons == MouseButtons.Left) fill = Primary ? AccentDown : GridLine;
                using var brush = new SolidBrush(fill); g.FillPath(brush, path);
                using var pen = new Pen(Primary && Enabled ? fill : GridLine); g.DrawPath(pen, path);
                TextRenderer.DrawText(g, Text, Font, ClientRectangle,
                    !Enabled ? TextDim : Primary ? Color.White : TextMain,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                if (Focused && ShowFocusCues)
                    ControlPaint.DrawFocusRectangle(g, Rectangle.Inflate(ClientRectangle, -7, -7));
            }
        }

        internal sealed class StudioCombo : Control
        {
            private readonly List<string> _items = new();
            private int _selectedIndex = -1;
            private bool hovered;

            [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
            [System.ComponentModel.Browsable(false)]
            public IList<string> Items => _items;

            [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
            [System.ComponentModel.DefaultValue(-1)]
            public int SelectedIndex
            {
                get => _selectedIndex;
                set { _selectedIndex = Math.Max(-1, Math.Min(value, _items.Count - 1)); Invalidate(); }
            }

            public StudioCombo()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
                Font = new Font("Segoe UI", 9.5F); Cursor = Cursors.Hand; TabStop = true;
            }

            public void AddRange(IEnumerable<string> items)
            {
                _items.AddRange(items); if (_selectedIndex < 0 && _items.Count > 0) _selectedIndex = 0; Invalidate();
            }

            protected override void OnMouseEnter(EventArgs e) { hovered = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { hovered = false; Invalidate(); base.OnMouseLeave(e); }
            protected override void OnClick(EventArgs e)
            {
                base.OnClick(e); Focus();
                if (Enabled && _items.Count > 0) SelectedIndex = (SelectedIndex + 1) % _items.Count;
            }
            protected override void OnKeyDown(KeyEventArgs e)
            {
                if (_items.Count > 0 && (e.KeyCode == Keys.Space || e.KeyCode == Keys.Enter || e.KeyCode == Keys.Down))
                { SelectedIndex = (SelectedIndex + 1) % _items.Count; e.Handled = true; }
                else if (_items.Count > 0 && e.KeyCode == Keys.Up)
                { SelectedIndex = (SelectedIndex - 1 + _items.Count) % _items.Count; e.Handled = true; }
                base.OnKeyDown(e);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias; g.Clear(Parent?.BackColor ?? BgCard);
                float s = DeviceDpi / 96f; int P(float n) => (int)Math.Round(n * s);
                var r = new RectangleF(0.5F, 0.5F, Width - 1.5F, Height - 1.5F);
                using var path = Rounded(r, P(5));
                using var fill = new SolidBrush(hovered && Enabled ? AccentLight : BgInput);
                using var border = new Pen(Focused ? Accent : GridLine);
                g.FillPath(fill, path); g.DrawPath(border, path);
                string text = SelectedIndex >= 0 && SelectedIndex < _items.Count ? _items[SelectedIndex] : "";
                TextRenderer.DrawText(g, text, Font, new Rectangle(P(7), 0, Math.Max(1, Width - P(31)), Height),
                    Enabled ? TextMain : TextDim, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                float cx = Width - P(14), cy = Height / 2F;
                using var pen = new Pen(Enabled ? Color.FromArgb(183, 159, 232) : TextDim, 1.5F * s)
                    { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round };
                g.DrawLines(pen, new[] { new PointF(cx - P(3), cy - P(1)), new PointF(cx, cy + P(2)), new PointF(cx + P(3), cy - P(1)) });
            }
        }

        internal sealed class StudioSlider : Control
        {
            private int _minimum, _maximum = 100, _value;
            private bool dragging, hovered;
            public event EventHandler? ValueChanged;

            [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
            [System.ComponentModel.DefaultValue(0)]
            public int Minimum { get => _minimum; set { _minimum = value; if (_maximum < value) _maximum = value; Value = _value; } }

            [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
            [System.ComponentModel.DefaultValue(100)]
            public int Maximum { get => _maximum; set { _maximum = Math.Max(value, _minimum); Value = _value; } }

            [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
            [System.ComponentModel.DefaultValue(0)]
            public int Value
            {
                get => _value;
                set { int v = Math.Max(Minimum, Math.Min(Maximum, value)); if (_value == v) return; _value = v; Invalidate(); ValueChanged?.Invoke(this, EventArgs.Empty); }
            }
            public StudioSlider()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
                Cursor = Cursors.Hand; TabStop = true;
            }
            private void SetFromX(int x)
            {
                int pad = Math.Max(7, Height / 4), width = Math.Max(1, Width - pad * 2);
                double t = Math.Max(0, Math.Min(1, (x - pad) / (double)width));
                Value = Minimum + (int)Math.Round(t * (Maximum - Minimum));
            }
            protected override void OnMouseDown(MouseEventArgs e) { if (e.Button == MouseButtons.Left && Enabled) { dragging = true; Capture = true; Focus(); SetFromX(e.X); } base.OnMouseDown(e); }
            protected override void OnMouseMove(MouseEventArgs e) { if (dragging) SetFromX(e.X); base.OnMouseMove(e); }
            protected override void OnMouseUp(MouseEventArgs e) { dragging = false; Capture = false; base.OnMouseUp(e); }
            protected override void OnMouseEnter(EventArgs e) { hovered = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { hovered = false; Invalidate(); base.OnMouseLeave(e); }
            protected override void OnKeyDown(KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Down) { Value--; e.Handled = true; }
                if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Up) { Value++; e.Handled = true; }
                base.OnKeyDown(e);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias; g.Clear(Parent?.BackColor ?? BgCard);
                float s = DeviceDpi / 96f; int pad = Math.Max(7, Height / 4); int cy = Height / 2;
                int trackW = Math.Max(1, Width - pad * 2);
                float t = Maximum == Minimum ? 0 : (Value - Minimum) / (float)(Maximum - Minimum);
                int knobX = pad + (int)Math.Round(trackW * t);
                using var basePen = new Pen(GridLine, Math.Max(3, 3 * s)) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                using var activePen = new Pen(Enabled ? Accent : TextDim, Math.Max(3, 3 * s)) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                g.DrawLine(basePen, pad, cy, Width - pad, cy); g.DrawLine(activePen, pad, cy, knobX, cy);
                float radius = (hovered || dragging || Focused ? 6.5F : 5.5F) * s;
                using var knob = new SolidBrush(Enabled ? (dragging ? AccentDown : Accent) : TextDim);
                g.FillEllipse(knob, knobX - radius, cy - radius, radius * 2, radius * 2);
            }
        }

        internal sealed class StudioNumber : NumericUpDown
        {
            public StudioNumber()
            {
                BackColor = BgInput; ForeColor = TextMain; BorderStyle = BorderStyle.None;
                Font = new Font("Segoe UI", 9.5F);
                TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
                foreach (Control child in Controls)
                {
                    if (child is TextBoxBase) continue;
                    SetDoubleBuffered(child);
                    // Paint вызывается после штатной отрисовки кнопок NumericUpDown.
                    child.Paint += PaintArrows;
                }
            }

            private void PaintArrows(object? sender, PaintEventArgs e)
            {
                if (sender is not Control buttons || buttons.Width < 2 || buttons.Height < 2) return;
                var g = e.Graphics; g.Clear(BackColor); g.SmoothingMode = SmoothingMode.AntiAlias;
                float scale = DeviceDpi / 96f;
                var mouse = buttons.PointToClient(Cursor.Position);
                bool pressed = buttons.Capture && Control.MouseButtons == MouseButtons.Left;
                int half = buttons.Height / 2;
                for (int i = 0; i < 2; i++)
                {
                    var r = new Rectangle(0, i == 0 ? 0 : half, buttons.Width, i == 0 ? half : buttons.Height - half);
                    bool hot = Enabled && r.Contains(mouse);
                    using var fill = new SolidBrush(hot ? pressed ? AccentDown : AccentLight : BgInput);
                    g.FillRectangle(fill, r);
                    float x = r.Left + r.Width / 2F, y = r.Top + r.Height / 2F;
                    float dx = Math.Min(3 * scale, r.Width / 4F), dy = Math.Min(1.5F * scale, r.Height / 4F);
                    float direction = i == 0 ? -1 : 1;
                    using var pen = new Pen(!Enabled ? TextDim : hot ? CanvasText : _arrowColor, 1.4F * scale)
                        { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round };
                    g.DrawLines(pen, new[] { new PointF(x - dx, y - direction * dy),
                        new PointF(x, y + direction * dy), new PointF(x + dx, y - direction * dy) });
                }
                using var separator = new Pen(GridLine);
                g.DrawLine(separator, 0, 0, 0, buttons.Height - 1);
                g.DrawLine(separator, 1, half, buttons.Width - 1, half);
            }
            private static readonly Color _arrowColor = Color.FromArgb(183, 159, 232);
        }

        private sealed class AlgorithmTile : CheckBox
        {
            public Color Tint;
            public string Detail = "";
            public string Number = "";
            private bool hovered;
            public AlgorithmTile()
            {
                Appearance = Appearance.Button; FlatStyle = FlatStyle.Flat;
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                Cursor = Cursors.Hand; Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            }
            protected override void OnMouseEnter(EventArgs e) { hovered = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { hovered = false; Invalidate(); base.OnMouseLeave(e); }
            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Parent?.BackColor ?? BackColor);
                float s = DeviceDpi / 96f;
                int P(float n) => (int)Math.Round(n * s);
                using var path = Rounded(new RectangleF(1, 1, Width - 3, Height - 3), P(8));
                using var fill = new SolidBrush(Checked || hovered ? AccentLight : BgInput);
                using var border = new Pen(Checked ? Accent : GridLine, 1);
                g.FillPath(fill, path); g.DrawPath(border, path);
                using var small = new Font("Segoe UI", 8F);
                Color ink = Checked ? CanvasText : TextMain;
                TextRenderer.DrawText(g, Number, small, new Point(P(12), P(9)), Checked ? Tint : TextDim);
                var circle = new Rectangle(P(Width / s - 25), P(10), P(12), P(12));
                using var dot = new SolidBrush(Checked ? Tint : GridLine); g.FillEllipse(dot, circle);
                if (Checked)
                {
                    using var check = new Pen(Canvas, 1.6F * s);
                    g.DrawLines(check, new[] { new PointF(circle.Left + 3 * s, circle.Top + 6 * s),
                        new PointF(circle.Left + 5 * s, circle.Top + 8 * s), new PointF(circle.Left + 9 * s, circle.Top + 4 * s) });
                }
                TextRenderer.DrawText(g, Text, Font, new Rectangle(P(10), P(28), Width - P(20), P(24)), ink,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                TextRenderer.DrawText(g, Detail, small, new Rectangle(P(10), P(52), Width - P(20), P(20)),
                    Checked ? CanvasMuted : TextDim, TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
                if (Focused && ShowFocusCues) ControlPaint.DrawFocusRectangle(g, Rectangle.Inflate(ClientRectangle, -6, -6));
            }
        }

        private Label LabelAt(string text, float size, Color color, Control parent, bool bold = false)
        {
            var label = new Label { Text = text, AutoSize = true, ForeColor = color,
                BackColor = Color.Transparent, Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular) };
            parent.Controls.Add(label); return label;
        }

        private Button MakeSmallButton(string text) => new StudioButton { Text = text, BackColor = BgCard };

        private CheckBox MakeAlgoCheck(string name, string detail, int number, bool selected)
        {
            var tile = new AlgorithmTile { Text = name, Detail = detail, Number = number.ToString("00"),
                Tint = _algoColors[name], Checked = selected, BackColor = BgCard, AccessibleName = name };
            tile.CheckedChanged += (s, e) => headerPanel.Invalidate();
            optionsPanel.Controls.Add(tile); return tile;
        }

        private void BuildUi()
        {
            SuspendLayout();
            headerPanel = new Surface { Dock = DockStyle.Top, BackColor = BgDark };
            headerPanel.Paint += HeaderPanel_Paint;
            titleLabel = LabelAt("SORTIX", 27, TextMain, headerPanel, true);
            subtitleLabel = LabelAt("СТУДИЯ СОРТИРОВОК  /  Исследуйте порядок", 9, TextDim, headerPanel);
            leftPanel = new Surface { BackColor = BgCard };
            optionsPanel = new Surface { BackColor = BgCard };
            stagePanel = new Surface { BackColor = Canvas };
            foreach (var panel in new[] { leftPanel, optionsPanel, stagePanel })
            {
                panel.Resize += (s, e) => RoundControl((Control)s!, U(12));
                panel.Paint += CardPanel_Paint;
            }
            dataTitle = LabelAt("01  /  ДАННЫЕ", 10, TextMain, leftPanel, true);
            dataNote = LabelAt("Ваш исходный набор чисел", 9, TextDim, leftPanel);
            algoTitle = LabelAt("02  /  ВЫБЕРИТЕ АЛГОРИТМЫ", 9.5F, TextMain, optionsPanel, true);
            stageTitle = LabelAt("Визуализация", 13, CanvasText, stagePanel, true);
            stageNote = LabelAt("Сравнение  /  голубой     Обмен  /  розовый", 8, CanvasMuted, stagePanel);
            btnReport = MakeSmallButton("Отчёт  ↗"); btnReport.Enabled = false;
            btnReport.BackColor = Canvas; btnReport.Click += (s, e) => { _showReport = !_showReport; UpdateWorkspace(); };
            stagePanel.Controls.Add(btnReport);

            toolbarPanel = new Panel { BackColor = BgCard };
            btnGenerate = MakeSmallButton("＋  Генерация");
            btnExcel = MakeSmallButton("Excel  ↗"); btnGoogle = MakeSmallButton("Google Sheets");
            btnClear = MakeSmallButton("Очистить"); btnClear.CausesValidation = false;
            btnGenerate.Click += BtnGenerate_Click; btnExcel.Click += BtnExcel_Click;
            btnGoogle.Click += BtnGoogle_Click; btnClear.Click += BtnClear_Click;
            toolbarPanel.Controls.AddRange(new Control[] { btnGenerate, btnClear, btnExcel, btnGoogle });
            leftPanel.Controls.Add(toolbarPanel);
            dataGrid = new DataGridView
            {
                AllowUserToAddRows = true, AllowUserToDeleteRows = true, EditMode = DataGridViewEditMode.EditOnEnter,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = U(28), RowHeadersVisible = true, RowHeadersWidth = U(34),
                RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing,
                BackgroundColor = BgCard, GridColor = GridLine, BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.None, ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None, AllowUserToResizeRows = false,
                EnableHeadersVisualStyles = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ScrollBars = ScrollBars.Vertical, Font = new Font("Segoe UI", 9.5F), ForeColor = TextMain
            };
            dataGrid.RowTemplate.Height = U(24);
            dataGrid.DefaultCellStyle.BackColor = BgCard; dataGrid.DefaultCellStyle.ForeColor = TextMain;
            dataGrid.DefaultCellStyle.SelectionBackColor = AccentLight; dataGrid.DefaultCellStyle.SelectionForeColor = TextMain;
            dataGrid.DefaultCellStyle.Padding = new Padding(U(5), 0, U(5), 0);
            dataGrid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(28, 25, 39);
            dataGrid.ColumnHeadersDefaultCellStyle.BackColor = BgInput;
            dataGrid.ColumnHeadersDefaultCellStyle.ForeColor = TextDim;
            dataGrid.ColumnHeadersDefaultCellStyle.SelectionBackColor = BgInput;
            dataGrid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dataGrid.RowHeadersDefaultCellStyle.BackColor = BgInput;
            dataGrid.RowHeadersDefaultCellStyle.ForeColor = TextDim;
            dataGrid.RowHeadersDefaultCellStyle.SelectionBackColor = AccentLight;
            dataGrid.RowHeadersDefaultCellStyle.SelectionForeColor = TextMain;
            dataGrid.CellPainting += (s, e) =>
            {
                if (e.ColumnIndex != -1 || e.RowIndex < 0 || e.Graphics == null) return;
                using var background = new SolidBrush(BgCard);
                e.Graphics.FillRectangle(background, e.CellBounds);

                bool isNew = dataGrid.Rows[e.RowIndex].IsNewRow;
                bool active = !isNew && (e.RowIndex == _hoveredDeleteRow ||
                    dataGrid.CurrentCell?.RowIndex == e.RowIndex);

                if (active)
                {
                    var deleteRect = new Rectangle(e.CellBounds.Left + U(2), e.CellBounds.Top,
                        U(15), e.CellBounds.Height);
                    TextRenderer.DrawText(e.Graphics, "−", new Font("Segoe UI", 10F, FontStyle.Bold),
                        deleteRect, SwapClr,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }

                var numberRect = new Rectangle(e.CellBounds.Left + (isNew ? 0 : U(15)), e.CellBounds.Top,
                    Math.Max(1, e.CellBounds.Width - (isNew ? 0 : U(15))), e.CellBounds.Height);
                TextRenderer.DrawText(e.Graphics, isNew ? "+" : (e.RowIndex + 1).ToString(),
                    dataGrid.Font, numberRect, TextDim,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                e.Handled = true;
            };
            dataGrid.CellMouseMove += (s, e) =>
            {
                int row = e.ColumnIndex == -1 && e.RowIndex >= 0 && !dataGrid.Rows[e.RowIndex].IsNewRow
                    ? e.RowIndex : -1;
                if (_hoveredDeleteRow != row)
                {
                    int old = _hoveredDeleteRow; _hoveredDeleteRow = row;
                    if (old >= 0 && old < dataGrid.Rows.Count) dataGrid.InvalidateCell(-1, old);
                    if (row >= 0) dataGrid.InvalidateCell(-1, row);
                }
            };
            dataGrid.MouseLeave += (s, e) =>
            {
                int old = _hoveredDeleteRow; _hoveredDeleteRow = -1;
                if (old >= 0 && old < dataGrid.Rows.Count) dataGrid.InvalidateCell(-1, old);
            };
            dataGrid.CurrentCellChanged += (s, e) =>
            {
                // Row headers have ColumnIndex == -1, but -1 is not a real column and
                // InvalidateColumn(-1) throws ArgumentOutOfRangeException.
                // Repaint the row-header area directly instead.
                dataGrid.Invalidate();
            };
            dataGrid.CellMouseClick += (s, e) =>
            {
                if (_isRunning || e.Button != MouseButtons.Left || e.ColumnIndex != -1 || e.RowIndex < 0) return;
                if (dataGrid.Rows[e.RowIndex].IsNewRow || e.X > U(17)) return;

                dataGrid.CancelEdit();
                dataGrid.Rows.RemoveAt(e.RowIndex);
                _hoveredDeleteRow = -1;
                SyncDataFromGrid();
                dataGrid.Invalidate();
            };
            dataGrid.CellValidating += DataGrid_CellValidating;
            SetDoubleBuffered(dataGrid); leftPanel.Controls.Add(dataGrid);
            lblCount = LabelAt("Элементов: 0", 9, TextDim, leftPanel);
            cbBubble = MakeAlgoCheck("Пузырьковая", "Обмен соседей", 1, true);
            cbInsertion = MakeAlgoCheck("Вставками", "Сдвиг элементов", 2, false);
            cbShaker = MakeAlgoCheck("Шейкерная", "Вперёд и назад", 3, false);
            cbQuick = MakeAlgoCheck("Быстрая", "Разделение", 4, true);
            cbBogo = MakeAlgoCheck("BOGO", "Случайный поиск", 5, false);

            lblDirection = LabelAt("ПОРЯДОК", 8, TextDim, optionsPanel, true);
            cmbDirection = new StudioCombo { BackColor = BgInput, ForeColor = TextMain, Font = new Font("Segoe UI", 10F) };
            cmbDirection.AddRange(new[] { "По возрастанию", "По убыванию" }); cmbDirection.SelectedIndex = 0;
            lblDecimals = LabelAt("ЗНАКОВ ПОСЛЕ ЗАПЯТОЙ", 8, TextDim, optionsPanel, true);
            nudDecimals = new StudioNumber { Minimum = 0, Maximum = 10, Value = 3,
                BackColor = BgInput, ForeColor = TextMain, BorderStyle = BorderStyle.None,
                TextAlign = System.Windows.Forms.HorizontalAlignment.Center, Font = new Font("Segoe UI", 9.5F) };
            lblDelay = LabelAt("ЗАДЕРЖКА", 8, TextDim, optionsPanel, true);
            lblDelayValue = LabelAt("15 мс", 9, TextMain, optionsPanel, true);
            tbDelay = new StudioSlider { Minimum = 0, Maximum = 100, Value = 15,
                BackColor = BgCard, AccessibleName = "Задержка визуализации, миллисекунд" };
            tbDelay.ValueChanged += (s, e) => lblDelayValue.Text = $"{tbDelay.Value} мс";
            optionsPanel.Controls.AddRange(new Control[] { cmbDirection, nudDecimals, tbDelay });
            nudDecimals.ValueChanged += (s, e) =>
            {
                _decimalPlaces = (int)nudDecimals.Value;
                if (_currentData.Length > 0)
                {
                    var rounded = _currentData.Select(v => Math.Round(v, _decimalPlaces, MidpointRounding.AwayFromZero)).ToArray();
                    dataGrid.SuspendLayout();
                    for (int i = 0; i < dataGrid.Rows.Count && i < rounded.Length; i++)
                        if (!dataGrid.Rows[i].IsNewRow) dataGrid.Rows[i].Cells[0].Value = rounded[i];
                    dataGrid.ResumeLayout(); _currentData = rounded;
                }
                workspacePanel.Invalidate();
            };

            workspacePanel = new Surface { BackColor = Canvas, AutoScroll = true };
            workspacePanel.Paint += WorkspacePanel_Paint;
            workspacePanel.MouseMove += ChartMouseMove;
            workspacePanel.MouseLeave += (s, e) => chartTip.SetToolTip(workspacePanel, "");
            workspaceHint = new Label { Text = EmptyHint, Font = new Font("Segoe UI", 10F),
                ForeColor = CanvasMuted, BackColor = Canvas, TextAlign = ContentAlignment.MiddleCenter };
            workspacePanel.Controls.Add(workspaceHint); stagePanel.Controls.Add(workspacePanel);
            reportBox = new TextBox { Multiline = true, ReadOnly = true, WordWrap = false,
                ScrollBars = ScrollBars.Both, BorderStyle = BorderStyle.None, BackColor = Canvas,
                ForeColor = CanvasText, Font = new Font("Consolas", 9.5F), Visible = false };
            stagePanel.Controls.Add(reportBox);
            btnCalculate = new StudioButton { Text = "Запустить   →", Primary = true, BackColor = BgDark,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold) };
            btnStop = MakeSmallButton("■  Стоп"); btnStop.BackColor = BgDark; btnStop.Enabled = false;
            btnCalculate.Click += BtnCalculate_Click; btnStop.Click += BtnStop_Click;
            lblStatus = new Label { Text = "Готово к запуску", Font = new Font("Segoe UI", 9F),
                ForeColor = TextDim, BackColor = BgDark, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
            lblStatus.TextChanged += (s, e) => headerPanel.Invalidate();
            Controls.AddRange(new Control[] { leftPanel, stagePanel, optionsPanel, lblStatus, btnStop, btnCalculate, headerPanel });
            headerPanel.BringToFront(); ResumeLayout(false);
        }

        private static void SetDoubleBuffered(Control control)
        {
            typeof(Control).GetProperty("DoubleBuffered",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?.SetValue(control, true, null);
        }

        [System.Runtime.InteropServices.DllImport("dwmapi.dll", ExactSpelling = true)]
        private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);

        private static void SetDarkTitleBar(Form form)
        {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT) return;
            int dark = 1, background = ColorTranslator.ToWin32(BgDark), text = ColorTranslator.ToWin32(TextMain);
            // Неподдерживаемые атрибуты возвращают HRESULT: рамка остаётся системной.
            DwmSetWindowAttribute(form.Handle, 20, ref dark, sizeof(int));
            DwmSetWindowAttribute(form.Handle, 35, ref background, sizeof(int));
            DwmSetWindowAttribute(form.Handle, 36, ref text, sizeof(int));
        }

        internal static void ApplyDialogStyle(Form dialog)
        {
            dialog.BackColor = BgDark; dialog.ForeColor = TextMain;
            dialog.HandleCreated += (s, e) => SetDarkTitleBar(dialog);
            if (dialog.IsHandleCreated) SetDarkTitleBar(dialog);
            void Style(Control c)
            {
                c.ForeColor = Equals(c.Tag, "muted") ? TextDim : TextMain;
                c.BackColor = c is TextBoxBase || c is NumericUpDown ? BgInput : BgDark;
                if (c is Button b)
                {
                    if (b is StudioButton studio) studio.Primary = b.DialogResult == DialogResult.OK;
                    b.Cursor = Cursors.Hand;
                }
                // Внутренние кнопки и редактор NumericUpDown оформляет StudioNumber.
                if (c is NumericUpDown) return;
                foreach (Control child in c.Controls) Style(child);
            }
            foreach (Control c in dialog.Controls) Style(c);
        }

        internal static TableLayoutPanel CreateDialogBody(Form dialog, string title, Size size, int rows)
        {
            dialog.SuspendLayout();
            dialog.AutoScaleDimensions = new SizeF(96F, 96F);
            dialog.AutoScaleMode = AutoScaleMode.Dpi;
            dialog.Font = new Font("Segoe UI", 9F);
            dialog.Text = title; dialog.ClientSize = size; dialog.Padding = new Padding(18);
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.StartPosition = FormStartPosition.CenterParent;
            dialog.MaximizeBox = false; dialog.MinimizeBox = false; dialog.ShowInTaskbar = false;
            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = rows,
                Margin = Padding.Empty, Padding = Padding.Empty };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            dialog.Controls.Add(body); return body;
        }

        internal static FlowLayoutPanel CreateDialogActions(Button accept, Button cancel)
        {
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false, Margin = new Padding(0, 14, 0, 0), Padding = Padding.Empty };
            foreach (var button in new[] { cancel, accept })
            {
                button.Size = new Size(112, 34); button.Margin = new Padding(8, 0, 0, 0);
                actions.Controls.Add(button);
            }
            cancel.CausesValidation = false;
            return actions;
        }

        private static string FormatTime(double ms)
        {
            if (ms < 0) ms = 0;

            if (ms < 1000.0)
                return $"{ms:F4} мс";

            double sec = ms / 1000.0;
            if (sec < 60.0)
                return $"{sec:F4} с";

            double min = Math.Floor(sec / 60.0);
            double remSec = sec - min * 60.0;
            return $"{min:F0} мин {remSec:F2} с";
        }

        // ============================================================
        //  ХЕЛПЕР: парсинг double с запятой и точкой
        // ============================================================
        private static bool TryParseDouble(string s, out double value)
        {
            s = s.Trim();
            return
                double.TryParse(s, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.CurrentCulture, out value) ||
                double.TryParse(s, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out value);
        }

        // ============================================================
        //  РАСКЛАДКА
        // ============================================================
        private static void RoundControl(Control control, int radius)
        {
            if (control.Width < 2 || control.Height < 2) return;
            using var path = Rounded(new RectangleF(0, 0, control.Width, control.Height), radius);
            var previous = control.Region; control.Region = new Region(path); previous?.Dispose();
        }

        private void LayoutControls()
        {
            if (workspacePanel == null || btnCalculate == null) return;
            int m = U(22), gap = U(16), top = U(110), bottom = ClientSize.Height - m;
            headerPanel.Height = U(90);
            titleLabel.Location = new Point(U(86), U(13)); subtitleLabel.Location = new Point(U(90), U(59));
            int leftW = Math.Max(U(248), Math.Min(U(300), ClientSize.Width / 5));
            leftPanel.Bounds = new Rectangle(m, top, leftW, Math.Max(U(220), bottom - top));
            int x = leftPanel.Right + gap, w = Math.Max(U(680), ClientSize.Width - m - x);
            optionsPanel.Bounds = new Rectangle(x, bottom - U(52 + 14 + 206), w, U(206));
            stagePanel.Bounds = new Rectangle(x, top, w, Math.Max(U(100), optionsPanel.Top - gap - top));
            stageTitle.Location = new Point(U(20), U(15)); stageNote.Location = new Point(U(21), U(43));
            btnReport.Bounds = new Rectangle(w - U(136), U(17), U(116), U(34));
            workspacePanel.Bounds = new Rectangle(U(8), U(69), w - U(16), Math.Max(U(30), stagePanel.Height - U(80)));
            reportBox.Bounds = new Rectangle(U(20), U(73), w - U(40), Math.Max(U(30), stagePanel.Height - U(90)));
            dataTitle.Location = new Point(U(18), U(20)); dataNote.Location = new Point(U(18), U(46));
            toolbarPanel.Bounds = new Rectangle(U(14), U(81), leftW - U(28), U(87));
            int first = (toolbarPanel.Width - U(8)) * 3 / 5;
            btnGenerate.Bounds = new Rectangle(0, 0, first, U(39));
            btnClear.Bounds = new Rectangle(first + U(8), 0, toolbarPanel.Width - first - U(8), U(39));
            int half = (toolbarPanel.Width - U(8)) / 2;
            btnExcel.Bounds = new Rectangle(0, U(47), half, U(38));
            btnGoogle.Bounds = new Rectangle(half + U(8), U(47), toolbarPanel.Width - half - U(8), U(38));
            dataGrid.Bounds = new Rectangle(U(14), U(186), leftW - U(28), Math.Max(U(60), leftPanel.Height - U(232)));
            lblCount.Location = new Point(U(20), leftPanel.Height - U(31));

            algoTitle.Location = new Point(U(18), U(15));
            var choices = new[] { cbBubble, cbInsertion, cbShaker, cbQuick, cbBogo };
            int tileW = (w - U(36 + 8 * 4)) / 5;
            for (int i = 0; i < choices.Length; i++)
                choices[i].Bounds = new Rectangle(U(18) + i * (tileW + U(8)), U(43), tileW, U(77));
            int col1 = U(18), col2 = U(217), col3 = U(411);
            lblDirection.Location = new Point(col1, U(139));
            cmbDirection.Bounds = new Rectangle(col1, U(161), U(177), U(28));
            lblDecimals.Location = new Point(col2, U(139));
            nudDecimals.Bounds = new Rectangle(col2, U(161), U(60), U(24));
            lblDelay.Location = new Point(col3, U(139));
            lblDelayValue.Location = new Point(w - U(76), U(137));
            tbDelay.Bounds = new Rectangle(col3 - U(8), U(157), w - col3 - U(12), U(34));
            btnCalculate.Bounds = new Rectangle(ClientSize.Width - m - U(208), bottom - U(52), U(208), U(52));
            btnStop.Bounds = new Rectangle(btnCalculate.Left - U(110), btnCalculate.Top + U(5), U(96), U(42));
            lblStatus.Bounds = new Rectangle(x + U(3), btnCalculate.Top, Math.Max(U(100), btnStop.Left - x - U(20)), U(52));
            workspacePanel.Invalidate();
        }

        private void HeaderPanel_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
            int w = headerPanel.ClientSize.Width;
            using var accentBrush = new SolidBrush(Accent);
            for (int i = 0; i < 4; i++)
            {
                using var bar = Rounded(new RectangleF(U(26 + i * 11), U(53 - i * 9), U(7), U(14 + i * 9)), U(3));
                g.FillPath(accentBrush, bar);
            }
            using var line = new Pen(GridLine); g.DrawLine(line, U(22), U(89), w - U(22), U(89));
            if (cbBubble == null) return;
            int selected = new[] { cbBubble, cbInsertion, cbShaker, cbQuick, cbBogo }.Count(c => c.Checked);
            using var font = new Font("Segoe UI", 9F, FontStyle.Bold);
            using var small = new Font("Segoe UI", 8F);
            TextRenderer.DrawText(g, $"{_currentData.Length:N0} ЭЛЕМЕНТОВ    /    ВЫБРАНО: {selected} ИЗ 5", font,
                new Rectangle(w - U(590), U(21), U(375), U(24)), TextMain,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
            TextRenderer.DrawText(g, "F11  /  На весь экран", small,
                new Rectangle(w - U(580), U(50), U(365), U(23)), TextDim, TextFormatFlags.Right);
            var chip = new Rectangle(w - U(191), U(23), U(167), U(42));
            using var path = Rounded(chip, U(8));
            using var fill = new SolidBrush(_isRunning ? AccentLight : BgInput); g.FillPath(fill, path);
            TextRenderer.DrawText(g, _isRunning ? "●  В РАБОТЕ" : "●  ГОТОВО", font, chip,
                _isRunning ? CanvasText : TextDim, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        private void CardPanel_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not Panel p || p.Width < 2 || p.Height < 2) return;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var path = Rounded(new RectangleF(0.5F, 0.5F, p.Width - 1, p.Height - 1), U(12));
            using var pen = new Pen(p == stagePanel ? ChartGrid : GridLine);
            e.Graphics.DrawPath(pen, path);
        }

        private void UpdateWorkspace()
        {
            reportBox.Visible = _showReport; workspacePanel.Visible = !_showReport;
            btnReport.Text = _showReport ? "←  Графики" : "Отчёт  ↗";
            workspaceHint.Text = _states.Count > 0 ? "Визуализация доступна до 50 элементов.\nРезультаты расчёта — в отчёте." : EmptyHint;
            workspaceHint.Visible = !_states.Any(s => s.Visualize);
            workspacePanel.AutoScrollMinSize = Size.Empty;
            workspacePanel.Invalidate(); headerPanel.Invalidate();
        }

        private void WorkspacePanel_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias; g.Clear(Canvas);
            _chartAreas.Clear();
            var states = _states.Where(s => s.Visualize).ToList();
            if (states.Count == 0)
            {
                workspacePanel.AutoScrollMinSize = Size.Empty;
                int w = workspacePanel.ClientSize.Width, h = workspacePanel.ClientSize.Height;
                using var grid = new SolidBrush(ChartGrid);
                for (int x = U(18); x < w; x += U(24))
                    for (int y = U(6); y < h; y += U(24)) g.FillEllipse(grid, x, y, U(2), U(2));
                if (h > U(155))
                {
                    int center = w / 2, baseY = Math.Max(U(66), h / 2 - U(12));
                    int[] heights = { 21, 54, 32, 78, 44, 92, 64 };
                    for (int i = 0; i < heights.Length; i++)
                    {
                        using var bar = Rounded(new RectangleF(center + U((i - 3) * 24 - 9), baseY - U(heights[i]), U(15), U(heights[i])), U(6));
                        using var fill = new SolidBrush(i == 3 ? Accent : i == 5 ? _algoColors["Быстрая"] : ChartGrid);
                        g.FillPath(fill, bar);
                    }
                }
                workspaceHint.Bounds = new Rectangle(U(28), Math.Max(0, h / 2 + (h > U(155) ? U(10) : -U(35))), w - U(56), Math.Min(U(72), h));
                workspaceHint.Visible = true; return;
            }
            workspaceHint.Visible = false;
            int pad = U(10), available = workspacePanel.ClientSize.Width - SystemInformation.VerticalScrollBarWidth;
            int cols = Math.Min(states.Count, available >= U(930) ? 3 : available >= U(580) ? 2 : 1);
            int rows = (states.Count + cols - 1) / cols;
            int cellW = Math.Max(1, (available - pad * (cols + 1)) / cols);
            int cellH = Math.Max(U(194), (workspacePanel.ClientSize.Height - pad * (rows + 1)) / rows);
            var size = new Size(0, rows * (cellH + pad) + pad);
            if (workspacePanel.AutoScrollMinSize != size) workspacePanel.AutoScrollMinSize = size;
            for (int i = 0; i < states.Count; i++)
            {
                var area = new Rectangle(pad + i % cols * (cellW + pad) + workspacePanel.AutoScrollPosition.X,
                    pad + i / cols * (cellH + pad) + workspacePanel.AutoScrollPosition.Y, cellW, cellH);
                _chartAreas.Add((states[i], area));
                if (area.IntersectsWith(workspacePanel.ClientRectangle)) DrawAlgorithm(g, states[i], area);
            }
        }

        private Rectangle PlotBounds(Rectangle area) => new Rectangle(area.Left + U(13), area.Top + U(57),
            Math.Max(1, area.Width - U(26)), Math.Max(1, area.Height - U(119)));

        private void DrawAlgorithm(Graphics g, AlgorithmState st, Rectangle area)
        {
            var snap = st.Snapshot(); var data = snap.data;
            using var path = Rounded(area, U(14)); using var fill = new SolidBrush(ChartBg);
            using var border = new Pen(ChartGrid); g.FillPath(fill, path); g.DrawPath(border, path);
            using var titleFont = new Font("Segoe UI", 10F, FontStyle.Bold);
            using var infoFont = new Font("Segoe UI", 8F); using var valueFont = new Font("Segoe UI", 7.5F);
            TextRenderer.DrawText(g, st.Name, titleFont, new Rectangle(area.Left + U(13), area.Top + U(10), area.Width - U(40), U(23)),
                st.BarColor, TextFormatFlags.EndEllipsis);
            TextRenderer.DrawText(g, snap.finished ? snap.err == null ? "✓" : "!" : "•", titleFont,
                new Point(area.Right - U(27), area.Top + U(10)), snap.err == null ? st.BarColor : SwapClr);
            string timing = snap.err != null ? "Ошибка: " + snap.err : snap.runs > 0
                ? $"≈ {FormatTime(snap.avg)}  ·  прогонов: {snap.runs}" : "Подготовка замера…";
            TextRenderer.DrawText(g, timing, infoFont, new Rectangle(area.Left + U(13), area.Top + U(34), area.Width - U(26), U(18)),
                CanvasMuted, TextFormatFlags.EndEllipsis);
            if (data.Length == 0) return;
            var plot = PlotBounds(area);
            for (int k = 1; k <= 3; k++) g.DrawLine(border, plot.Left, plot.Top + plot.Height * k / 3, plot.Right, plot.Top + plot.Height * k / 3);
            double max = data.Max(), min = data.Min(); if (max == min) max = min + 1;
            float barW = plot.Width / (float)data.Length;
            string format = "F" + _decimalPlaces;
            float valueWidth = data.Max(v => g.MeasureString(v.ToString(format), valueFont).Width);
            bool showAll = barW >= valueWidth + U(3);
            float indexWidth = g.MeasureString(data.Length.ToString(), valueFont).Width + U(4);
            int step = Math.Max(1, (int)Math.Ceiling(indexWidth / barW));
            using var textBrush = new SolidBrush(CanvasText); using var dimBrush = new SolidBrush(CanvasMuted);
            for (int i = 0; i < data.Length; i++)
            {
                bool swap = i == snap.si || i == snap.sj, compare = i == snap.ci || i == snap.cj;
                Color color = swap ? SwapClr : compare ? CompareClr : st.BarColor;
                float height = Math.Max(U(2), (float)((data[i] - min + 0.5) / (max - min + 1)) * Math.Max(U(2), plot.Height - U(17)));
                float x = plot.Left + i * barW, y = plot.Bottom - height;
                var r = new RectangleF(x + 0.5F, y, Math.Max(1, barW - U(2)), height);
                using var bar = Rounded(r, Math.Min(U(3), r.Width / 2));
                using var brush = new LinearGradientBrush(new PointF(x, plot.Top), new PointF(x, plot.Bottom + 1), color, Color.FromArgb(125, color));
                g.FillPath(brush, bar);
                if (i == data.Length - 1 || (i % step == 0 && (data.Length - 1 - i) * barW >= indexWidth))
                {
                    string index = (i + 1).ToString(); var size = g.MeasureString(index, valueFont);
                    float ix = Math.Max(plot.Left, Math.Min(plot.Right - size.Width, x + (barW - size.Width) / 2));
                    g.DrawString(index, valueFont, dimBrush, ix, plot.Bottom + U(3));
                }
                if (showAll || ((swap || compare) && barW >= U(12)))
                {
                    string value = data[i].ToString(format); var size = g.MeasureString(value, valueFont);
                    float tx = Math.Max(plot.Left, Math.Min(plot.Right - size.Width, x + (barW - size.Width) / 2));
                    g.DrawString(value, valueFont, textBrush, tx, Math.Max(plot.Top, y - size.Height));
                }
            }
            long[] counts = { snap.cmp, snap.swp, snap.iter };
            string[] labels = { "сравнений", "обменов", "итераций" };
            for (int i = 0; i < 3; i++)
            {
                int x = area.Left + U(13) + i * (area.Width - U(26)) / 3, w = (area.Width - U(26)) / 3;
                TextRenderer.DrawText(g, counts[i].ToString("N0"), titleFont, new Rectangle(x, area.Bottom - U(38), w, U(20)),
                    CanvasText, TextFormatFlags.EndEllipsis);
                TextRenderer.DrawText(g, labels[i], infoFont, new Rectangle(x, area.Bottom - U(19), w, U(16)),
                    CanvasMuted, TextFormatFlags.EndEllipsis);
            }
        }

        private void ChartMouseMove(object? sender, MouseEventArgs e)
        {
            string text = "";
            foreach (var entry in _chartAreas)
            {
                var plot = PlotBounds(entry.area); if (!plot.Contains(e.Location)) continue;
                var data = entry.state.Snapshot().data; if (data.Length == 0) break;
                int index = Math.Min(data.Length - 1, (e.X - plot.Left) * data.Length / plot.Width);
                text = $"{entry.state.Name} · элемент №{index + 1}: {data[index].ToString("F" + _decimalPlaces)}"; break;
            }
            if (chartTip.GetToolTip(workspacePanel) != text) chartTip.SetToolTip(workspacePanel, text);
        }

        // ============================================================
        //  ДАННЫЕ
        // ============================================================
        private void SetupGrid()
        {
            dataGrid.Columns.Clear();
            dataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Value",
                HeaderText = "Значение",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                ReadOnly = false,
                Visible = true
            });
        }

        private void DataGrid_CellValidating(object? sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.ColumnIndex != 0) return;
            if (e.RowIndex < 0) return;

            var text = e.FormattedValue?.ToString();
            if (string.IsNullOrWhiteSpace(text)) return;

            if (!TryParseDouble(text, out _))
            {
                MessageBox.Show(
                    $"Некорректное значение \"{text}\" в строке {e.RowIndex + 1}. Ожидается число.",
                    "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
        }

        private void SyncDataFromGrid()
        {
            try
            {
                var list = new List<double>();
                foreach (DataGridViewRow row in dataGrid.Rows)
                {
                    if (row.IsNewRow) continue;
                    var cell = row.Cells[0];
                    var val = cell.Value;
                    if (val == null) continue;
                    var s = val.ToString();
                    if (string.IsNullOrWhiteSpace(s)) continue;

                    if (!TryParseDouble(s, out double num)) continue;
                    list.Add(num);
                }
                _currentData = list.ToArray();
                lblCount.Text = $"Элементов: {_currentData.Length:N0}";
                int numberWidth = Math.Max(U(42), TextRenderer.MeasureText(Math.Max(1, list.Count).ToString(), dataGrid.Font).Width + U(23));
                if (dataGrid.RowHeadersWidth != numberWidth) dataGrid.RowHeadersWidth = numberWidth;
                headerPanel.Invalidate();
            }
            catch { }
        }

        private void LoadDataToGrid(IEnumerable<double> data)
        {
            dataGrid.SuspendLayout();
            bool prevAllow = dataGrid.AllowUserToAddRows;
            dataGrid.AllowUserToAddRows = false;
            dataGrid.Rows.Clear();

            var list = data.ToList();
            for (int i = 0; i < list.Count; i++)
                dataGrid.Rows.Add();

            for (int i = 0; i < list.Count; i++)
            {
                double rounded = Math.Round(list[i], _decimalPlaces, MidpointRounding.AwayFromZero);
                dataGrid.Rows[i].Cells[0].Value = rounded;
            }

            dataGrid.AllowUserToAddRows = prevAllow;
            dataGrid.ResumeLayout();
            dataGrid.Refresh();
            dataGrid.Invalidate();

            SyncDataFromGrid();
        }

        private void BtnClear_Click(object? sender, EventArgs e)
        {
            if (_isRunning) return;
            dataGrid.CancelEdit();
            bool allowAdd = dataGrid.AllowUserToAddRows;
            dataGrid.AllowUserToAddRows = false;
            dataGrid.Rows.Clear();
            dataGrid.AllowUserToAddRows = allowAdd;
            _currentData = Array.Empty<double>();
            _states.Clear(); _chartAreas.Clear(); reportBox.Clear();
            _showReport = false; btnReport.Enabled = false;
            lblCount.Text = "Элементов: 0";
            UpdateWorkspace();
            lblStatus.Text = "Данные очищены.";
        }

        // ============================================================
        //  ФАБРИКА
        // ============================================================
        private static SortingAlgorithm CreateAlgoByName(string name)
        {
            switch (name)
            {
                case "Пузырьковая": return new BubbleSort();
                case "Вставками":   return new InsertionSort();
                case "Шейкерная":   return new ShakerSort();
                case "Быстрая":     return new QuickSort();
                case "BOGO":        return new BogoSort();
                default: throw new ArgumentException("Неизвестный алгоритм: " + name);
            }
        }

        // ============================================================
        //  ГЕНЕРАЦИЯ
        // ============================================================
        private void BtnGenerate_Click(object? sender, EventArgs e)
        {
            using var dlg = new GenerateForm();
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            var rnd = new Random();
            int count = (int)dlg.Count;
            var data = new double[count];
            for (int i = 0; i < count; i++)
            {
                double value = dlg.Min + rnd.NextDouble() * (dlg.Max - dlg.Min);
                data[i] = Math.Round(value, _decimalPlaces, MidpointRounding.AwayFromZero);
            }
            LoadDataToGrid(data);
        }

        // ============================================================
        //  EXCEL
        // ============================================================
        private void BtnExcel_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Filter = "Excel файлы (*.xlsx;*.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*",
                Title = "Выберите Excel-файл"
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                var data = LoadFromExcel(dlg.FileName);
                if (data.Count == 0)
                {
                    MessageBox.Show("Файл не содержит числовых данных.",
                        "Пустой файл", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                LoadDataToGrid(data);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки Excel: {ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<double> LoadFromExcel(string path)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var result = new List<double>();

            using var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            do
            {
                while (reader.Read())
                {
                    if (reader.FieldCount == 0) continue;
                    var val = reader.GetValue(0);
                    if (val == null) continue;
                    var s = val.ToString();
                    if (s == null) continue;

                    if (TryParseDouble(s, out double num))
                        result.Add(num);
                    else
                        throw new FormatException($"Значение \"{s}\" не является числом.");
                }
            } while (reader.NextResult());

            return result;
        }

        // ============================================================
        //  GOOGLE SHEETS
        // ============================================================
        private async void BtnGoogle_Click(object? sender, EventArgs e)
        {
            using var dlg = new GoogleLinkForm();
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                var data = await LoadFromGoogleSheetsByLink(dlg.SheetUrl, dlg.UseHtml);
                if (data.Count == 0)
                {
                    MessageBox.Show("Таблица не содержит числовых данных.",
                        "Пустая таблица", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                LoadDataToGrid(data);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки Google Sheets: {ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task<List<double>> LoadFromGoogleSheetsByLink(string url, bool useHtml)
        {
            if (string.IsNullOrWhiteSpace(url))
                throw new ArgumentException("Пустая ссылка.");

            string spreadsheetId = ExtractSpreadsheetId(url);
            string gid = ExtractGid(url) ?? "0";

            string csvUrl = useHtml
                ? $"https://docs.google.com/spreadsheets/d/{spreadsheetId}/gviz/tq?tqx=out:html&gid={gid}"
                : $"https://docs.google.com/spreadsheets/d/{spreadsheetId}/export?format=csv&gid={gid}";

            string content;
            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(30);
                content = await client.GetStringAsync(csvUrl);
            }

            var result = new List<double>();

            if (useHtml)
            {
                var matches = Regex.Matches(content, @"<td[^>]*>(.*?)</td>",
                    RegexOptions.Singleline | RegexOptions.IgnoreCase);
                foreach (Match m in matches)
                {
                    var text = WebUtility.HtmlDecode(m.Groups[1].Value).Trim();
                    if (TryParseDouble(text, out double num))
                        result.Add(num);
                }
            }
            else
            {
                var lines = content.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var first = line.Split(',')[0].Trim().Trim('"');
                    if (TryParseDouble(first, out double num))
                        result.Add(num);
                }
            }

            return result;
        }

        private static string ExtractSpreadsheetId(string url)
        {
            var m = Regex.Match(url, @"/spreadsheets/d/([a-zA-Z0-9-_]+)");
            if (m.Success) return m.Groups[1].Value;
            throw new FormatException(
                "Не удалось извлечь ID таблицы из ссылки. Проверьте, что ссылка вида " +
                "https://docs.google.com/spreadsheets/d/XXXXXXXXX/edit");
        }

        private static string? ExtractGid(string url)
        {
            var m = Regex.Match(url, @"[#&]gid=(\d+)");
            return m.Success ? m.Groups[1].Value : null;
        }

        // ============================================================
        //  СТОП
        // ============================================================
        private void BtnStop_Click(object? sender, EventArgs e)
        {
            _cts?.Cancel();
            lblStatus.Text = "Остановка...";
        }

        // ============================================================
        //  РАССЧИТАТЬ
        // ============================================================
        private async void BtnCalculate_Click(object? sender, EventArgs e)
        {
            if (_isRunning) return;

            SyncDataFromGrid();
            if (_currentData.Length == 0)
            {
                MessageBox.Show("Нет данных для сортировки.",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cbBogo.Checked && _currentData.Length > MaxBogoElements)
            {
                MessageBox.Show(
                    $"BOGO-сортировку нельзя запускать, если элементов больше {MaxBogoElements}.\n" +
                    $"Сейчас элементов: {_currentData.Length}.",
                    "Ограничение BOGO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedNames = new List<string>();
            if (cbBubble.Checked)    selectedNames.Add("Пузырьковая");
            if (cbInsertion.Checked) selectedNames.Add("Вставками");
            if (cbShaker.Checked)    selectedNames.Add("Шейкерная");
            if (cbQuick.Checked)     selectedNames.Add("Быстрая");
            if (cbBogo.Checked)      selectedNames.Add("BOGO");

            if (selectedNames.Count == 0)
            {
                MessageBox.Show("Выберите хотя бы один алгоритм сортировки.",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _isRunning = true;
            btnCalculate.Enabled = false;
            btnStop.Enabled = true;
            btnClear.Enabled = false;
            leftPanel.Enabled = optionsPanel.Enabled = false;
            tbDelay.Enabled = false;
            nudDecimals.Enabled = false;

            _cts?.Dispose();
            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            bool asc = cmbDirection.SelectedIndex == 0;
            bool visualize = _currentData.Length <= MaxVisualizeElements;
            int delayMs = tbDelay.Value;

            lblStatus.Text = visualize
                ? $"Замер ({RunsPerAlgorithm} прогонов, кроме BOGO)..."
                : $"Сортировка выполняется (визуализация отключена, элементов > {MaxVisualizeElements}).";

            // ---------- Инициализация состояний ----------
            _states = new List<AlgorithmState>();
            foreach (var name in selectedNames)
            {
                var st = new AlgorithmState
                {
                    Name = name,
                    BarColor = _algoColors.TryGetValue(name, out var c) ? c : Accent,
                    Data = (double[])_currentData.Clone(),
                    Visualize = visualize
                };
                _states.Add(st);
            }

            _showReport = false; reportBox.Clear(); btnReport.Enabled = false;
            UpdateWorkspace();

            // ============================================================
            //  ФАЗА 1: ЧИСТЫЙ ЗАМЕР
            // ============================================================
            var measureTasks = new List<Task>();

            foreach (var name in selectedNames)
            {
                var algoName = name;
                var st = _states.First(s => s.Name == algoName);
                bool isBogo = (algoName == "BOGO");
                int runs = isBogo ? 1 : RunsPerAlgorithm;

                measureTasks.Add(Task.Run(() =>
                {
                    var times = new List<double>(runs);
                    long firstCmp = 0, firstSwp = 0, firstIter = 0;
                    string? error = null;

                    for (int run = 0; run < runs; run++)
                    {
                        if (token.IsCancellationRequested) break;

                        var source = (double[])_currentData.Clone();
                        Array.Copy(source, st.Data, source.Length);

                        var algo = CreateAlgoByName(algoName);
                        algo.Ascending = asc;

                        bool collectThis = (run == 0);
                        long localCmp = 0, localSwp = 0;

                        if (collectThis && visualize)
                        {
                            algo.OnCompare += (i, j) =>
                            {
                                System.Threading.Interlocked.Increment(ref localCmp);
                                lock (st.LogSync) st.EventLog.Add((0, i, j));
                            };
                            algo.OnSwap += (i, j) =>
                            {
                                System.Threading.Interlocked.Increment(ref localSwp);
                                lock (st.LogSync) st.EventLog.Add((1, i, j));
                            };
                        }
                        else if (collectThis)
                        {
                            algo.OnCompare += (i, j) => System.Threading.Interlocked.Increment(ref localCmp);
                            algo.OnSwap    += (i, j) => System.Threading.Interlocked.Increment(ref localSwp);
                        }

                        var sw = System.Diagnostics.Stopwatch.StartNew();
                        try { algo.Sort(st.Data); }
                        catch (Exception ex)
                        {
                            error = ex.Message;
                            sw.Stop();
                            break;
                        }
                        sw.Stop();

                        times.Add(sw.Elapsed.TotalMilliseconds);

                        if (collectThis)
                        {
                            firstCmp = localCmp;
                            firstSwp = localSwp;
                            firstIter = algo.Iterations;
                        }

                        lock (st.Sync) st.Runs = run + 1;
                    }

                    lock (st.Sync)
                    {
                        if (times.Count > 0)
                        {
                            st.AvgMs = times.Average();
                            st.MinMs = times.Min();
                            st.MaxMs = times.Max();
                        }
                        st.Comparisons = firstCmp;
                        st.Swaps = firstSwp;
                        st.Iterations = firstIter;
                        st.Error = error;
                    }
                }, token));
            }

            try { await Task.WhenAll(measureTasks); }
            catch (OperationCanceledException) { }

            // ============================================================
            //  ФАЗА 2: ВОСПРОИЗВЕДЕНИЕ
            // ============================================================
            if (visualize && !token.IsCancellationRequested)
            {
                lblStatus.Text = "Воспроизведение визуализации...";

                var playTasks = new List<Task>();

                foreach (var name in selectedNames)
                {
                    var st = _states.First(s => s.Name == name);

                    playTasks.Add(Task.Run(() =>
                    {
                        var working = (double[])_currentData.Clone();
                        lock (st.Sync)
                        {
                            Array.Copy(working, st.Data, working.Length);
                            st.CompareI = st.CompareJ = st.SwapI = st.SwapJ = -1;
                        }

                        List<(int type, int i, int j)> log;
                        lock (st.LogSync) log = new List<(int, int, int)>(st.EventLog);

                        foreach (var ev in log)
                        {
                            if (token.IsCancellationRequested) break;

                            if (ev.type == 0)
                            {
                                lock (st.Sync)
                                {
                                    st.CompareI = ev.i;
                                    st.CompareJ = ev.j;
                                }
                            }
                            else
                            {
                                lock (st.Sync)
                                {
                                    if (ev.i >= 0 && ev.j >= 0 &&
                                        ev.i < st.Data.Length && ev.j < st.Data.Length)
                                    {
                                        (st.Data[ev.i], st.Data[ev.j]) = (st.Data[ev.j], st.Data[ev.i]);
                                    }
                                    st.SwapI = ev.i;
                                    st.SwapJ = ev.j;
                                }
                            }

                            if (delayMs > 0)
                            {
                                int slept = 0;
                                while (slept < delayMs && !token.IsCancellationRequested)
                                {
                                    int chunk = Math.Min(10, delayMs - slept);
                                    Thread.Sleep(chunk);
                                    slept += chunk;
                                }
                            }
                        }

                        lock (st.Sync)
                        {
                            st.CompareI = st.CompareJ = st.SwapI = st.SwapJ = -1;
                            st.Finished = true;
                        }
                    }, token));
                }

                try { await Task.WhenAll(playTasks); }
                catch (OperationCanceledException) { }
            }
            else
            {
                foreach (var st in _states)
                {
                    lock (st.Sync) st.Finished = true;
                }
            }

            // ============================================================
            //  ОТЧЁТ
            // ============================================================
            var ordered = _states.OrderBy(s => s.AvgMs).ToList();
            var sb = new StringBuilder();

            if (token.IsCancellationRequested)
                sb.AppendLine("=== Результаты (остановлено пользователем) ===");
            else
                sb.AppendLine("=== Результаты ===");

            sb.AppendLine(
                $"{"Алгоритм",-16}{"Среднее",-20}{"Мин.",-20}{"Макс.",-20}" +
                $"{"Прогонов",10}{"Итерации",12}{"Сравн.",14}{"Обмены",12}");
            sb.AppendLine(new string('-', 124));

            foreach (var s in ordered)
            {
                var snap = s.Snapshot();
                if (snap.err != null)
                {
                    sb.AppendLine(
                        $"{s.Name,-16}{"ошибка",-20}{"",-20}{"",-20}" +
                        $"{snap.runs,10}{snap.iter,12}{snap.cmp,14}{snap.swp,12}  {snap.err}");
                }
                else
                {
                    sb.AppendLine(
                        $"{s.Name,-16}{FormatTime(snap.avg),-20}{FormatTime(snap.min),-20}{FormatTime(snap.max),-20}" +
                        $"{snap.runs,10}{snap.iter,12}{snap.cmp,14}{snap.swp,12}");
                }
            }

            var fastest = ordered.FirstOrDefault(s => s.Error == null && s.Runs > 0);
            if (fastest != null)
                sb.AppendLine($"\nСамый быстрый (по среднему): {fastest.Name} ({FormatTime(fastest.AvgMs)})");

            if (!visualize)
                sb.AppendLine($"\nВизуализация отключена: элементов больше {MaxVisualizeElements}.");

            if (token.IsCancellationRequested)
                sb.AppendLine("\n*** Остановлено пользователем ***");

            reportBox.Text = sb.ToString();
            _showReport = !visualize; btnReport.Enabled = true;
            UpdateWorkspace();

            lblStatus.Text = token.IsCancellationRequested
                ? "Остановлено пользователем."
                : $"Готово. {_states.Count} алгоритм(ов). По {RunsPerAlgorithm} прогонов (BOGO — 1).";

            _isRunning = false;
            btnCalculate.Enabled = true;
            btnStop.Enabled = false;
            btnClear.Enabled = true;
            leftPanel.Enabled = optionsPanel.Enabled = true;
            headerPanel.Invalidate();
            tbDelay.Enabled = true;
            nudDecimals.Enabled = true;
            workspacePanel.Invalidate();
        }

        // ============================================================
        //  КЛАВИШИ
        // ============================================================
        private void Form1_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                if (_isRunning)
                {
                    _cts?.Cancel();
                    lblStatus.Text = "Остановка...";
                }
                else if (_isFullscreen) ExitFullscreen();
                else Close();
            }
            else if (e.KeyCode == Keys.F11)
            {
                if (_isFullscreen) ExitFullscreen();
                else EnterFullscreen();
            }
        }
    }

    // ============================================================
    //  ФОРМА: ГЕНЕРАЦИЯ
    // ============================================================
    public class GenerateForm : Form
    {
        private NumericUpDown nudCount = null!, nudMin = null!, nudMax = null!;
        public double Count => (double)nudCount.Value;
        public double Min => (double)nudMin.Value;
        public double Max => (double)nudMax.Value;

        public GenerateForm()
        {
            var body = Form1.CreateDialogBody(this, "Генерация данных", new Size(320, 208), 2);
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            body.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var fields = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 2, RowCount = 3, Margin = Padding.Empty };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            nudCount = new Form1.StudioNumber { Minimum = 1, Maximum = 100000, Value = 15 };
            nudMin = new Form1.StudioNumber { Minimum = -1000000, Maximum = 1000000, Value = 0,
                DecimalPlaces = 3, Increment = 0.1m };
            nudMax = new Form1.StudioNumber { Minimum = -1000000, Maximum = 1000000, Value = 100,
                DecimalPlaces = 3, Increment = 0.1m };
            var numbers = new[] { nudCount, nudMin, nudMax };
            string[] titles = { "Количество", "Минимум", "Максимум" };
            for (int i = 0; i < numbers.Length; i++)
            {
                fields.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var label = new Label { Text = titles[i], AutoSize = true, Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 5, 12, 5) };
                numbers[i].Width = 144; numbers[i].Anchor = AnchorStyles.Right;
                numbers[i].Margin = new Padding(0, 6, 0, 6);
                numbers[i].AccessibleName = titles[i]; numbers[i].TabIndex = i;
                fields.Controls.Add(label, 0, i); fields.Controls.Add(numbers[i], 1, i);
            }
            body.Controls.Add(fields, 0, 0);
            var btnOk = new Form1.StudioButton { Text = "Создать", DialogResult = DialogResult.OK, Primary = true };
            var btnCancel = new Form1.StudioButton { Text = "Отмена", DialogResult = DialogResult.Cancel };
            btnOk.Click += (s, e) =>
            {
                if (nudMin.Value >= nudMax.Value)
                {
                    MessageBox.Show(this, "Минимум должен быть меньше максимума.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                }
            };
            body.Controls.Add(Form1.CreateDialogActions(btnOk, btnCancel), 0, 1);
            AcceptButton = btnOk; CancelButton = btnCancel;
            Form1.ApplyDialogStyle(this);
            ResumeLayout(true);
        }
    }

    // ============================================================
    //  ФОРМА: GOOGLE SHEETS
    // ============================================================
    public class GoogleLinkForm : Form
    {
        private TextBox txtUrl = null!;
        private CheckBox chkHtml = null!;
        public string SheetUrl => txtUrl.Text.Trim();
        public bool UseHtml => chkHtml.Checked;

        public GoogleLinkForm()
        {
            var body = Form1.CreateDialogBody(this, "Импорт из Google Sheets", new Size(580, 250), 6);
            for (int i = 0; i < 6; i++)
                body.RowStyles.Add(new RowStyle(i == 4 ? SizeType.Percent : SizeType.AutoSize, i == 4 ? 100 : 0));
            var intro = new Label { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10),
                Text = "Вставьте ссылку на Google-таблицу.\nТаблица должна быть опубликована или доступна по ссылке." };
            txtUrl = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 10),
                Text = "https://docs.google.com/spreadsheets/d/", Font = new Font("Consolas", 9F),
                BorderStyle = BorderStyle.FixedSingle, AccessibleName = "Ссылка на Google-таблицу", TabIndex = 0 };
            chkHtml = new CheckBox { AutoSize = true, Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat,
                Text = "Использовать HTML-формат (если CSV не работает)", Margin = new Padding(0, 6, 0, 10), TabIndex = 1 };
            var hint = new Label { AutoSize = true, Dock = DockStyle.Fill, Tag = "muted", Margin = Padding.Empty,
                Font = new Font("Segoe UI", 8.5F), Text = "Пример: https://docs.google.com/spreadsheets/d/1AbCdEf123.../edit#gid=0" };
            body.Controls.Add(intro, 0, 0); body.Controls.Add(txtUrl, 0, 1);
            body.Controls.Add(chkHtml, 0, 2); body.Controls.Add(hint, 0, 3);
            var btnOk = new Form1.StudioButton { Text = "Загрузить", DialogResult = DialogResult.OK, Primary = true };
            var btnCancel = new Form1.StudioButton { Text = "Отмена", DialogResult = DialogResult.Cancel };
            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(SheetUrl))
                {
                    MessageBox.Show(this, "Вставьте ссылку на таблицу.", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }
                var m = Regex.Match(SheetUrl, @"/spreadsheets/d/([a-zA-Z0-9-_]+)");
                if (!m.Success)
                {
                    MessageBox.Show(this,
                        "Не удалось найти ID таблицы в ссылке.\n\n" +
                        "Убедитесь, что ссылка вида:\nhttps://docs.google.com/spreadsheets/d/XXXXXXXXX/edit",
                        "Неверная ссылка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                }
            };
            body.Controls.Add(Form1.CreateDialogActions(btnOk, btnCancel), 0, 5);
            AcceptButton = btnOk; CancelButton = btnCancel;
            Form1.ApplyDialogStyle(this);
            ResumeLayout(true);
        }
    }
}
