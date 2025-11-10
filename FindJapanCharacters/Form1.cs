using ClosedXML.Excel;
using System.Drawing.Drawing2D;
using System.Text;
using System.Text.RegularExpressions;

namespace FindJapanCharacters
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Paint += Form1_Paint;
        }
        private static readonly Regex JapaneseRegex = new Regex(
    @"[\u3040-\u309F\u30A0-\u30FF\u31F0-\u31FF\uFF65-\uFF9F\u3400-\u4DBF\u4E00-\u9FFF]",
    RegexOptions.Compiled);

        string _lastFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); // mặc định lần đầu
        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Chọn file Excel";
                openFileDialog.Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*";
                openFileDialog.InitialDirectory = _lastFolder; // 💡 mở lại folder trước đó

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _excelPath = openFileDialog.FileName;

                    // ✅ Lưu lại folder hiện tại để lần sau mở lại ở đây
                    _lastFolder = Path.GetDirectoryName(_excelPath)!;

                    richTextBox1.Clear();
                    richTextBox1.AppendText("Đã chọn: " + _excelPath + Environment.NewLine);
                }
            }
        }


        private void ProcessFile(string filePath)
        {
            richTextBox1.Text = filePath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Lấy đường dẫn từ biến bạn lưu, hoặc từ richTextBox nếu bạn đang hiển thị ở đó
            var path = string.IsNullOrWhiteSpace(_excelPath) ? richTextBox1.Text.Trim() : _excelPath;

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show("Chưa có đường dẫn file hợp lệ. Hãy chọn file trước.",
                                "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var report = ScanExcelEvenIfLocked(path);
                richTextBox1.AppendText("\r\n" + report);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi quét Excel: " + ex.Message, "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ĐỌC KỂ CẢ KHI EXCEL ĐANG MỞ FILE
        private string ScanExcelEvenIfLocked(string filePath)
        {
            // Mở với FileShare.ReadWrite để không bị Excel chặn
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var ms = new MemoryStream();
            fs.CopyTo(ms);
            ms.Position = 0;

            int hit = 0;
            var sb = new StringBuilder();

            using var wb = new XLWorkbook(ms);
            foreach (var ws in wb.Worksheets)
            {
                foreach (var cell in ws.CellsUsed())
                {
                    string text = cell.Value.ToString() ?? string.Empty;
                    if (ContainsJapanese(text))
                    {
                        hit++;
                        sb.AppendLine($"{hit,4}. Sheet={ws.Name} | Ô={cell.Address} | \"{Truncate(text, 120)}\"");
                    }
                }
            }

            if (hit == 0) return "Không tìm thấy ô nào có ký tự tiếng Nhật.";
            return $"Tìm thấy {hit} ô có ký tự tiếng Nhật:\r\n" + sb.ToString();
        }

        private static string Truncate(string input, int max)
        {
            if (string.IsNullOrEmpty(input)) return input;
            return input.Length <= max ? input : input.Substring(0, max) + "...";
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }




        private static bool ContainsJapanese(string? s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            return JapaneseRegex.IsMatch(s);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        // VẼ TRÁI TIM BÊN TRÁI FORM
        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

            // ======= CỤM TRÁI TIM GÓC TRÊN TRÁI (DỊCH SANG PHẢI) =======
            float offsetX = 25f; // chỉnh con số này để dịch sang phải nhiều hay ít
            DrawHeart(g, new RectangleF(40 + offsetX, 90, 70, 70), Color.MistyRose, Color.Red);
            DrawHeart(g, new RectangleF(130 + offsetX, 100, 50, 50), Color.Pink, Color.DeepPink);
            DrawHeart(g, new RectangleF(90 + offsetX, 160, 40, 40), Color.MistyRose, Color.HotPink);

            // ======= CỤM TRÁI TIM GÓC DƯỚI (THÊM MỚI, CŨNG DỊCH SANG PHẢI) =======
            float bottomY = this.ClientSize.Height - 180;
            DrawHeart(g, new RectangleF(40 + offsetX, bottomY, 70, 70), Color.MistyRose, Color.Red);
            DrawHeart(g, new RectangleF(120 + offsetX, bottomY + 20, 50, 50), Color.Pink, Color.DeepPink);
            DrawHeart(g, new RectangleF(80 + offsetX, bottomY + 60, 40, 40), Color.MistyRose, Color.HotPink);


            // ======= VIỀN BÊN TRÁI =======
            for (int y = 60; y <= this.ClientSize.Height - 80; y += 70)
            {
                DrawHeart(g, new RectangleF(10, y, 28, 28), Color.MistyRose, Color.HotPink);
            }

            // ======= VIỀN PHÍA TRÊN =======
            for (int x = 260; x <= this.ClientSize.Width - 80; x += 90)
            {
                DrawHeart(g, new RectangleF(x, 25, 24, 24), Color.MistyRose, Color.DeepPink);
            }

            // ======= VIỀN BÊN PHẢI =======
            for (int y = 80; y <= this.ClientSize.Height - 100; y += 80)
            {
                DrawHeart(g, new RectangleF(this.ClientSize.Width - 60, y, 26, 26), Color.Pink, Color.HotPink);
            }

            // ======= VIỀN DƯỚI =======
            for (int x = 260; x <= this.ClientSize.Width - 80; x += 90)
            {
                DrawHeart(g, new RectangleF(x, this.ClientSize.Height - 80, 24, 24), Color.MistyRose, Color.DeepPink);
            }

            // ======= QUANH NÚT "Ting ting" =======
            if (button2 != null)
            {
                var r = button2.Bounds;

                DrawHeart(g, new RectangleF(r.Left - 55, r.Top - 25, 28, 28), Color.Pink, Color.Red);
                DrawHeart(g, new RectangleF(r.Right + 15, r.Top - 20, 26, 26), Color.MistyRose, Color.DeepPink);
                DrawHeart(g, new RectangleF(r.Left - 40, r.Bottom + 5, 22, 22), Color.MistyRose, Color.HotPink);
                DrawHeart(g, new RectangleF(r.Right + 10, r.Bottom + 5, 22, 22), Color.Pink, Color.HotPink);
            }
        }

        private void DrawHeart(Graphics g, RectangleF rect, Color fillColor, Color borderColor)
        {
            float x = rect.X;
            float y = rect.Y;
            float w = rect.Width;
            float h = rect.Height;

            PointF topCenter = new PointF(x + w / 2f, y + h * 0.25f);
            PointF bottom = new PointF(x + w / 2f, y + h);

            using (var path = new System.Drawing.Drawing2D.GraphicsPath())
            {
                path.StartFigure();

                // Nửa trái
                path.AddBezier(
                    topCenter,
                    new PointF(x + w * 0.05f, y + h * 0.05f),
                    new PointF(x - w * 0.10f, y + h * 0.60f),
                    bottom
                );

                // Nửa phải
                path.AddBezier(
                    bottom,
                    new PointF(x + w + w * 0.10f, y + h * 0.60f),
                    new PointF(x + w * 0.95f, y + h * 0.05f),
                    topCenter
                );

                path.CloseFigure();

                using (var brush = new SolidBrush(fillColor))
                using (var pen = new Pen(borderColor, 2f))
                {
                    g.FillPath(brush, path);
                    g.DrawPath(pen, path);
                }
            }
        }
    }
}
