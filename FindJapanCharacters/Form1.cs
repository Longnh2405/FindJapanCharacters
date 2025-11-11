using ClosedXML.Excel;
using System.Drawing.Drawing2D;
using System.Text;
using System.Text.RegularExpressions;

namespace FindJapanCharacters
{
    public partial class Form1 : Form
    {
        private readonly Random _rand = new Random(); // thêm dòng này
        public Form1()
        {
            InitializeComponent();
            this.Paint += Form1_Paint;
        }

        private static readonly Regex JapaneseRegex = new Regex(
            @"[\u3040-\u309F\u30A0-\u30FF\u31F0-\u31FF\uFF65-\uFF9F\u3400-\u4DBF\u4E00-\u9FFF]",
            RegexOptions.Compiled);

        // Tiếng Việt có dấu
        private static readonly Regex VietnameseRegex = new Regex(
            @"[ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚĂĐĨŨƠƯàáâãèéêìíòóôõùúăđĩũơư"
          + @"Ạ-ỹ]",
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
            // Tìm ký tự tiếng Nhật
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

        // ĐỌC KỂ CẢ KHI EXCEL ĐANG MỞ FILE – tìm ký tự tiếng Nhật
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
                        sb.AppendLine($"{hit,4}. Sheet={ws.Name} | Ô={cell.Address}");
                    }
                }
            }

            if (hit == 0) return "Không tìm thấy ô nào có ký tự tiếng Nhật.";
            return $"Tìm thấy {hit} ô có ký tự tiếng Nhật:\r\n" + sb.ToString();
        }

        // ĐỌC – tìm ký tự KHÁC tiếng Nhật (có ký tự non-ASCII nhưng không thuộc vùng Nhật)
        private string ScanExcelNonJapanese(string filePath)
        {
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
                    if (ContainsNonJapanese(text))
                    {
                        hit++;
                        sb.AppendLine($"{hit,4}. Sheet={ws.Name} | Ô={cell.Address}");
                    }
                }
            }

            if (hit == 0) return "Không tìm thấy ô nào có ký tự khác tiếng Nhật.";
            return $"Tìm thấy {hit} ô có ký tự khác tiếng Nhật:\r\n" + sb.ToString();
        }

        // ĐỌC – tìm ký tự tiếng Việt có dấu
        private string ScanExcelVietnamese(string filePath)
        {
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
                    if (ContainsVietnamese(text))
                    {
                        hit++;
                        sb.AppendLine($"{hit,4}. Sheet={ws.Name} | Ô={cell.Address}");
                    }
                }
            }

            if (hit == 0) return "Không tìm thấy ô nào có ký tự tiếng Việt có dấu.";
            return $"Tìm thấy {hit} ô có ký tự tiếng Việt có dấu:\r\n" + sb.ToString();
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

        // Ký tự KHÁC tiếng Nhật: có ký tự non-ASCII nhưng không match JapaneseRegex
        private static bool ContainsNonJapanese(string? s)
        {
            if (string.IsNullOrEmpty(s)) return false;

            foreach (var ch in s)
            {
                // Bỏ qua số (vì tiếng Nhật cũng dùng 0-9)
                if (char.IsDigit(ch))
                    continue;

                // Nếu không phải ký tự tiếng Nhật thì trả về true
                if (!JapaneseRegex.IsMatch(ch.ToString()))
                    return true;
            }

            return false;
        }



        private static bool ContainsVietnamese(string? s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            return VietnameseRegex.IsMatch(s);
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

            // ======= QUANH NÚT "Ting ting" (giữ nguyên nếu có button2) =======
            if (button2 != null)
            {
                var r = button2.Bounds;

                DrawHeart(g, new RectangleF(r.Left - 55, r.Top - 25, 28, 28), Color.Pink, Color.Red);
                DrawHeart(g, new RectangleF(r.Right + 15, r.Top - 20, 26, 26), Color.MistyRose, Color.DeepPink);
                DrawHeart(g, new RectangleF(r.Left - 40, r.Bottom + 5, 22, 22), Color.MistyRose, Color.HotPink);
                DrawHeart(g, new RectangleF(r.Right + 10, r.Bottom + 5, 22, 22), Color.Pink, Color.HotPink);
            }

            // ======= THÊM 1 VÀI CÂY NẤM =======
            DrawMushroom(g, new RectangleF(60, this.ClientSize.Height - 160, 80, 110),
                         Color.LightPink, Color.WhiteSmoke, Color.SaddleBrown);

            DrawMushroom(g, new RectangleF(this.ClientSize.Width - 160, this.ClientSize.Height - 160, 80, 110),
                         Color.LightSalmon, Color.MistyRose, Color.Sienna);

            DrawMushroom(g, new RectangleF(40, 40, 70, 100),
                         Color.PaleVioletRed, Color.White, Color.Peru);

            // ======= RẢI NHIỀU NGÔI SAO RANDOM KHẮP MÀN HÌNH =======
            int starCount = 120; // tăng số lượng ngôi sao
            for (int i = 0; i < starCount; i++)
            {
                float x = _rand.Next(10, Math.Max(20, this.ClientSize.Width - 10));
                float y = _rand.Next(10, Math.Max(20, this.ClientSize.Height - 10));
                float size = _rand.Next(8, 26); // kích thước sao ngẫu nhiên

                Color fill = Color.FromArgb(
                    _rand.Next(220, 255),
                    _rand.Next(220, 255),
                    _rand.Next(150, 255)
                );

                Color border = Color.FromArgb(
                    _rand.Next(180, 255),
                    _rand.Next(180, 255),
                    _rand.Next(180, 255)
                );

                DrawStar(g, new PointF(x, y), size, fill, border);
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

        private void DrawStar(Graphics g, PointF center, float radius, Color fillColor, Color borderColor)
        {
            const int points = 5;
            PointF[] starPoints = new PointF[points * 2];
            double angle = -Math.PI / 2;
            double step = Math.PI / points;

            for (int i = 0; i < points * 2; i++)
            {
                double r = (i % 2 == 0) ? radius : radius * 0.45f;
                starPoints[i] = new PointF(
                    center.X + (float)(Math.Cos(angle) * r),
                    center.Y + (float)(Math.Sin(angle) * r)
                );
                angle += step;
            }

            using (GraphicsPath path = new GraphicsPath())
            {
                path.AddPolygon(starPoints);
                using (var brush = new SolidBrush(fillColor))
                using (var pen = new Pen(borderColor, 1.5f))
                {
                    g.FillPath(brush, path);
                    g.DrawPath(pen, path);
                }
            }
        }

        private void DrawMushroom(Graphics g, RectangleF rect, Color capColor, Color dotColor, Color stemColor)
        {
            float x = rect.X;
            float y = rect.Y;
            float w = rect.Width;
            float h = rect.Height;

            // Mũ nấm
            RectangleF cap = new RectangleF(x, y, w, h * 0.6f);
            using (GraphicsPath capPath = new GraphicsPath())
            {
                capPath.AddArc(cap.X, cap.Y + cap.Height / 2, cap.Width, cap.Height, 180, 180);
                capPath.AddLine(cap.X, cap.Y + cap.Height, cap.X + cap.Width, cap.Y + cap.Height);
                capPath.CloseFigure();

                using (var brush = new SolidBrush(capColor))
                using (var pen = new Pen(Color.Firebrick, 2f))
                {
                    g.FillPath(brush, capPath);
                    g.DrawPath(pen, capPath);
                }

                // Chấm trắng trên mũ
                using (var dotBrush = new SolidBrush(dotColor))
                {
                    float dotSize = w * 0.12f;
                    g.FillEllipse(dotBrush, x + w * 0.25f, y + cap.Height * 0.3f, dotSize, dotSize);
                    g.FillEllipse(dotBrush, x + w * 0.55f, y + cap.Height * 0.2f, dotSize * 1.1f, dotSize * 1.1f);
                    g.FillEllipse(dotBrush, x + w * 0.40f, y + cap.Height * 0.55f, dotSize * 0.9f, dotSize * 0.9f);
                }
            }

            // Thân nấm
            float stemWidth = w * 0.35f;
            float stemHeight = h * 0.45f;
            RectangleF stem = new RectangleF(
                x + (w - stemWidth) / 2,
                y + h * 0.55f,
                stemWidth,
                stemHeight
            );

            using (GraphicsPath stemPath = new GraphicsPath())
            {
                stemPath.AddRectangle(stem);
                using (var brush = new SolidBrush(stemColor))
                using (var pen = new Pen(Color.SaddleBrown, 1.5f))
                {
                    g.FillPath(brush, stemPath);
                    g.DrawPath(pen, stemPath);
                }
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            // Tìm ký tự khác tiếng Nhật
            var path = string.IsNullOrWhiteSpace(_excelPath) ? richTextBox1.Text.Trim() : _excelPath;

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show("Chưa có đường dẫn file hợp lệ. Hãy chọn file trước.",
                                "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var report = ScanExcelNonJapanese(path);
                richTextBox1.AppendText("\r\n" + report);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi quét Excel: " + ex.Message, "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Tìm ký tự tiếng Việt
            var path = string.IsNullOrWhiteSpace(_excelPath) ? richTextBox1.Text.Trim() : _excelPath;

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show("Chưa có đường dẫn file hợp lệ. Hãy chọn file trước.",
                                "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var report = ScanExcelVietnamese(path);
                richTextBox1.AppendText("\r\n" + report);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi quét Excel: " + ex.Message, "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
