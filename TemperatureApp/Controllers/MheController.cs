using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using TemperatureApp.Services;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.IO;
using TemperatureApp.Data;

namespace TemperatureApp.Controllers
{
    public class MheController : Controller
    {
        private readonly ILogger<MheController> _logger;
        private readonly DatabaseService _databaseService;
        private readonly ApplicationDbContext _context;

        public MheController(ILogger<MheController> logger, DatabaseService databaseService, ApplicationDbContext context)
        {
            _logger = logger;
            _databaseService = databaseService;
            _context = context;
        }

        // Điểm dữ liệu chung cho tất cả các máy độ nhớt
        public record VisPoint(DateTime Time, int PPM, int ThucTe, int TieuChuan);

        private const string DbErr = "Không thể kết nối đến cơ sở dữ liệu.";

        // ---------------- Máy Độ Nhớt 1.1 ----------------
        public async Task<IActionResult> May1_1(DateTime? startDate, DateTime? endDate)
        {
            if (!await _databaseService.CanConnectAsync()) { TempData["Error"] = DbErr; return View("Machine"); }
            var end = endDate?.AddSeconds(59).AddMilliseconds(999);
            var raw = startDate.HasValue && endDate.HasValue
                ? await _context.VIS_MHE1_1.Where(x => x.VISMHE_1_Viscosity_1_1_TIME >= startDate && x.VISMHE_1_Viscosity_1_1_TIME <= end).OrderByDescending(x => x.VISMHE_1_Viscosity_1_1_TIME).ToListAsync()
                : await _context.VIS_MHE1_1.OrderByDescending(x => x.VISMHE_1_Viscosity_1_1_TIME).Take(50).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_1_Viscosity_1_1_TIME, x.VISMHE_1_Viscosity_1_1_PPM, x.VISMHE_1_Viscosity_1_1_Thuc_te, x.VISMHE_1_Viscosity_1_1_Tieu_chuan)).ToList();
            return RenderMachine("Máy Độ Nhớt 1.1", "May1_1", "Export1_1", data, startDate, endDate);
        }

        public async Task<IActionResult> Export1_1(DateTime? startDate, DateTime? endDate)
        {
            if (!startDate.HasValue || !endDate.HasValue || startDate > endDate) return BadRequest("Ngày không hợp lệ.");
            var end = endDate.Value.AddSeconds(59).AddMilliseconds(999);
            var raw = await _context.VIS_MHE1_1.Where(x => x.VISMHE_1_Viscosity_1_1_TIME >= startDate && x.VISMHE_1_Viscosity_1_1_TIME <= end).OrderBy(x => x.VISMHE_1_Viscosity_1_1_TIME).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_1_Viscosity_1_1_TIME, x.VISMHE_1_Viscosity_1_1_PPM, x.VISMHE_1_Viscosity_1_1_Thuc_te, x.VISMHE_1_Viscosity_1_1_Tieu_chuan)).ToList();
            return BuildExcel(data, "MHE 1.1", "MHE_1_1");
        }

        // ---------------- Máy Độ Nhớt 1.2 ----------------
        public async Task<IActionResult> May1_2(DateTime? startDate, DateTime? endDate)
        {
            if (!await _databaseService.CanConnectAsync()) { TempData["Error"] = DbErr; return View("Machine"); }
            var end = endDate?.AddSeconds(59).AddMilliseconds(999);
            var raw = startDate.HasValue && endDate.HasValue
                ? await _context.VIS_MHE1_2.Where(x => x.VISMHE_1_Viscosity_1_2_TIME >= startDate && x.VISMHE_1_Viscosity_1_2_TIME <= end).OrderByDescending(x => x.VISMHE_1_Viscosity_1_2_TIME).ToListAsync()
                : await _context.VIS_MHE1_2.OrderByDescending(x => x.VISMHE_1_Viscosity_1_2_TIME).Take(50).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_1_Viscosity_1_2_TIME, x.VISMHE_1_Viscosity_1_2_PPM, x.VISMHE_1_Viscosity_1_2_Thuc_te, x.VISMHE_1_Viscosity_1_2_Tieu_chuan)).ToList();
            return RenderMachine("Máy Độ Nhớt 1.2", "May1_2", "Export1_2", data, startDate, endDate);
        }

        public async Task<IActionResult> Export1_2(DateTime? startDate, DateTime? endDate)
        {
            if (!startDate.HasValue || !endDate.HasValue || startDate > endDate) return BadRequest("Ngày không hợp lệ.");
            var end = endDate.Value.AddSeconds(59).AddMilliseconds(999);
            var raw = await _context.VIS_MHE1_2.Where(x => x.VISMHE_1_Viscosity_1_2_TIME >= startDate && x.VISMHE_1_Viscosity_1_2_TIME <= end).OrderBy(x => x.VISMHE_1_Viscosity_1_2_TIME).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_1_Viscosity_1_2_TIME, x.VISMHE_1_Viscosity_1_2_PPM, x.VISMHE_1_Viscosity_1_2_Thuc_te, x.VISMHE_1_Viscosity_1_2_Tieu_chuan)).ToList();
            return BuildExcel(data, "MHE 1.2", "MHE_1_2");
        }

        // ---------------- Máy Độ Nhớt 2.1 ----------------
        // Lưu ý: bảng VIS#MHE2_1 trong DB có cột đặt tên theo mẫu 1_2 (giữ nguyên theo DB)
        public async Task<IActionResult> May2_1(DateTime? startDate, DateTime? endDate)
        {
            if (!await _databaseService.CanConnectAsync()) { TempData["Error"] = DbErr; return View("Machine"); }
            var end = endDate?.AddSeconds(59).AddMilliseconds(999);
            var raw = startDate.HasValue && endDate.HasValue
                ? await _context.VIS_MHE2_1.Where(x => x.VISMHE_1_Viscosity_1_2_TIME >= startDate && x.VISMHE_1_Viscosity_1_2_TIME <= end).OrderByDescending(x => x.VISMHE_1_Viscosity_1_2_TIME).ToListAsync()
                : await _context.VIS_MHE2_1.OrderByDescending(x => x.VISMHE_1_Viscosity_1_2_TIME).Take(50).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_1_Viscosity_1_2_TIME, x.VISMHE_1_Viscosity_1_2_PPM, x.VISMHE_1_Viscosity_1_2_Thuc_te, x.VISMHE_1_Viscosity_1_2_Tieu_chuan)).ToList();
            return RenderMachine("Máy Độ Nhớt 2.1", "May2_1", "Export2_1", data, startDate, endDate);
        }

        public async Task<IActionResult> Export2_1(DateTime? startDate, DateTime? endDate)
        {
            if (!startDate.HasValue || !endDate.HasValue || startDate > endDate) return BadRequest("Ngày không hợp lệ.");
            var end = endDate.Value.AddSeconds(59).AddMilliseconds(999);
            var raw = await _context.VIS_MHE2_1.Where(x => x.VISMHE_1_Viscosity_1_2_TIME >= startDate && x.VISMHE_1_Viscosity_1_2_TIME <= end).OrderBy(x => x.VISMHE_1_Viscosity_1_2_TIME).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_1_Viscosity_1_2_TIME, x.VISMHE_1_Viscosity_1_2_PPM, x.VISMHE_1_Viscosity_1_2_Thuc_te, x.VISMHE_1_Viscosity_1_2_Tieu_chuan)).ToList();
            return BuildExcel(data, "MHE 2.1", "MHE_2_1");
        }

        // ---------------- Máy Độ Nhớt 2.2 ----------------
        public async Task<IActionResult> May2_2(DateTime? startDate, DateTime? endDate)
        {
            if (!await _databaseService.CanConnectAsync()) { TempData["Error"] = DbErr; return View("Machine"); }
            var end = endDate?.AddSeconds(59).AddMilliseconds(999);
            var raw = startDate.HasValue && endDate.HasValue
                ? await _context.VIS_MHE2_2.Where(x => x.VISMHE_2_Viscosity_2_2_TIME >= startDate && x.VISMHE_2_Viscosity_2_2_TIME <= end).OrderByDescending(x => x.VISMHE_2_Viscosity_2_2_TIME).ToListAsync()
                : await _context.VIS_MHE2_2.OrderByDescending(x => x.VISMHE_2_Viscosity_2_2_TIME).Take(50).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_2_Viscosity_2_2_TIME, x.VISMHE_2_Viscosity_2_2_PPM, x.VISMHE_2_Viscosity_2_2_Thuc_te, x.VISMHE_2_Viscosity_2_2_Tieu_chuan)).ToList();
            return RenderMachine("Máy Độ Nhớt 2.2", "May2_2", "Export2_2", data, startDate, endDate);
        }

        public async Task<IActionResult> Export2_2(DateTime? startDate, DateTime? endDate)
        {
            if (!startDate.HasValue || !endDate.HasValue || startDate > endDate) return BadRequest("Ngày không hợp lệ.");
            var end = endDate.Value.AddSeconds(59).AddMilliseconds(999);
            var raw = await _context.VIS_MHE2_2.Where(x => x.VISMHE_2_Viscosity_2_2_TIME >= startDate && x.VISMHE_2_Viscosity_2_2_TIME <= end).OrderBy(x => x.VISMHE_2_Viscosity_2_2_TIME).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_2_Viscosity_2_2_TIME, x.VISMHE_2_Viscosity_2_2_PPM, x.VISMHE_2_Viscosity_2_2_Thuc_te, x.VISMHE_2_Viscosity_2_2_Tieu_chuan)).ToList();
            return BuildExcel(data, "MHE 2.2", "MHE_2_2");
        }

        // ---------------- Máy Độ Nhớt 3.1 ----------------
        public async Task<IActionResult> May3_1(DateTime? startDate, DateTime? endDate)
        {
            if (!await _databaseService.CanConnectAsync()) { TempData["Error"] = DbErr; return View("Machine"); }
            var end = endDate?.AddSeconds(59).AddMilliseconds(999);
            var raw = startDate.HasValue && endDate.HasValue
                ? await _context.VIS_MHE3_1.Where(x => x.VISMHE_3_Viscosity_3_1_TIME >= startDate && x.VISMHE_3_Viscosity_3_1_TIME <= end).OrderByDescending(x => x.VISMHE_3_Viscosity_3_1_TIME).ToListAsync()
                : await _context.VIS_MHE3_1.OrderByDescending(x => x.VISMHE_3_Viscosity_3_1_TIME).Take(50).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_3_Viscosity_3_1_TIME, x.VISMHE_3_Viscosity_3_1_PPM, x.VISMHE_3_Viscosity_3_1_Thuc_te, x.VISMHE_3_Viscosity_3_1_Tieu_chuan)).ToList();
            return RenderMachine("Máy Độ Nhớt 3.1", "May3_1", "Export3_1", data, startDate, endDate);
        }

        public async Task<IActionResult> Export3_1(DateTime? startDate, DateTime? endDate)
        {
            if (!startDate.HasValue || !endDate.HasValue || startDate > endDate) return BadRequest("Ngày không hợp lệ.");
            var end = endDate.Value.AddSeconds(59).AddMilliseconds(999);
            var raw = await _context.VIS_MHE3_1.Where(x => x.VISMHE_3_Viscosity_3_1_TIME >= startDate && x.VISMHE_3_Viscosity_3_1_TIME <= end).OrderBy(x => x.VISMHE_3_Viscosity_3_1_TIME).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_3_Viscosity_3_1_TIME, x.VISMHE_3_Viscosity_3_1_PPM, x.VISMHE_3_Viscosity_3_1_Thuc_te, x.VISMHE_3_Viscosity_3_1_Tieu_chuan)).ToList();
            return BuildExcel(data, "MHE 3.1", "MHE_3_1");
        }

        // ---------------- Máy Độ Nhớt 3.2 ----------------
        public async Task<IActionResult> May3_2(DateTime? startDate, DateTime? endDate)
        {
            if (!await _databaseService.CanConnectAsync()) { TempData["Error"] = DbErr; return View("Machine"); }
            var end = endDate?.AddSeconds(59).AddMilliseconds(999);
            var raw = startDate.HasValue && endDate.HasValue
                ? await _context.VIS_MHE3_2.Where(x => x.VISMHE_3_Viscosity_3_2_TIME >= startDate && x.VISMHE_3_Viscosity_3_2_TIME <= end).OrderByDescending(x => x.VISMHE_3_Viscosity_3_2_TIME).ToListAsync()
                : await _context.VIS_MHE3_2.OrderByDescending(x => x.VISMHE_3_Viscosity_3_2_TIME).Take(50).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_3_Viscosity_3_2_TIME, x.VISMHE_3_Viscosity_3_2_PPM, x.VISMHE_3_Viscosity_3_2_Thuc_te, x.VISMHE_3_Viscosity_3_2_Tieu_chuan)).ToList();
            return RenderMachine("Máy Độ Nhớt 3.2", "May3_2", "Export3_2", data, startDate, endDate);
        }

        public async Task<IActionResult> Export3_2(DateTime? startDate, DateTime? endDate)
        {
            if (!startDate.HasValue || !endDate.HasValue || startDate > endDate) return BadRequest("Ngày không hợp lệ.");
            var end = endDate.Value.AddSeconds(59).AddMilliseconds(999);
            var raw = await _context.VIS_MHE3_2.Where(x => x.VISMHE_3_Viscosity_3_2_TIME >= startDate && x.VISMHE_3_Viscosity_3_2_TIME <= end).OrderBy(x => x.VISMHE_3_Viscosity_3_2_TIME).ToListAsync();
            var data = raw.Select(x => new VisPoint(x.VISMHE_3_Viscosity_3_2_TIME, x.VISMHE_3_Viscosity_3_2_PPM, x.VISMHE_3_Viscosity_3_2_Thuc_te, x.VISMHE_3_Viscosity_3_2_Tieu_chuan)).ToList();
            return BuildExcel(data, "MHE 3.2", "MHE_3_2");
        }

        // ---------------- Helpers dùng chung ----------------
        private IActionResult RenderMachine(string title, string action, string exportAction, List<VisPoint> data, DateTime? startDate, DateTime? endDate)
        {
            data.Reverse(); // đổi từ giảm dần -> tăng dần theo thời gian

            ViewData["Title"] = title;
            ViewData["Action"] = action;
            ViewData["ExportAction"] = exportAction;
            ViewData["Labels"] = data.Select(x => x.Time.ToString("yyyy-MM-dd HH:mm:ss")).ToList();
            ViewData["ThucTe"] = data.Select(x => x.ThucTe).ToList();
            ViewData["TieuChuan"] = data.Select(x => x.TieuChuan).ToList();
            ViewData["PPM"] = data.Select(x => x.PPM).ToList();
            ViewData["StartDate"] = startDate?.ToString("yyyy-MM-ddTHH:mm");
            ViewData["EndDate"] = endDate?.ToString("yyyy-MM-ddTHH:mm");

            return View("Machine");
        }

        private IActionResult BuildExcel(List<VisPoint> data, string sheetName, string fileName)
        {
            if (!data.Any()) return BadRequest("Không có dữ liệu.");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add(sheetName);

            ws.Cells[1, 1].Value = "Thời gian";
            ws.Cells[1, 2].Value = "Áp lực";
            ws.Cells[1, 3].Value = "Thực tế";
            ws.Cells[1, 4].Value = "Tiêu chuẩn";

            for (int i = 0; i < data.Count; i++)
            {
                var r = i + 2;
                ws.Cells[r, 1].Value = data[i].Time.ToString("yyyy-MM-dd HH:mm:ss");
                ws.Cells[r, 2].Value = data[i].PPM;
                ws.Cells[r, 3].Value = data[i].ThucTe;
                ws.Cells[r, 4].Value = data[i].TieuChuan;
            }

            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{fileName}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        }
    }
}
