using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using TemperatureApp.Services;
using TemperatureApp.Models;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.IO;
using TemperatureApp.Data;

namespace TemperatureApp.Controllers
{
    public class BangTaiController : Controller
    {
        private readonly ILogger<BangTaiController> _logger;
        private readonly DatabaseService _databaseService;
        private readonly ApplicationDbContext _context;

        public BangTaiController(ILogger<BangTaiController> logger, DatabaseService databaseService, ApplicationDbContext context)
        {
            _logger = logger;
            _databaseService = databaseService;
            _context = context;
        }

        public async Task<IActionResult> IndexAsync(DateTime? startDate, DateTime? endDate)
        {
            if (!await _databaseService.CanConnectAsync())
            {
                TempData["Error"] = "Không thể kết nối đến cơ sở dữ liệu.";
                return View();
            }

            var query = _context.CBstellcord_1.AsQueryable();

            if (startDate.HasValue && endDate.HasValue)
            {
                if (startDate > endDate)
                {
                    TempData["Error"] = "Ngày bắt đầu phải nhỏ hơn hoặc bằng ngày kết thúc.";
                    return View();
                }
                var adjustedEnd = endDate.Value.AddSeconds(59).AddMilliseconds(999);
                query = query.Where(x => x.CB1_TIMES >= startDate && x.CB1_TIMES <= adjustedEnd);
            }

            var data = await (startDate.HasValue && endDate.HasValue
                ? query.OrderByDescending(x => x.CB1_TIMES).ToListAsync()
                : query.OrderByDescending(x => x.CB1_TIMES).Take(50).ToListAsync());

            data.Reverse();

            ViewData["Title"] = "Press #1";
            ViewData["Action"] = "Index";
            ViewData["ExportAction"] = "ExportToExcel";
            ViewData["Labels"] = data.Select(x => x.CB1_TIMES.ToString("yyyy-MM-dd HH:mm:ss")).ToList();
            ViewData["ApLucDa"] = data.Select(x => x.CB1_POLE_STOPER).ToList();      // Áp lực đá
            ViewData["ApLucBanKep"] = data.Select(x => x.CB1_TENSION_CLAMP).ToList(); // Áp lực bàn kẹp
            ViewData["CB1_STATUS"] = data.Select(x => x.CB1_STATUS ? 1 : 0).ToList();

            ViewData["StartDate"] = startDate?.ToString("yyyy-MM-ddTHH:mm");
            ViewData["EndDate"] = endDate?.ToString("yyyy-MM-ddTHH:mm");

            return View();
        }

        // ---------------- Press #5 ----------------
        public async Task<IActionResult> Press5(DateTime? startDate, DateTime? endDate)
        {
            if (!await _databaseService.CanConnectAsync())
            {
                TempData["Error"] = "Không thể kết nối đến cơ sở dữ liệu.";
                return View("Index");
            }

            var query = _context.CBstellcord_5.AsQueryable();

            if (startDate.HasValue && endDate.HasValue)
            {
                if (startDate > endDate)
                {
                    TempData["Error"] = "Ngày bắt đầu phải nhỏ hơn hoặc bằng ngày kết thúc.";
                    return View("Index");
                }
                var adjustedEnd = endDate.Value.AddSeconds(59).AddMilliseconds(999);
                query = query.Where(x => x.CB_5_TIMES >= startDate && x.CB_5_TIMES <= adjustedEnd);
            }

            var data = await (startDate.HasValue && endDate.HasValue
                ? query.OrderByDescending(x => x.CB_5_TIMES).ToListAsync()
                : query.OrderByDescending(x => x.CB_5_TIMES).Take(50).ToListAsync());

            data.Reverse();

            ViewData["Title"] = "Press #5";
            ViewData["Action"] = "Press5";
            ViewData["ExportAction"] = "ExportPress5";
            ViewData["Labels"] = data.Select(x => x.CB_5_TIMES.ToString("yyyy-MM-dd HH:mm:ss")).ToList();
            ViewData["ApLucDa"] = data.Select(x => x.CB_5_POLE).ToList();
            ViewData["ApLucBanKep"] = data.Select(x => x.CB_5_CLAMP).ToList();
            ViewData["CB1_STATUS"] = data.Select(x => x.CB_5_STATUS ? 1 : 0).ToList();

            ViewData["StartDate"] = startDate?.ToString("yyyy-MM-ddTHH:mm");
            ViewData["EndDate"] = endDate?.ToString("yyyy-MM-ddTHH:mm");

            return View("Index");
        }

        public async Task<IActionResult> ExportPress5(DateTime? startDate, DateTime? endDate)
        {
            if (!startDate.HasValue || !endDate.HasValue || startDate > endDate)
                return BadRequest("Ngày không hợp lệ.");

            var adjustedEnd = endDate.Value.AddSeconds(59).AddMilliseconds(999);
            var data = await _context.CBstellcord_5
                .Where(x => x.CB_5_TIMES >= startDate && x.CB_5_TIMES <= adjustedEnd)
                .OrderBy(x => x.CB_5_TIMES)
                .ToListAsync();

            if (!data.Any()) return BadRequest("Không có dữ liệu.");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Press #5");

            ws.Cells[1, 1].Value = "Thời gian";
            ws.Cells[1, 2].Value = "Áp lực đá (CB_5_POLE)";
            ws.Cells[1, 3].Value = "Trạng thái (ON/OFF)";
            ws.Cells[1, 4].Value = "Áp lực bàn kẹp (CB_5_CLAMP)";

            for (int i = 0; i < data.Count; i++)
            {
                var r = i + 2;
                ws.Cells[r, 1].Value = data[i].CB_5_TIMES.ToString("yyyy-MM-dd HH:mm:ss");
                ws.Cells[r, 2].Value = data[i].CB_5_POLE;
                ws.Cells[r, 3].Value = data[i].CB_5_STATUS ? "ON" : "OFF";
                ws.Cells[r, 4].Value = data[i].CB_5_CLAMP;
            }

            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Press5_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        }

        // ---------------- Press #7 ----------------
        public async Task<IActionResult> Press7(DateTime? startDate, DateTime? endDate)
        {
            if (!await _databaseService.CanConnectAsync())
            {
                TempData["Error"] = "Không thể kết nối đến cơ sở dữ liệu.";
                return View("Index");
            }

            var query = _context.CBstellcord_7.AsQueryable();

            if (startDate.HasValue && endDate.HasValue)
            {
                if (startDate > endDate)
                {
                    TempData["Error"] = "Ngày bắt đầu phải nhỏ hơn hoặc bằng ngày kết thúc.";
                    return View("Index");
                }
                var adjustedEnd = endDate.Value.AddSeconds(59).AddMilliseconds(999);
                query = query.Where(x => x.CB_7_TIMES >= startDate && x.CB_7_TIMES <= adjustedEnd);
            }

            var data = await (startDate.HasValue && endDate.HasValue
                ? query.OrderByDescending(x => x.CB_7_TIMES).ToListAsync()
                : query.OrderByDescending(x => x.CB_7_TIMES).Take(50).ToListAsync());

            data.Reverse();

            ViewData["Title"] = "Press #7";
            ViewData["Action"] = "Press7";
            ViewData["ExportAction"] = "ExportPress7";
            ViewData["Labels"] = data.Select(x => x.CB_7_TIMES.ToString("yyyy-MM-dd HH:mm:ss")).ToList();
            ViewData["ApLucDa"] = data.Select(x => x.CB_7_POLE).ToList();
            ViewData["ApLucBanKep"] = data.Select(x => x.CB_7_CLAMP).ToList();
            ViewData["CB1_STATUS"] = data.Select(x => x.CB_7_STATUS ? 1 : 0).ToList();

            ViewData["StartDate"] = startDate?.ToString("yyyy-MM-ddTHH:mm");
            ViewData["EndDate"] = endDate?.ToString("yyyy-MM-ddTHH:mm");

            return View("Index");
        }

        public async Task<IActionResult> ExportPress7(DateTime? startDate, DateTime? endDate)
        {
            if (!startDate.HasValue || !endDate.HasValue || startDate > endDate)
                return BadRequest("Ngày không hợp lệ.");

            var adjustedEnd = endDate.Value.AddSeconds(59).AddMilliseconds(999);
            var data = await _context.CBstellcord_7
                .Where(x => x.CB_7_TIMES >= startDate && x.CB_7_TIMES <= adjustedEnd)
                .OrderBy(x => x.CB_7_TIMES)
                .ToListAsync();

            if (!data.Any()) return BadRequest("Không có dữ liệu.");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Press #7");

            ws.Cells[1, 1].Value = "Thời gian";
            ws.Cells[1, 2].Value = "Áp lực đá (CB_7_POLE)";
            ws.Cells[1, 3].Value = "Trạng thái (ON/OFF)";
            ws.Cells[1, 4].Value = "Áp lực bàn kẹp (CB_7_CLAMP)";

            for (int i = 0; i < data.Count; i++)
            {
                var r = i + 2;
                ws.Cells[r, 1].Value = data[i].CB_7_TIMES.ToString("yyyy-MM-dd HH:mm:ss");
                ws.Cells[r, 2].Value = data[i].CB_7_POLE;
                ws.Cells[r, 3].Value = data[i].CB_7_STATUS ? "ON" : "OFF";
                ws.Cells[r, 4].Value = data[i].CB_7_CLAMP;
            }

            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Press7_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        }

        public async Task<IActionResult> ExportToExcel(DateTime? startDate, DateTime? endDate)
        {
            if (!startDate.HasValue || !endDate.HasValue || startDate > endDate)
                return BadRequest("Ngày không hợp lệ.");

            var adjustedEnd = endDate.Value.AddSeconds(59).AddMilliseconds(999);
            var data = await _context.CBstellcord_1
                .Where(x => x.CB1_TIMES >= startDate && x.CB1_TIMES <= adjustedEnd)
                .OrderBy(x => x.CB1_TIMES)
                .ToListAsync();

            if (!data.Any()) return BadRequest("Không có dữ liệu.");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Băng Tải");

            ws.Cells[1, 1].Value = "Thời gian";
            ws.Cells[1, 2].Value = "Áp lực đá (CB1_POLE_STOPER)";
            ws.Cells[1, 3].Value = "Trạng thái (ON/OFF)";
            ws.Cells[1, 4].Value = "Áp lực bàn kẹp (CB1_TENSION_CLAMP)";

            for (int i = 0; i < data.Count; i++)
            {
                var r = i + 2;
                ws.Cells[r, 1].Value = data[i].CB1_TIMES.ToString("yyyy-MM-dd HH:mm:ss");
                ws.Cells[r, 2].Value = data[i].CB1_POLE_STOPER;
                ws.Cells[r, 3].Value = data[i].CB1_STATUS ? "ON" : "OFF";
                ws.Cells[r, 4].Value = data[i].CB1_TENSION_CLAMP;
            }

            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"BangTai_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        }
    }
}