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

            ViewData["Labels"] = data.Select(x => x.CB1_TIMES.ToString("yyyy-MM-dd HH:mm:ss")).ToList();
            ViewData["ApLucDa"] = data.Select(x => x.CB1_POLE_STOPER).ToList();      // Áp lực đá
            ViewData["ApLucBanKep"] = data.Select(x => x.CB1_TENSION_CLAMP).ToList(); // Áp lực bàn kẹp
            ViewData["CB1_STATUS"] = data.Select(x => x.CB1_STATUS ? 1 : 0).ToList();

            ViewData["StartDate"] = startDate?.ToString("yyyy-MM-ddTHH:mm");
            ViewData["EndDate"] = endDate?.ToString("yyyy-MM-ddTHH:mm");

            return View();
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