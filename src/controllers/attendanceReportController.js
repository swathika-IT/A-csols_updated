const service = require('../services/attendanceReportService');

function pickMonthYear(req) {
  const monthRaw = req.query.month;
  const yearRaw = req.query.year;
  const monthParam = req.query.monthParam;
  if (monthParam && /^\d{4}-\d{2}$/.test(String(monthParam))) return service.parseMonthYear({ month: String(monthParam) });
  if (monthRaw && /^\d{4}-\d{2}$/.test(String(monthRaw).trim())) return service.parseMonthYear({ month: String(monthRaw).trim() });
  const month = monthRaw != null ? parseInt(monthRaw, 10) : NaN;
  const year = yearRaw != null ? parseInt(yearRaw, 10) : NaN;
  return service.parseMonthYear({ month, year });
}

async function downloadMonthly(req, res, next) {
  try {
    const my = pickMonthYear(req);
    const report = await service.buildMonthlyReport({
      month: my.month,
      year: my.year,
      category: req.query.category || 'all',
      employeeId: req.query.employeeId || undefined
    });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${report.filename}"`);
    res.setHeader('Content-Length', report.buffer.length);
    res.setHeader('X-Content-Type-Options', 'nosniff');
    return res.end(report.buffer);
  } catch (err) {
    return next(err);
  }
}

async function downloadDaily(req, res, next) {
  try {
    const report = await service.buildDailyReport({
      date: req.query.date,
      category: req.query.category || 'all',
      employeeId: req.query.employeeId || undefined
    });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${report.filename}"`);
    res.setHeader('Content-Length', report.buffer.length);
    res.setHeader('X-Content-Type-Options', 'nosniff');
    return res.end(report.buffer);
  } catch (err) {
    return next(err);
  }
}

async function previewMonthly(req, res, next) {
  try {
    const my = pickMonthYear(req);
    const report = await service.buildMonthlyReport({
      month: my.month,
      year: my.year,
      category: req.query.category || 'all',
      employeeId: req.query.employeeId || undefined
    });
    return res.json({
      success: true,
      monthYear: report.meta.monthYear,
      category: report.meta.category,
      count: report.meta.count,
      totals: report.meta.totals,
      summary: report.summary
    });
  } catch (err) {
    return next(err);
  }
}

module.exports = { downloadMonthly, downloadDaily, previewMonthly };
