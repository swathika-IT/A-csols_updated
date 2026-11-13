const fs = require('fs');
const os = require('os');
const path = require('path');

function createMultipartParser({ uploadDir, maxFileSize = 8 * 1024 * 1024, allowedExtensions = [] } = {}) {
  // Never create directories at module import / parser creation time.
  // Serverless filesystems (e.g. Vercel /var/task) are read-only, so the
  // target directory is resolved lazily per upload request. Temporary
  // processing files live under os.tmpdir(); only callers that explicitly
  // need durable local storage pass uploadDir.
  const temporary = !uploadDir;
  const dir = temporary
    ? path.join(os.tmpdir(), 'ac-solutions-uploads')
    : path.resolve(uploadDir);

  return function parseMultipart(req, res, next) {
    try {
      const contentType = String(req.headers['content-type'] || '');
      if (!contentType.toLowerCase().startsWith('multipart/form-data')) return next();
      const raw = req.body;
      if (!Buffer.isBuffer(raw)) return next(new Error('Invalid multipart body.'));

      const match = contentType.match(/boundary=(?:(?:"([^"]+)")|([^;]+))/i);
      if (!match) return next(new Error('Multipart boundary missing.'));
      const boundary = Buffer.from('--' + (match[1] || match[2]));
      const body = raw;
      const fields = {};
      const files = {};
      const written = [];
      let pos = 0;
      let dirReady = false;

      const ensureDir = () => {
        if (dirReady) return true;
        try {
          fs.mkdirSync(dir, { recursive: true });
          dirReady = true;
          return true;
        } catch (err) {
          return err;
        }
      };

      const cleanupWritten = () => {
        for (const p of written) { try { fs.unlinkSync(p); } catch (_) {} }
      };

      while ((pos = body.indexOf(boundary, pos)) !== -1) {
        pos += boundary.length;
        if (body.slice(pos, pos + 2).toString() === '--') break;
        if (body.slice(pos, pos + 2).toString() === '\r\n') pos += 2;
        const nextBoundary = body.indexOf(boundary, pos);
        if (nextBoundary < 0) break;
        const part = body.slice(pos, nextBoundary - 2);
        const sep = part.indexOf(Buffer.from('\r\n\r\n'));
        if (sep < 0) continue;

        const headers = part.slice(0, sep).toString('utf8');
        const content = part.slice(sep + 4);
        const nameMatch = headers.match(/name="([^"]+)"/i);
        if (!nameMatch) continue;
        const fieldName = nameMatch[1];
        const fileMatch = headers.match(/filename="([^"]*)"/i);

        if (fileMatch && fileMatch[1]) {
          const originalName = path.basename(fileMatch[1]);
          const ext = path.extname(originalName).toLowerCase();
          if (!allowedExtensions.includes(ext)) {
            cleanupWritten();
            return res.status(400).json({ success: false, message: `Unsupported file type for ${fieldName}. Allowed: ${allowedExtensions.join(', ')}` });
          }
          if (content.length > maxFileSize) {
            cleanupWritten();
            return res.status(400).json({ success: false, message: `${fieldName} must be ${Math.round(maxFileSize / 1024 / 1024)} MB or smaller.` });
          }
          const ready = ensureDir();
          if (ready !== true) {
            cleanupWritten();
            return res.status(503).json({ success: false, message: 'Upload storage is unavailable on this server. Please contact the administrator.' });
          }
          const filename = `${Date.now()}-${Math.random().toString(36).slice(2)}${ext}`;
          const filePath = path.join(dir, filename);
          try {
            fs.writeFileSync(filePath, content);
          } catch (err) {
            cleanupWritten();
            return res.status(503).json({ success: false, message: 'Upload storage is unavailable on this server. Please contact the administrator.' });
          }
          written.push(filePath);
          files[fieldName] = { path: filePath, originalname: originalName, filename, size: content.length, extension: ext };
        } else {
          fields[fieldName] = content.toString('utf8').trim();
        }
      }

      req.body = fields;
      req.files = files;
      next();
    } catch (err) {
      next(err);
    }
  };
}

module.exports = { createMultipartParser };
