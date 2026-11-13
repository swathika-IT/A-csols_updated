/**
 * A&C SOLUTIONS PVT. LTD. - TECHNICAL SEO & SITEMAP VERIFICATION SUITE
 */

const assert = require('assert');
const http = require('http');
const fs = require('fs');
const path = require('path');
const app = require('../src/app');

async function test(name, fn) {
  try {
    process.stdout.write(`• Testing SEO: ${name}... `);
    await fn();
    console.log('✓ PASS');
  } catch (err) {
    console.log('✕ FAIL');
    console.error(err);
    process.exitCode = 1;
  }
}

function request(server, pathStr) {
  return new Promise((resolve, reject) => {
    const addr = server.address();
    const req = http.get({
      host: '127.0.0.1',
      port: addr.port,
      path: pathStr
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        resolve({
          statusCode: res.statusCode,
          headers: res.headers,
          body: data
        });
      });
    });
    req.on('error', reject);
  });
}

async function runSeoTests() {
  console.log('\n=============================================================');
  console.log('STARTING TECHNICAL SEO & SITEMAP VERIFICATION');
  console.log('=============================================================');

  const server = http.createServer(app);
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));

  try {
    // 1. Robots.txt
    await test('/robots.txt returns HTTP 200 with text/plain', async () => {
      const res = await request(server, '/robots.txt');
      assert.strictEqual(res.statusCode, 200);
      assert.ok(res.headers['content-type'].includes('text/plain'), `Expected text/plain but got ${res.headers['content-type']}`);
      assert.ok(res.body.includes('User-agent: *'));
      assert.ok(res.body.includes('Allow: /'));
      assert.ok(res.body.includes('Sitemap: https://www.ancsolutionspvtltd.com/sitemap.xml'));
      assert.ok(res.body.includes('Disallow: /api/'));
      assert.ok(res.body.includes('Disallow: /admin'));
      assert.ok(res.body.includes('Disallow: /login'));
      assert.ok(res.body.includes('Disallow: /register'));
      assert.ok(res.body.includes('Disallow: /employee'));
      assert.ok(res.body.includes('Disallow: /intern'));
    });

    // 2. Sitemap.xml HTTP status and MIME type
    await test('/sitemap.xml returns HTTP 200 with xml content-type', async () => {
      const res = await request(server, '/sitemap.xml');
      assert.strictEqual(res.statusCode, 200);
      assert.ok(/xml/.test(res.headers['content-type']), `Expected xml content-type but got ${res.headers['content-type']}`);
      assert.ok(res.body.startsWith('<?xml version="1.0" encoding="UTF-8"?>'));
      assert.ok(res.body.includes('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'));
    });

    // 3. Sitemap XML validity and canonical URLs
    const expectedUrls = [
      'https://www.ancsolutionspvtltd.com/',
      'https://www.ancsolutionspvtltd.com/domain-cctv.html',
      'https://www.ancsolutionspvtltd.com/it-technology.html',
      'https://www.ancsolutionspvtltd.com/Real%20Estate%20Promotions.html',
      'https://www.ancsolutionspvtltd.com/digital%20marketing.html',
      'https://www.ancsolutionspvtltd.com/content-creation.html',
      'https://www.ancsolutionspvtltd.com/content-moderation.html'
    ];

    await test('sitemap.xml contains exact 7 public canonical URLs and valid tags', async () => {
      const res = await request(server, '/sitemap.xml');
      const locMatches = [...res.body.matchAll(/<loc>(.*?)<\/loc>/g)].map(m => m[1]);
      assert.strictEqual(locMatches.length, 7);

      for (const expected of expectedUrls) {
        assert.ok(locMatches.includes(expected), `Missing expected URL in sitemap: ${expected}`);
      }

      // Check no private routes in sitemap
      const lower = res.body.toLowerCase();
      assert.ok(!lower.includes('/api/'), 'Sitemap must not contain API routes');
      assert.ok(!lower.includes('/admin'), 'Sitemap must not contain admin routes');
      assert.ok(!lower.includes('/employee'), 'Sitemap must not contain employee routes');
      assert.ok(!lower.includes('/intern'), 'Sitemap must not contain intern routes');
      assert.ok(!lower.includes('/login'), 'Sitemap must not contain login routes');
      assert.ok(!lower.includes('/register'), 'Sitemap must not contain register routes');
    });

    // 4. Verify each URL in sitemap resolves with HTTP 200
    for (const url of expectedUrls) {
      const parsedPath = new URL(url).pathname;
      await test(`Sitemap page ${parsedPath} resolves with HTTP 200`, async () => {
        const res = await request(server, parsedPath);
        assert.strictEqual(res.statusCode, 200, `Expected 200 for ${parsedPath} but got ${res.statusCode}`);
        assert.ok(res.body.includes('<!DOCTYPE html>') || res.body.includes('<!doctype html>'));
      });
    }

    // 5. Verify SEO tags on all public pages
    const publicPages = [
      { file: 'index.html', path: '/', canonical: 'https://www.ancsolutionspvtltd.com/' },
      { file: 'domain-cctv.html', path: '/domain-cctv.html', canonical: 'https://www.ancsolutionspvtltd.com/domain-cctv.html' },
      { file: 'it-technology.html', path: '/it-technology.html', canonical: 'https://www.ancsolutionspvtltd.com/it-technology.html' },
      { file: 'Real Estate Promotions.html', path: '/Real%20Estate%20Promotions.html', canonical: 'https://www.ancsolutionspvtltd.com/Real%20Estate%20Promotions.html' },
      { file: 'digital marketing.html', path: '/digital%20marketing.html', canonical: 'https://www.ancsolutionspvtltd.com/digital%20marketing.html' },
      { file: 'content-creation.html', path: '/content-creation.html', canonical: 'https://www.ancsolutionspvtltd.com/content-creation.html' },
      { file: 'content-moderation.html', path: '/content-moderation.html', canonical: 'https://www.ancsolutionspvtltd.com/content-moderation.html' }
    ];

    for (const p of publicPages) {
      await test(`Public page ${p.file} has unique SEO title, meta description, canonical and OG metadata`, async () => {
        const res = await request(server, p.path);
        const html = res.body;

        // Title
        const titleMatch = html.match(/<title>(.*?)<\/title>/i);
        assert.ok(titleMatch && titleMatch[1].trim().length > 10, `${p.file} missing or short title`);

        // Meta description
        const descMatch = html.match(/<meta[^>]*name=["']description["'][^>]*content=["'](.*?)["']/i);
        assert.ok(descMatch && descMatch[1].trim().length > 20, `${p.file} missing or short meta description`);

        // Canonical URL
        const canonicalMatch = html.match(/<link[^>]*rel=["']canonical["'][^>]*href=["'](.*?)["']/i);
        assert.ok(canonicalMatch, `${p.file} missing canonical link`);
        assert.strictEqual(canonicalMatch[1], p.canonical, `${p.file} canonical link mismatch`);

        // Open Graph URL
        const ogUrlMatch = html.match(/<meta[^>]*property=["']og:url["'][^>]*content=["'](.*?)["']/i);
        assert.ok(ogUrlMatch, `${p.file} missing og:url`);
        assert.strictEqual(ogUrlMatch[1], p.canonical, `${p.file} og:url mismatch`);

        // Open Graph Image
        const ogImgMatch = html.match(/<meta[^>]*property=["']og:image["'][^>]*content=["'](.*?)["']/i);
        assert.ok(ogImgMatch && ogImgMatch[1].startsWith('https://www.ancsolutionspvtltd.com/'), `${p.file} missing or invalid og:image`);

        // Heading Structure: exactly one <h1>
        const h1Matches = html.match(/<h1[\s>]/gi);
        assert.ok(h1Matches && h1Matches.length === 1, `${p.file} must have exactly one <h1> heading (found ${h1Matches ? h1Matches.length : 0})`);
      });
    }

    // 6. Verify private pages have noindex, nofollow
    const privatePages = [
      'login.html',
      'register.html',
      'admin-login.html',
      'employee-login.html',
      'intern-login.html',
      'admin.html',
      'admin-workspace.html',
      'employee.html'
    ];

    for (const file of privatePages) {
      await test(`Private page ${file} contains robots noindex, nofollow`, async () => {
        const content = fs.readFileSync(path.join(__dirname, '..', 'public', file), 'utf8');
        assert.ok(
          content.includes('content="noindex, nofollow"') || content.includes('content=\'noindex, nofollow\''),
          `${file} must contain robots noindex, nofollow directive`
        );
      });
    }

  } finally {
    server.close();
  }

  console.log('=============================================================');
  console.log('TECHNICAL SEO & SITEMAP VERIFICATION COMPLETED SUCCESSFULLY');
  console.log('=============================================================\n');
}

if (require.main === module) {
  runSeoTests().then(() => {
    if (process.exitCode) process.exit(process.exitCode);
  });
}

module.exports = { runSeoTests };

