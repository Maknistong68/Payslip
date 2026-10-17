#!/usr/bin/env node
/* Builds icons/ (SVG + PNG app icons) from the logo. Needs Playwright with Chromium (dev only):
 * NODE_PATH=$(npm root -g) node tools/build-icons.js */
'use strict';
var path = require('path');
var fs = require('fs');
var out = path.join(__dirname, '..', 'icons');
if (!fs.existsSync(out)) fs.mkdirSync(out);

var DEFS = '<defs><linearGradient id="a" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#1e3a8a"/><stop offset=".55" stop-color="#2563eb"/><stop offset="1" stop-color="#0ea5e9"/></linearGradient>' +
  '<linearGradient id="b" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff"/><stop offset="1" stop-color="#e2e8f0"/></linearGradient></defs>';
var MARK = '<g transform="rotate(-8 24 25)">' +
  '<path d="M15 10.5h14.5l5 5V35l-2.5 2.5-2.5-2.5-2.5 2.5-2.5-2.5-2.5 2.5-2.5-2.5-2.5 2.5-2.5-2.5V12.5a2 2 0 0 1 2-2z" fill="url(#b)"/>' +
  '<path d="M29.5 10.5v3.5a1.5 1.5 0 0 0 1.5 1.5h3.5z" fill="#cbd5e1"/>' +
  '<rect x="17" y="17" width="9" height="2.4" rx="1.2" fill="#93c5fd"/><rect x="17" y="22" width="13" height="2.4" rx="1.2" fill="#cbd5e1"/>' +
  '<rect x="17" y="27" width="7" height="2.4" rx="1.2" fill="url(#a)"/></g>' +
  '<circle cx="35" cy="34" r="7.5" fill="#10b981" stroke="#fff" stroke-width="2.2"/>' +
  '<path d="m31.6 34 2.3 2.3 4.4-4.6" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>';

// rounded tile (browsers, desktop)
var rounded = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">' + DEFS +
  '<rect x="2" y="2" width="44" height="44" rx="14" fill="url(#a)"/><rect x="2" y="2" width="44" height="22" rx="14" fill="#fff" opacity=".12"/>' + MARK + '</svg>';
// full bleed, artwork inside the safe zone (Android maskable, iOS rounds corners itself)
function bleed(scale) {
  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">' + DEFS + '<rect width="48" height="48" fill="url(#a)"/>' +
    '<g transform="translate(24 24) scale(' + scale + ') translate(-24.5 -24.5)">' + MARK + '</g></svg>';
}

fs.writeFileSync(path.join(out, 'icon.svg'), rounded + '\n');

var chromium = require('playwright').chromium;
(async function () {
  var browser = await chromium.launch();
  var page = await browser.newPage();
  async function png(svg, size, file) {
    await page.setViewportSize({ width: size, height: size });
    await page.setContent('<html><body style="margin:0;background:transparent">' + svg.replace('<svg ', '<svg width="' + size + '" height="' + size + '" ') + '</body></html>');
    await page.screenshot({ path: path.join(out, file), omitBackground: true, clip: { x: 0, y: 0, width: size, height: size } });
    console.log('Wrote icons/' + file);
  }
  await png(rounded, 192, 'icon-192.png');
  await png(rounded, 512, 'icon-512.png');
  await png(bleed(0.78), 512, 'icon-maskable-512.png');
  await png(bleed(0.9), 180, 'apple-touch-icon.png');
  await browser.close();
})().catch(function (e) { console.error(e); process.exit(1); });
