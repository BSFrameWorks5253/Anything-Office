const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');

function createIcoFromPngBuffers(pngBuffers) {
  const numImages = pngBuffers.length;
  const headerSize = 6;
  const dirEntrySize = 16;
  const totalDirSize = headerSize + numImages * dirEntrySize;

  let currentOffset = totalDirSize;
  const dirBuffers = [];
  const imageBuffers = [];

  for (const item of pngBuffers) {
    const { width, height, buffer } = item;
    const entry = Buffer.alloc(16);
    entry.writeUInt8(width >= 256 ? 0 : width, 0);
    entry.writeUInt8(height >= 256 ? 0 : height, 1);
    entry.writeUInt8(0, 2);
    entry.writeUInt8(0, 3);
    entry.writeUInt16LE(1, 4);
    entry.writeUInt16LE(32, 6);
    entry.writeUInt32LE(buffer.length, 8);
    entry.writeUInt32LE(currentOffset, 12);

    dirBuffers.push(entry);
    imageBuffers.push(buffer);
    currentOffset += buffer.length;
  }

  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(numImages, 4);

  return Buffer.concat([header, ...dirBuffers, ...imageBuffers]);
}

app.whenReady().then(async () => {
  try {
    const win = new BrowserWindow({
      width: 512,
      height: 512,
      show: false,
      useContentSize: true,
      frame: false,
      transparent: true,
    });

    const svgPath = path.join(__dirname, 'icon.svg');
    const svgContent = fs.readFileSync(svgPath, 'utf8');
    const html = `<!DOCTYPE html>
<html>
<head>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    html, body { width: 512px; height: 512px; background: transparent; overflow: hidden; }
    svg { width: 512px; height: 512px; display: block; }
  </style>
</head>
<body>
  ${svgContent}
</body>
</html>`;

    await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);

    // Allow layout and filters to render
    await new Promise((resolve) => setTimeout(resolve, 500));

    const masterImage = await win.webContents.capturePage({ x: 0, y: 0, width: 512, height: 512 });

    const sizes = [256, 128, 64, 48, 32, 16];
    const pngBuffersForIco = [];

    // Save 512x512
    const buf512 = masterImage.toPNG();
    fs.writeFileSync(path.join(__dirname, 'icon.png'), buf512);
    fs.writeFileSync(path.join(__dirname, '../public/icon.png'), buf512);
    fs.writeFileSync(path.join(__dirname, '../public/favicon.svg'), svgContent);

    // Generate mipmaps
    for (const size of sizes) {
      const resized = masterImage.resize({ width: size, height: size, quality: 'best' });
      const buf = resized.toPNG();
      fs.writeFileSync(path.join(__dirname, `icon-${size}.png`), buf);
      pngBuffersForIco.push({ width: size, height: size, buffer: buf });
    }

    // Generate Windows .ico
    const icoBuffer = createIcoFromPngBuffers(pngBuffersForIco);
    fs.writeFileSync(path.join(__dirname, 'icon.ico'), icoBuffer);
    fs.writeFileSync(path.join(__dirname, '../public/favicon.ico'), icoBuffer);

    console.log('✅ Successfully generated icon.png, icon.ico, and favicon mipmaps!');
    win.destroy();
    app.quit();
  } catch (err) {
    console.error('❌ Error generating icons:', err);
    process.exit(1);
  }
});
