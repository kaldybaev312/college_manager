// Подготовка фото для терминала СКУД: уменьшение, проверка качества, при поддержке браузера поиск лица и кадрирование.
// Используется на страницах регистрации и на публичной анкете. Ничего не отправляет на сервер.
(function () {
  const MAX_SIDE = 1024, MAX_BYTES = 190 * 1024, MIN_SIDE = 300;

  async function faces(bitmap) {
    try {
      if (!("FaceDetector" in window)) return null;
      return await new window.FaceDetector({ fastMode: true, maxDetectedFaces: 5 }).detect(bitmap);
    } catch { return null; }
  }

  // Резкость (дисперсия лапласиана) и яркость по уменьшенной копии
  function metrics(bitmap) {
    const k = Math.min(1, 256 / Math.max(bitmap.width, bitmap.height));
    const w = Math.max(8, Math.round(bitmap.width * k)), h = Math.max(8, Math.round(bitmap.height * k));
    const c = document.createElement("canvas"); c.width = w; c.height = h;
    const x = c.getContext("2d", { willReadFrequently: true }); x.drawImage(bitmap, 0, 0, w, h);
    const d = x.getImageData(0, 0, w, h).data;
    const g = new Float32Array(w * h); let sum = 0;
    for (let i = 0; i < w * h; i++) { g[i] = 0.299 * d[i * 4] + 0.587 * d[i * 4 + 1] + 0.114 * d[i * 4 + 2]; sum += g[i]; }
    const mean = sum / (w * h);
    let m = 0, n = 0, v = 0;
    const lap = [];
    for (let yy = 1; yy < h - 1; yy++) for (let xx = 1; xx < w - 1; xx++) {
      const i = yy * w + xx;
      const l = 4 * g[i] - g[i - 1] - g[i + 1] - g[i - w] - g[i + w];
      lap.push(l); m += l; n++;
    }
    m /= n; for (const l of lap) v += (l - m) * (l - m);
    let sd = 0; for (let i = 0; i < w * h; i++) sd += (g[i] - mean) ** 2;
    return { sharp: v / n, mean, contrast: Math.sqrt(sd / (w * h)) };
  }

  async function toJpeg(canvas) {
    for (const q of [0.88, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3]) {
      const b = await new Promise((r) => canvas.toBlob(r, "image/jpeg", q));
      if (b && b.size <= MAX_BYTES) return b;
    }
    throw new Error("Не удалось уменьшить фото, выберите другое");
  }

  // Возвращает { blob, issues: [{level: "error"|"warn", text}], faceChecked }
  async function prepare(file, opts = {}) {
    if (!/^image\//.test(file.type) && !/\.(jpe?g|png|webp)$/i.test(file.name || ""))
      throw new Error("Нужен файл-изображение (JPG, PNG)");
    if (file.size > 25 * 1024 * 1024) throw new Error("Файл слишком большой (больше 25 МБ)");
    let bmp;
    try { bmp = await createImageBitmap(file, { imageOrientation: "from-image" }); }
    catch { throw new Error(/HEIC|HEIF/i.test(file.type + file.name) ? "Формат HEIC не поддерживается, сохраните как JPG" : "Не удалось открыть фото"); }

    const issues = [];
    const fs = await faces(bmp);
    let sx = 0, sy = 0, sw = bmp.width, sh = bmp.height;
    if (fs) {
      if (fs.length === 0) issues.push({ level: "warn", text: "Лицо не найдено. Проверьте, что лицо крупно и видно целиком." });
      else if (fs.length > 1) issues.push({ level: "warn", text: "На фото несколько лиц. Нужен один человек." });
      else if (opts.crop !== false) {
        // Кадрируем вокруг лица (портрет 3:4), чтобы лицо занимало основную часть кадра
        const f = fs[0].boundingBox;
        const ch = Math.min(bmp.height, f.height * 2.6), cw = Math.min(bmp.width, ch * 0.75);
        if (f.width / bmp.width < 0.55) {
          sw = cw; sh = ch;
          sx = Math.min(Math.max(0, f.x + f.width / 2 - cw / 2), bmp.width - cw);
          sy = Math.min(Math.max(0, f.y + f.height / 2 - ch * 0.5), bmp.height - ch);
        }
        if (f.width < Math.min(bmp.width, bmp.height) * 0.18) issues.push({ level: "warn", text: "Лицо слишком мелкое. Подойдите ближе." });
      }
    }
    const k = Math.min(1, MAX_SIDE / Math.max(sw, sh));
    const c = document.createElement("canvas");
    c.width = Math.round(sw * k); c.height = Math.round(sh * k);
    c.getContext("2d").drawImage(bmp, sx, sy, sw, sh, 0, 0, c.width, c.height);
    if (Math.min(c.width, c.height) < MIN_SIDE)
      issues.push({ level: "error", text: `Фото слишком маленькое (${c.width}×${c.height}). Нужно не меньше ${MIN_SIDE} px по короткой стороне.` });

    const m = metrics(c);
    if (m.sharp < 35) issues.push({ level: "warn", text: "Фото похоже на размытое. Переснимите при хорошем свете, не двигая камеру." });
    if (m.mean < 70) issues.push({ level: "warn", text: "Слишком тёмное фото. Нужен ровный свет на лицо." });
    else if (m.mean > 205) issues.push({ level: "warn", text: "Слишком светлое фото (засветка)." });
    else if (m.contrast < 25) issues.push({ level: "warn", text: "Низкий контраст. Добавьте света." });

    const blob = await toJpeg(c);
    bmp.close && bmp.close();
    return { blob, issues, faceChecked: !!fs, width: c.width, height: c.height };
  }

  window.SkudPhoto = { prepare, MIN_SIDE };
})();
