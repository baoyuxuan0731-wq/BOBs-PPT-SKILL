function serializeSlide(slide) {
  const page = slide.getBoundingClientRect();
  const segmenter = new Intl.Segmenter("zh-CN", { granularity: "grapheme" });
  const layers = [];
  const styles = new WeakMap();
  const ascentCache = new Map();
  const fixture = document.createElement("span");
  fixture.style.cssText = "position:fixed;left:-5000px;top:0;white-space:pre;pointer-events:none";
  document.body.append(fixture);

  function xml(value) {
    return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;").replaceAll('"', "&quot;");
  }
  function round(value) {
    return +value.toFixed(3);
  }
  function css(element) {
    if (!styles.has(element)) styles.set(element, getComputedStyle(element));
    return styles.get(element);
  }
  function ink(color) {
    return color !== "transparent" && color !== "rgba(0, 0, 0, 0)";
  }
  function paint(color) {
    const components = /^rgb\((\d+), (\d+), (\d+)\)$/.exec(color);
    if (!components) throw new Error("Use opaque solid colors for slide text and shapes.");
    return "#" + components.slice(1).map(part => Number(part).toString(16).padStart(2, "0")).join("");
  }
  function fillRect(x, y, w, h, color) {
    if (w > 0 && h > 0 && ink(color)) {
      layers.push(`<rect x="${round(x)}" y="${round(y)}" width="${round(w)}" height="${round(h)}" fill="${paint(color)}"/>`);
    }
  }
  function elementBox(element, style) {
    const rect = element.getBoundingClientRect();
    const x = rect.left - page.left;
    const y = rect.top - page.top;
    const w = rect.width;
    const h = rect.height;
    fillRect(x, y, w, h, style.backgroundColor);
    const top = parseFloat(style.borderTopWidth);
    const bottom = parseFloat(style.borderBottomWidth);
    const left = parseFloat(style.borderLeftWidth);
    const right = parseFloat(style.borderRightWidth);
    fillRect(x, y, w, top, style.borderTopColor);
    fillRect(x, y + h - bottom, w, bottom, style.borderBottomColor);
    fillRect(x, y, left, h, style.borderLeftColor);
    fillRect(x + w - right, y, right, h, style.borderRightColor);
  }
  function embedImage(element) {
    const rect = element.getBoundingClientRect();
    if (!/^data:image\/(png|jpeg);base64,/i.test(element.currentSrc || element.src)) {
      throw new Error("Embed result images as PNG or JPEG data URIs before exporting.");
    }
    if (!element.naturalWidth || !element.naturalHeight) throw new Error("An image has not loaded.");
    const scale = Math.min(rect.width / element.naturalWidth, rect.height / element.naturalHeight);
    const w = element.naturalWidth * scale;
    const h = element.naturalHeight * scale;
    const x = rect.left - page.left + (rect.width - w) / 2;
    const y = rect.top - page.top + (rect.height - h) / 2;
    layers.push(`<image x="${round(x)}" y="${round(y)}" width="${round(w)}" height="${round(h)}" xlink:href="${xml(element.currentSrc || element.src)}"><title>${xml(element.alt)}</title></image>`);
  }
  function ascent(style, content) {
    const cjk = /[\p{Script=Han}\u3000-\u303f\uff00-\uffef]/u.test(content);
    const key = [style.fontFamily, style.fontSize, style.fontWeight, style.fontStyle, cjk].join("|");
    if (ascentCache.has(key)) return ascentCache.get(key);
    fixture.style.fontFamily = style.fontFamily;
    fixture.style.fontSize = style.fontSize;
    fixture.style.fontWeight = style.fontWeight;
    fixture.style.fontStyle = style.fontStyle;
    const sample = document.createTextNode(cjk ? "测" : "Ag");
    const mark = document.createElement("i");
    mark.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline";
    fixture.replaceChildren(sample, mark);
    const range = document.createRange();
    range.selectNode(sample);
    const offset = mark.getBoundingClientRect().top - range.getBoundingClientRect().top;
    ascentCache.set(key, offset);
    return offset;
  }
  function textNode(node) {
    if (!node.data.trim()) return;
    const range = document.createRange();
    const style = css(node.parentElement);
    const font = style.fontFamily.split(",")[0].trim().replace(/^["']|["']$/g, "");
    let run;
    function emit() {
      if (!run?.content.trim()) return;
      const baseline = run.top - page.top + ascent(style, run.content);
      const x = run.left - page.left;
      const letter = style.letterSpacing === "normal" ? "" : ` letter-spacing="${parseFloat(style.letterSpacing)}"`;
      layers.push(`<text x="${round(x)}" y="${round(baseline)}" fill="${paint(style.color)}" font-family="${xml(font)}" font-size="${parseFloat(style.fontSize)}" font-weight="${style.fontWeight}" font-style="${style.fontStyle}"${letter}>${xml(run.content)}</text>`);
    }
    for (const { segment, index } of segmenter.segment(node.data)) {
      range.setStart(node, index);
      range.setEnd(node, index + segment.length);
      const rect = range.getClientRects()[0];
      if (!rect || rect.width < 0.01 || rect.height < 0.01) continue;
      const character = /\s/u.test(segment) && segment !== "\u00a0" && segment !== "\u3000" ? " " : segment;
      if (!run || Math.abs(rect.top - run.top) > 0.45 || Math.abs(rect.left - run.right) > 3) {
        emit();
        run = { left: rect.left, top: rect.top, right: rect.right, content: character };
      } else {
        run.content += character;
        run.right = rect.right;
      }
    }
    emit();
  }
  function draw(element) {
    if (element.matches("template,script,style,[data-export-ignore]")) return;
    const style = css(element);
    if (style.display === "none" || style.visibility === "hidden") return;
    if (element.namespaceURI !== "http://www.w3.org/1999/xhtml" || style.transform !== "none" || style.backgroundImage !== "none") {
      throw new Error("Simplify SVG elements, transforms, and background images or mark them data-export-ignore.");
    }
    elementBox(element, style);
    if (element.tagName === "IMG") {
      embedImage(element);
      return;
    }
    for (const child of element.childNodes) {
      if (child.nodeType === Node.ELEMENT_NODE) draw(child);
      else if (child.nodeType === Node.TEXT_NODE) textNode(child);
    }
  }
  try {
    draw(slide);
    const name = slide.getAttribute("aria-label") || "Slide";
    return ['<?xml version="1.0" encoding="UTF-8"?>',
      `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" version="1.1" width="${page.width}" height="${page.height}" viewBox="0 0 ${page.width} ${page.height}" xml:space="preserve">`,
      `<title>${xml(name)}</title>`,
      ...layers, "</svg>"].join("\n");
  } finally {
    fixture.remove();
  }
}

async function exportLocalSlides() {
  if (location.protocol !== "file:" && !["localhost", "127.0.0.1", "[::1]"].includes(location.hostname)) {
    throw new Error("Run this snippet on a local presentation.");
  }
  const slides = [...document.querySelectorAll("[data-slide]")];
  if (!slides.length) throw new Error("Mark the slide containers with data-slide.");

  const stage = document.createElement("div");
  stage.style.cssText = "position:fixed;left:-20000px;top:0;width:1280px;height:720px;pointer-events:none";
  document.body.append(stage);
  const files = [];
  try {
    await document.fonts.ready;
    for (const [index, original] of slides.entries()) {
      const copy = original.cloneNode(true);
      const styles = {
        display: original.dataset.exportDisplay || "flex",
        position: "relative", inset: "auto", width: "1280px", height: "720px",
        "box-sizing": "border-box", transform: "none", visibility: "visible",
      };
      for (const [key, value] of Object.entries(styles)) copy.style.setProperty(key, value, "important");
      stage.replaceChildren(copy);
      const figures = [...copy.querySelectorAll("img")].filter(img => !img.closest("[data-export-ignore]"));
      await Promise.all(figures.map(img => img.decode()));
      await document.fonts.ready;
      const geometry = copy.getBoundingClientRect();
      if (geometry.width !== 1280 || geometry.height !== 720) throw new Error("Unexpected slide dimensions.");
      files.push({
        name: String(index + 1).padStart(Math.max(2, String(slides.length).length), "0") + ".svg",
        contents: serializeSlide(copy),
      });
    }
  } finally {
    stage.remove();
  }

  document.querySelector("[data-bob-ppt-downloads] button")?.click();
  const panel = document.createElement("aside");
  panel.setAttribute("data-bob-ppt-downloads", "");
  panel.style.cssText = "position:fixed;inset:auto 1rem 1rem auto;z-index:2147483647;max-width:85vw;max-height:65vh;overflow:auto;background:#fff;color:#15281e;padding:.8rem;border:2px solid #164a37;font:16px system-ui";
  const close = document.createElement("button");
  close.textContent = "关闭下载面板";
  const links = files.map(file => {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([file.contents], { type: "image/svg+xml;charset=utf-8" }));
    link.download = file.name;
    link.textContent = "下载 " + file.name;
    link.style.display = "block";
    return link;
  });
  close.addEventListener("click", () => {
    for (const link of links) URL.revokeObjectURL(link.href);
    panel.remove();
  });
  panel.append(close, ...links);
  document.body.append(panel);
}

exportLocalSlides().catch(error => console.error("SVG export:", error.message));
