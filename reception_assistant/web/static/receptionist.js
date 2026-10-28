/* Pixel-art receptionist. Drawn entirely in code on an 80x80 canvas (scaled up with crisp pixels).
   API: const r = Receptionist.mount(canvas); r.setState("idle|listening|thinking|speaking");
        r.setName("Aria"); r.ding();                                                            */
(function (global) {
"use strict";
const W = 80, H = 80;
const K = {
  skin:"#f6cfae", skinShade:"#e5ab8c", skinDark:"#cf9479", blush:"#f2909b",
  lip:"#c9434f", lipDark:"#9d2d3c", mouthIn:"#4f1622", tongue:"#e0707c",
  hair:"#3a2218", hairMid:"#5f3a24", hairHi:"#9a6640", hairDark:"#25140d",
  lash:"#1b1220", white:"#fbfaf7", iris:"#2f9e8f", irisDark:"#1b655e", irisLight:"#86e3d2", pupil:"#0d1b22",
  blazer:"#8e2f4c", blazerDark:"#6c2038", blazerHi:"#ab4466", blouse:"#f5f1ec", blouseSh:"#d6cec6",
  gold:"#e4b84f", goldDark:"#a67c24",
  head:"#2a2d36", headHi:"#555b6b", headDark:"#101116",
  deskTop:"#c08050", deskHi:"#e0aa78", deskFront:"#7a4a2c", deskDark:"#5a3620",
  ok:"#58e08c", busy:"#ffb23e", teal:"#6fe6d8",
};

/* tiny 3x5 pixel font for the nameplate */
const FONT = {A:"010101111101101",B:"110101110101110",C:"011100100100011",D:"110101101101110",E:"111100110100111",
F:"111100110100100",G:"011100101101011",H:"101101111101101",I:"111010010010111",J:"001001001101010",
K:"101101110101101",L:"100100100100111",M:"101111111101101",N:"110101101101101",O:"010101101101010",
P:"110101110100100",Q:"010101101111011",R:"111101111110101",S:"011100010001110",T:"111010010010010",
U:"101101101101111",V:"101101101101010",W:"101101111111101",X:"101101010101101",Y:"101101010010010",Z:"111001010100111"};

function painter(g) {
  const px = (x, y, c) => { g.fillStyle = c; g.fillRect(Math.round(x), Math.round(y), 1, 1); };
  return {
    px,
    rect(x, y, w, h, c) { g.fillStyle = c; g.fillRect(x, y, w, h); },
    disc(cx, cy, r, c) {
      for (let y = -r; y <= r; y++) for (let x = -r; x <= r; x++)
        if (x * x + y * y <= r * r + 0.6) px(cx + x, cy + y, c);
    },
    ellipse(cx, cy, rx, ry, c) {
      for (let y = Math.floor(cy - ry); y <= Math.ceil(cy + ry); y++)
        for (let x = Math.floor(cx - rx); x <= Math.ceil(cx + rx); x++) {
          const dx = (x - cx) / rx, dy = (y - cy) / ry;
          if (dx * dx + dy * dy <= 1) px(x, y, c);
        }
    },
    line(x0, y0, x1, y1, c) {
      x0 = Math.round(x0); y0 = Math.round(y0); x1 = Math.round(x1); y1 = Math.round(y1);
      const dx = Math.abs(x1 - x0), dy = -Math.abs(y1 - y0), sx = x0 < x1 ? 1 : -1, sy = y0 < y1 ? 1 : -1;
      let err = dx + dy;
      for (;;) {
        px(x0, y0, c);
        if (x0 === x1 && y0 === y1) break;
        const e2 = 2 * err;
        if (e2 >= dy) { err += dy; x0 += sx; }
        if (e2 <= dx) { err += dx; y0 += sy; }
      }
    },
    text(str, x, y, c) {
      for (const ch of str) {
        const gl = FONT[ch];
        if (!gl) { x += 2; continue; }
        for (let i = 0; i < 15; i++) if (gl[i] === "1") px(x + (i % 3), y + Math.floor(i / 3), c);
        x += 4;
      }
    },
  };
}

const hash = n => Math.abs(Math.sin(n * 12.9898) * 43758.5453) % 1;

/* ---------- static scene (wall, window, plant, lamp, desk) ---------- */
function buildBackground(g) {
  const p = painter(g);
  for (let y = 0; y < 67; y++) for (let x = 0; x < W; x++) {
    const v = y + ((x + y) & 1) * 3;
    p.px(x, y, v < 20 ? "#2a505e" : v < 42 ? "#22434f" : "#1b3640");
  }
  // window
  p.rect(4, 5, 24, 26, "#0e1d24");
  const sky = ["#5eb6d0", "#86cdd8", "#b6dcc8", "#efd49a", "#f5b97a"];
  for (let y = 7; y < 29; y++) for (let x = 6; x < 26; x++) {
    const t = (y - 7) / 22 * sky.length + (((x + y) & 1) ? 0.25 : -0.25);
    p.px(x, y, sky[Math.max(0, Math.min(sky.length - 1, Math.floor(t)))]);
  }
  const heights = [6, 9, 5, 11, 7, 10, 6, 8, 5, 9];
  let bx = 6;
  heights.forEach((h, i) => {
    const w = 2 + (i % 2);
    p.rect(bx, 28 - h, Math.min(w, 26 - bx), h, "#1a3a46");
    for (let k = 0; k < h; k += 2) if (hash(i * 7 + k) > 0.7 && bx + 1 < 26) p.px(bx + 1, 26 - k, "#f4c86a");
    bx += w;
  });
  p.rect(15, 7, 1, 22, "#0e1d24"); p.rect(6, 18, 20, 1, "#0e1d24");
  p.rect(2, 30, 28, 2, "#3b6b78"); p.rect(2, 30, 28, 1, "#5d93a0");
  // plant
  p.rect(70, 60, 7, 6, "#b5603f"); p.rect(69, 59, 9, 2, "#cf7a52"); p.rect(70, 65, 7, 1, "#8d4a30");
  const leaves = [[73, 52, 2, 6, "#2f8a5a"], [69, 54, 2, 4, "#3aa06a"], [77, 54, 2, 4, "#2a7a50"],
                  [73, 47, 2, 4, "#4cb27a"], [71, 50, 1.5, 4, "#41a972"], [75, 50, 1.5, 4, "#35966a"]];
  leaves.forEach(([x, y, rx, ry, c]) => p.ellipse(x, y, rx, ry, c));
  // lamp (left)
  p.rect(3, 65, 8, 2, "#8b93a6"); p.rect(3, 65, 8, 1, "#b4bccd");
  p.line(6, 65, 6, 57, "#8b93a6"); p.line(7, 65, 7, 57, "#69708a"); p.line(6, 57, 10, 52, "#8b93a6"); p.line(7, 57, 11, 52, "#69708a");
  p.rect(8, 49, 9, 2, "#e0a238"); p.rect(9, 51, 7, 1, "#b87a1c"); p.rect(8, 49, 9, 1, "#f2c066");
  // desk
  p.rect(0, 67, W, 1, K.deskHi); p.rect(0, 68, W, 2, K.deskTop);
  p.rect(0, 70, W, 10, K.deskFront);
  for (let y = 70; y < 80; y++) { p.px(18, y, K.deskDark); p.px(62, y, K.deskDark); }
  p.rect(0, 79, W, 1, K.deskDark);
}

/* ---------- face geometry ---------- */
const FRX = 10.5, FRY = 12.5;
function faceHalfW(y, cy) {
  const fcy = cy + 1, fy = (y - fcy) / FRY;
  if (Math.abs(fy) > 1) return -1;
  let w = FRX * Math.sqrt(1 - fy * fy);
  if (y > fcy + 3) w *= 1 - 0.22 * ((y - (fcy + 3)) / (FRY - 3));
  return w;
}
function hairColor(x, y, cx, cy) {
  const l = -(x - cx) * 0.55 - (y - cy) * 0.75 + (((x + y) & 1) * 1.4 - 0.7);
  if (l > 8) return K.hairHi;
  if (l > 2.5) return K.hairMid;
  if (l < -10) return K.hairDark;
  return K.hair;
}
function fringeBottom(d, cy) {
  return d < -3 ? cy - 7 + (-3 - d) * 0.4 : cy - 7 + (d + 3) * 0.26;
}

/* ---------- mount ---------- */
function mount(canvas, opts) {
  opts = opts || {};
  canvas.width = W; canvas.height = H;
  const g = canvas.getContext("2d");
  g.imageSmoothingEnabled = false;
  const make = opts.makeCanvas || ((w, h) => { const c = document.createElement("canvas"); c.width = w; c.height = h; return c; });
  const bgCanvas = make(W, H);
  buildBackground(bgCanvas.getContext("2d"));
  const p = painter(g);

  const S = { state: "idle", name: "ARIA", dingAt: -10, nextBlink: 2, blinkAt: -10 };

  function blinkLevel(t) {          // 3 open, 2 half, 0 closed
    if (t >= S.nextBlink) { S.blinkAt = t; S.nextBlink = t + 2.4 + hash(t) * 3.2; }
    const e = t - S.blinkAt;
    if (e < 0.07) return 2; if (e < 0.14) return 0; if (e < 0.21) return 2; return 3;
  }

  function drawEye(ex, ey, side, open, look, sleeping) {
    if (sleeping) {
      [[-2, 0], [-1, 1], [0, 1], [1, 1], [2, 0]].forEach(([dx, dy]) => p.px(ex + dx, ey + dy, K.lash));
      p.px(ex + side * 3, ey - 1, K.lash);
      return;
    }
    if (open === 0) {
      for (let dx = -2; dx <= 2; dx++) p.px(ex + dx, ey, K.lash);
      p.px(ex + side * 3, ey - 1, K.lash);
      return;
    }
    const rows = open === 3 ? [[-2, 3], [-1, 5], [0, 5], [1, 5], [2, 3]] : [[-1, 5], [0, 5], [1, 3]];
    const lashY = ey + rows[0][0] - 1, set = new Set();
    rows.forEach(([dy, w]) => { for (let dx = -(w >> 1); dx <= (w >> 1); dx++) { p.px(ex + dx, ey + dy, K.white); set.add((ex + dx) + "," + (ey + dy)); } });
    const ix = ex + look.x, iy = ey + look.y;
    for (let r = -2; r <= 1; r++) for (let dx = -1; dx <= 1; dx++) {
      const x = ix + dx, y = iy + r;
      if (!set.has(x + "," + y)) continue;
      p.px(x, y, r === -2 ? K.irisDark : r === 1 ? K.irisLight : K.iris);
    }
    [[0, -1], [0, 0]].forEach(([dx, dy]) => { const x = ix + dx, y = iy + dy; if (set.has(x + "," + y)) p.px(x, y, K.pupil); });
    if (set.has((ix - 1) + "," + (iy - 1))) p.px(ix - 1, iy - 1, K.white);
    if (set.has((ix + 1) + "," + (iy + 1))) p.px(ix + 1, iy + 1, "#d9fff6");
    for (let dx = -2; dx <= 2; dx++) p.px(ex + dx, lashY, K.lash);
    p.px(ex + side * 3, lashY, K.lash); p.px(ex + side * 3, lashY - 1, K.lash);
  }

  function drawMouth(cx, my, shape, state) {
    if (shape === "closed") {
      p.px(cx - 3, my - 1, K.skinShade); p.px(cx + 3, my - 1, K.skinShade);
      for (let dx = -2; dx <= 2; dx++) p.px(cx + dx, my, Math.abs(dx) === 2 ? K.lipDark : K.lip);
      p.px(cx - 1, my + 1, K.lipDark); p.px(cx, my + 1, K.lip); p.px(cx + 1, my + 1, K.lipDark);
      return;
    }
    if (shape === "smirk") {
      p.px(cx + 4, my - 1, K.lipDark);
      for (let dx = -1; dx <= 3; dx++) p.px(cx + dx, my, K.lip);
      return;
    }
    const [w, h] = shape;
    for (let dx = -w; dx <= w; dx++) p.px(cx + dx, my, Math.abs(dx) === w ? K.lipDark : K.lip);
    for (let r = 1; r <= h; r++) for (let dx = -w + 1; dx <= w - 1; dx++)
      p.px(cx + dx, my + r, (r === h && h >= 3 && Math.abs(dx) < w - 1) ? K.tongue : K.mouthIn);
    for (let dx = -w + 1; dx <= w - 1; dx++) p.px(cx + dx, my + h + 1, K.lip);
    p.px(cx - w + 1, my + 1 , K.white); p.px(cx + w - 1, my + 1, K.white);
  }

  function drawBody(cx) {
    for (let y = 51; y <= 66; y++) {
      const hw = Math.min(27, 12 + (y - 51) * 3.4), x0 = Math.ceil(cx - hw), x1 = Math.floor(cx + hw);
      for (let x = x0; x <= x1; x++) {
        let c = x > cx + 14 ? K.blazerDark : (y <= 52 ? K.blazerHi : K.blazer);
        if (x === x0 || x === x1) c = K.blazerDark;
        p.px(x, y, c);
      }
    }
    p.line(cx - 19, 57, cx - 19, 66, K.blazerDark); p.line(cx + 19, 57, cx + 19, 66, K.blazerDark);
    for (let y = 51; y <= 60; y++) {
      const hw = 6 - (y - 51) * 0.62;
      for (let x = Math.ceil(cx - hw); x <= Math.floor(cx + hw); x++) p.px(x, y, (x === Math.ceil(cx - hw) || x === Math.floor(cx + hw)) ? K.blouseSh : K.blouse);
    }
    for (let y = 51; y <= 54; y++) {
      const hw = 3.6 - (y - 51) * 0.9;
      for (let x = Math.ceil(cx - hw); x <= Math.floor(cx + hw); x++) p.px(x, y, K.skin);
    }
    p.line(cx - 7, 51, cx - 1, 62, K.blazerDark); p.line(cx + 7, 51, cx + 1, 62, K.blazerDark);
    p.px(cx, 56, K.gold); p.px(cx, 57, K.goldDark);
    p.px(cx, 63, K.gold);
  }

  function frame(t) {
    const st = S.state, idle = st === "idle";
    g.drawImage(bgCanvas, 0, 0);

    // clock
    const d = new Date(), cl = [68, 12];
    p.disc(cl[0], cl[1], 6, "#d9c9a0"); p.disc(cl[0], cl[1], 5, "#f7f1e3");
    [[0, -4], [4, 0], [0, 4], [-4, 0]].forEach(([dx, dy]) => p.px(cl[0] + dx, cl[1] + dy, "#5a4a30"));
    const ma = (d.getMinutes() + d.getSeconds() / 60) / 60 * Math.PI * 2, ha = ((d.getHours() % 12) + d.getMinutes() / 60) / 12 * Math.PI * 2;
    p.line(cl[0], cl[1], cl[0] + 4.4 * Math.sin(ma), cl[1] - 4.4 * Math.cos(ma), "#2a2d36");
    p.line(cl[0], cl[1], cl[0] + 2.6 * Math.sin(ha), cl[1] - 2.6 * Math.cos(ha), "#8e2f4c");

    // lamp light (on while on duty)
    if (!idle) {
      p.rect(11, 52, 3, 1, "#fff3b0"); p.px(12, 53, "#ffe08a");
      for (let y = 54; y < 67; y++) {
        const half = 1 + (y - 54) * 0.95;
        for (let x = Math.round(11 - half * 0.4); x <= Math.round(11 + half * 1.5); x++)
          if (((x + y) & 1) === 0) { g.globalAlpha = 0.3; p.px(x, y, "#ffd27a"); }
      }
      g.globalAlpha = 1;
    }

    // pose
    let ox = 0, oy = 0;
    if (idle) { ox = -1; oy = 1; }
    else if (st === "speaking") oy = Math.sin(t * 5) > 0.45 ? 1 : 0;
    else if (st === "thinking") { ox = 1; oy = Math.sin(t * 1.3) > 0.55 ? 1 : 0; }
    else oy = Math.sin(t * 1.3) > 0.55 ? 1 : 0;
    const cx = 40 + ox, cy = 34 + oy;

    drawBody(40);

    // back hair (bob)
    const swayAmp = idle ? 0 : st === "speaking" ? 1 : 0.6;
    for (let y = cy - 18; y <= cy + 16; y++) for (let x = cx - 17; x <= cx + 17; x++) {
      let inside;
      if (y <= cy + 4) { const a = (x - cx) / 14.8, b = (y - (cy - 1)) / 16.5; inside = a * a + b * b <= 1; }
      else {
        const t2 = (y - (cy + 4)) / 12, hw = 14.3 - t2 * t2 * 3.2;
        const sway = Math.round(Math.sin(t * 1.8 + y * 0.18) * ((y - (cy + 6)) / 10) * swayAmp);
        inside = Math.abs(x - cx - sway) <= hw;
      }
      if (inside) p.px(x, y, hairColor(x, y, cx, cy));
    }

    // neck
    for (let y = cy + 12; y <= 51; y++) for (let dx = -3; dx <= 3; dx++) {
      let c = dx >= 2 ? K.skinShade : K.skin;
      if (y <= cy + 14) c = K.skinDark;
      p.px(cx + dx, y, c);
    }

    // face
    const fringeEdge = {};
    for (let y = cy - 12; y <= cy + 14; y++) {
      const w = faceHalfW(y, cy); if (w < 0) continue;
      for (let x = Math.ceil(cx - w); x <= Math.floor(cx + w); x++) {
        const dd = x - cx, fb = fringeBottom(dd, cy);
        let c = K.skin;
        if (y < fb) c = hairColor(x, y, cx, cy);
        else if (y < fb + 1) c = K.skinShade;
        else if (dd >= w - 1.6) c = K.skinShade;
        else if (y >= cy + 13) c = K.skinShade;
        if (y >= cy - 6 && y <= cy + 3 && Math.abs(dd) >= w - 1.4 && y >= fb) c = hairColor(x, y, cx, cy); // framing locks
        p.px(x, y, c);
      }
    }
    // cheeks + nose
    [[-8, 6], [-7, 6], [7, 6], [8, 6]].forEach(([dx, dy]) => { g.globalAlpha = 0.75; p.px(cx + dx, cy + dy + 1, K.blush); });
    g.globalAlpha = 1;
    p.px(cx, cy + 6, K.skinDark); p.px(cx - 1, cy + 7, K.skinShade);

    // eyes & brows
    const blink = blinkLevel(t);
    const ey = cy + 2, look = { x: 0, y: 0 };
    let browL = cy - 4, browR = cy - 4;
    if (st === "listening") { look.x = Math.round(Math.sin(t * 0.8) * 0.9); browL = browR = cy - 5; }
    else if (st === "thinking") { look.x = 1; look.y = -1; browL = cy - 5; browR = cy - 3; }
    else if (st === "speaking") { look.x = Math.round(Math.sin(t * 1.1) * 0.6); }
    else if (idle) { browL = browR = cy - 3; }
    const speakIdx = Math.floor(t * 9);
    const pick = Math.floor(hash(speakIdx) * 5);
    const shapes = [[2, 0], [2, 1], [3, 2], [3, 3], [2, 2]];
    if (st === "speaking" && pick >= 3) { browL = browR = cy - 5; }
    drawEye(cx - 6, ey, -1, idle ? 0 : blink, look, idle);
    drawEye(cx + 6, ey, 1, idle ? 0 : blink, look, idle);
    [[cx - 6, browL, -1], [cx + 6, browR, 1]].forEach(([bx, by, s]) => {
      p.px(bx - 1, by + (st === "thinking" && s === 1 ? 0 : 0), K.hair);
      p.px(bx, by - 0, K.hair); p.px(bx + 1, by + (s === -1 ? 0 : 0), K.hair);
      p.px(bx - s * 2, by + 1, K.hairMid);
    });

    // mouth
    let mouth = "closed";
    if (st === "speaking") mouth = shapes[pick];
    else if (st === "thinking") mouth = "smirk";
    drawMouth(cx, cy + 9, mouth, st);

    // headset
    for (let a = 12; a <= 168; a += 3) {
      const r = a * Math.PI / 180;
      p.px(cx + 14.2 * Math.cos(r), cy - 2 - 15.8 * Math.sin(r), K.head);
    }
    for (let a = 30; a <= 150; a += 4) { const r = a * Math.PI / 180; p.px(cx + 13.4 * Math.cos(r), cy - 2 - 15 * Math.sin(r), K.headHi); }
    [-1, 1].forEach(s => {
      p.line(cx + s * 14, cy - 4, cx + s * 14, cy + 1, K.head);
      p.rect(s < 0 ? cx - 17 : cx + 14, cy + 1, 3, 7, K.head);
      p.rect(s < 0 ? cx - 16 : cx + 15, cy + 2, 1, 5, K.headHi);
    });
    p.line(cx - 15, cy + 8, cx - 15, cy + 11, K.head);
    p.line(cx - 15, cy + 11, cx - 12, cy + 14, K.head);
    p.line(cx - 12, cy + 14, cx - 8, cy + 14, K.head);
    p.line(cx - 8, cy + 14, cx - 5, cy + 11, K.head);
    p.rect(cx - 6, cy + 10, 2, 2, K.headDark); p.px(cx - 6, cy + 10, K.headHi);
    const blinkOn = Math.floor(t * 3) % 2 === 0;
    const led = idle ? "#7a3030" : st === "listening" ? (blinkOn ? K.ok : "#2a7a4a") : st === "speaking" ? K.busy : st === "thinking" ? (blinkOn ? K.busy : "#7a5a20") : K.ok;
    p.px(cx - 16, cy + 7, led);

    // desk objects
    const dingE = t - S.dingAt, dingging = dingE < 0.5, by = dingging ? -1 : 0;
    p.rect(60, 65, 8, 2, K.goldDark);
    p.rect(61, 63 + by, 6, 2, K.gold); p.rect(62, 62 + by, 4, 1, K.gold); p.px(64, 61 + by, K.goldDark);
    p.px(62, 63 + by, "#fff0b8");
    if (dingging && Math.floor(dingE * 14) % 2 === 0) { p.px(59, 60, "#fff0b8"); p.px(69, 60, "#fff0b8"); p.px(64, 58, "#fff0b8"); }
    // nameplate
    p.rect(26, 71, 28, 7, K.gold); p.rect(27, 72, 26, 5, "#2a1d17");
    const nm = S.name.slice(0, 6), wpx = nm.length * 4 - 1;
    p.text(nm, 40 - Math.floor(wpx / 2), 73, K.gold);

    // state props
    if (idle) {
      g.fillStyle = "rgba(8,18,32,0.42)"; g.fillRect(0, 0, W, H);
      for (let k = 0; k < 2; k++) {
        const ph = ((t * 0.5) + k * 0.5) % 1, zx = cx + 13 + ph * 6, zy = cy - 8 - ph * 10 - k * 2, s = k ? 2 : 3;
        g.globalAlpha = 1 - ph * 0.9;
        p.text("Z", Math.round(zx), Math.round(zy), "#e8f1ff");
        g.globalAlpha = 1;
      }
    }
    if (st === "thinking") {
      p.px(cx + 13, cy - 9, K.white); p.disc(cx + 16, cy - 12, 1, K.white);
      p.rect(cx + 17, cy - 24, 13, 8, K.white); p.rect(cx + 18, cy - 25, 11, 10, K.white);
      for (let i = 0; i < 3; i++) {
        const on = Math.floor(t * 3) % 4 > i;
        p.rect(cx + 20 + i * 3, cy - 21 + (on ? -1 : 0), 2, 2, on ? K.head : "#b7bcc7");
      }
    }
    if (st === "listening") {
      [[-1, 4], [-1, 7], [1, 4], [1, 7]].forEach(([s, r], i) => {
        const pulse = (Math.sin(t * 4 - (r - 4) * 0.5) + 1) / 2;
        g.globalAlpha = 0.25 + pulse * 0.65;
        for (let a = -45; a <= 45; a += 22) {
          const rr = a * Math.PI / 180;
          p.px(cx + s * (16 + r * Math.cos(rr)), cy + 4 + r * Math.sin(rr), K.teal);
        }
      });
      g.globalAlpha = 1;
    }
  }

  let last = 0, raf = 0;
  function loop(ts) {
    raf = requestAnimationFrame(loop);
    if (document.hidden || ts - last < 66) return;      // ~15 fps keeps the pixel feel
    last = ts; frame(ts / 1000);
  }
  if (!opts.manual) raf = requestAnimationFrame(loop);

  return {
    setState(s) { S.state = s; },
    setName(n) { S.name = String(n || "").toUpperCase().replace(/[^A-Z ]/g, ""); },
    ding() { S.dingAt = performance.now() / 1000; },
    frame,                                              // exposed for tests
    stop() { cancelAnimationFrame(raf); },
    _S: S,
  };
}

const api = { mount };
if (typeof module !== "undefined") module.exports = api; else global.Receptionist = api;
})(typeof window !== "undefined" ? window : globalThis);
