/* ============================================================
   theme.js — Advance Plus Shared Theme Engine
   โหลดในทุกหน้า: <script src="theme.js"></script>
   ต้องโหลดก่อน Firebase scripts
   ============================================================ */

const THEME = {
  /* ── defaults ── */
  defaults: {
    mode:      'light',
    bg1:       '#bbf7d0',
    bg2:       '#f0fdf4',
    bgMid:     '#dcfce7',
    bgAngle:   '150',
    accent:    '#16a34a',
    cardStyle: 'solid',   // solid | glass | flat
    fontSize:  '14',
    radius:    '10',
  },

  get(k)  { return localStorage.getItem('adv_' + k) ?? this.defaults[k]; },
  set(k,v){ localStorage.setItem('adv_' + k, v); },

  /* ── Apply all settings to current page ── */
  apply() {
    const mode      = this.get('mode');
    const bg1       = this.get('bg1');
    const bg2       = this.get('bg2');
    const bgMid     = this.get('bgMid');
    const bgAngle   = this.get('bgAngle');
    const accent    = this.get('accent');
    const cardStyle = this.get('cardStyle');
    const fontSize  = this.get('fontSize');
    const radius    = this.get('radius');

    const root = document.documentElement;
    root.setAttribute('data-theme', mode);
    root.setAttribute('data-card', cardStyle);

    // Background gradient (overrides CSS var)
    document.body.style.background =
      `linear-gradient(${bgAngle}deg, ${bg1} 0%, ${bgMid} 40%, ${bg2} 100%)`;
    document.body.style.backgroundAttachment = 'fixed';

    // Accent CSS vars
    root.style.setProperty('--accent',    accent);
    root.style.setProperty('--accent-dk', this._shade(accent, -25));
    root.style.setProperty('--accent-lt', this._shade(accent, 30));
    root.style.setProperty('--accent-xlt', this._shade(accent, 70));
    root.style.setProperty('--g-md',  accent);
    root.style.setProperty('--g-dk',  this._shade(accent, -25));
    root.style.setProperty('--g-lt',  this._shade(accent, 30));
    root.style.setProperty('--border', this._shade(accent, 55) + 'b0');

    // Dark mode accent adjustments
    if (mode === 'dark') {
      document.body.style.background =
        `linear-gradient(${bgAngle}deg, ${this._shade(bg1,-60)} 0%, ${this._shade(bgMid,-55)} 40%, ${this._shade(bg2,-65)} 100%)`;
    }

    // Font size & radius on inputs (global)
    const styleId = 'adv-theme-overrides';
    let styleEl = document.getElementById(styleId);
    if (!styleEl) { styleEl = document.createElement('style'); styleEl.id = styleId; document.head.appendChild(styleEl); }
    styleEl.textContent = `
      input[type=text],input[type=tel],input[type=email],input[type=password],
      input[type=url],textarea,select {
        font-size: ${fontSize}px !important;
        border-radius: ${radius}px !important;
      }
      [data-card="glass"] .card, [data-card="glass"] .sec-card,
      [data-card="glass"] .hero, [data-card="glass"] .stat-card {
        background: rgba(255,255,255,0.6) !important;
        backdrop-filter: blur(16px) saturate(180%) !important;
        border: 1px solid rgba(255,255,255,0.4) !important;
      }
      [data-theme="dark"][data-card="glass"] .card,
      [data-theme="dark"][data-card="glass"] .sec-card,
      [data-theme="dark"][data-card="glass"] .hero,
      [data-theme="dark"][data-card="glass"] .stat-card {
        background: rgba(30,60,40,0.55) !important;
        backdrop-filter: blur(16px) !important;
        border: 1px solid rgba(110,231,183,0.25) !important;
      }
      [data-card="flat"] .card, [data-card="flat"] .sec-card,
      [data-card="flat"] .hero, [data-card="flat"] .stat-card {
        box-shadow: none !important;
        border: 2px solid var(--sec-div) !important;
      }
    `;

    // Update theme toggle button if present
    const fab = document.getElementById('themeFab') || document.getElementById('themeBtn');
    if (fab) fab.textContent = mode === 'dark' ? '☀️' : '🌙';
  },

  toggle() {
    const next = this.get('mode') === 'dark' ? 'light' : 'dark';
    this.set('mode', next);
    this.apply();
  },

  _shade(hex, pct) {
    hex = hex.replace('#','');
    if (hex.length === 3) hex = hex.split('').map(c=>c+c).join('');
    const num = parseInt(hex, 16);
    const r = Math.min(255, Math.max(0, ((num>>16)&0xff) + pct));
    const g = Math.min(255, Math.max(0, ((num>>8)&0xff)  + pct));
    const b = Math.min(255, Math.max(0, (num&0xff)        + pct));
    return '#' + [r,g,b].map(x => x.toString(16).padStart(2,'0')).join('');
  },
};

/* Auto-apply on script load (before DOMContentLoaded for no-flash) */
(function() {
  const root = document.documentElement;
  root.setAttribute('data-theme', THEME.get('mode'));
})();

/* Full apply after DOM ready */
document.addEventListener('DOMContentLoaded', () => THEME.apply());