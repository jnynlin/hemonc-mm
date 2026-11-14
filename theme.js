/* hemonc-mm theme.js — 主題切換（跟隨系統 → 深色 → 淺色 → 跟隨系統）
   放在 <head> 的樣式表之後、不加 defer：第一次繪製前就設好 data-theme，不閃爍。
   - 跟隨系統時不寫 localStorage，作業系統換主題會即時跟著變（matchMedia change）
   - 另一個分頁切了主題，這頁同步（storage 事件）
   - 按鈕由這支程式自己長出來：儀表板放進 #nav，其餘頁放 h1 右側；頁面不用加任何標記
   - 只用於 github.io／本機檔；claude.ai artifact 版由 viewer 自己管 data-theme，不要載入這支
   出處：Antigravity 初稿（2026-11-14）→ Claude 改為三態、即時跟隨、跨分頁、無 emoji */
(() => {
  const KEY = 'hemonc-mm-theme', root = document.documentElement, mq = matchMedia('(prefers-color-scheme: dark)');
  const NAME = { auto: '跟隨系統', light: '淺色', dark: '深色' }, NEXT = { auto: 'dark', dark: 'light', light: 'auto' };
  let btn = null;
  const pref = () => { try { const v = localStorage.getItem(KEY); if (v === 'dark' || v === 'light') return v; } catch (_) {} return 'auto'; };
  const apply = () => {
    const p = pref(), t = p === 'auto' ? (mq.matches ? 'dark' : 'light') : p;
    root.dataset.theme = t;
    if (btn) { btn.textContent = '主題：' + NAME[p]; btn.title = btn.ariaLabel = `目前${NAME[p]}，按一下換成${NAME[NEXT[p]]}`; btn.setAttribute('aria-label', btn.title); }
  };
  apply();
  mq.addEventListener('change', apply);
  addEventListener('storage', e => { if (e.key === KEY) apply(); });
  addEventListener('DOMContentLoaded', () => {
    const nav = document.getElementById('nav'), h1 = document.querySelector('h1');
    const host = nav || h1; if (!host) return;
    const st = document.createElement('style');
    st.textContent = '.theme-toggle{font:inherit;font-size:.8rem;line-height:1.4;background:none;border:1px solid currentColor;border-radius:6px;padding:2px 9px;color:inherit;opacity:.6;cursor:pointer;margin-left:10px;white-space:nowrap}.theme-toggle:hover,.theme-toggle:focus-visible{opacity:1}h1 .theme-toggle{float:right;font-weight:400;margin-top:.35em}@media print{.theme-toggle{display:none}}';
    document.head.appendChild(st);
    btn = document.createElement('button'); btn.type = 'button'; btn.id = 'theme-toggle'; btn.className = 'theme-toggle';
    btn.addEventListener('click', () => { const n = NEXT[pref()]; try { n === 'auto' ? localStorage.removeItem(KEY) : localStorage.setItem(KEY, n); } catch (_) {} apply(); });
    host.appendChild(btn); apply();
  });
})();
