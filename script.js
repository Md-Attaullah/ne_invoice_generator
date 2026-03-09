(function(){
  /* ----------------------------- 1) THEME ----------------------------- */
  const THEME_KEY='nusrat.theme';
  const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  function applyTheme(t){ document.documentElement.setAttribute('data-theme', t); }
  function curTheme(){ return document.documentElement.getAttribute('data-theme') || 'light'; }
  function updateThemeBtn(){
    const b=document.getElementById('themeToggle'); if(!b) return;
    const isDark = curTheme()==='dark';
    b.textContent = isDark ? '☀️ Light Mode' : '🌙 Dark Mode';
  }
  (function initTheme(){
    const saved = localStorage.getItem(THEME_KEY);
    applyTheme(saved || (prefersDark ? 'dark' : 'light'));
    updateThemeBtn();
  })();
  const themeToggleBtn = document.getElementById('themeToggle');
  if (themeToggleBtn){
    themeToggleBtn.addEventListener('click', ()=>{
      const next = curTheme()==='dark' ? 'light' : 'dark';
      applyTheme(next); localStorage.setItem(THEME_KEY, next); updateThemeBtn();
    });
  }

  /* ----------------------------- 2) CONFIG ----------------------------- */
  const LINKS = {
    instagram: "https://www.instagram.com/nusrat_enterprises/",
    waChannel: "https://whatsapp.com/channel/0029VaTqsgN5Ui2gpzD62D06",
    googleFeedback: "https://search.google.com/local/writereview?placeid=ChIJebZSWdmnGToRWa9S2iTwlgo"
  };
  const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwDh3hVYOpAOa12sOPpfpC0IORb6HswDBiAS-OMIZ9cu2GH4evGu5nld5_2HAC3ayOg/exec';
  const AUTH_TOKEN  = '9705322252';
  const ITEM_OPTIONS=["Shirt","T-Shirt","Jean","Cargo Pant","Formal Pant","Track","Half Pant","Under Wear","Vest"];
  const NAME_MAX = 15;

  /* ----------------------------- 3) EL REFS ----------------------------- */
  const els={
    itemsList:document.getElementById('itemsList'),
    quickline:document.getElementById('quickline'),
    quicklineForm:document.getElementById('quicklineForm'),
    fabAdd:document.getElementById('fabAdd'),
    invoiceNumber:document.getElementById('invoiceNumber'),
    invoiceDate:document.getElementById('invoiceDate'),
    customerName:document.getElementById('customerName'),
    customerPhone:document.getElementById('customerPhone'),
    phoneError:document.getElementById('phoneError'),
    pickContactBtn:document.getElementById('pickContactBtn'),
    saleType:document.getElementById('saleType'),
    subtotal:document.getElementById('subtotal'),
    itemsBreakdown:document.getElementById('itemsBreakdown'),
    itemsCountQty:document.getElementById('itemsCountQty'),
    taxRate:document.getElementById('taxRate'),
    flatDiscount:document.getElementById('flatDiscount'),
    grandTotal:document.getElementById('grandTotal'),
    paymentMode:document.getElementById('paymentMode'), // hidden input that stores selected mode
    paidAmount:document.getElementById('paidAmount'),
    paidError:document.getElementById('paidError'),
    dueAmount:document.getElementById('dueAmount'),
    previewCard:document.getElementById('previewCard'),
    invoicePreview:document.getElementById('invoicePreview'),
    monoRender:document.getElementById('monoRender'),
    whatsAppBtn: document.getElementById('whatsAppBtn'),
    waInvoiceBtn: document.getElementById('waInvoiceBtn'),
    saveBtn: document.getElementById('saveToSheetBtn'),
    printBtn: document.getElementById('printBtn'),
    previewBtn: document.getElementById('previewBtn'),
    copyBtn: document.getElementById('copyTextBtn'),
    clearItemsBtn: document.getElementById('clearItemsBtn'),
    vcfBtn: document.getElementById('vcfBtn'),
    /* Stats elements */
    previewCustomerStats: document.getElementById('previewCustomerStats'),
    inlineCustomerStats: document.getElementById('inlineCustomerStats') // optional; if not present, we fall back to previewCustomerStats
  };

  /* ----------------------------- 3a) COUPON TOGGLE (FIX) ----------------------------- */
  const couponToggle = document.getElementById('couponToggle');
  function couponIsOn() {
    return !!couponToggle && !!couponToggle.checked;
  }
  if (couponToggle) {
    couponToggle.addEventListener('change', () => {
      if (els.previewCard && els.previewCard.style.display !== 'none') {
        renderPreview();
      }
    });
  }
  // Unified coupon value rule:
  // >= 1000 → 50; >= 500 → 50; else 0
  function couponValueFor(grand) {
    if (grand >= 1000) return 50;
    if (grand >= 500) return 50;
    return 0;
  }

  /* ----------------------------- 4) SALE TYPE ----------------------------- */
  (function wireSaleTypeSegments(){
    const group = document.getElementById('saleTypeGroup');
    if (!group) return;
    const labels = Array.from(group.querySelectorAll('label.seg'));
    const radios  = Array.from(group.querySelectorAll('input[type="radio"][name="saleTypeRadio"]'));
    function setActiveByValue(val){
      labels.forEach(l => l.classList.remove('active'));
      const match = radios.find(r => r.value === val);
      if (match){
        const label = match.closest('label.seg');
        if (label) label.classList.add('active');
        const hidden = document.getElementById('saleType');
        if (hidden) hidden.value = val;
      }
    }
    radios.forEach(r => { r.checked = false; });
    els.saleType.value = '';
    labels.forEach(l => l.classList.remove('active'));
    radios.forEach(radio => {
      radio.addEventListener('change', () =>{
        if (radio.checked){ setActiveByValue(radio.value); persist(); }
      });
    });
  })();

  /* ----------------------------- 5) HELPERS ----------------------------- */
  const pad2=n=>String(n).padStart(2,'0');
  function genInvoice() {
    // e.g., NEddyyyymmhhmm
    const d = new Date();
    const dd = pad2(d.getDate());
       const mm = pad2(d.getMonth() + 1);
    const yyyy = d.getFullYear();
    const hh = pad2(d.getHours());
    const min = pad2(d.getMinutes());
    return `NE${dd}${yyyy}${mm}${hh}${min}`;
  }
  function toNumber(v){ const n=parseFloat(v); return Number.isFinite(n) ? n : 0; }
  function fit(text, len){ let t=String(text||'').replace(/\s+/g,' ').trim(); return t.length>len ? t.slice(0,len-1)+'…' : t.padEnd(len,' '); }
  function p2(n, width){ return String(n.toFixed(2)).padStart(width,' '); }
  const RECEIPT_MM = 107;
  const mmToPx = mm => Math.round(mm * 96 / 25.4);
  const mmToPt = mm => (mm * 72 / 25.4);
  function setBtnLoading(btn, loading, loadingText, normalText){
    if(!btn) return;
    const textEl = btn.querySelector('.btn-text');
    const spin = btn.querySelector('.spinner');
    if(loading){
      btn.classList.add('loading'); btn.disabled = true;
      if (textEl && loadingText) textEl.textContent = loadingText;
      if (spin) spin.style.display='inline-block';
    } else {
      btn.classList.remove('loading'); btn.disabled = false;
      if (textEl && normalText) textEl.textContent = normalText;
      if (spin) spin.style.display='none';
    }
  }

  // Date-only formatter (DD-MM-YYYY)
  function formatDateOnly(isoOrStr) {
    if (!isoOrStr) return '';
    const d = new Date(isoOrStr);
    if (isNaN(d)) return String(isoOrStr);
    const dd = pad2(d.getDate());
    const mm = pad2(d.getMonth() + 1);
    const yyyy = d.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  }

  /* ----------------------------- 6) INIT HEADER ----------------------------- */
  els.invoiceNumber.value = genInvoice();
  const today = new Date();
  els.invoiceDate.value = `${today.getFullYear()}-${pad2(today.getMonth()+1)}-${pad2(today.getDate())}`;

  /* ----------------------------- Payment Mode (Emoji Icons) ------------- */
  function initPaymentModeIcons(root = document) {
    const group = root.querySelector('.pm-buttons');
    const hidden = root.getElementById
      ? root.getElementById('paymentMode')
      : root.querySelector('#paymentMode');
    if (!group || !hidden) return;

    const buttons = Array.from(group.querySelectorAll('.pm-btn'));

    function setSelected(btn) {
      buttons.forEach(b => {
        const sel = b === btn;
        b.classList.toggle('selected', sel);
        b.setAttribute('aria-checked', sel ? 'true' : 'false');
        b.tabIndex = sel ? 0 : -1; // roving tabindex
      });
      hidden.value = btn.dataset.value || '';
      // Fire a change so existing listeners (recalc/persist) react
      hidden.dispatchEvent(new Event('change', { bubbles: true }));
    }

    // Click selection
    buttons.forEach(btn => btn.addEventListener('click', () => setSelected(btn)));

    // Keyboard navigation (Arrow keys, Enter/Space)
    group.addEventListener('keydown', (e) => {
      const active = document.activeElement;
      const i = buttons.indexOf(active);
      if (i === -1) return;

      if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
        e.preventDefault();
        const next = buttons[(i + 1) % buttons.length];
        next.focus();
        setSelected(next);
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
        e.preventDefault();
        const prev = buttons[(i - 1 + buttons.length) % buttons.length];
        prev.focus();
        setSelected(prev);
      } else if (e.key === ' ' || e.key === 'Enter') {
        e.preventDefault();
        setSelected(active);
      }
    });

    // Initialize focusability
    buttons.forEach((b, idx) => (b.tabIndex = idx === 0 ? 0 : -1));

    // Optional default selection:
    // const def = buttons.find(b => b.dataset.value === 'UPI');
    // if (def) setSelected(def);
  }
  initPaymentModeIcons();

  /* ----------------------------- 6a) VISIT STATS HELPERS ----------------------------- */
  const preferInlineEl = () => (els.inlineCustomerStats || els.previewCustomerStats || null);

  function showStatsLoading(el){
    if (!el) return;
    el.textContent = 'Checking visit history…';
    el.style.display = 'block';
  }
  function hideStats(el){
    if (!el) return;
    el.textContent = '';
    el.style.display = 'none';
  }
  function setStats(el, stats, phone10){
    if (!el) return;
    if (!/^\d{10}$/.test(phone10)){
      hideStats(el);
      return;
    }
    if (!stats){
      el.textContent = 'New customer';
      el.style.display = 'block';
      return;
    }
    const parts = [`Visits: ${Number(stats.count || 0)}`];
    const lastOnly = formatDateOnly(stats.lastVisit);
    if (lastOnly) parts.push(`Last: ${lastOnly}`);
    if (Number.isFinite(Number(stats.totalSpend))) parts.push(`Spent: ₹${Number(stats.totalSpend).toFixed(2)}`);
    el.textContent = parts.join(' • ');
    el.style.display = 'block';
  }

  /* ----------------------------- 6b) FETCH VISIT STATS (server) ----------------------------- */
  function debounce(fn, ms = 450) {
    let t; return (...args) => { clearTimeout(t); t = setTimeout(()=>fn(...args), ms); };
  }

  async function fetchVisitStatsByPhone(phone10) {
    if (!/^\d{10}$/.test(phone10)) return null;
    const url = `${WEB_APP_URL}?action=getVisitStats&token=${encodeURIComponent(AUTH_TOKEN)}&phone=${encodeURIComponent(phone10)}`;
    try{
      const res = await fetch(url, { method:'GET' });
      const data = await res.json();
      if (!data || !data.ok) return null;
      return {
        phone10: data.phone || phone10,
        name: data.name || '',
        count: Number(data.count || 0),
        totalSpend: Number(data.totalSpent || data.totalSpend || 0),
        lastVisit: data.lastVisit || ''
      };
    }catch(e){
      console.error('Visit stats fetch error:', e);
      return null;
    }
  }

  async function pullAndRenderVisitStats(phone10){
    if (/^\d{10}$/.test(phone10)) showStatsLoading(preferInlineEl());
    else { hideStats(preferInlineEl()); }

    const stats = await fetchVisitStatsByPhone(phone10);

    setStats(preferInlineEl(), stats, phone10);

    if (els.previewCustomerStats){
      if (!stats){
        els.previewCustomerStats.textContent = '';
        els.previewCustomerStats.dataset.count = '0';
        els.previewCustomerStats.dataset.lastVisit = '';
        els.previewCustomerStats.dataset.totalSpend = '0';
      } else {
        setStats(els.previewCustomerStats, stats, phone10);
        els.previewCustomerStats.dataset.count = String(stats.count || 0);
        els.previewCustomerStats.dataset.lastVisit = formatDateOnly(stats.lastVisit) || '';
        els.previewCustomerStats.dataset.totalSpend = String(stats.totalSpend || 0);
      }
    }
  }

  const fetchVisitStatsDebounced = debounce(pullAndRenderVisitStats, 500);

  // Phone input (10 digits strict)
  els.customerPhone.addEventListener('input', (e)=>{
    e.target.value = e.target.value.replace(/\D/g,'').slice(0,10);
    validatePhone();

    const phone10 = e.target.value;
    if (/^\d{10}$/.test(phone10)) {
      showStatsLoading(preferInlineEl());
      fetchVisitStatsDebounced(phone10);
    } else {
      hideStats(preferInlineEl());
    }
  });

  function validatePhone(){
    const ok = /^\d{10}$/.test(els.customerPhone.value);
    if (els.phoneError) {
      els.phoneError.style.display = ok || els.customerPhone.value.length===0 ? 'none' : 'block';
    }
    return ok;
  }

  // Contact Picker
  if (els.pickContactBtn){
    els.pickContactBtn.addEventListener('click', async ()=>{
      try{
        if (!('contacts' in navigator) || !('select' in navigator.contacts)) {
          alert('Contact picking is not supported on this browser. Try Chrome on Android with HTTPS.');
          return;
        }
        const selected = await navigator.contacts.select(['name','tel'], { multiple:false });
        if (!selected || !selected.length) return;
        const c = selected[0];

        let phones = [];
        if (Array.isArray(c.tel)) phones = c.tel;
        else if (Array.isArray(c.phoneNumbers)) phones = c.phoneNumbers.map(p => p.number || p.value || p);
        else if (c.tel) phones = [c.tel];
        else if (c.phoneNumber) phones = [c.phoneNumber];

        const raw = (phones && phones[0]) ? String(phones[0]) : '';
        const digits = raw.replace(/\D/g,'');
        const last10 = digits.slice(-10);
        if (last10.length === 10) { els.customerPhone.value = last10; validatePhone(); }

        if (c.name) {
          if (Array.isArray(c.name)) els.customerName.value = c.name[0] || els.customerName.value;
          else if (typeof c.name === 'string') els.customerName.value = c.name || els.customerName.value;
        }

        const phone10 = els.customerPhone.value;
        if (/^\d{10}$/.test(phone10)) {
          showStatsLoading(preferInlineEl());
          fetchVisitStatsDebounced(phone10);
        } else {
          hideStats(preferInlineEl());
        }
      }catch(err){
        console.error(err);
        alert('Contact picking was cancelled or failed.');
      }
    });
  }

  /* ----------------------------- 7) ITEMS ----------------------------- */
  const LS='nusrat.vertical.items.v7';
  function compute(obj){ return Math.max(0, obj.qty * obj.rate); }
  function liToObj(li){
    const obj = {
      type: li.querySelector('[data-f="type"]').value || '',
      name: (li.querySelector('[data-f="name"]').value || '').trim() || (li.querySelector('[data-f="type"]').value || 'Item'),
      qty:  toNumber(li.querySelector('[data-f="qty"]').value),
      rate: toNumber(li.querySelector('[data-f="rate"]').value)
    };
    if (obj.name.length > NAME_MAX) obj.name = obj.name.slice(0, NAME_MAX);
    obj.amount = compute(obj);
    return obj;
  }
  function getItems(){ return Array.from(els.itemsList.children).map(liToObj); }
  function updateLi(li){
    const amt = liToObj(li).amount;
    li.querySelector('[data-f="amt"]').textContent = amt.toFixed(2);
  }
  function updateNameCounter(nameEl, ccEl){
    if(!nameEl) return;
    const len = nameEl.value.length;
    if (ccEl) ccEl.textContent = `${len}/${NAME_MAX}`;
    if (len === NAME_MAX){ showErrorBelow(nameEl, `Max ${NAME_MAX} characters reached.`); }
    else { hideErrorBelow(nameEl); }
  }
  function attachLi(li){
    const typeEl = li.querySelector('[data-f="type"]');
    const nameEl = li.querySelector('[data-f="name"]');
    const qtyEl  = li.querySelector('[data-f="qty"]');
    const ccEl   = li.querySelector('[data-f="cc"]');

    updateNameCounter(nameEl, ccEl);

    nameEl.addEventListener('input', () => {
      if (nameEl.value.length > NAME_MAX) nameEl.value = nameEl.value.slice(0, NAME_MAX);
      updateNameCounter(nameEl, ccEl); updateLi(li); recalc(); persist();
    });

    li.querySelector('[data-act="minus"]').addEventListener('click', ()=>{
      qtyEl.value = Math.max(1, (parseInt(qtyEl.value||'1',10)-1));
      updateLi(li); recalc(); persist();
    });
    li.querySelector('[data-act="plus"]').addEventListener('click', ()=>{
      qtyEl.value = Math.max(1, (parseInt(qtyEl.value||'1',10)+1));
      updateLi(li); recalc(); persist();
    });

    li.addEventListener('input', (e)=>{
      if (e.target === typeEl){
        if ((!nameEl.value || ITEM_OPTIONS.includes(nameEl.value)) && typeEl.value) {
          nameEl.value = typeEl.value.slice(0, NAME_MAX);
          updateNameCounter(nameEl, ccEl);
        }
      }
      updateLi(li); recalc(); persist();
    });
    li.addEventListener('change', ()=>{ updateLi(li); recalc(); persist(); });

    li.querySelector('.del').addEventListener('click', ()=>{ li.remove(); recalc(); persist(); });
  }
  function addItem(obj={type:'',name:'',qty:1,rate:''}, { atTop = true } = {}){
    const safeName = (obj.name || obj.type || '').slice(0, NAME_MAX).replace(/"/g,'&quot;');
    const li = document.createElement('li'); li.className='item';
    li.innerHTML = `
      <div class="top">
        <div class="type">
          <select data-f="type">
            <option value="" disabled ${obj.type?'':'selected'}>Select item</option>
            ${ITEM_OPTIONS.map(x=>`<option value="${x}" ${x===obj.type?'selected':''}>${x}</option>`).join('')}
          </select>
        </div>
        <div class="name name-wrap">
          <input class="name" data-f="name" placeholder="Name" maxlength="15" value="${safeName}"/>
          <div class="char-counter" data-f="cc" aria-live="polite">0/15</div>
        </div>
        <button type="button" class="btn danger del" title="Delete">✕</button>
      </div>
      <div class="bot">
        <div class="stepper">
          <button type="button" class="sbtn" data-act="minus">−</button>
          <input class="qty" data-f="qty" type="number" min="1" step="1" value="${obj.qty||1}"/>
          <button type="button" class="sbtn" data-act="plus">+</button>
        </div>
        <input class="rate" data-f="rate" type="number" min="0" step="0.01" inputmode="decimal" placeholder="Rate" value="${obj.rate === '' ? '' : (obj.rate||0)}"/>
        <div class="amt">₹ <span data-f="amt">0.00</span></div>
      </div>
    `;
    if (atTop && els.itemsList.firstChild) els.itemsList.insertBefore(li, els.itemsList.firstChild);
    else els.itemsList.appendChild(li);
    attachLi(li); updateLi(li); recalc();
  }

  // Quickline parse & actions
  function parseQuickline(s){
    const parts = (s||'').trim().split(/\s+/);
    if(!parts.length) return null;
    let name=[], qty=1, rate='';
    for(let i=parts.length-1;i>=0;i--){
      const t=parts[i];
      if(rate==='' && /^[0-9]+(\.[0-9]+)?$/.test(t)){ rate = parseFloat(t); continue; }
      if(qty===1 && /^[0-9]+$/.test(t)){ qty = parseInt(t,10); continue; }
      name = parts.slice(0, i+1); break;
    }
    if(name.length===0) name = [parts[0]];
    const nm = name.join(' ').slice(0, NAME_MAX);
    return { name: nm, qty, rate: rate===''?'':rate };
  }
  function addFromQuickline(){
    const parsed = parseQuickline(els.quickline.value);
    if(!parsed || parsed.rate===''){ alert('Use: Item  Qty  Rate. Example: "Shirt 2 300"'); return; }
    let typeMatch = ITEM_OPTIONS.find(x => parsed.name.toLowerCase().includes(x.toLowerCase()));
    addItem({ type:typeMatch||'', name:parsed.name, qty:parsed.qty, rate:parsed.rate }, { atTop: true });
    els.quickline.value=''; persist();
  }
  if (els.quicklineForm){
    els.quicklineForm.addEventListener('submit', (e) => { e.preventDefault(); addFromQuickline(); });
  }
  if (els.quickline){
    els.quickline.addEventListener('keydown', (e) => {
      if ((e.key && e.key.toLowerCase() === 'enter') || e.keyCode === 13) { e.preventDefault(); addFromQuickline(); }
    });
    els.quickline.addEventListener('input', () => {
      if (/\n/.test(els.quickline.value)) { els.quickline.value = els.quickline.value.replace(/\n/g, '').trim(); }
    });
  }
  if (els.fabAdd){
    els.fabAdd.addEventListener('click', () => addItem({}, { atTop: true }));
  }

  // Clear items
  if (els.clearItemsBtn){
    els.clearItemsBtn.addEventListener('click', ()=>{
      if(!confirm('Clear all items?')) return;
      els.itemsList.innerHTML = '';
      try { localStorage.removeItem(LS); } catch(_) {}
      recalc();
    });
  }

  // Persistence
  function persist(){ try{ localStorage.setItem(LS, JSON.stringify(getItems())); }catch(_){} }
  function load(){
    const raw=localStorage.getItem(LS);
    if(!raw){ recalc(); return; }
    try{
      const arr=JSON.parse(raw);
      if(Array.isArray(arr) && arr.length){
        els.itemsList.innerHTML='';
        arr.forEach(it=> addItem({
          type: it.type||'',
          name: (it.name||'').slice(0, NAME_MAX),
          qty: it.qty||1,
          rate: (it.rate===''? '': it.rate)
        }, { atTop: false }));
      }
    }catch(e){}
    recalc();
  }

  /* ----------------------------- 8) TOTALS ----------------------------- */
  function buildItemsBreakdown(items){
    const map = new Map();
    items.forEach(it=>{
      const qty=toNumber(it.qty);
      if(!qty) return;
      const key=(it.name||it.type||'Item').toString().trim();
      map.set(key,(map.get(key)||0)+qty);
    });
    const parts=[]; for (const [name,qty] of map.entries()) parts.push(`${qty}x ${name}`);
    return parts.join(', ');
  }
  function recalc(){
    const items=getItems();
    const countQty = items.reduce((s,i)=> s + toNumber(i.qty), 0);
    const breakdown = buildItemsBreakdown(items);

    const subtotal=items.reduce((s,i)=> s + toNumber(i.amount), 0);
    const gstRate=toNumber(els.taxRate.value);
    const flat=toNumber(els.flatDiscount.value);
    const gstAmt=subtotal*(gstRate/100);
    let grand=subtotal+gstAmt-flat; if(grand<0) grand=0;

    els.subtotal.value=subtotal.toFixed(2);
    els.grandTotal.value=grand.toFixed(2);
    els.itemsBreakdown.value = breakdown;
    els.itemsCountQty.value = String(countQty);

    const paid=toNumber(els.paidAmount.value);
    let negotiated = grand - paid; if(negotiated < 0) negotiated = 0;
    els.dueAmount.value = negotiated.toFixed(2);

    validatePaid();
  }

  // Paid validation input constraints
  if (els.paidAmount){
    els.paidAmount.addEventListener('keydown', (e)=>{
      const blocked = ['e','E','+','-'];
      if (blocked.includes(e.key)) e.preventDefault();
    });
    els.paidAmount.addEventListener('input', (e)=>{
      let v = e.target.value || '';
      v = v.replace(/[^\d.]/g, '');
      const parts = v.split('.');
      if (parts.length > 2) v = parts[0] + '.' + parts.slice(1).join('');
      if (parts.length === 2) { parts[1] = parts[1].slice(0, 2); v = parts[0] + '.' + parts[1]; }
      e.target.value = v;
      recalc(); persist();
    });
  }

  // === MOD: Paid must be > 0 and <= Grand (global rule).
  function validatePaid(){
    const paid = toNumber(els.paidAmount.value);
    const grand = toNumber(els.grandTotal.value);

    if (els.paidAmount.value === '') {
      els.paidError.style.display='none';
      return false;
    }
    if (paid <= 0) {
      els.paidError.textContent='Paid must be greater than 0.';
      els.paidError.style.display='block';
      return false;
    }
    if (paid > grand) {
      els.paidError.textContent='Paid cannot exceed Grand Total.';
      els.paidError.style.display='block';
      return false;
    }
    els.paidError.style.display='none';
    return true;
  }

  ['taxRate','flatDiscount','paidAmount','paymentMode','saleType'].forEach(id=>{
    const el = document.getElementById(id);
    if(!el) return;
    el.addEventListener('input', ()=>{ recalc(); persist(); });
    el.addEventListener('change', ()=>{ recalc(); persist(); });
  });

  // Validation helpers
  function showErrorBelow(el, msg){
    let e=el.parentElement.querySelector('.error');
    if(!e){ e=document.createElement('div'); e.className='error'; el.parentElement.appendChild(e);}
    e.textContent=msg; e.style.display='block';
  }
  function hideErrorBelow(el){
    const e=el.parentElement.querySelector('.error'); if(e) e.style.display='none';
  }

  // === MOD: Validation policy
  // ONLY for 'preview' we can ignore Payment Mode and Paid Amount;
  // Phone remains strict for all actions.
  function validateAll(forAction='general'){
    let ok=true;

    const isPreview = (forAction === 'preview'); // ONLY preview is relaxed for payment fields

    // Require name
    if(!els.customerName.value.trim()){
      showErrorBelow(els.customerName,'Customer name is required.');
      ok=false;
    } else {
      hideErrorBelow(els.customerName);
    }

    // Phone: strict for ALL actions
    if(!/^\d{10}$/.test(els.customerPhone.value)){
      els.phoneError.style.display='block';
      ok=false;
    } else {
      els.phoneError.style.display='none';
    }

    // Sale type required for all
    if(!els.saleType.value){
      showErrorBelow(els.saleType,'Select sale type.');
      ok=false;
    } else {
      hideErrorBelow(els.saleType);
    }

    // Items
    const rows = Array.from(els.itemsList.children);
    if(!rows.length){ ok=false; alert('Add at least one item.'); }
    let hasValid=false;
    rows.forEach(li=>{
      const sel = li.querySelector('[data-f="type"]');
      const qty = li.querySelector('[data-f="qty"]');
      const rate= li.querySelector('[data-f="rate"]');
      const name= li.querySelector('[data-f="name"]');
      if(!sel.value){ showErrorBelow(sel,'Select an item.'); ok=false; } else hideErrorBelow(sel);
      if(!qty.value){ showErrorBelow(qty,'Enter quantity.'); ok=false; } else hideErrorBelow(qty);
      if(rate.value===''){ showErrorBelow(rate,'Enter rate.'); ok=false; } else hideErrorBelow(rate);
      if (name.value.length > NAME_MAX){ showErrorBelow(name, `Item name cannot exceed ${NAME_MAX} characters.`); ok=false; }
      else if (name.value.length === NAME_MAX) { showErrorBelow(name, `Max ${NAME_MAX} characters reached.`); }
      else { hideErrorBelow(name); }
      if (sel.value && qty.value && rate.value!=='') hasValid=true;
    });
    if(!hasValid) ok=false;

    // Payment Mode + Paid:
    if (isPreview) {
      hideErrorBelow(els.paymentMode);
      hideErrorBelow(els.paidAmount);
    } else {
      if(!els.paymentMode.value){
        showErrorBelow(els.paymentMode,'Select payment mode.');
        ok=false;
      } else {
        hideErrorBelow(els.paymentMode);
      }
      if(els.paidAmount.value===''){
        showErrorBelow(els.paidAmount,'Enter paid amount (must be > 0).');
        ok=false;
      } else {
        hideErrorBelow(els.paidAmount);
        if(!validatePaid()) ok = false;
      }
    }

    if(!ok && forAction!=='general'){ alert('Please correct highlighted fields.'); }
    return ok;
  }

  /* ----------------------------- 9) PREVIEW ----------------------------- */
  function renderPreview(){
    const items = getItems();

    const rowsHtml = items.map(it => `
      <div style="display:flex;justify-content:space-between;border-bottom:1px solid #f3f4f6;padding:6px 0">
        <div>${it.name}</div>
        <div>Qty ${it.qty} × ₹${toNumber(it.rate).toFixed(2)} = <b>₹${toNumber(it.amount).toFixed(2)}</b></div>
      </div>
    `).join('');

    const subtotal = toNumber(els.subtotal.value);
    const gstRate = toNumber(els.taxRate.value);
    const gstAmt = subtotal * (gstRate / 100);
    const flat = toNumber(els.flatDiscount.value);
    const grand = toNumber(els.grandTotal.value);
    const paid = toNumber(els.paidAmount.value);
    const negotiated = toNumber(els.dueAmount.value);

    let totalsHtml = "";
    if (gstRate > 0 && gstAmt > 0)
      totalsHtml += `<tr><td>GST ${gstRate}%</td><td style="text-align:right">₹${gstAmt.toFixed(2)}</td></tr>`;
    if (flat > 0)
      totalsHtml += `<tr><td>Discount</td><td style="text-align:right">-₹${flat.toFixed(2)}</td></tr>`;
    totalsHtml += `<tr><td><b>Grand Total</b></td><td style="text-align:right"><b>₹${grand.toFixed(2)}</b></td></tr>`;
    if (els.paymentMode.value)
      totalsHtml += `<tr><td>Payment Mode</td><td style="text-align:right">${els.paymentMode.value}</td></tr>`;
    if (paid > 0)
      totalsHtml += `<tr><td>Paid</td><td style="text-align:right">₹${paid.toFixed(2)}</td></tr>`;
    if (negotiated > 0)
      totalsHtml += `<tr><td><b>Negotiated</b></td><td style="text-align:right"><b>₹${negotiated.toFixed(2)}</b></td></tr>`;

    // COUPON (Unified rule)
    let couponHtml = "";
    if (els.saleType.value === "Retail" && couponIsOn()) {
      const couponValue = couponValueFor(grand);
      if (couponValue > 0) {
        const couponNumber = els.invoiceNumber.value;
        const p2d = n => (n < 10 ? "0" + n : n);
        const dt = new Date(); dt.setMonth(dt.getMonth() + 2);
        const validTill = `${p2d(dt.getDate())}-${p2d(dt.getMonth() + 1)}-${dt.getFullYear()}`;
        couponHtml = `
          <tr>
            <td><b>Coupon (${couponNumber})</b><br><small>Valid till ${validTill}</small></td>
            <td style="text-align:right"><b>₹${couponValue}</b></td>
          </tr>
        `;
      }
    }

    // Visit stats from dataset (date-only already stored)
    let visitStatsHtml = '';
    if (els.previewCustomerStats) {
      const cnt = Number(els.previewCustomerStats.dataset.count || 0);
      const last = els.previewCustomerStats.dataset.lastVisit || '';
      const spent = Number(els.previewCustomerStats.dataset.totalSpend || 0);
      if (cnt > 0) {
        const bits = [`Visits: ${cnt}`];
        if (last) bits.push(`Last: ${last}`);
        if (Number.isFinite(spent)) bits.push(`Spent: ₹${spent.toFixed(2)}`);
        visitStatsHtml = `<div class="print-meta" style="margin:4px 0 8px 0;color:#4b5563;">${bits.join(' • ')}</div>`;
      } else if (/^\d{10}$/.test(els.customerPhone.value)) {
        visitStatsHtml = `<div class="print-meta" style="margin:4px 0 8px 0;color:#4b5563;">New customer</div>`;
      }
    }

    const dateDisplay = els.invoiceDate.value ? formatDateOnly(els.invoiceDate.value) : '';

    els.invoicePreview.innerHTML = `
      <div class="print-header">
        <div>
          <div class="print-title">Nusrat Enterprises</div>
          <div class="print-meta">📞 7978830017, 9330066455, 9040366455</div>
          <div class="print-meta">📍Plot No. 53, Goutam Nagar, Bhubaneswar, Odisha — 751014</div>
        </div>

        <div style="text-align:right" class="print-meta">
          <div><b>Invoice #:</b> ${els.invoiceNumber.value}</div>
          <div><b>Date:</b> ${dateDisplay || ''}</div>
          <div><b>Customer:</b> ${els.customerName.value || ''}</div>
        </div>
        
      </div>

      ${visitStatsHtml}

      ${rowsHtml}

      <div style="display:flex;justify-content:flex-end;margin-top:8px">
        <table>
          ${totalsHtml}
          ${couponHtml}
        </table>
      </div>
    `;
  }

  /* ----------------------------- 10) WA SUMMARY ----------------------------- */
  function summaryMonospace() {
    const items = getItems();
    const toAscii = s => String(s || '').replace(/[^\x00-\x7F]/g, '');
    const padEnd = (s, n) => (s.length > n ? s.slice(0, n) : s.padEnd(n, ' '));
    const padStart = (s, n) => (s.length > n ? s.slice(-n) : s.padStart(n, ' '));

    function fmtNoDecimal(n) {
      if (n == null || n === '') return '';
      const num = Number(n);
      if (Number.isNaN(num)) return String(n);
      const raw = Number.isInteger(num) ? String(num) : num.toFixed(2).replace(/\.?0+$/, '');
      return toAscii(raw);
    }

    function makeLine40(idx, name, qty, rate, amount) {
      const W = { idx: 2, name: 14, qty: 3, rate: 3, amt: 5 };
      const fIdx = padEnd(toAscii(idx), W.idx);
      const fName = padEnd(toAscii(name).slice(0, W.name), W.name);
      const fQty = padEnd(fmtNoDecimal(qty), W.qty);
      const fRate = padEnd(fmtNoDecimal(rate), W.rate);
      const fAmt = padEnd(fmtNoDecimal(amount), W.amt);
      return `${fIdx} ${fName} ${fQty}X${fRate}=${fAmt}`;
    }

    const rows = items.map((it, i) => {
      const qty = Number(it.qty) || 0;
      const rate = Number(it.rate) || 0;
      const amount = qty * rate;
      return makeLine40(i + 1, it.name, qty, rate, amount);
    });

    const subtotal = Number(els.subtotal.value || 0);
    const gstRate = Number(els.taxRate.value || 0);
    const flat = Number(els.flatDiscount.value || 0);
    const gstAmt = subtotal * (gstRate / 100);
    const grand = Number(els.grandTotal.value || 0);
    const paid = Number(els.paidAmount.value || 0);

    const totals = [
      ...(gstRate > 0 ? [`GST ${gstRate}%: ₹${gstAmt.toFixed(2)}`] : []),
      ...(flat > 0 ? [`Discount: -₹${flat.toFixed(2)}`] : []),
      `Grand Total: ₹${grand.toFixed(2)}`,
      ...(paid > 0 ? [`Paid: ₹${paid.toFixed(2)}`] : []),
    ];

    /* COUPON (Unified rule) */
    let couponLine = "";
    if (els.saleType.value === "Retail" && couponIsOn()) {
      const cVal = couponValueFor(grand);
      if (cVal > 0) {
        const couponNumber = els.invoiceNumber.value;
        const dt = new Date();
        dt.setMonth(dt.getMonth() + 2);
        const validTill = `${pad2(dt.getDate())}-${pad2(dt.getMonth() + 1)}-${dt.getFullYear()}`;
        couponLine = `Coupon: ${couponNumber} | Value: Rs ${cVal} (Redeem before ${validTill})`;
      }
    }

    const columnHeader = "#  Name           Qty Rs  Amt ";

    const lines = [
      "Nusrat Enterprises",
      "📞: 7978830017, 9330066455, 9040366455",
      "📍: Plot 53, Goutam Nagar, Bhubaneswar - 751014",
      "Trusted Since 2008",
      "",
      `Invoice: ${els.invoiceNumber.value} | Date: ${els.invoiceDate.value}`,
      `Name: ${els.customerName.value}`,
      "-------------------------------",
      columnHeader,
      ...rows,
      "-------------------------------",
      ...totals,
      "Thank you for shopping with Nusrat Enterprises!",
      ...(couponLine ? [couponLine] : []),
      "",
      `Review: ${LINKS.googleFeedback}`,
      `WA Channel: ${LINKS.waChannel}`,
      `Instagram: ${LINKS.instagram}`
    ];

    return "```\n" + lines.join("\n") + "\n```";
  }

  /* ----------------------------- 11) PDF (monospace) ----------------------------- */
  function buildMonospaceHTML() {
    const items=getItems();
    const COLS = { idx:3, name:22, qty:3, rate:8, amt:10 };
    const lineSep = '-'.repeat(COLS.idx+1+COLS.name+1+COLS.qty+1+COLS.rate+1+COLS.amt);
    const header = [ fit('#',COLS.idx), fit('ITEM',COLS.name), fit('QTY',COLS.qty), fit('RATE',COLS.rate), fit('AMOUNT',COLS.amt) ].join(' ');
    const rows = items.map((it,i)=>{
      const idx=String(i+1).padStart(COLS.idx,' ');
      const nm =fit(it.name,COLS.name);
      const qt =String(toNumber(it.qty)).padStart(COLS.qty,' ');
      const rt =p2(toNumber(it.rate),COLS.rate);
      const am =p2(toNumber(it.amount),COLS.amt);
      return `${idx} ${nm} ${qt} ${rt} ${am}`;
    });

    // totals
    const subtotal=toNumber(els.subtotal.value);
    const gstRate=toNumber(els.taxRate.value);
    const gstAmt=subtotal*(gstRate/100);
    const flat=toNumber(els.flatDiscount.value);
    const grand=toNumber(els.grandTotal.value);
    const paid=toNumber(els.paidAmount.value);
    function kv(k,v){ const key=(k+':').padEnd(13,' '); const val=String(v); return key + val; }

    /* COUPON TEXT (Unified rule) */
    let couponText = "";
    if (els.saleType.value === "Retail" && couponIsOn()) {
      const cVal = couponValueFor(grand);
      if (cVal > 0) {
        const couponNumber = els.invoiceNumber.value;
        const p2d = n => (n < 10 ? "0" + n : n);
        const dt = new Date();
        dt.setMonth(dt.getMonth() + 2);
        const validTill = `${p2d(dt.getDate())}-${p2d(dt.getMonth() + 1)}-${dt.getFullYear()}`;
        couponText =
          kv("Coupon", couponNumber) + "\n" +
          kv("Value", `₹${cVal}`) + "\n" +
          kv("Valid Till", validTill) + "\n";
      }
    }

    return `
      <div id="monoBox" style="
        width:${mmToPx(RECEIPT_MM)}px; padding:12px; border:1px solid #e5e7eb; border-radius:8px; background:#fff; color:#111827;
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono','Courier New', monospace; font-size:12px; line-height:1.4;">
        <pre style="margin:0; white-space:pre-wrap;">${
          [
            'Nusrat Enterprises',
            '📞 7978830017, 9330066455, 9040366455',
            '📍Plot No. 53, Goutam Nagar, Bhubaneswar, Odisha — 751014',
            '',
            `Invoice: ${els.invoiceNumber.value}` + (els.invoiceDate.value ? ` | Date: ${els.invoiceDate.value}` : ''),
            els.customerName.value ? `Customer: ${els.customerName.value}` : '',
            '',
            header, lineSep, ...rows, lineSep,
            ...(gstRate>0 && gstAmt>0 ? [kv(`GST ${gstRate}%`, `₹${gstAmt.toFixed(2)}`)] : []),
            ...(flat>0 ? [kv('Discount', `-₹${flat.toFixed(2)}`)] : []),
            kv('Grand Total', `₹${grand.toFixed(2)}`),
            ...(els.paymentMode.value ? [kv('Payment Mode', els.paymentMode.value)] : []),
            ...(paid>0 ? [kv('Paid', `₹${paid.toFixed(2)}`)] : []),
            '',
            couponText,
            'Thank you for shopping with Nusrat Enterprises!'
          ].join('\n')
        }</pre>
        <div style="height:8px"></div>
        <div style="font-size:11px;color:#374151">
          <div style="font-weight:600;margin-bottom:4px">Follow & Feedback:</div>
          <div><a id="gLink" href="${LINKS.googleFeedback}" target="_blank" rel="noopener">Google Feedback: ${LINKS.googleFeedback}</a></div>
          <div><a id="igLink" href="${LINKS.instagram}" target="_blank" rel="noopener">Instagram: ${LINKS.instagram}</a></div>
          <div><a id="waLink" href="${LINKS.waChannel}" target="_blank" rel="noopener">WhatsApp Channel: ${LINKS.waChannel}</a></div>
        </div>
      </div>`;
  }

  async function generatePdf() {
    els.monoRender.innerHTML = buildMonospaceHTML();
    await new Promise(r => setTimeout(r, 120));
    const box = els.monoRender.querySelector('#monoBox');
    const ig = els.monoRender.querySelector('#igLink');
    const wa = els.monoRender.querySelector('#waLink');
    const gl = els.monoRender.querySelector('#gLink');

    const canvas = await html2canvas(box, { scale: 2, backgroundColor: '#ffffff', useCORS: true });

    const pageWpt = mmToPt(RECEIPT_MM);
    const margin = 8;
    const targetWidth = pageWpt - margin * 2;
    const ratio = canvas.height / canvas.width;
    const targetHeight = targetWidth * ratio;
    const pageHpt = targetHeight + margin * 2;

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ unit: 'pt', format: [pageWpt, pageHpt] });

    const x = margin, y = margin, w = targetWidth, h = targetHeight;
    const imgData = canvas.toDataURL('image/png');
    pdf.addImage(imgData, 'PNG', x, y, w, h, undefined, 'FAST');

    const scaleX = w / canvas.width;
    const scaleY = h / canvas.height;
    function addLinkHotspot(el, url) {
      if (!el) return;
      const r = el.getBoundingClientRect();
      const base = box.getBoundingClientRect();
      const dx = r.left - base.left;
      const dy = r.top  - base.top;
      const lx = x + dx * scaleX;
      const ly = y + dy * scaleY;
      const lW = r.width * scaleX;
      const lH = r.height * scaleY;
      try { pdf.link(lx, ly, lW, lH, { url }); } catch(_) {}
    }
    addLinkHotspot(ig, LINKS.instagram);
    addLinkHotspot(wa, LINKS.waChannel);
    addLinkHotspot(gl, LINKS.googleFeedback);

    const cust = (els.customerName.value || 'Customer').replace(/[^\w\s-]+/g,'').trim().replace(/\s+/g,'_');
    const filename = `Invoice_${cust}_${els.invoiceDate.value}.pdf`;
    pdf.save(filename);
  }

  /* ----------------------------- 12) VCF ----------------------------- */
  function sanitizeName(s){ return (s||'').replace(/[\n\r;:,]/g,' ').trim(); }
  function makeVCard(){
    const first = sanitizeName(els.customerName.value || '');
    const last = sanitizeName(els.saleType.value || '');
    const phone10 = (els.customerPhone.value || '').replace(/\D/g,'').slice(-10);
    const msisdn = phone10 ? `+91${phone10}` : '';
    const org = 'Nusrat Enterprises';
    const title = last || '';
    const fn = (first + (last ? (' ' + last) : '')).trim() || 'Customer';
    const lines = [
      'BEGIN:VCARD','VERSION:3.0',
      `N:${last};${first};;;`,`FN:${fn}`,
      ...(title ? [`TITLE:${title}`] : []),
      `ORG:${org}`,
      ...(msisdn ? [`TEL;TYPE=CELL,VOICE:${msisdn}`] : []),
      ...(phone10 ? [`X-WA-ID:91${phone10}`] : []),
      `NOTE:Invoice ${els.invoiceNumber.value || ''}${els.paymentMode.value ? (' | Payment ' + els.paymentMode.value) : ''}`,
      'END:VCARD'
    ];
    return lines.join('\r\n');
  }
  function downloadVcfNow(){
    const vcf = makeVCard();
    const blob = new Blob([vcf], { type: 'text/vcard;charset=utf-8' });
    const a = document.createElement('a');
    const safeName = (els.customerName.value || 'Customer').replace(/[^\w\s-]+/g,'').trim().replace(/\s+/g,'_');
    const safeSale = (els.saleType.value || 'Sale').replace(/[^\w\s-]+/g,'').trim().replace(/\s+/g,'_');
    a.href = URL.createObjectURL(blob);
    a.download = `${safeName}_${safeSale}.vcf`;
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 300);
  }

  /* ----------------------------- 13) GOOGLE SHEET ----------------------------- */
  const DEBUG_POST = false; // set true temporarily to see outgoing payload in console

  function normalizeItemsForServer(rawItems) {
    const arr = Array.isArray(rawItems) ? rawItems : (rawItems ? Object.values(rawItems) : []);
    return arr.map(it => {
      const qty  = Number(it?.qty ?? it?.quantity ?? 0);
      const rate = Number(it?.rate ?? it?.price ?? 0);
      const amount = Number.isFinite(Number(it?.amount)) ? Number(it.amount) : Math.max(0, qty * rate);
      const nm = (it?.name || it?.itemName || it?.type || 'Item');
      return {
        type:   String(it?.type || ''),
        name:   String(nm).slice(0, NAME_MAX),
        qty:    Number.isFinite(qty) ? qty : 0,
        rate:   Number.isFinite(rate) ? rate : 0,
        amount: Number.isFinite(amount) ? amount : 0
      };
    });
  }

  async function collectInvoicePayload() {
    const rawItems = getItems();        // your existing DOM -> items collector
    const items = normalizeItemsForServer(rawItems);

    const payload = {
      token: AUTH_TOKEN,
      invoiceNumber: els.invoiceNumber.value || '',
      invoiceDate: els.invoiceDate.value || '',
      customerName: els.customerName.value || '',
      customerPhone: els.customerPhone.value || '',
      itemsBreakdown: els.itemsBreakdown.value || '',
      itemsCountQty: els.itemsCountQty.value || '0',
      subtotal: els.subtotal.value || '0',
      taxRate: els.taxRate.value || '0',
      flatDiscount: els.flatDiscount.value || '0',
      grandTotal: els.grandTotal.value || '0',
      paymentMode: els.paymentMode.value || '',
      paidAmount: els.paidAmount.value || '0',
      dueAmount: els.dueAmount.value || '0',
      saleType: els.saleType.value || '',
      items
    };

    if (DEBUG_POST) {
      console.log('[POST] payload:', JSON.stringify(payload, null, 2));
      console.log('[POST] items: isArray=%s, len=%d', Array.isArray(items), items.length);
    }

    return payload;
  }

  async function pushToGoogleSheet(opts = {}) {
    const { alertOnResult = true } = opts;
    if (!validateAll('push')) { if (alertOnResult) alert('Please correct highlighted fields.'); return false; }

    const payload = await collectInvoicePayload();
    try {
      const form = new FormData();
      form.append('payload', JSON.stringify(payload));

      const res = await fetch(WEB_APP_URL, { method: 'POST', body: form });

      let ok = false, data = null;
      try { data = await res.json(); ok = !!(data && data.ok); } catch (_) { ok = res.ok; }

      if (!ok && alertOnResult) {
        const txt = data?.error ? String(data.error) : (await res.text().catch(() => '')) || '';
        alert('Failed to save to Google Sheet ❌' + (txt ? '\n' + txt.slice(0, 500) : ''));
      }
      return ok;
    } catch (err) {
      console.error(err);
      if (alertOnResult) alert('Network error while saving to Google Sheet.');
      return false;
    }
  }
  /* ----------------------------- 14) ACTIONS ----------------------------- */
  if (els.previewBtn){
    els.previewBtn.addEventListener('click', async ()=>{
      if (!validateAll('preview')) return;
      els.previewCard.style.display='block';
      if (/^\d{10}$/.test(els.customerPhone.value)) {
        await pullAndRenderVisitStats(els.customerPhone.value);
      } else {
        hideStats(els.previewCustomerStats);
      }
      renderPreview();
    });
  }

  if (els.printBtn){
    els.printBtn.addEventListener('click', async (e)=>{
      e.preventDefault();
      if (!validateAll('print')) return;
      try { await generatePdf(); } catch (err) { console.error(err); alert('PDF generation failed. Please try again.'); }
    });
  }

  if (els.saveBtn){
    els.saveBtn.addEventListener('click', async ()=>{
      if (!validateAll('push')) return;
      setBtnLoading(els.saveBtn, true, 'Saving…', null);
      try {
        const ok = await pushToGoogleSheet({ alertOnResult: true });
        if (ok && /^\d{10}$/.test(els.customerPhone.value)) {
          await pullAndRenderVisitStats(els.customerPhone.value);
          if (els.previewCard.style.display !== 'none') renderPreview();
        }
      } finally {
        setBtnLoading(els.saveBtn, false, null, 'Save to Google Sheet');
      }
    });
  }

  if (els.copyBtn){
    els.copyBtn.addEventListener('click', ()=>{
      if (!validateAll('copy')) return;
      const items=getItems(); const lines=[];
      lines.push('Nusrat Enterprises');
      lines.push('📞 7978830017, 9330066455, 9040366455');
      lines.push('📍Plot No. 53, Goutam Nagar, Bhubaneswar, Odisha — 751014');
      lines.push('');
      lines.push(`Invoice: ${els.invoiceNumber.value}`);
      if(els.invoiceDate.value) lines.push(`Date: ${els.invoiceDate.value}`);
      if(els.customerName.value) lines.push(`Customer: ${els.customerName.value}`);
      lines.push(''); lines.push('Items:');
      items.forEach((it,i)=>lines.push(`${i+1}. ${it.name} — Qty ${it.qty} × ₹${toNumber(it.rate).toFixed(2)} = ₹${toNumber(it.amount).toFixed(2)}`));
      const subtotal=toNumber(els.subtotal.value), gstRate=toNumber(els.taxRate.value), flat=toNumber(els.flatDiscount.value), grand=toNumber(els.grandTotal.value);
      const gstAmt=subtotal*(gstRate/100); const paid=toNumber(els.paidAmount.value); const negotiated=toNumber(els.dueAmount.value);
      lines.push('');
      if(gstRate>0 && gstAmt>0) lines.push(`GST ${gstRate}%: ₹${gstAmt.toFixed(2)}`);
      if(flat>0) lines.push(`Discount: -₹${flat.toFixed(2)}`);
      lines.push(`Grand Total: ₹${grand.toFixed(2)}`);
      if(els.paymentMode.value) lines.push(`Payment Mode: ${els.paymentMode.value}`);
      if(paid>0) lines.push(`Paid: ₹${paid.toFixed(2)}`);
      if(negotiated>0) lines.push(`Negotiated: ₹${negotiated.toFixed(2)}`);
      lines.push('');
      lines.push('Thank you for shopping with Nusrat Enterprises!');
      lines.push('');
      lines.push('Follow & Feedback:');
      lines.push(`Instagram: ${LINKS.instagram}`);
      lines.push(`WhatsApp Channel: ${LINKS.waChannel}`);
      lines.push(`⭐ Write a Review: ${LINKS.googleFeedback}`);
      navigator.clipboard.writeText(lines.join('\n'))
        .catch(()=>alert('Copy failed (permission denied).'));
    });
  }

  if (els.vcfBtn){
    els.vcfBtn.addEventListener('click', ()=>{
      const nameOk = !!els.customerName.value.trim();
      const phoneOk = /^\d{10}$/.test(els.customerPhone.value || '');
      const saleOk = !!els.saleType.value;
      if (!nameOk || !phoneOk || !saleOk) {
        validateAll('vcf');
        alert('Please ensure Customer Name, 10-digit WhatsApp number, and Sale Type are filled.');
        return;
      }
      downloadVcfNow();
    });
  }

  function setWaState(stateClass){
    const btn = els.whatsAppBtn;
    if(!btn) return;
    btn.classList.remove('is-saving','is-downloading','is-sending','is-success');
    btn.removeAttribute('aria-busy');
    if(stateClass){
      btn.classList.add(stateClass);
      if(stateClass !== 'is-success'){ btn.setAttribute('aria-busy','true'); }
    }
  }
  if (els.whatsAppBtn){
    els.whatsAppBtn.addEventListener('click', async ()=>{
      if (!validateAll('whatsapp')) return;
      try {
        setWaState('is-saving');
        const saved = await pushToGoogleSheet({ alertOnResult: false });
        if (!saved) { setWaState(''); alert('Saving to Google Sheet failed. Please try again.'); return; }

        if (/^\d{10}$/.test(els.customerPhone.value)) {
          pullAndRenderVisitStats(els.customerPhone.value).catch(()=>{});
        }

        setWaState('is-downloading');
        downloadVcfNow();
        setWaState('is-sending');
        const phone = `91${els.customerPhone.value}`;
        const msg = summaryMonospace();
        setTimeout(()=>{
          window.open(`https://wa.me/${phone}?text=${encodeURIComponent(msg)}`,'_blank','noopener,noreferrer');
          setWaState('is-success');
          setTimeout(()=> setWaState(''), 1400);
        }, 500);
      } catch (err) {
        console.error(err);
        setWaState('');
        alert('Something went wrong while processing the WhatsApp flow.');
      }
    });
  }
  if (els.waInvoiceBtn){
    els.waInvoiceBtn.addEventListener('click', ()=>{
      if (!validateAll('whatsappInvoice')) return;
      const phone = `91${els.customerPhone.value}`;
      const msg = summaryMonospace();
      window.open(`https://wa.me/${phone}?text=${encodeURIComponent(msg)}`,'_blank','noopener,noreferrer');

      if (/^\d{10}$/.test(els.customerPhone.value)) {
        pullAndRenderVisitStats(els.customerPhone.value).catch(()=>{});
      }
    });
  }

  // Boot
  load();
})();