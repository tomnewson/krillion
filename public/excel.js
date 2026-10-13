// Excel chrome for Krillion. Draws the title bar, ribbon, formula bar, row and
// column headings, sheet tabs and status bar around the real (proxied) game,
// and wires the chrome's controls to the game's own buttons and links so every
// feature still works through the original React code.
(() => {
  "use strict";

  const WORKBOOK = "Opex_Forecast_FY27.xlsx";
  const COL_W = 64;
  const ROW_H = 20;
  const ZOOMS = [50, 75, 90, 100, 110, 125, 150, 200];
  const PREFS_KEY = "xl-prefs";
  const PACKS = { movies: "Movies", sports: "Sports", geography: "Geography", lexicon: "Lexicon", music: "Music" };

  const prefs = loadPrefs();
  let zoom = prefs.zoom || 100;
  let selection = { col: 0, row: 1, visible: false };
  let lastEsc = 0;

  // ---------- helpers ----------
  function loadPrefs() {
    try {
      return JSON.parse(localStorage.getItem(PREFS_KEY)) || {};
    } catch {
      return {};
    }
  }
  function savePrefs() {
    try {
      localStorage.setItem(PREFS_KEY, JSON.stringify({ ...prefs, zoom }));
    } catch {}
  }
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
  const h = (html) => {
    const t = document.createElement("template");
    t.innerHTML = html.trim();
    return t.content.firstElementChild;
  };
  const icon = (name, cls = "") => `<i class="xl-ico ${cls}" style="--ico:url('/__excel/icons/${name}.svg')"></i>`;
  const colName = (n) => {
    let s = "";
    for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  };
  const inChrome = (el) => !!el?.closest?.("[id^=xl-]");
  const guessInput = () => $(".ds-input");
  const gameRoot = () => $(".ds-root");
  const phase = () => gameRoot()?.dataset.phase || "";
  const visible = (el) => el && el.getClientRects().length > 0;
  const byText = (sel, re, root = document) => $$(sel, root).find((el) => !inChrome(el) && re.test(el.textContent.trim()));
  const text = (el) => (el?.textContent || "").replace(/\s+/g, " ").trim();
  // App dialogs: the site menu and dive log are <dialog>s, settings and the
  // wardrobe are module-styled backdrops.
  const dialogOpen = () => !!$('dialog[open], [class*="__backdrop"]');

  function status(msg, ms = 2500) {
    const el = $("#xl-status-msg");
    if (!el) return;
    el.textContent = msg;
    clearTimeout(status.t);
    status.t = setTimeout(() => {
      status.t = 0;
      updateMode();
    }, ms);
  }

  // Prefer clicking one of the app's own <a> tags so Next.js does a client-side
  // navigation; fall back to a normal page load.
  function navigate(path) {
    const link = $$(`a[href="${path}"]`).find((a) => !inChrome(a));
    if (link) link.click();
    else location.assign(path);
  }

  function clickOr(el, fallback) {
    if (el && !el.disabled) el.click();
    else if (typeof fallback === "function") fallback();
    else if (fallback) status(fallback);
  }

  // Settings and the wardrobe have buttons on game pages only; elsewhere the
  // site menu has a Settings entry.
  function openSettings() {
    const btn = $('button[aria-label="Open settings"]');
    if (btn) return btn.click();
    $('button[aria-label="Open menu"]')?.click();
    setTimeout(() => {
      const item = byText("#site-menu button", /^settings$/i);
      item ? item.click() : status("Settings are not available on this sheet");
    }, 60);
  }

  // ---------- actions available to the ribbon, backstage and sheet tabs ----------
  const actions = {
    today: () => navigate("/"),
    unlimited: () => navigate("/unlimited/classic"),
    archive: () => navigate("/archive"),
    packs: () => navigate("/packs"),
    account: () => navigate("/account"),
    friends: () => navigate("/friends"),
    faq: () => navigate("/faq"),
    privacy: () => navigate("/privacy"),
    restore: () => navigate("/restore"),
    menu: () => clickOr($('button[aria-label="Open menu"]'), "Menu is not available"),
    settings: openSettings,
    wardrobe: () => clickOr($('button[aria-label="Open wardrobe"]'), "The wardrobe opens from a game sheet"),
    mute: () => {
      const btn = $(".ds-mute");
      if (!btn) return status("Sound is only on game sheets");
      btn.click();
      setTimeout(syncPage, 50);
    },
    howto: () => {
      const d = $("details.ds-howto");
      if (d && visible(d)) {
        d.open = !d.open;
        d.scrollIntoView({ block: "nearest" });
      } else navigate("/faq");
    },
    // The one big button on a game sheet: Begin descent, Descend, Surface.
    go: () => {
      const btn = $$(".ds-begin-btn, .ds-next-btn").find(visible);
      if (btn) return btn.click();
      if (phase() === "answering") return actions.submit();
      status("Nothing to run right now");
    },
    submit: () => {
      const input = guessInput();
      if (!input || !visible(input)) return status("No prompt is open");
      if (!input.value.trim()) return status("Type an answer first");
      input.form.requestSubmit();
      setTimeout(() => ($("#xl-fx-input").value = guessInput()?.value ?? ""), 400);
    },
    copyResult: () => {
      const btn = byText(".ds-done button", /^copy result/i);
      btn ? (btn.click(), status("Result copied")) : status("Finish the dive to copy your result");
    },
    share: () => {
      const btn = $('.ds-done button[aria-label="Share via apps"]') || byText(".ds-done button", /^copy result/i);
      btn ? btn.click() : status("Finish the dive to share your result");
    },
    catch: () => {
      const sec = byText(".ds-done section", /^the catch/i);
      sec ? sec.scrollIntoView({ block: "start", behavior: "smooth" }) : status("Answers appear once you finish");
    },
    diveLog: () => clickOr($('[class*="GameHistoryModal-module"][class*="launcher"]'), "The dive log is on the Unlimited and pack sheets"),
    playAgain: () => clickOr(byText(".ds-done button", /play this one again/i), "Nothing to replay yet"),
    paste: async () => {
      const input = guessInput();
      if (!input || !visible(input)) return status("Nothing to paste into");
      try {
        setGuess((await navigator.clipboard.readText()).slice(0, 80));
        input.focus();
      } catch {
        status("Clipboard access was blocked");
      }
    },
    refresh: () => location.reload(),
    gridlines: () => toggleClass("xl-no-grid", "gridlines"),
    headings: () => toggleClass("xl-no-headings", "headings"),
    zoomIn: () => setZoom(ZOOMS.find((z) => z > zoom) ?? zoom),
    zoomOut: () => setZoom([...ZOOMS].reverse().find((z) => z < zoom) ?? zoom),
    zoom100: () => setZoom(100),
    boss: () => toggleBoss(),
    file: () => openBackstage(),
    privacyChoices: () => clickOr($(".consumer-privacy-btn"), () => navigate("/privacy")),
  };

  function toggleClass(cls, key) {
    const on = document.documentElement.classList.toggle(cls);
    prefs[key] = !on;
    savePrefs();
    syncChecks();
  }

  // ---------- ribbon definition ----------
  // Buttons: [label, icon, action?, size] — size "lg" is a tall button.
  const RIBBON = {
    Home: [
      ["Clipboard", [["Paste", "clipboard_paste", "paste", "lg"], ["Cut", "cut"], ["Copy", "copy", "copyResult"], ["Format Painter", "paint_brush", "wardrobe"]]],
      ["Font", "font"],
      ["Alignment", "align"],
      ["Number", "number"],
      ["Styles", [["Conditional Formatting", "table_lightning", "catch", "lg"], ["Format as Table", "table_simple", null, "lg"], ["Cell Styles", "color_fill", "wardrobe", "lg"]]],
      ["Cells", [["Insert", "table_insert_row", null, "lg"], ["Delete", "table_delete_row", null, "lg"], ["Format", "table_settings", "settings", "lg"]]],
      ["Editing", [["AutoSum", "autosum", "submit"], ["Fill", "arrow_down", "go"], ["Clear", "eraser"], ["Sort & Filter", "arrow_sort", "packs", "lg"], ["Find & Select", "search", "howto", "lg"]]],
      ["Dive", [["Run", "play", "go", "lg"], ["Analyze Data", "data_bar_vertical", "catch", "lg"], ["Share", "share", "share", "lg"]]],
    ],
    Insert: [
      ["Tables", [["PivotTable", "table_simple", "catch", "lg"], ["Table", "grid", null, "lg"]]],
      ["Illustrations", [["Pictures", "image", "wardrobe", "lg"], ["Shapes", "shapes", null, "lg"], ["Icons", "star", null, "lg"]]],
      ["Charts", [["Recommended Charts", "chart_multiple", "catch", "lg"], ["Column", "data_bar_vertical"], ["Line", "data_trending"], ["Pie", "data_pie"]]],
      ["Links", [["Link", "link", "packs", "lg"]]],
      ["Comments", [["Comment", "comment_add", "howto", "lg"]]],
      ["Text", [["Text Box", "text_field", null, "lg"], ["Header & Footer", "document_header_footer", null, "lg"]]],
    ],
    "Page Layout": [
      ["Themes", [["Themes", "color", "wardrobe", "lg"], ["Colors", "color_fill", "settings"], ["Fonts", "text_font"]]],
      ["Page Setup", [["Margins", "document_margins", null, "lg"], ["Orientation", "document_landscape", null, "lg"], ["Size", "document", null, "lg"], ["Print Area", "print", null, "lg"]]],
      ["Sheet Options", "sheetopts"],
    ],
    Formulas: [
      ["Function Library", [["Insert Function", "math_formula", "go", "lg"], ["AutoSum", "autosum", "submit", "lg"], ["Recently Used", "history", "archive", "lg"], ["Financial", "money", "restore", "lg"], ["Logical", "branch", null, "lg"], ["Text", "text_t", null, "lg"], ["Date & Time", "calendar_ltr", "today", "lg"]]],
      ["Formula Auditing", [["Trace Precedents", "arrow_trending", "diveLog"], ["Show Formulas", "code"], ["Error Checking", "error_circle", "faq"]]],
      ["Calculation", [["Calculate Now", "calculator", "refresh", "lg"], ["Calculation Options", "settings", "settings", "lg"]]],
    ],
    Data: [
      ["Get & Transform Data", [["Get Data", "database_arrow_down", "packs", "lg"], ["From Web", "globe", "today"], ["Recent Sources", "history", "archive"]]],
      ["Queries & Connections", [["Refresh All", "arrow_clockwise", "refresh", "lg"], ["Queries", "database", "diveLog"]]],
      ["Sort & Filter", [["Sort", "arrow_sort", "packs", "lg"], ["Filter", "filter", "packs", "lg"]]],
      ["Forecast", [["What-If Analysis", "lightbulb", "unlimited", "lg"], ["Forecast Sheet", "data_trending", "catch", "lg"]]],
    ],
    Review: [
      ["Proofing", [["Spelling", "text_proofing_tools", "howto", "lg"], ["Thesaurus", "book", "faq", "lg"]]],
      ["Accessibility", [["Check Accessibility", "accessibility", "settings", "lg"]]],
      ["Comments", [["New Comment", "comment_add", "howto", "lg"], ["Show Comments", "comment_multiple", "catch", "lg"]]],
      ["Protect", [["Protect Sheet", "shield_lock", "privacy", "lg"], ["Allow Edit Ranges", "lock_closed", "privacyChoices", "lg"]]],
      ["Share", [["Share Workbook", "share", "share", "lg"], ["Collaborators", "people", "friends", "lg"]]],
    ],
    View: [
      ["Workbook Views", [["Normal", "grid", null, "lg"], ["Page Layout", "document", null, "lg"]]],
      ["Show", "show"],
      ["Zoom", [["Zoom In", "zoom_in", "zoomIn", "lg"], ["Zoom Out", "zoom_out", "zoomOut", "lg"], ["100%", "zoom_fit", "zoom100", "lg"]]],
      ["Window", [["Sound", "speaker_2", "mute", "lg"], ["Navigation", "navigation", "menu", "lg"], ["Switch Windows", "window_multiple", "boss", "lg"]]],
    ],
    Help: [
      ["Help", [["Help", "question_circle", "howto", "lg"], ["FAQ", "book_question_mark", "faq", "lg"], ["Settings", "settings", "settings", "lg"]]],
      ["About", [["My Krillion", "person", "account", "lg"], ["Privacy", "shield_lock", "privacy", "lg"], ["Privacy Choices", "lock_closed", "privacyChoices", "lg"]]],
    ],
  };

  const TITLES = {
    paste: "Paste clipboard into your answer",
    share: "Share your result",
    copyResult: "Copy your result",
    submit: "Submit answer (Enter)",
    go: "Begin / continue the dive",
    howto: "How to play",
    packs: "Themed packs",
    catch: "See every answer",
    refresh: "Reload",
    boss: "Switch windows (or press Esc twice)",
    settings: "Settings",
    wardrobe: "Wardrobe",
    mute: "Sound on/off",
  };

  function btnHtml([label, ico, action, size]) {
    const tip = TITLES[action] || label;
    return `<button class="xl-rb ${size === "lg" ? "xl-rb-lg" : "xl-rb-sm"}" ${action ? `data-act="${action}"` : ""} title="${tip}">${icon(ico)}<span>${label}</span></button>`;
  }

  const SPECIAL = {
    font: () => `
      <div class="xl-font-grp">
        <div class="xl-row"><div class="xl-combo" style="width:118px">Aptos Narrow ${icon("chevron_down", "xl-chev")}</div><div class="xl-combo" style="width:44px">11 ${icon("chevron_down", "xl-chev")}</div>${btnHtml(["Increase Font Size", "text_font_size"])}</div>
        <div class="xl-row">${btnHtml(["", "text_bold"])}${btnHtml(["", "text_italic"])}${btnHtml(["", "text_underline"])}<span class="xl-vsep"></span>${btnHtml(["", "border_bottom"])}${btnHtml(["", "paint_bucket", null, "fill"])}${btnHtml(["", "text_color"])}</div>
      </div>`,
    align: () => `
      <div class="xl-font-grp">
        <div class="xl-row">${btnHtml(["", "align_top"])}${btnHtml(["", "align_center_vertical"])}${btnHtml(["", "align_bottom"])}<span class="xl-vsep"></span>${btnHtml(["Wrap Text", "text_wrap"])}</div>
        <div class="xl-row">${btnHtml(["", "text_align_left"])}${btnHtml(["", "text_align_center"])}${btnHtml(["", "text_align_right"])}<span class="xl-vsep"></span>${btnHtml(["Merge & Center", "table_cells_merge"])}</div>
      </div>`,
    number: () => `
      <div class="xl-font-grp">
        <div class="xl-row"><div class="xl-combo" style="width:124px">General ${icon("chevron_down", "xl-chev")}</div></div>
        <div class="xl-row">${btnHtml(["", "currency_dollar_euro"])}${btnHtml(["", "text_number_format"])}<span class="xl-mini">%</span><span class="xl-mini">,</span><span class="xl-vsep"></span><span class="xl-mini">.00</span><span class="xl-mini">.0</span></div>
      </div>`,
    sheetopts: () => `
      <div class="xl-font-grp xl-checks">
        <label><input type="checkbox" data-pref="gridlines"> Gridlines</label>
        <label><input type="checkbox" data-pref="headings"> Headings</label>
      </div>`,
    show: () => `
      <div class="xl-font-grp xl-checks">
        <label><input type="checkbox" checked disabled> Ruler</label>
        <label><input type="checkbox" data-pref="gridlines"> Gridlines</label>
        <label><input type="checkbox" checked disabled> Formula Bar</label>
        <label><input type="checkbox" data-pref="headings"> Headings</label>
      </div>`,
  };

  function ribbonPanel(name) {
    return RIBBON[name]
      .map(([group, content]) => {
        const body = typeof content === "string" ? SPECIAL[content]() : groupButtons(content);
        return `<div class="xl-grp" data-group="${group}"><div class="xl-grp-body">${body}</div><div class="xl-grp-label">${group}</div></div>`;
      })
      .join("");
  }

  // Tall buttons sit side by side; small buttons stack three to a column.
  function groupButtons(buttons) {
    let out = "";
    let stack = [];
    const flush = () => {
      if (stack.length) out += `<div class="xl-stack">${stack.join("")}</div>`;
      stack = [];
    };
    for (const b of buttons) {
      if (b[3] === "lg") {
        flush();
        out += btnHtml(b);
      } else {
        stack.push(btnHtml(b));
        if (stack.length === 3) flush();
      }
    }
    flush();
    return out;
  }

  // ---------- build chrome ----------
  function build() {
    const tabs = ["File", ...Object.keys(RIBBON)];
    const top = h(`
      <div id="xl-top">
        <div id="xl-titlebar">
          <div class="xl-tb-left">
            <span class="xl-autosave">AutoSave <span class="xl-toggle"><span></span></span> On</span>
            <button class="xl-qat" title="Save">${icon("save")}</button>
            <button class="xl-qat" title="Undo">${icon("arrow_undo")}</button>
            <button class="xl-qat" title="Redo">${icon("arrow_redo")}</button>
            <button class="xl-qat" title="Customize Quick Access Toolbar">${icon("chevron_down")}</button>
          </div>
          <div class="xl-tb-title"><span id="xl-docname">${WORKBOOK}</span><span class="xl-saved"> - Saved ${icon("chevron_down", "xl-chev")}</span></div>
          <div class="xl-tb-right">
            <div class="xl-search">${icon("search")}<span>Search (Alt+Q)</span></div>
            <button id="xl-avatar" data-act="account" title="My Krillion">K</button>
            <span class="xl-winbtns"><span>&#x2013;</span><span>&#x25A1;</span><span>&#x2715;</span></span>
          </div>
        </div>
        <div id="xl-tabs">
          ${tabs.map((t) => `<button class="xl-tab${t === "Home" ? " is-active" : ""}${t === "File" ? " xl-tab-file" : ""}" data-tab="${t}">${t}</button>`).join("")}
          <span class="xl-tabs-spacer"></span>
          <button class="xl-tab-extra" data-act="catch">${icon("comment_multiple")} Comments</button>
          <button class="xl-tab-share" data-act="share">${icon("share")} Share</button>
        </div>
        <div id="xl-ribbon">${Object.keys(RIBBON).map((t) => `<div class="xl-panel${t === "Home" ? " is-active" : ""}" data-panel="${t}">${ribbonPanel(t)}</div>`).join("")}</div>
        <div id="xl-formulabar">
          <div id="xl-namebox"><input id="xl-namebox-input" value="A1" spellcheck="false">${icon("chevron_down", "xl-chev")}</div>
          <div class="xl-fx-btns"><span class="xl-dots">&#x22EE;</span><button id="xl-fx-cancel" title="Cancel">&#x2715;</button><button id="xl-fx-enter" title="Enter">&#x2713;</button><span class="xl-fx"><i>f</i>x</span></div>
          <input id="xl-fx-input" spellcheck="false" autocomplete="off" placeholder="">
          <button class="xl-fx-expand">${icon("chevron_down")}</button>
        </div>
        <div id="xl-colheads"><div id="xl-corner"></div><div id="xl-cols"></div></div>
      </div>`);

    const rows = h(`<div id="xl-rowheads"><div id="xl-rows"></div></div>`);
    const sel = h(`<div id="xl-selection"><span class="xl-handle"></span></div>`);
    const bottom = h(`
      <div id="xl-bottom">
        <div id="xl-sheetbar">
          <div class="xl-sheetnav"><button>&#x25C0;</button><button>&#x25B6;</button></div>
          <div id="xl-sheettabs"></div>
          <button class="xl-addsheet" title="New sheet" data-act="boss">+</button>
          <div class="xl-hscroll"><div></div></div>
        </div>
        <div id="xl-statusbar">
          <span id="xl-status-msg">Ready</span>
          <span class="xl-acc">${icon("accessibility")} Accessibility: Good to go</span>
          <span class="xl-status-spacer"></span>
          <span id="xl-stats"></span>
          <span class="xl-views">${icon("grid")}${icon("document")}${icon("document_page_break")}</span>
          <span class="xl-zoom"><button data-act="zoomOut">&#x2212;</button><span class="xl-zoom-track"><span id="xl-zoom-thumb"></span></span><button data-act="zoomIn">+</button><button id="xl-zoom-pct" data-act="zoom100" title="Reset zoom">100%</button></span>
        </div>
      </div>`);
    const backstage = h(`<div id="xl-backstage" hidden></div>`);
    const boss = h(`<div id="xl-boss" hidden></div>`);

    for (const el of [top, rows, sel, bottom, backstage, boss]) document.body.appendChild(el);

    // React may re-render <body> on hydration errors; put the chrome back if so.
    new MutationObserver(() => {
      for (const el of [top, rows, sel, bottom, backstage, boss]) if (!el.isConnected) document.body.appendChild(el);
    }).observe(document.body, { childList: true });

    wire(top);
  }

  function wire(top) {
    document.addEventListener("click", (e) => {
      const act = e.target.closest("[data-act]");
      if (act && inChrome(act)) {
        e.preventDefault();
        closeBackstage();
        actions[act.dataset.act]?.();
      }
    });

    top.querySelectorAll(".xl-tab").forEach((tab) =>
      tab.addEventListener("click", () => {
        if (tab.dataset.tab === "File") return openBackstage();
        top.querySelectorAll(".xl-tab").forEach((t) => t.classList.toggle("is-active", t === tab));
        top.querySelectorAll(".xl-panel").forEach((p) => p.classList.toggle("is-active", p.dataset.panel === tab.dataset.tab));
      }),
    );

    top.querySelectorAll("input[data-pref]").forEach((cb) => cb.addEventListener("change", () => actions[cb.dataset.pref]()));

    // Formula bar <-> answer input.
    const fx = $("#xl-fx-input");
    const editable = () => {
      const input = guessInput();
      return input && visible(input) && !input.disabled;
    };
    fx.addEventListener("input", () => {
      if (editable()) setGuess(fx.value);
    });
    fx.addEventListener("focus", () => {
      fx.readOnly = !editable();
      if (!fx.readOnly) fx.value = guessInput().value;
      setMode(fx.readOnly ? "Ready" : "Edit");
    });
    fx.addEventListener("blur", () => updateMode());
    fx.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        fx.readOnly ? actions.go() : actions.submit();
      } else if (e.key === "Escape") {
        if (!fx.readOnly) setGuess("");
        fx.blur();
      }
    });
    $("#xl-fx-enter").addEventListener("click", () => (editable() ? actions.submit() : actions.go()));
    $("#xl-fx-cancel").addEventListener("click", () => {
      setGuess("");
      guessInput()?.focus();
    });

    // Name box: type a cell reference to jump there.
    const nb = $("#xl-namebox-input");
    nb.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      const m = nb.value.trim().toUpperCase().match(/^([A-Z]{1,3})(\d{1,5})$/);
      if (!m) return status("That isn't a valid reference");
      const col = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
      select(col, Number(m[2]), true);
      nb.blur();
    });
    nb.addEventListener("focus", () => nb.select());

    document.addEventListener(
      "input",
      (e) => {
        if (e.target === guessInput() && document.activeElement !== fx) fx.value = e.target.value;
      },
      true,
    );
    document.addEventListener("focusin", (e) => {
      if (e.target === guessInput()) {
        selection.visible = false;
        placeSelection();
        const r = e.target.getBoundingClientRect();
        const cell = cellAt(r.left + 2, r.top + 2);
        if (cell) setNameBox(cell.col, cell.row);
        fx.value = e.target.value;
      }
      updateMode();
    });
    document.addEventListener("focusout", () => setTimeout(updateMode, 0));

    // Clicking the sheet selects the cell under the pointer.
    document.addEventListener("mousedown", (e) => {
      if (e.button !== 0 || inChrome(e.target) || e.target === document.documentElement) return;
      if (e.target.closest('dialog, [class*="__backdrop"], #cmpwrapper, input, textarea, select')) return;
      const cell = cellAt(e.clientX, e.clientY);
      if (!cell) return;
      select(cell.col, cell.row, false, e.target);
    });

    document.addEventListener("keydown", onKey, true);
    addEventListener("scroll", onViewportChange, { passive: true });
    addEventListener("resize", onViewportChange);
    $("#xl-sheettabs").addEventListener("click", (e) => {
      const tab = e.target.closest("[data-path]");
      if (!tab) return;
      if (!$("#xl-boss").hidden) toggleBoss(false);
      navigate(tab.dataset.path);
    });
  }

  // Set the React-controlled answer input's value so React sees the change.
  function setGuess(value) {
    const input = guessInput();
    if (!input) return;
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set;
    setter.call(input, value);
    input.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function onKey(e) {
    if (e.key === "Escape") {
      const now = Date.now();
      if (!$("#xl-backstage").hidden) return closeBackstage();
      if (now - lastEsc < 400) {
        lastEsc = 0;
        return toggleBoss();
      }
      lastEsc = now;
      return;
    }
    const t = e.target;
    const typing = t.closest?.("input, textarea, select, [contenteditable]");
    if (typing || e.metaKey || e.ctrlKey || e.altKey) return;
    if (dialogOpen() || !$("#xl-boss").hidden) return;

    const moves = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0], Tab: [e.shiftKey ? -1 : 1, 0] };
    // Enter runs the game's Begin/Descend/Surface button while there is one.
    const goBtn = $$(".ds-begin-btn, .ds-next-btn").find(visible);
    if (e.key === "Enter" && goBtn && !e.repeat) {
      e.preventDefault();
      return goBtn.click();
    }
    moves.Enter = [0, e.shiftKey ? -1 : 1];
    if (moves[e.key] && selection.visible) {
      e.preventDefault();
      const [dc, dr] = moves[e.key];
      select(Math.max(0, selection.col + dc), Math.max(1, selection.row + dr), true);
      return;
    }
    // Start typing anywhere → type into the answer, like entering a cell.
    if (e.key.length === 1 && !e.repeat) {
      const input = guessInput();
      if (input && !input.disabled && visible(input)) {
        input.focus();
        input.scrollIntoView({ block: "nearest" });
      }
    }
  }

  // ---------- grid geometry ----------
  // The sheet starts at the body's content box, where the app content sits.
  function metrics() {
    const cs = getComputedStyle(document.body);
    const z = zoom / 100;
    return {
      left: parseFloat(cs.paddingLeft) || 0,
      top: document.body.getBoundingClientRect().top + (parseFloat(cs.paddingTop) || 0),
      cw: COL_W * z,
      rh: ROW_H * z,
    };
  }
  function cellAt(x, y) {
    const m = metrics();
    if (x < m.left || y < m.top) return null;
    return { col: Math.floor((x - m.left) / m.cw), row: Math.floor((y - m.top) / m.rh) + 1 };
  }
  function setNameBox(col, row) {
    $("#xl-namebox-input").value = colName(col) + row;
    selection.col = col;
    selection.row = row;
    renderHeaders();
  }
  function select(col, row, scroll, target) {
    selection.visible = true;
    setNameBox(col, row);
    placeSelection();
    if (scroll) {
      const m = metrics();
      const y = m.top + (row - 1) * m.rh;
      const topBound = $("#xl-top").offsetHeight;
      const bottomBound = innerHeight - $("#xl-bottom").offsetHeight;
      if (y < topBound) scrollBy(0, y - topBound);
      else if (y + m.rh > bottomBound) scrollBy(0, y + m.rh - bottomBound);
    }
    $("#xl-fx-input").value = target ? cellText(target) : cellTextAt(col, row);
  }
  function cellTextAt(col, row) {
    const m = metrics();
    const x = m.left + col * m.cw + m.cw / 2;
    const y = m.top + (row - 1) * m.rh + m.rh / 2;
    if (y < 0 || y > innerHeight) return "";
    const els = document.elementsFromPoint(x, y).filter((el) => !inChrome(el));
    return els[0] ? cellText(els[0]) : "";
  }
  function cellText(el) {
    const leaf = el.closest("h1,h2,h3,p,span,a,button,li,label,dt,dd,summary") || el;
    const t = (leaf.innerText || "").replace(/\s+/g, " ").trim();
    if (!t || /^(MAIN|BODY)$/.test(leaf.tagName) || leaf.children.length > 6) return "";
    const p = currentPrompt();
    if (p && t === p) return krillionFormula(p);
    if (leaf.closest(".ds-score-pod")) return `=SUM(Dive!C2:C8)`;
    if (leaf.closest(".ds-depth-num")) return `=SUM(Dive!C2:C8)*10`;
    return t.slice(0, 300);
  }
  function placeSelection() {
    const box = $("#xl-selection");
    if (!selection.visible) return (box.style.display = "none");
    const m = metrics();
    box.style.display = "block";
    box.style.left = `${m.left + scrollX + selection.col * m.cw - 1}px`;
    box.style.top = `${m.top + scrollY + (selection.row - 1) * m.rh - 1}px`;
    box.style.width = `${m.cw + 1}px`;
    box.style.height = `${m.rh + 1}px`;
  }

  function renderHeaders() {
    const m = metrics();
    const cols = $("#xl-cols");
    const rows = $("#xl-rows");
    if (!cols || !rows) return;
    const nCols = Math.ceil((innerWidth - m.left) / m.cw) + 1;
    const offsetX = m.left - $("#xl-corner").offsetWidth;
    cols.style.transform = `translateX(${offsetX}px)`;
    cols.innerHTML = Array.from({ length: nCols }, (_, i) => `<span style="width:${m.cw}px" class="${selection.col === i ? "is-sel" : ""}">${colName(i)}</span>`).join("");

    const topBound = $("#xl-top").offsetHeight;
    const first = Math.max(1, Math.floor((topBound - m.top) / m.rh) + 1);
    const y0 = m.top + (first - 1) * m.rh - topBound;
    const nRows = Math.ceil(innerHeight / m.rh) + 2;
    rows.style.transform = `translateY(${y0}px)`;
    rows.innerHTML = Array.from({ length: nRows }, (_, i) => {
      const r = first + i;
      return `<span style="height:${m.rh}px" class="${selection.row === r ? "is-sel" : ""}">${r}</span>`;
    }).join("");
  }

  let raf = 0;
  function onViewportChange() {
    if (raf) return;
    raf = requestAnimationFrame(() => {
      raf = 0;
      renderHeaders();
      placeSelection();
    });
  }

  function setZoom(z) {
    zoom = z;
    savePrefs();
    document.documentElement.style.setProperty("--xl-zoom", z / 100);
    $("#xl-zoom-pct").textContent = `${z}%`;
    $("#xl-zoom-thumb").style.left = `${(ZOOMS.indexOf(z) / (ZOOMS.length - 1)) * 100}%`;
    onViewportChange();
  }

  function syncChecks() {
    const cls = document.documentElement.classList;
    $$("input[data-pref=gridlines]").forEach((c) => (c.checked = !cls.contains("xl-no-grid")));
    $$("input[data-pref=headings]").forEach((c) => (c.checked = !cls.contains("xl-no-headings")));
  }

  // ---------- status bar ----------
  function setMode(t) {
    const el = $("#xl-status-msg");
    if (el) el.textContent = t;
  }
  function updateMode() {
    const a = document.activeElement;
    if (a === guessInput()) return setMode("Enter");
    if (a?.id === "xl-fx-input" && !a.readOnly) return setMode("Edit");
    setMode(phase() === "answering" ? "Calculating..." : "Ready");
  }

  // Tier → conditional formatting preset, by points.
  const cfFor = (pts) => (pts >= 100 ? "krill" : pts >= 85 ? "deep" : pts >= 60 ? "rare" : pts >= 30 ? "school" : pts > 0 ? "plankton" : "miss");

  // Tag result rows so the CSS can style them with conditional formatting, and
  // feed the status bar's Average/Count/Sum.
  function scanSheet() {
    const scores = [];
    const landed = $(".ds-landed");
    if (landed) {
      const pts = Number((text($(".tpts b", landed)).match(/\d+/) || [0])[0]);
      landed.dataset.cf = landed.classList.contains("miss") ? "miss" : cfFor(pts);
    }
    for (const tag of $$(".ds-tierline .tag")) {
      const pts = Number((text(tag).match(/(\d+)\s*$/) || [0, 0])[1]);
      tag.dataset.cf = cfFor(pts);
    }
    const catchSec = byText(".ds-done section", /^the catch/i);
    if (catchSec) {
      catchSec.classList.add("xl-catch");
      for (const li of $$(":scope > ul > li", catchSec)) {
        const btn = $(":scope > button", li);
        const ptsEl = btn && $$(":scope > span.tabular-nums", btn)[0];
        if (!ptsEl) continue;
        const v = Number(text(ptsEl)) || 0;
        ptsEl.dataset.cf = cfFor(v);
        scores.push(v);
        // Expanded answer list: score badges on each answer.
        for (const s of $$("span", li)) {
          const m = text(s).match(/^\+(\d+)$/);
          if (m && !s.children.length) s.dataset.cf = cfFor(Number(m[1]));
        }
      }
    }
    // "▼ BEGIN DESCENT ▼" → "Begin descent", shown by the CSS from an attribute
    // so React's own text node is left alone.
    for (const btn of $$(".ds-begin-btn, .ds-next-btn")) {
      const raw = text(btn).replace(/[▼▲▶◀►◄]/g, "").trim();
      const label = raw === raw.toUpperCase() ? raw.charAt(0) + raw.slice(1).toLowerCase() : raw;
      if (btn.dataset.xlLabel !== label) btn.dataset.xlLabel = label;
    }
    byText(".ds-done section", /^the bearing/i)?.classList.add("xl-bearing");
    byText(".ds-done > * > div > div", /^dive log/i)?.classList.add("xl-divelog");

    const stats = $("#xl-stats");
    if (!stats) return;
    const done = $$(".ds-dots i.done").length;
    if (scores.length) {
      const sum = scores.reduce((a, b) => a + b, 0);
      stats.textContent = `Average: ${Math.round((sum / scores.length) * 10) / 10}    Count: ${scores.length}    Sum: ${sum}`;
    } else if (done) {
      const sum = Number(text($(".ds-score-num")).replace(/\D/g, "")) || 0;
      stats.textContent = `Average: ${Math.round((sum / done) * 10) / 10}    Count: ${done}    Sum: ${sum}`;
    } else stats.textContent = "";
  }

  function currentPrompt() {
    const p = $(".ds-card .prompt");
    return p && visible(p) ? text(p) : "";
  }
  const krillionFormula = (p) => `=KRILLION("${p.replace(/"/g, '""')}")`;

  function diveNumber() {
    const meta = text($(".ds-intro-meta > span")) || text(byText(".ds-done span", /dive #\s*\d+/i));
    return (meta.match(/#\s*(\d+)/) || [])[1] || "";
  }

  // ---------- sheet tabs ----------
  function renderSheetTabs() {
    const path = location.pathname;
    const seg = path.split("/")[1] || "";
    const tabs = [
      ["Today", "/", path === "/"],
      ["Unlimited", "/unlimited/classic", seg === "unlimited"],
      ["Packs", "/packs", path === "/packs"],
    ];
    if (PACKS[seg]) tabs.push([PACKS[seg], `/${seg}`, true]);
    tabs.push(
      ["Archive", "/archive", seg === "archive"],
      ["My Krillion", "/account", path === "/account"],
      ["Friends", "/friends", path === "/friends"],
      ["FAQ", "/faq", path === "/faq"],
    );
    const known = new Set(["", "unlimited", "packs", "archive", "account", "friends", "faq", ...Object.keys(PACKS)]);
    if (!known.has(seg)) tabs.push([seg.charAt(0).toUpperCase() + seg.slice(1), path, true]);
    const html = tabs.map(([n, p, a]) => `<button class="xl-sheettab${a ? " is-active" : ""}" data-path="${p}">${n}</button>`).join("");
    const el = $("#xl-sheettabs");
    if (el && el.innerHTML !== html) el.innerHTML = html;
  }

  // ---------- backstage (File menu) ----------
  function openBackstage(pane = "home") {
    const bs = $("#xl-backstage");
    const hour = new Date().getHours();
    const panes = {
      home: `
        <h1>Good ${hour < 12 ? "morning" : hour < 18 ? "afternoon" : "evening"}</h1>
        <div class="xl-bs-new">
          <button data-act="today"><span class="xl-bs-thumb xl-thumb-grid"></span>Today's dive</button>
          <button data-act="unlimited"><span class="xl-bs-thumb xl-thumb-list"></span>Unlimited</button>
          <button data-act="packs"><span class="xl-bs-thumb xl-thumb-tour"></span>Themed packs</button>
        </div>
        <h2>Recent</h2>
        <table class="xl-bs-recent"><thead><tr><th>Name</th><th>Location</th></tr></thead><tbody>
          <tr data-act="today"><td>${icon("grid")} Today's dive</td><td>krillion.io</td></tr>
          <tr data-act="unlimited"><td>${icon("arrow_repeat_all")} The endless dive</td><td>krillion.io › unlimited</td></tr>
          <tr data-act="packs"><td>${icon("library")} Themed packs</td><td>krillion.io › packs</td></tr>
          <tr data-act="archive"><td>${icon("archive")} Archive</td><td>krillion.io › archive</td></tr>
          <tr data-act="account"><td>${icon("person")} My Krillion</td><td>krillion.io › account</td></tr>
        </tbody></table>`,
      account: `
        <h1>Account</h1>
        <p>My Krillion keeps your streaks, best scores, completed dives and dive buddies.</p>
        <p><button class="xl-bs-primary" data-act="account">Open My Krillion</button> <button class="xl-bs-secondary" data-act="friends">Friends</button></p>
        <h2>Signing in</h2>
        <p class="xl-bs-note">Krillion signs in with Google through Firebase, which only accepts sign-ins on krillion.io itself. On this workbook your dives are kept in this browser instead. Purchases can be brought over with Restore.</p>
        <p><button class="xl-bs-secondary" data-act="restore">Restore a purchase</button></p>`,
      share: `
        <h1>Share</h1>
        <p>Copy your result grid, or share it with another app.</p>
        <p><button class="xl-bs-primary" data-act="copyResult">Copy result</button> <button class="xl-bs-secondary" data-act="share">Share…</button></p>`,
      options: `
        <h1>Options</h1>
        <div class="xl-bs-list">
          <button data-act="settings">${icon("settings")}<b>Settings</b><span>Low-power, colorblind, answer corrections, share link</span></button>
          <button data-act="wardrobe">${icon("shirt")}<b>Wardrobe</b><span>Pick your dive buddy</span></button>
          <button data-act="mute">${icon("speaker_2")}<b>Sound</b><span>Turn sound on or off</span></button>
          <button data-act="privacyChoices">${icon("lock_closed")}<b>Privacy choices</b><span>Update cookie and ad consent</span></button>
        </div>`,
      help: `
        <h1>Help</h1>
        <div class="xl-bs-list">
          <button data-act="howto">${icon("question_circle")}<b>How to play</b><span>7 prompts, 25 seconds each, rarer answers sink deeper</span></button>
          <button data-act="faq">${icon("book_question_mark")}<b>FAQ</b><span>Scoring, rarity and answer sets</span></button>
          <button data-act="menu">${icon("navigation")}<b>Krillion menu</b><span>Every page, merch and more</span></button>
          <button data-act="privacy">${icon("shield_lock")}<b>Privacy</b></button>
        </div>
        <h2>Keyboard shortcuts</h2>
        <p class="xl-bs-note">Start typing anywhere to answer · Enter to dive · Enter in the formula bar runs the next step · Arrow keys move the selection · Esc Esc switches to a work-safe sheet</p>`,
    };
    const nav = [
      ["home", "Home", "home"],
      ["today", "New", "document_add", "today"],
      ["packs", "Open", "folder_open", "packs"],
      ["account", "Account", "person"],
      ["share", "Share", "share"],
      ["options", "Options", "settings"],
      ["help", "Help", "question_circle"],
    ];
    bs.innerHTML = `
      <nav>
        <button class="xl-bs-back" title="Back">${icon("arrow_left")}</button>
        ${nav.map(([id, label, ico, act]) => `<button class="xl-bs-nav${id === pane ? " is-active" : ""}" ${act ? `data-act="${act}"` : `data-pane="${id}"`}>${icon(ico)}${label}</button>`).join("")}
        <span class="xl-bs-spacer"></span>
        <button class="xl-bs-nav" data-act="faq">FAQ</button>
        <button class="xl-bs-nav" data-act="privacy">Privacy</button>
      </nav>
      <section>${panes[pane]}</section>`;
    bs.hidden = false;
    bs.querySelector(".xl-bs-back").onclick = closeBackstage;
    bs.querySelectorAll("[data-pane]").forEach((b) => (b.onclick = () => openBackstage(b.dataset.pane)));
  }
  function closeBackstage() {
    const bs = $("#xl-backstage");
    if (bs) bs.hidden = true;
  }

  // ---------- boss mode: a plausible, boring expense sheet ----------
  function toggleBoss(force) {
    const boss = $("#xl-boss");
    const show = force ?? boss.hidden;
    if (show && !boss.innerHTML) boss.innerHTML = bossSheet();
    boss.hidden = !show;
    document.documentElement.classList.toggle("xl-boss-on", show);
    $("#xl-docname").textContent = WORKBOOK;
    status("Ready", 10);
  }
  function bossSheet() {
    const depts = ["Sales", "Marketing", "Engineering", "Operations", "Finance", "People", "Legal", "IT Services", "Facilities", "Customer Success", "Procurement", "Data"];
    let seed = 11;
    const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
    const fmt = (n) => n.toLocaleString("en-GB", { maximumFractionDigits: 0 });
    const rows = depts.map((d) => {
      const q = [0, 1, 2, 3].map(() => Math.round(35000 + rnd() * 240000));
      const budget = Math.round(q.reduce((a, b) => a + b) * (0.9 + rnd() * 0.2));
      return { d, q, total: q.reduce((a, b) => a + b), budget };
    });
    const tot = [0, 1, 2, 3].map((i) => rows.reduce((a, r) => a + r.q[i], 0));
    const grand = rows.reduce((a, r) => a + r.total, 0);
    const gbud = rows.reduce((a, r) => a + r.budget, 0);
    const cell = (v, cls = "") => `<td class="${cls}">${v}</td>`;
    return `<table>
      <tr class="b-title"><td colspan="3">FY27 Operating Expense Forecast</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
      <tr class="b-sub"><td colspan="3">All figures GBP · Version 2 (draft)</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
      <tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
      <tr class="b-head"><td colspan="2">Cost centre</td><td>Q1</td><td>Q2</td><td>Q3</td><td>Q4</td><td>Total</td><td>Budget</td><td>Var %</td></tr>
      ${rows
        .map((r) => {
          const v = ((r.total - r.budget) / r.budget) * 100;
          return `<tr>${cell(r.d, "b-l")}<td></td>${r.q.map((x) => cell(fmt(x), "n")).join("")}${cell(fmt(r.total), "n b")}${cell(fmt(r.budget), "n")}${cell(`${v.toFixed(1)}%`, `n ${v > 0 ? "neg" : "pos"}`)}</tr>`;
        })
        .join("")}
      <tr class="b-total">${cell("Total", "b-l")}<td></td>${tot.map((x) => cell(fmt(x), "n")).join("")}${cell(fmt(grand), "n")}${cell(fmt(gbud), "n")}${cell(`${(((grand - gbud) / gbud) * 100).toFixed(1)}%`, "n")}</tr>
    </table>`;
  }

  // ---------- page sync ----------
  function syncPage() {
    renderSheetTabs();
    scanSheet();
    if (document.activeElement !== guessInput() && document.activeElement?.id !== "xl-fx-input" && !status.t) updateMode();
    const fx = $("#xl-fx-input");
    if (!selection.visible && document.activeElement !== fx && document.activeElement !== guessInput()) {
      const p = currentPrompt();
      const landed = $(".ds-landed");
      fx.value = p ? krillionFormula(p) : landed ? text($(".tans", landed)).replace(/[“”"]/g, "") : "";
    }
    const muted = /sound on/i.test($(".ds-mute")?.getAttribute("aria-label") || "");
    $$('[data-act="mute"] .xl-ico').forEach((i) => i.style.setProperty("--ico", `url('/__excel/icons/${muted ? "speaker_mute" : "speaker_2"}.svg')`));
    document.documentElement.dataset.xlPhase = phase();
    const n = diveNumber();
    const title = `${WORKBOOK}${n ? ` - Dive ${n}` : ""} - Excel`;
    if (document.title !== title) document.title = title;
  }

  function watchNavigation() {
    for (const fn of ["pushState", "replaceState"]) {
      const orig = history[fn];
      history[fn] = function (...args) {
        const r = orig.apply(this, args);
        queueMicrotask(onNavigate);
        return r;
      };
    }
    addEventListener("popstate", onNavigate);
  }
  let lastPath = location.pathname;
  function onNavigate() {
    if (location.pathname === lastPath) return;
    lastPath = location.pathname;
    selection.visible = false;
    placeSelection();
    closeBackstage();
    toggleBoss(false);
    scrollTo(0, 0);
    syncPage();
  }

  function init() {
    if (prefs.gridlines === false) document.documentElement.classList.add("xl-no-grid");
    if (prefs.headings === false) document.documentElement.classList.add("xl-no-headings");
    document.documentElement.classList.add("xl-on");
    build();
    syncChecks();
    setZoom(zoom);
    watchNavigation();
    syncPage();
    renderHeaders();

    let pending = false;
    new MutationObserver((records) => {
      if (pending || records.every((r) => inChrome(r.target.nodeType === 1 ? r.target : r.target.parentElement))) return;
      pending = true;
      setTimeout(() => {
        pending = false;
        syncPage();
        onViewportChange();
      }, 120);
    }).observe(document.body, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ["data-phase", "aria-label", "open"] });

    // Keep the tab title in disguise when Next.js updates it.
    const titleEl = $("title");
    if (titleEl) new MutationObserver(syncPage).observe(titleEl, { childList: true });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
