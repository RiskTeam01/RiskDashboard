def base_css() -> str:
    return """
    <style>
        :root {
            --pc-blue: #003b7f;
            --pc-blue-dark: #002d62;
            --pc-blue-soft: #eaf2fb;
            --pc-orange: #f59e0b;
            --pc-orange-dark: #d97706;
            --bg: #f5f7fb;
            --card: #ffffff;
            --text: #172033;
            --muted: #667085;
            --border: #d7dde5;
            --green: #15803d;
            --red: #b91c1c;
            --soft-green: #dcfce7;
            --soft-red: #fee2e2;
        }

        * { box-sizing: border-box; }

        body {
            margin: 0;
            padding: 22px;
            font-family: Segoe UI, Arial, sans-serif;
            background:
                radial-gradient(circle at 0% 0%, rgba(0, 59, 127, 0.08), transparent 28%),
                linear-gradient(180deg, #f8fafc 0%, var(--bg) 100%);
            color: var(--text);
        }

        .shell { max-width: 1180px; margin: 0 auto; }

        .topbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 18px;
            position: relative;
            z-index: 20;
        }

        .brand { display: flex; align-items: center; gap: 16px; min-width: 0; }

        .brand-logo {
            height: 76px;
            width: auto;
            max-width: 180px;
            object-fit: contain;
            display: block;
            flex-shrink: 0;
        }

        .brand-text-wrap {
            display: flex;
            flex-direction: column;
            justify-content: center;
            line-height: 1.05;
        }

        .brand-title { font-weight: 950; color: var(--pc-blue); font-size: 27px; letter-spacing: -.04em; }
        .brand-subtitle { margin-top: 6px; font-size: 19px; font-weight: 900; color: var(--pc-orange-dark); letter-spacing: -.02em; }

        .topbar-right { display: flex; align-items: center; gap: 12px; }

        .user-pill {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: white;
            border: 1px solid rgba(0, 59, 127, .18);
            border-radius: 999px;
            padding: 8px 14px;
            font-size: 13px;
            font-weight: 800;
            color: var(--pc-blue-dark);
            box-shadow: 0 8px 24px rgba(0, 59, 127, .08);
            max-width: 260px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .user-pill .dot { width: 9px; height: 9px; border-radius: 50%; background: var(--green); flex-shrink: 0; }

        .menu-wrap { position: relative; display: inline-block; }

        .menu-button {
            width: 48px;
            height: 48px;
            border-radius: 14px;
            border: 1px solid rgba(0, 59, 127, .18);
            background: white;
            color: var(--pc-blue);
            font-size: 25px;
            font-weight: 900;
            cursor: pointer;
            box-shadow: 0 8px 24px rgba(0, 59, 127, .10);
            margin: 0;
            padding: 0;
        }

        .menu-button:hover { background: #f8fafc; }

        .menu-dropdown {
            display: none;
            position: absolute;
            top: 58px;
            right: 0;
            min-width: 210px;
            background: white;
            border: 1px solid var(--border);
            border-radius: 16px;
            box-shadow: 0 18px 44px rgba(15, 23, 42, 0.18);
            overflow: hidden;
            z-index: 999;
        }

        .menu-dropdown.show { display: block; }

        .menu-dropdown a {
            display: block;
            padding: 14px 16px;
            text-decoration: none;
            color: var(--text);
            font-weight: 800;
            border-bottom: 1px solid #edf2f7;
        }

        .menu-dropdown a:last-child { border-bottom: none; }
        .menu-dropdown a:hover { background: #f8fafc; }
        .menu-dropdown a.active { background: var(--pc-blue); color: white; }
        .menu-dropdown a.logout { color: var(--red); }

        .hero {
            background: white;
            border: 1px solid var(--border);
            border-radius: 20px;
            box-shadow: 0 10px 28px rgba(15, 23, 42, 0.08);
            margin-bottom: 18px;
            overflow: hidden;
        }

        .hero-topline {
            height: 8px;
            background: linear-gradient(90deg, var(--pc-blue) 0%, var(--pc-blue-dark) 60%, var(--pc-orange) 100%);
        }

        .hero-inner {
            padding: 24px 28px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 24px;
        }

        .hero-left { max-width: 760px; }

        .hero-kicker {
            color: var(--pc-orange-dark);
            font-weight: 900;
            text-transform: uppercase;
            letter-spacing: .12em;
            font-size: 12px;
            margin-bottom: 8px;
        }

        .hero h1 { margin: 0 0 8px 0; font-size: 34px; color: var(--pc-blue-dark); line-height: 1.1; letter-spacing: -.035em; }
        .hero p { margin: 0; color: var(--muted); font-size: 15px; line-height: 1.55; }

        .hero-right { min-width: 310px; text-align: right; display: flex; justify-content: flex-end; }

        .hero-right-brand { display: inline-flex; flex-direction: column; align-items: flex-end; line-height: 1.04; }
        .hero-right-title { color: var(--pc-blue); font-size: 24px; font-weight: 900; letter-spacing: -.03em; }
        .hero-right-subtitle { margin-top: 8px; color: #5d6f8b; font-size: 19px; font-weight: 400; letter-spacing: -.02em; line-height: 1.08; }

        .card {
            background: rgba(255, 255, 255, .96);
            border: 1px solid var(--border);
            border-radius: 18px;
            padding: 22px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.07);
            margin-bottom: 18px;
        }

        .center-card { max-width: 800px; margin: 0 auto; }
        h2, h3 { margin-top: 0; letter-spacing: -.02em; }
        .muted { color: var(--muted); font-size: 14px; }

        .status-pill { display: inline-block; padding: 6px 11px; border-radius: 999px; font-weight: 800; font-size: 13px; }
        .good { background: var(--soft-green); color: var(--green); }
        .bad { background: var(--soft-red); color: var(--red); }

        .upload-zone {
            border: 1px dashed #9fb4cf;
            border-radius: 16px;
            background: linear-gradient(180deg, #fbfdff, #f8fbff);
            padding: 18px;
            margin: 12px 0 16px 0;
        }

        input[type=file] { width: 100%; font-weight: 650; }

        input[type=text], input[type=password], input[type=number], input[type=email] {
            width: 100%;
            padding: 12px 14px;
            border: 1px solid var(--border);
            border-radius: 12px;
            font-size: 15px;
            font-family: inherit;
            background: white;
        }

        input:focus { outline: none; border-color: var(--pc-blue); box-shadow: 0 0 0 3px rgba(0, 59, 127, .12); }

        button, .button-link {
            border: 0;
            border-radius: 12px;
            padding: 12px 17px;
            font-weight: 900;
            cursor: pointer;
            background: var(--pc-blue);
            color: white;
            margin: 4px 4px 4px 0;
            text-decoration: none;
            display: inline-block;
            box-shadow: 0 8px 18px rgba(0, 59, 127, .18);
            font-family: inherit;
            font-size: 14px;
        }

        button:hover, .button-link:hover { filter: brightness(1.04); }
        button.orange, .button-link.orange { background: var(--pc-orange-dark); }
        button.secondary, .button-link.secondary { background: #475467; }
        button.danger, .button-link.danger { background: var(--red); box-shadow: 0 8px 18px rgba(185, 28, 28, .18); }

        code { background: #eef2f7; padding: 3px 6px; border-radius: 5px; word-break: break-all; }

        .output-box { border-left: 5px solid var(--green); padding: 14px; background: #f0fdf4; border-radius: 12px; margin-top: 14px; display: none; }
        .error-box { border-left: 5px solid var(--red); padding: 14px; background: #fff1f2; border-radius: 12px; margin-top: 14px; display: none; white-space: pre-wrap; }

        .console {
            background: #07111f;
            color: #d1e7ff;
            padding: 16px;
            border-radius: 14px;
            min-height: 590px;
            overflow: auto;
            font-family: Consolas, monospace;
            font-size: 12px;
            white-space: pre;
            border: 1px solid rgba(118, 183, 255, .25);
        }

        .summary-list {
            background: #f8fafc;
            padding: 14px;
            border-radius: 14px;
            border: 1px solid var(--border);
            font-family: Consolas, monospace;
            white-space: pre-wrap;
            font-size: 13px;
        }

        .metric-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 16px; }

        .metric { background: #f8fafc; border: 1px solid var(--border); padding: 14px; border-radius: 14px; }
        .metric b { color: var(--pc-blue); display: block; font-size: 20px; margin-bottom: 4px; }

        .outputs-grid { display: grid; grid-template-columns: 1fr; gap: 14px; }

        .output-item {
            border: 1px solid var(--border);
            border-radius: 16px;
            background: linear-gradient(180deg, #ffffff 0%, #fbfcfe 100%);
            padding: 18px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 18px;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
        }

        .output-left { min-width: 0; flex: 1; }
        .output-title { font-size: 16px; font-weight: 800; color: var(--pc-blue-dark); margin-bottom: 6px; word-break: break-word; }
        .output-meta { font-size: 13px; color: var(--muted); display: flex; flex-wrap: wrap; gap: 14px; }

        .empty-state {
            text-align: center;
            padding: 40px 20px;
            border: 1px dashed var(--border);
            border-radius: 16px;
            background: #fafcff;
            color: var(--muted);
        }

        /* Outputs page folder/file view */
        .outputs-list { display: flex; flex-direction: column; }

        .date-divider {
            display: flex;
            align-items: baseline;
            gap: 12px;
            margin: 28px 0 12px 0;
            padding: 0 4px 10px 4px;
            border-bottom: 1px solid var(--border);
        }
        .date-divider:first-child { margin-top: 0; }
        .date-divider .date-label { font-size: 18px; font-weight: 800; color: var(--pc-blue-dark); letter-spacing: -0.2px; }
        .date-divider .date-meta { font-size: 11px; color: var(--muted); font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; }
        .date-divider .date-count { margin-left: auto; font-size: 12px; color: var(--muted); background: var(--pc-blue-soft); border-radius: 999px; padding: 3px 10px; font-weight: 700; }

        .folder-card {
            border: 1px solid var(--border);
            border-radius: 14px;
            background: white;
            margin-bottom: 12px;
            overflow: hidden;
            transition: box-shadow 0.15s ease, border-color 0.15s ease;
        }
        .folder-card:hover { border-color: rgba(0, 59, 127, 0.25); }
        .folder-card[open] { box-shadow: 0 6px 18px rgba(15, 23, 42, 0.06); border-color: rgba(0, 59, 127, 0.18); }
        .folder-card > summary {
            display: flex;
            align-items: center;
            gap: 14px;
            padding: 14px 16px;
            cursor: pointer;
            list-style: none;
            background: linear-gradient(180deg, #fcfdfe 0%, #f4f7fc 100%);
        }
        .folder-card > summary::-webkit-details-marker { display: none; }
        .folder-card > summary::before {
            content: "";
            display: inline-block;
            width: 0;
            height: 0;
            border-top: 5px solid transparent;
            border-bottom: 5px solid transparent;
            border-left: 6px solid var(--pc-blue);
            transition: transform 0.15s ease;
            flex-shrink: 0;
        }
        .folder-card[open] > summary::before { transform: rotate(90deg); }

        .file-icon {
            flex-shrink: 0;
            width: 36px;
            height: 36px;
            border-radius: 8px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 18px;
            color: white;
        }
        .file-icon.folder { background: var(--pc-orange); }
        .file-icon.xlsx { background: #16a34a; }
        .file-icon.audit { background: var(--pc-blue); }
        .file-icon svg { width: 20px; height: 20px; }

        .file-row .file-icon { width: 28px; height: 28px; font-size: 11px; font-weight: 800; letter-spacing: 0.4px; border-radius: 6px; }
        .file-row .file-icon svg { width: 16px; height: 16px; }

        .folder-info, .file-info { flex: 1; min-width: 0; }
        .folder-name { font-weight: 800; font-size: 15px; color: var(--pc-blue-dark); word-break: break-word; }
        .folder-meta, .file-meta { font-size: 12px; color: var(--muted); margin-top: 3px; display: flex; flex-wrap: wrap; gap: 4px 12px; }
        .folder-meta .dot, .file-meta .dot { color: #c0c8d4; }
        .folder-meta code { background: #f3f5f9; padding: 1px 6px; border-radius: 4px; font-family: var(--mono, monospace); font-size: 11px; color: var(--text); }
        .folder-actions, .file-actions { display: flex; gap: 6px; flex-shrink: 0; align-items: center; }

        .folder-contents { padding: 4px 16px 12px 56px; background: #fafbfd; border-top: 1px solid var(--border); }

        .file-row { display: flex; align-items: center; gap: 12px; padding: 10px 0; border-bottom: 1px solid #eef1f7; }
        .file-row:last-child { border-bottom: none; }
        .file-row .file-name { font-weight: 700; color: var(--pc-blue-dark); font-size: 13px; word-break: break-all; }
        .file-row .file-actions a { padding: 6px 10px; font-size: 12px; }

        .single-card {
            display: flex;
            align-items: center;
            gap: 14px;
            padding: 14px 16px;
            border: 1px solid var(--border);
            border-radius: 14px;
            background: white;
            margin-bottom: 12px;
            transition: box-shadow 0.15s ease, border-color 0.15s ease;
        }
        .single-card:hover { border-color: rgba(0, 59, 127, 0.25); box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05); }
        .single-card.audit-only { background: #fafcff; }

        .batch-note {
            background: var(--pc-blue-soft);
            border: 1px solid rgba(0, 59, 127, .12);
            color: var(--pc-blue-dark);
            padding: 12px 14px;
            border-radius: 14px;
            font-size: 14px;
            font-weight: 650;
            margin-top: 12px;
        }

        /* Login */
        .login-shell { min-height: 84vh; display: flex; align-items: center; justify-content: center; }

        .login-card {
            width: 420px;
            max-width: 100%;
            background: white;
            border: 1px solid var(--border);
            border-radius: 22px;
            box-shadow: 0 24px 70px rgba(15, 23, 42, .16);
            overflow: hidden;
        }

        .login-topline { height: 8px; background: linear-gradient(90deg, var(--pc-blue) 0%, var(--pc-blue-dark) 60%, var(--pc-orange) 100%); }
        .login-body { padding: 30px 30px 34px 30px; }
        .login-logo { text-align: center; margin-bottom: 16px; }
        .login-logo img { height: 70px; max-width: 200px; object-fit: contain; }
        .login-title { text-align: center; color: var(--pc-blue-dark); font-size: 24px; font-weight: 950; letter-spacing: -.03em; margin-bottom: 4px; }
        .login-sub { text-align: center; color: var(--muted); font-size: 14px; margin-bottom: 22px; }
        .field-group { margin-bottom: 14px; }
        .field-group label { display: block; font-weight: 800; font-size: 13px; color: var(--pc-blue-dark); margin-bottom: 6px; }
        .login-error { background: var(--soft-red); color: var(--red); border-radius: 12px; padding: 11px 14px; font-size: 14px; font-weight: 700; margin-bottom: 16px; }
        .full-btn { width: 100%; text-align: center; margin: 8px 0 0 0; padding: 14px; font-size: 15px; }

        /* Progress */
        .progress-wrap { margin-top: 18px; display: none; }

        .progress-head { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 12px; }
        .progress-counts { display: flex; gap: 10px; flex-wrap: wrap; }

        .count-chip { border-radius: 999px; padding: 6px 12px; font-weight: 800; font-size: 13px; border: 1px solid var(--border); background: #f8fafc; }
        .count-chip.ok { background: var(--soft-green); color: var(--green); border-color: transparent; }
        .count-chip.fail { background: var(--soft-red); color: var(--red); border-color: transparent; }

        .progress-list { display: flex; flex-direction: column; gap: 10px; }

        .progress-row { border: 1px solid var(--border); border-radius: 14px; padding: 14px 16px; background: white; display: flex; justify-content: space-between; align-items: center; gap: 14px; }
        .progress-row .pr-left { min-width: 0; flex: 1; }
        .progress-row .pr-name { font-weight: 800; color: var(--pc-blue-dark); word-break: break-word; }
        .progress-row .pr-detail { font-size: 12.5px; color: var(--muted); margin-top: 4px; word-break: break-word; }
        .progress-row .pr-detail.err { color: var(--red); font-weight: 700; }

        .pstat { display: inline-flex; align-items: center; gap: 7px; border-radius: 999px; padding: 6px 12px; font-weight: 800; font-size: 12.5px; white-space: nowrap; }
        .pstat .pdot { width: 8px; height: 8px; border-radius: 50%; }
        .pstat.queued { background: #eef2f7; color: #475467; }
        .pstat.queued .pdot { background: #94a3b8; }
        .pstat.processing { background: var(--pc-blue-soft); color: var(--pc-blue); }
        .pstat.processing .pdot { background: var(--pc-blue); animation: pulse 1s infinite; }
        .pstat.complete { background: var(--soft-green); color: var(--green); }
        .pstat.complete .pdot { background: var(--green); }
        .pstat.failed { background: var(--soft-red); color: var(--red); }
        .pstat.failed .pdot { background: var(--red); }

        @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: .35; } }

        .pr-actions { display: flex; flex-direction: column; align-items: flex-end; gap: 6px; }
        .pr-actions a { font-size: 12px; padding: 7px 12px; margin: 0; }

        .download-bar { display: none; margin-top: 16px; padding: 16px; border-radius: 14px; background: #f0fdf4; border: 1px solid #bbf7d0; }
        .download-bar.has-fail { background: #fffbeb; border-color: #fde68a; }

        /* Settings */
        .settings-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }

        .user-row { display: flex; flex-direction: column; gap: 10px; border: 1px solid var(--border); border-radius: 12px; padding: 12px 14px; margin-bottom: 8px; background: white; }
        .user-row .ur-top { display: flex; justify-content: space-between; align-items: center; gap: 12px; }
        .user-row .ur-pwform { display: flex; gap: 6px; align-items: stretch; flex-wrap: wrap; }
        .user-row .ur-pwform input[type="password"] { flex: 1; min-width: 180px; padding: 8px 10px; font-size: 13px; border: 1px solid var(--border); border-radius: 8px; font-family: inherit; background: #fff; }
        .user-row .ur-pwform button { padding: 8px 12px; font-size: 12px; white-space: nowrap; }
        .user-row .ur-name { font-weight: 800; color: var(--pc-blue-dark); word-break: break-all; }
        .user-row .ur-meta { font-size: 12px; color: var(--muted); margin-top: 2px; }
        .user-row .ur-self-tag { display: inline-block; margin-left: 6px; padding: 1px 8px; background: var(--pc-blue); color: white; border-radius: 10px; font-size: 10px; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase; vertical-align: middle; }

        .flash { border-radius: 12px; padding: 12px 14px; font-weight: 700; font-size: 14px; margin-bottom: 16px; }
        .flash.ok { background: var(--soft-green); color: var(--green); }
        .flash.err { background: var(--soft-red); color: var(--red); }

        .limit-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }

        .loading-overlay { display: none; position: fixed; inset: 0; background: rgba(15, 23, 42, 0.38); z-index: 2000; align-items: center; justify-content: center; }
        .loading-card { width: 380px; background: white; border-radius: 18px; padding: 24px; text-align: center; box-shadow: 0 24px 70px rgba(15, 23, 42, .25); }
        .spinner { width: 42px; height: 42px; border: 5px solid #e2e8f0; border-top-color: var(--pc-orange); border-radius: 50%; margin: 0 auto 14px auto; animation: spin 1s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }

        @media (max-width: 900px) {
            body { padding: 14px; }
            .brand-logo { height: 58px; max-width: 140px; }
            .brand-title { font-size: 22px; }
            .brand-subtitle { font-size: 16px; }
            .hero-inner { flex-direction: column; align-items: flex-start; }
            .hero-right { min-width: 0; width: 100%; justify-content: flex-start; text-align: left; }
            .hero-right-brand { align-items: flex-start; }
            .hero-right-title { font-size: 22px; }
            .hero-right-subtitle { font-size: 17px; }
            .hero h1 { font-size: 28px; }
            .metric-grid { grid-template-columns: 1fr; }
            .output-item { flex-direction: column; align-items: flex-start; }
            .settings-grid { grid-template-columns: 1fr; }
            .limit-grid { grid-template-columns: 1fr; }
            .user-pill { max-width: 150px; }
            .folder-card > summary { flex-wrap: wrap; gap: 10px; }
            .single-card { flex-wrap: wrap; gap: 10px; }
            .folder-info, .file-info { flex-basis: 100%; min-width: 0; order: 2; }
            .folder-actions, .file-actions { flex-basis: 100%; order: 3; }
            .folder-contents { padding: 4px 12px 12px 12px; }
            .date-divider { flex-wrap: wrap; }
            .date-divider .date-count { margin-left: 0; }
        }
    </style>
    """
