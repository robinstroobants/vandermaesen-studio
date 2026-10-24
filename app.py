#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VDM Prijslijst Studio — Lokale Web/Desktop Interface (v1.2.0)
Draait 100% lokaal op 127.0.0.1 (Mac & Windows HP EliteBook, geen externe pakketten nodig).
Inclusief Markdown Mail Template Studio, Variabele Datums & Auto-Updater.
"""

import os
import sys
import json
import time
import socket
import subprocess
import webbrowser
from urllib.parse import parse_qs, urlparse
from http.server import HTTPServer, BaseHTTPRequestHandler

import base64
from vdm_engine import (
    APP_VERSION,
    BASE_DIR,
    RAW_DIR,
    OUT_DIR,
    LOGO_PATH,
    CONFIG_JSON_PATH,
    VOORWAARDEN_PATH,
    MASTER_EXCEL_PATH,
    ALL_TAB_LABELS,
    get_files_status,
    load_config_json,
    save_config_json,
    load_campaign_settings,
    save_campaign_settings,
    get_date_variables,
    format_smart_greeting,
    list_mail_templates,
    save_mail_template,
    render_customer_mail_content,
    parse_voorwaarden_excel,
    parse_master_klantenconfig_complete,
    generate_master_klantenconfig_excel,
    generate_customer_workbook,
    generate_customer_eml,
    generate_navision_export,
    audit_workbook_security,
    check_and_apply_updates,
)

PORT = 8765

HTML_UI = r"""<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>VDM Prijslijst Studio — 100% Lokaal & Discreet</title>
<style>
  :root {
    --vdm-green: #1b5e20;
    --vdm-green-light: #e8f5e9;
    --vdm-accent: #2e7d32;
    --bg: #f4f6f8;
    --card: #ffffff;
    --text: #1e293b;
    --muted: #64748b;
    --border: #e2e8f0;
    --danger: #dc2626;
    --blue: #0284c7;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.45;
  }
  header {
    background: linear-gradient(135deg, #144318 0%, #1b5e20 60%, #2e7d32 100%);
    color: #fff;
    padding: 16px 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 2px 10px rgba(0,0,0,0.12);
  }
  .brand { display: flex; align-items: center; gap: 14px; }
  .brand-badge {
    background: rgba(255,255,255,0.15);
    border: 1px solid rgba(255,255,255,0.3);
    padding: 8px 14px;
    border-radius: 8px;
    font-weight: 800;
    letter-spacing: 0.5px;
    font-size: 18px;
  }
  .brand h1 { font-size: 20px; font-weight: 700; }
  .brand p { font-size: 12.5px; opacity: 0.88; }
  .header-right { display: flex; align-items: center; gap: 10px; }
  .privacy-pill {
    background: rgba(16, 185, 129, 0.2);
    border: 1px solid rgba(167, 243, 208, 0.5);
    color: #ecfdf5;
    padding: 6px 14px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 600;
  }
  .btn-update {
    background: rgba(255,255,255,0.16);
    border: 1px solid rgba(255,255,255,0.35);
    color: #fff;
    padding: 6px 12px;
    border-radius: 8px;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
  }
  .btn-update:hover { background: rgba(255,255,255,0.28); }

  .container { max-width: 1460px; margin: 16px auto; padding: 0 24px; }

  /* Compacte Top-rij: Prijsronde + 2 Compacte Tegels */
  .top-bar-grid {
    display: grid;
    grid-template-columns: 1fr 220px 230px;
    gap: 14px;
    margin-bottom: 14px;
    align-items: stretch;
  }
  .campaign-box {
    background: #fff;
    border: 1px solid #cbd5e1;
    border-left: 4px solid var(--vdm-green);
    border-radius: 10px;
    padding: 12px 16px;
    min-width: 0;
  }
  .campaign-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 8px;
  }
  .campaign-header h3 { font-size: 13.5px; color: var(--vdm-green); display: flex; align-items: center; gap: 6px; }
  .campaign-grid {
    display: grid;
    grid-template-columns: 150px minmax(0, 1fr) minmax(0, 1fr);
    gap: 10px;
    align-items: end;
  }
  .field-col { display: flex; flex-direction: column; gap: 3px; min-width: 0; }
  .field-col label { font-size: 11px; font-weight: 700; color: var(--muted); text-transform: uppercase; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .field-col select, .field-col input { width: 100%; max-width: 100%; text-overflow: ellipsis; }
  .date-preview {
    font-size: 11px;
    color: #166534;
    background: #f0fdf4;
    padding: 3px 8px;
    border-radius: 5px;
    margin-top: 6px;
    display: inline-block;
  }

  .stat-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 12px 16px;
    display: flex;
    flex-direction: column;
    justify-content: center;
  }
  .stat-card.clickable {
    cursor: pointer;
    transition: all 0.15s ease;
    border-color: #bae6fd;
    background: linear-gradient(180deg, #ffffff 0%, #f0f9ff 100%);
  }
  .stat-card.clickable:hover {
    border-color: var(--blue);
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(2, 132, 199, 0.12);
  }
  .stat-card .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.4px; color: var(--muted); font-weight: 700; }
  .stat-card .val { font-size: 22px; font-weight: 800; margin: 2px 0; color: var(--text); }
  .stat-card .sub { font-size: 11.5px; color: var(--muted); }

  .action-tiles-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 14px;
    margin-bottom: 14px;
  }
  .action-tile {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 12px 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
  }
  .action-tile.tile-green { border-left: 4px solid var(--vdm-green); }
  .action-tile.tile-blue { border-left: 4px solid var(--blue); }
  .action-tile.tile-slate { border-left: 4px solid #64748b; }
  .action-tile .tile-info { min-width: 0; }
  .action-tile .tile-title { font-size: 13px; font-weight: 700; color: var(--text); }
  .action-tile .tile-desc { font-size: 11.5px; color: var(--muted); margin-top: 2px; }
  .action-tile button { flex-shrink: 0; white-space: nowrap; }
  button, .btn {
    border: none;
    border-radius: 8px;
    padding: 9px 15px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.15s ease;
    display: inline-flex;
    align-items: center;
    gap: 7px;
  }
  .btn-primary { background: var(--vdm-green); color: #fff; }
  .btn-primary:hover { background: #144318; }
  .btn-secondary { background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1; }
  .btn-secondary:hover { background: #e2e8f0; }
  .btn-blue { background: var(--blue); color: #fff; }
  .btn-blue:hover { background: #0369a1; }
  .btn-sm { padding: 5px 10px; font-size: 12px; border-radius: 6px; }

  .filters {
    display: flex;
    gap: 10px;
    align-items: center;
    margin-bottom: 12px;
    flex-wrap: wrap;
  }
  input[type="text"], input[type="date"], select, input[type="number"], textarea {
    padding: 8px 11px;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    font-size: 13px;
    background: #fff;
    font-family: inherit;
  }
  #search-input { min-width: 260px; }

  .table-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: hidden;
  }
  table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  th {
    background: #f8fafc;
    color: #475569;
    font-weight: 700;
    text-align: left;
    padding: 10px 11px;
    border-bottom: 1px solid var(--border);
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.4px;
  }
  td {
    padding: 9px 11px;
    border-bottom: 1px solid #f1f5f9;
    vertical-align: middle;
  }
  tr:hover td { background: #f8fafc; }
  .badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
  }
  .badge-rood { background: #fee2e2; color: #991b1b; }
  .badge-blauw { background: #dbeafe; color: #1e40af; }
  .badge-special { background: #fef3c7; color: #92400e; }
  .badge-inactief { background: #f1f5f9; color: #64748b; }
  .badge-nl { background: #f1f5f9; color: #334155; }
  .badge-fr { background: #ede9fe; color: #5b21b6; }
  .flag-pill {
    display: inline-block;
    background: #e0f2fe;
    color: #0369a1;
    padding: 2px 6px;
    border-radius: 4px;
    font-size: 10.5px;
    font-weight: 600;
    margin: 1px;
  }

  /* Modals */
  .modal-backdrop {
    display: none;
    position: fixed;
    inset: 0;
    background: rgba(15, 23, 42, 0.55);
    z-index: 100;
    align-items: center;
    justify-content: center;
    padding: 20px;
  }
  .modal-backdrop.open { display: flex; }
  .modal {
    background: #fff;
    border-radius: 14px;
    max-width: 960px;
    width: 100%;
    max-height: 92vh;
    overflow-y: auto;
    padding: 24px;
    box-shadow: 0 20px 40px rgba(0,0,0,0.22);
  }
  .modal h2 { font-size: 18px; margin-bottom: 14px; color: var(--vdm-green); }
  .form-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 12px;
    margin-bottom: 14px;
  }
  .form-group { display: flex; flex-direction: column; gap: 4px; }
  .form-group label { font-size: 12px; font-weight: 600; color: var(--muted); }
  .tabs-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 6px;
    background: #f8fafc;
    padding: 12px;
    border-radius: 8px;
    border: 1px solid var(--border);
    margin-bottom: 16px;
  }
  .tabs-grid label { font-size: 12px; display: flex; align-items: center; gap: 6px; cursor: pointer; }

  .tpl-split {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    margin-top: 12px;
  }
  .var-chip {
    display: inline-block;
    background: #f1f5f9;
    border: 1px solid #cbd5e1;
    padding: 2px 7px;
    border-radius: 5px;
    font-family: monospace;
    font-size: 11.5px;
    cursor: pointer;
    margin: 2px;
  }
  .var-chip:hover { background: #e2e8f0; }
  .preview-pane {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 14px;
    max-height: 420px;
    overflow-y: auto;
    font-size: 13px;
  }
  #toast {
    position: fixed;
    bottom: 24px;
    right: 24px;
    background: #0f172a;
    color: #fff;
    padding: 14px 20px;
    border-radius: 10px;
    font-size: 13.5px;
    font-weight: 600;
    box-shadow: 0 10px 25px rgba(0,0,0,0.2);
    display: none;
    z-index: 200;
  }
</style>
</head>
<body>

<header>
  <div class="brand">
    <div class="brand-badge">VDM</div>
    <div>
      <h1>Vandermaesen Prijslijst & Mail Studio <span style="font-size:12px;opacity:0.8;font-weight:500;">v__APP_VERSION__</span></h1>
      <p id="header-sub">Bronfile: Art. Prijzen_BRON_20260601.xlsx • 100% Lokaal & Offline</p>
    </div>
  </div>
  <div class="header-right">
    <div class="privacy-pill">🔒 100% Lokaal & Discreet</div>
  </div>
</header>

<div id="update-banner" style="display:none; background:#ecfdf5; border-bottom:1px solid #6ee7b7; color:#065f46; padding:10px 28px; font-size:13px; font-weight:600; text-align:center;">
  ✨ Er is zojuist automatisch een nieuwe software-update geïnstalleerd (<span id="update-files-txt"></span>). Je werkt nu met versie v__APP_VERSION__!
</div>

<div class="container">

  <!-- Onboarding / Bestanden Beheer Paneel (opent automatisch als Bronfile of Master Klantenconfig ontbreekt!) -->
  <div id="missing-files-box" style="display:none; background:#fffbeb; border:2px solid #f59e0b; border-radius:12px; padding:18px 22px; margin-bottom:16px;">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
      <h3 id="files-box-title" style="color:#92400e; font-size:15px;">⚠️ Gevoelige Bronbestanden Koppelen of Uploaden</h3>
      <button class="btn-secondary btn-sm" id="btn-close-files-box" onclick="toggleFilesBox(false)" style="display:none;">✖ Sluiten</button>
    </div>
    <p style="font-size:12.5px; color:#78350f; margin-bottom:14px;">
      Om veiligheidsredenen worden je gevoelige Excel-bestanden niet meegeleverd met de software. Je kunt ze hieronder <strong>rechtstreeks inladen/uploaden</strong> (of handmatig in de vermelde map op je laptop zetten):
    </p>
    <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:14px;">
      <div style="background:#fff; border:1px solid #fde68a; border-radius:8px; padding:12px;">
        <div style="font-size:11.5px; font-weight:700; color:#475569;">1. PRIJSBEREKENING BRONFILE (NL)</div>
        <div id="status-bron-nl" style="font-size:12.5px; margin:6px 0; font-weight:600;">Controleren...</div>
        <div style="font-size:11px; color:#64748b; margin-bottom:8px;">Map: <code>confidential/Art. Prijzen_BRON_....xlsx</code></div>
        <label class="btn btn-secondary btn-sm" style="cursor:pointer;">
          📤 Selecteer Bronfile (NL)
          <input type="file" accept=".xlsx" style="display:none;" onchange="uploadSensitiveFile(this, 'bron_nl')">
        </label>
      </div>
      <div style="background:#fff; border:1px solid #fde68a; border-radius:8px; padding:12px;">
        <div style="font-size:11.5px; font-weight:700; color:#475569;">2. MASTER KLANTENCONFIG (.XLSX)</div>
        <div id="status-master-cfg" style="font-size:12.5px; margin:6px 0; font-weight:600;">Controleren...</div>
        <div style="font-size:11px; color:#64748b; margin-bottom:8px;">Map: <code>confidential/00_MASTER_KLANTENCONFIG.xlsx</code></div>
        <label class="btn btn-secondary btn-sm" style="cursor:pointer;">
          📤 Selecteer Master Klantenconfig
          <input type="file" accept=".xlsx" style="display:none;" onchange="uploadSensitiveFile(this, 'master_config')">
        </label>
      </div>
      <div style="background:#fff; border:1px solid #fde68a; border-radius:8px; padding:12px;">
        <div style="font-size:11.5px; font-weight:700; color:#475569;">3. BRONFILE FRANS (OPTIONEEL)</div>
        <div id="status-bron-fr" style="font-size:12.5px; margin:6px 0; font-weight:600;">Controleren...</div>
        <div style="font-size:11px; color:#64748b; margin-bottom:8px;">Map: <code>confidential/Art. Prijzen_BRON_FRANS_....xlsx</code></div>
        <label class="btn btn-secondary btn-sm" style="cursor:pointer;">
          📤 Selecteer Bronfile (FR)
          <input type="file" accept=".xlsx" style="display:none;" onchange="uploadSensitiveFile(this, 'bron_fr')">
        </label>
      </div>
    </div>
  </div>

  <!-- Compacte Bovenbalk: Prijsronde + Actieve Klanten + Klikbare Markdown Mailer -->
  <div class="top-bar-grid">
    <div class="campaign-box">
      <div class="campaign-header">
        <h3>📅 Prijsronde & Variabelen</h3>
        <div style="display:flex; gap:6px;">
          <button class="btn-secondary btn-sm" onclick="toggleFilesBox(true)">📁 Bronbestanden Beheren</button>
          <button class="btn-secondary btn-sm" onclick="openOutputFolder()">📂 Open Output Map</button>
        </div>
      </div>
      <div class="campaign-grid">
        <div class="field-col">
          <label>Ingangsdatum</label>
          <input type="date" id="camp-date" onchange="saveCampaignSettings()">
        </div>
        <div class="field-col">
          <label>Mail Template (NL)</label>
          <select id="camp-tpl-nl" onchange="saveCampaignSettings()"></select>
        </div>
        <div class="field-col">
          <label>Mail Template (FR)</label>
          <select id="camp-tpl-fr" onchange="saveCampaignSettings()"></select>
        </div>
      </div>
      <div class="date-preview" id="date-vars-preview">
        Variabelen: <code>{{ingangsdatum_lang}}</code> = 1 juni 2026 • <code>{{ingangsdatum_kort}}</code> = 1/6/26 • Code = <code>20260601</code>
      </div>
    </div>

    <div class="stat-card">
      <div class="label">Actieve Klanten</div>
      <div class="val" id="stat-active">-</div>
      <div class="sub" id="stat-breakdown">Standaard, Blauw & Specials</div>
    </div>

    <div class="stat-card clickable" onclick="openTemplateModal()" title="Klik om de Markdown e-mail templates te bekijken en bewerken">
      <div class="label" style="color:var(--blue);">✉️ Markdown Mailer ↗</div>
      <div class="val" style="color:var(--blue);" id="stat-tpls">4 Templates</div>
      <div class="sub" style="color:#0369a1;font-weight:600;">Klik om templates te beheren</div>
    </div>
  </div>

  <!-- 3 Compacte Actie-Tegels -->
  <div class="action-tiles-grid">
    <div class="action-tile tile-green">
      <div class="tile-info">
        <div class="tile-title">⚡ Prijslijsten & Outlook Mails</div>
        <div class="tile-desc" id="status-text">Excels + .eml mails voor alle actieve klanten</div>
      </div>
      <button class="btn-primary" id="btn-gen-all" onclick="generateAllActive()">⚡ Genereer</button>
    </div>

    <div class="action-tile tile-blue">
      <div class="tile-info">
        <div class="tile-title">📊 Navision Import</div>
        <div class="tile-desc">Export van RGK-kortingen &amp; speciale prijzen</div>
      </div>
      <button class="btn-blue" onclick="generateNavision()">📊 Maak Export</button>
    </div>

    <div class="action-tile tile-slate">
      <div class="tile-info">
        <div class="tile-title">📋 Master Klantenconfig</div>
        <div class="tile-desc">Update <code>00_MASTER_KLANTENCONFIG.xlsx</code></div>
      </div>
      <button class="btn-secondary" onclick="exportMasterExcel()">📋 Exporteer</button>
    </div>
  </div>

  <div class="filters">
    <input type="text" id="search-input" placeholder="🔍 Zoek op klantnaam, contactpersoon, klantnr of opmerking..." oninput="renderTable()">
    <select id="group-filter" onchange="renderTable()">
      <option value="ACTIEF">Enkel Actieve Klanten (Standaard + Blauw + Specials)</option>
      <option value="ALL">Alle Klanten (incl. Inactief/Stopgezet)</option>
      <option value="STANDAARD_ROOD">Groep: Standaard Rood</option>
      <option value="CATEGORIE_BLAUW">Groep: Voormalig Blauw (-6%)</option>
      <option value="SPECIALS">Groep: Specials (Maatwerk tabbladen)</option>
      <option value="INACTIEF">Groep: Inactief / Sporadisch / Stopgezet</option>
    </select>
    <select id="lang-filter" onchange="renderTable()">
      <option value="ALL">Taal: NL + FR</option>
      <option value="NL">Enkel Nederlands (NL)</option>
      <option value="FR">Enkel Frans (FR)</option>
    </select>
  </div>

  <div class="table-card">
    <table>
      <thead>
        <tr>
          <th>Klantnr</th>
          <th>Klantnaam & Aanspreking Mail</th>
          <th>Groep</th>
          <th>Taal</th>
          <th>% RP/SP</th>
          <th>% Band</th>
          <th>% Poorten (&lt;10 / &ge;10)</th>
          <th>Speciale Voorwaarden & Vlaggen</th>
          <th>Tabs</th>
          <th style="text-align:right;">Acties</th>
        </tr>
      </thead>
      <tbody id="customers-tbody"></tbody>
    </table>
  </div>
</div>

<!-- Modal 1: Klantvoorwaarden & Contactpersoon aanpassen -->
<div class="modal-backdrop" id="edit-modal">
  <div class="modal" style="max-width:780px;">
    <h2 id="modal-title">Klantvoorwaarden & Aanspreking aanpassen</h2>
    <input type="hidden" id="edit-id">
    <div class="form-grid">
      <div class="form-group">
        <label>Klantnaam</label>
        <input type="text" id="edit-klantnaam">
      </div>
      <div class="form-group">
        <label>Navision Klantnr.</label>
        <input type="text" id="edit-klantnr">
      </div>
      <div class="form-group">
        <label>Voornaam Contactpersoon (bv. Anja en Mario)</label>
        <input type="text" id="edit-contact" placeholder="bv. Anja en Mario" oninput="autoUpdateGreeting()">
      </div>
      <div class="form-group">
        <label>E-mailadres (voor Outlook .eml)</label>
        <input type="text" id="edit-email" placeholder="bv. anmabeton@hotmail.be">
      </div>
      <div class="form-group" style="grid-column: span 2;">
        <label>Exacte Aanspreking bovenaan de mail (gebruik / voor nieuwe regel)</label>
        <input type="text" id="edit-aanspreking" placeholder="bv. Goedemiddag Anja, / Goedemiddag Mario">
      </div>
      <div class="form-group">
        <label>Taal (NL / FR)</label>
        <select id="edit-taal" onchange="autoUpdateGreeting()">
          <option value="NL">Nederlands (NL)</option>
          <option value="FR">Frans (FR)</option>
        </select>
      </div>
      <div class="form-group">
        <label>% Korting RP & SP (bv. 8 of -6)</label>
        <input type="number" step="0.01" id="edit-rp">
      </div>
      <div class="form-group">
        <label>% Korting Poorten EP/DP (&lt; 10 stuks)</label>
        <input type="number" step="0.1" id="edit-ep1">
      </div>
      <div class="form-group">
        <label>% Korting Poorten EP/DP (&ge; 10 stuks)</label>
        <input type="number" step="0.1" id="edit-ep10">
      </div>
      <div class="form-group">
        <label>% Korting VDM Band (Volle Pallet)</label>
        <input type="number" step="0.5" id="edit-band-vp">
      </div>
      <div class="form-group">
        <label>MEHA Prijskolom</label>
        <select id="edit-meha">
          <option value="STANDAARD">Standaard MEHA Prijzen</option>
          <option value="LENAERS">Speciale kolom LENAERS</option>
          <option value="LMDP">Speciale kolom LM - DE PRIL</option>
        </select>
      </div>
    </div>

    <div style="margin-bottom:14px; display:flex; gap:16px; flex-wrap:wrap; font-size:13px;">
      <label><input type="checkbox" id="edit-toebeh-pal"> Toebehoren Palletprijzen tonen</label>
      <label><input type="checkbox" id="edit-spec-vpu"> Speciale VPU-tabel gebruiken</label>
      <label><input type="checkbox" id="edit-ral7016"> Geen meerkost voor RAL 7016</label>
      <label><input type="checkbox" id="edit-actief"> Actieve klant</label>
    </div>

    <div style="font-size:12px; font-weight:700; color:var(--muted); margin-bottom:6px;">
      ACTIEVE TABBLADEN VOOR DEZE KLANT (WORDEN AUTOMATISCH HERNUMMERD OP 'WELKOM'):
    </div>
    <div class="tabs-grid" id="edit-tabs-container"></div>

    <div style="display:flex; justify-content:flex-end; gap:10px;">
      <button class="btn-secondary" onclick="closeModal()">Annuleren</button>
      <button class="btn-primary" onclick="saveCustomer()">💾 Opslaan & Direct Prijslijst + Mail Genereren</button>
    </div>
  </div>
</div>

<!-- Modal 2: Markdown Mail Template Editor & Live Preview -->
<div class="modal-backdrop" id="tpl-modal">
  <div class="modal" style="max-width:1120px;">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
      <h2>✉️ Markdown E-mail Templates (Kiezen, Bewerken of Nieuw Aanmaken)</h2>
      <button class="btn-secondary btn-sm" onclick="closeTemplateModal()">✖ Sluiten</button>
    </div>

    <input type="hidden" id="tpl-filename">
    <div style="display:flex; gap:12px; align-items:end; margin-bottom:12px; flex-wrap:wrap;">
      <div class="form-group" style="flex:1; min-width:240px;">
        <label>Selecteer bestaande template</label>
        <select id="tpl-select" onchange="loadSelectedTemplateIntoEditor()"></select>
      </div>
      <div class="form-group" style="flex:1; min-width:220px;">
        <label>Templatenaam</label>
        <input type="text" id="tpl-naam" placeholder="bv. Najaarsactie">
      </div>
      <div class="form-group" style="width:90px;">
        <label>Taal</label>
        <select id="tpl-taal">
          <option value="NL">NL</option>
          <option value="FR">FR</option>
        </select>
      </div>
      <button class="btn-secondary" onclick="newBlankTemplate()">➕ Nieuwe Template</button>
    </div>

    <div class="form-group" style="margin-bottom:10px;">
      <label>E-mail Onderwerp (Subject)</label>
      <input type="text" id="tpl-onderwerp" oninput="updateLiveMailPreview()">
    </div>

    <div style="font-size:12px; color:var(--muted); margin-bottom:8px;">
      <strong>Klik om variabele in te voegen:</strong>
      <span class="var-chip" onclick="insertVar('{{aanspreking}}')">{{aanspreking}}</span>
      <span class="var-chip" onclick="insertVar('{{klantnaam}}')">{{klantnaam}}</span>
      <span class="var-chip" onclick="insertVar('{{contactpersoon}}')">{{contactpersoon}}</span>
      <span class="var-chip" onclick="insertVar('{{ingangsdatum_lang}}')">{{ingangsdatum_lang}}</span>
      <span class="var-chip" onclick="insertVar('{{ingangsdatum_kort}}')">{{ingangsdatum_kort}}</span>
      <span class="var-chip" onclick="insertVar('{{bijlage_naam}}')">{{bijlage_naam}}</span>
      <span class="var-chip" onclick="insertVar('{{handtekening}}')">{{handtekening}}</span>
    </div>

    <div class="tpl-split">
      <div>
        <div style="font-size:12px; font-weight:700; color:var(--muted); margin-bottom:4px;">MARKDOWN BRONTEKST:</div>
        <textarea id="tpl-body" rows="16" style="width:100%; font-family:monospace; font-size:12.5px; line-height:1.5;" oninput="updateLiveMailPreview()"></textarea>
      </div>
      <div>
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
          <span style="font-size:12px; font-weight:700; color:var(--muted);">LIVE VOORBEELD VOOR KLANT:</span>
          <select id="preview-customer-select" style="padding:3px 8px; font-size:12px;" onchange="updateLiveMailPreview()"></select>
        </div>
        <div class="preview-pane">
          <div style="border-bottom:1px solid #e2e8f0; padding-bottom:8px; margin-bottom:10px; font-size:12px; color:#475569;">
            <div><strong>Onderwerp:</strong> <span id="prev-subj"></span></div>
            <div><strong>Bijlage:</strong> <code id="prev-att"></code></div>
          </div>
          <div id="prev-html"></div>
        </div>
      </div>
    </div>

    <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:16px;">
      <button class="btn-secondary" onclick="closeTemplateModal()">Sluiten</button>
      <button class="btn-primary" onclick="saveTemplateFromEditor()">💾 Template Opslaan in templates/mails/</button>
    </div>
  </div>
</div>

<div id="toast"></div>

<script>
let CUSTOMERS = [];
let ALL_TABS = [];
let SETTINGS = {};
let TEMPLATES = [];
let DATE_VARS = {};
let FILES_STATUS = {};

async function init() {
  const res = await fetch('/api/customers');
  const data = await res.json();
  CUSTOMERS = data.customers;
  ALL_TABS = data.all_tabs;
  SETTINGS = data.settings;
  TEMPLATES = data.templates;
  DATE_VARS = data.date_vars;
  FILES_STATUS = data.files_status || {};

  if (data.just_updated && data.just_updated.length > 0) {
    const b = document.getElementById('update-banner');
    document.getElementById('update-files-txt').textContent = data.just_updated.join(', ');
    b.style.display = 'block';
    showToast(`✨ Software automatisch bijgewerkt naar v__APP_VERSION__!`);
  }

  renderFilesStatus();
  populateCampaignControls();
  updateStats();
  renderTable();
}

function renderFilesStatus() {
  const nlEl = document.getElementById('status-bron-nl');
  const cfgEl = document.getElementById('status-master-cfg');
  const frEl = document.getElementById('status-bron-fr');
  const box = document.getElementById('missing-files-box');
  const closeBtn = document.getElementById('btn-close-files-box');
  const titleEl = document.getElementById('files-box-title');

  nlEl.innerHTML = FILES_STATUS.has_bron_nl
    ? `<span style="color:#166534;">✅ Aanwezig (${FILES_STATUS.bron_nl_name})</span>`
    : `<span style="color:#dc2626;">❌ Ontbreekt! Drop in confidential/ of upload hieronder</span>`;

  cfgEl.innerHTML = FILES_STATUS.has_master_config
    ? `<span style="color:#166534;">✅ Aanwezig (${FILES_STATUS.master_config_name || 'Ingeladen'})</span>`
    : `<span style="color:#dc2626;">❌ Ontbreekt! Drop in confidential/ of upload hieronder</span>`;

  frEl.innerHTML = FILES_STATUS.has_bron_fr
    ? `<span style="color:#166534;">✅ Aanwezig (${FILES_STATUS.bron_fr_name})</span>`
    : `<span style="color:#64748b;">ℹ️ Nog niet geladen (gebruikt voorlopig NL bronfile)</span>`;

  const subTxt = FILES_STATUS.has_bron_nl ? `Bronfile: ${FILES_STATUS.bron_nl_name}` : '⚠️ Bronfile ontbreekt';
  document.getElementById('header-sub').textContent = `${subTxt} • 100% Lokaal & Offline`;

  if (!FILES_STATUS.ready) {
    box.style.display = 'block';
    box.style.background = '#fef2f2';
    box.style.borderColor = '#ef4444';
    titleEl.innerHTML = '🚨 Actie vereist: Gevoelige Bronbestanden Ontbreken';
    closeBtn.style.display = 'none';
    document.getElementById('btn-gen-all').disabled = true;
  } else {
    box.style.background = '#f8fafc';
    box.style.borderColor = '#cbd5e1';
    titleEl.innerHTML = '📁 Bronbestanden & Master Klantenconfig Beheren / Vervangen';
    closeBtn.style.display = 'inline-flex';
    document.getElementById('btn-gen-all').disabled = false;
  }
}

function toggleFilesBox(show) {
  document.getElementById('missing-files-box').style.display = show ? 'block' : 'none';
}

function uploadSensitiveFile(inputEl, fileType) {
  const file = inputEl.files[0];
  if (!file) return;
  const reader = new FileReader();
  showToast(`⏳ Bezig met inladen van ${file.name}...`);
  reader.onload = async function(e) {
    const base64Data = e.target.result.split(',')[1];
    const res = await fetch('/api/upload_file', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        file_type: fileType,
        filename: file.name,
        content_base64: base64Data
      })
    });
    const out = await res.json();
    if (out.ok) {
      showToast(`✅ ${out.message}`);
      await init();
    } else {
      showToast(`❌ Fout: ${out.error}`);
    }
  };
  reader.readAsDataURL(file);
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.style.display = 'block';
  setTimeout(() => { t.style.display = 'none'; }, 4200);
}

function populateCampaignControls() {
  document.getElementById('camp-date').value = SETTINGS.ingangsdatum_iso || '2026-06-01';
  const nlSel = document.getElementById('camp-tpl-nl');
  const frSel = document.getElementById('camp-tpl-fr');

  nlSel.innerHTML = TEMPLATES.map(t =>
    `<option value="${t.filename}" ${t.filename === SETTINGS.template_nl ? 'selected' : ''}>${t.naam}</option>`
  ).join('');
  frSel.innerHTML = TEMPLATES.map(t =>
    `<option value="${t.filename}" ${t.filename === SETTINGS.template_fr ? 'selected' : ''}>${t.naam}</option>`
  ).join('');

  document.getElementById('stat-tpls').textContent = `${TEMPLATES.length} Templates`;
  document.getElementById('date-vars-preview').innerHTML =
    `Actieve variabelen: <code>{{ingangsdatum_lang}}</code> = <strong>${DATE_VARS.ingangsdatum_lang_nl}</strong> (FR: <strong>${DATE_VARS.ingangsdatum_lang_fr}</strong>) • ` +
    `<code>{{ingangsdatum_kort}}</code> = <strong>${DATE_VARS.ingangsdatum_kort}</strong> • ` +
    `Excel & Navision datum = <strong>${DATE_VARS.ingangsdatum_ddmmyyyy}</strong> (Bestandscode: <code>${DATE_VARS.datum_code}</code>)`;
}

async function saveCampaignSettings() {
  SETTINGS.ingangsdatum_iso = document.getElementById('camp-date').value;
  SETTINGS.template_nl = document.getElementById('camp-tpl-nl').value;
  SETTINGS.template_fr = document.getElementById('camp-tpl-fr').value;
  const res = await fetch('/api/save_settings', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(SETTINGS)
  });
  const out = await res.json();
  DATE_VARS = out.date_vars;
  populateCampaignControls();
  showToast(`✅ Prijsronde-instellingen bijgewerkt (Vanaf ${DATE_VARS.ingangsdatum_ddmmyyyy})`);
}

function updateStats() {
  const active = CUSTOMERS.filter(c => c.actief);
  const rood = active.filter(c => c.groep === 'STANDAARD_ROOD').length;
  const blauw = active.filter(c => c.groep === 'CATEGORIE_BLAUW').length;
  const spec = active.filter(c => c.groep === 'SPECIALS').length;
  document.getElementById('stat-active').textContent = active.length;
  document.getElementById('stat-breakdown').textContent = `${rood} Rood • ${blauw} Blauw • ${spec} Specials`;
}

function groupBadge(grp) {
  if (grp === 'STANDAARD_ROOD') return '<span class="badge badge-rood">Standaard Rood</span>';
  if (grp === 'CATEGORIE_BLAUW') return '<span class="badge badge-blauw">Blauw (-6%)</span>';
  if (grp === 'SPECIALS') return '<span class="badge badge-special">Specials</span>';
  return `<span class="badge badge-inactief">${grp}</span>`;
}

function renderTable() {
  const q = document.getElementById('search-input').value.toLowerCase().trim();
  const gf = document.getElementById('group-filter').value;
  const lf = document.getElementById('lang-filter').value;

  const filtered = CUSTOMERS.filter(c => {
    if (gf === 'ACTIEF' && !c.actief) return false;
    if (gf === 'STANDAARD_ROOD' && c.groep !== 'STANDAARD_ROOD') return false;
    if (gf === 'CATEGORIE_BLAUW' && c.groep !== 'CATEGORIE_BLAUW') return false;
    if (gf === 'SPECIALS' && c.groep !== 'SPECIALS') return false;
    if (gf === 'INACTIEF' && c.actief) return false;
    if (lf !== 'ALL' && c.taal !== lf) return false;
    if (q) {
      const hay = `${c.klantnaam} ${c.klantnr} ${c.contactpersoon || ''} ${c.aanspreking || ''} ${c.opmerking} ${c.email}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  const tbody = document.getElementById('customers-tbody');
  tbody.innerHTML = filtered.map(c => {
    const rpPct = (c.korting_rp * 100).toFixed(2).replace('.00', '') + '%';
    const bandTxt = c.band_op_aanvraag ? 'Op aanvr.' : (
      c.korting_band_stuk > 0
        ? `(1)${(c.korting_band_stuk*100).toFixed(0)}% / (VP)${(c.korting_band_pallet*100).toFixed(0)}%`
        : `(VP) ${(c.korting_band_pallet*100).toFixed(0)}%`
    );
    const epTxt = (c.korting_ep_lt10 > 0 || c.korting_ep_ge10 > 0)
      ? `${(c.korting_ep_lt10*100).toFixed(1).replace('.0','')}% / ${(c.korting_ep_ge10*100).toFixed(1).replace('.0','')}%`
      : '-';

    let flags = [];
    if (c.toebeh_palletprijzen) flags.push('<span class="flag-pill">Toebeh. Pallet</span>');
    if (c.special_vpu) flags.push('<span class="flag-pill">Spec. VPU</span>');
    if (c.meha_mode !== 'STANDAARD') flags.push(`<span class="flag-pill">MEHA ${c.meha_mode}</span>`);
    if (c.pk_apart) flags.push('<span class="flag-pill">PK/PP/Sloten</span>');
    if (c.special_pres) flags.push('<span class="flag-pill">RP PRES</span>');
    if (c.special_verpakt) flags.push('<span class="flag-pill">RP VERPAKT</span>');
    if (c.special_vpb) flags.push('<span class="flag-pill">VPB</span>');
    if (c.special_schets) flags.push('<span class="flag-pill">RP SCHETS</span>');
    if (c.geen_meerkost_7016) flags.push('<span class="flag-pill" style="background:#dcfce7;color:#166534;">7016 = 6005</span>');
    if (c.opmerking) flags.push(`<div style="font-size:11px;color:#64748b;margin-top:2px;">${c.opmerking}</div>`);

    const aansprPreview = (c.aanspreking || '').replace(/\n/g, ' • ');

    return `<tr>
      <td style="font-family:monospace;font-weight:600;color:#475569;">${c.klantnr || '-'}</td>
      <td>
        <strong>${c.klantnaam}</strong>
        <div style="font-size:11px;color:#166534;">✉️ ${aansprPreview}</div>
      </td>
      <td>${groupBadge(c.groep)}</td>
      <td><span class="badge ${c.taal === 'FR' ? 'badge-fr' : 'badge-nl'}">${c.taal}</span></td>
      <td style="font-weight:700;color:${c.korting_rp < 0 ? '#dc2626' : '#0f172a'};">${rpPct}</td>
      <td>${bandTxt}</td>
      <td style="font-weight:600;">${epTxt}</td>
      <td>${flags.join(' ') || '<span style="color:#94a3b8;">Standaard</span>'}</td>
      <td><strong>${c.tabbladen.length}</strong> bladen</td>
      <td style="text-align:right;white-space:nowrap;">
        <button class="btn-secondary btn-sm" onclick="openEditModal('${c.id}')">⚙️ Voorwaarden</button>
        <button class="btn-primary btn-sm" onclick="generateSingle('${c.id}')">⚡ Genereer</button>
      </td>
    </tr>`;
  }).join('');
}

function autoUpdateGreeting() {
  const contact = document.getElementById('edit-contact').value.trim();
  const klantnaam = document.getElementById('edit-klantnaam').value.trim().replace(' (SPECIALS)', '');
  const lang = document.getElementById('edit-taal').value;
  if (contact) {
    const parts = contact.split(/[,&/]|\s+en\s+|\s+et\s+/i).map(s => s.trim()).filter(Boolean);
    const prefix = lang === 'FR' ? 'Bonjour' : 'Goedemiddag';
    document.getElementById('edit-aanspreking').value = parts.map(p => `${prefix} ${p},`).join(' / ');
  } else {
    document.getElementById('edit-aanspreking').value = lang === 'FR'
      ? `Bonjour l'équipe de ${klantnaam},`
      : `Beste team van ${klantnaam},`;
  }
}

function openEditModal(id) {
  const c = CUSTOMERS.find(x => x.id === id);
  if (!c) return;
  document.getElementById('edit-id').value = c.id;
  document.getElementById('modal-title').textContent = `Voorwaarden & Aanspreking: ${c.klantnaam}`;
  document.getElementById('edit-klantnaam').value = c.klantnaam;
  document.getElementById('edit-klantnr').value = c.klantnr;
  document.getElementById('edit-contact').value = c.contactpersoon || '';
  document.getElementById('edit-email').value = c.email || '';
  document.getElementById('edit-aanspreking').value = (c.aanspreking || '').replace(/\n/g, ' / ');
  document.getElementById('edit-taal').value = c.taal;
  document.getElementById('edit-rp').value = +(c.korting_rp * 100).toFixed(2);
  document.getElementById('edit-ep1').value = +(c.korting_ep_lt10 * 100).toFixed(2);
  document.getElementById('edit-ep10').value = +(c.korting_ep_ge10 * 100).toFixed(2);
  document.getElementById('edit-band-vp').value = +(c.korting_band_pallet * 100).toFixed(2);
  document.getElementById('edit-meha').value = c.meha_mode;
  document.getElementById('edit-toebeh-pal').checked = c.toebeh_palletprijzen;
  document.getElementById('edit-spec-vpu').checked = c.special_vpu;
  document.getElementById('edit-ral7016').checked = c.geen_meerkost_7016;
  document.getElementById('edit-actief').checked = c.actief;

  const tabsCont = document.getElementById('edit-tabs-container');
  tabsCont.innerHTML = ALL_TABS.map(t => {
    const checked = c.tabbladen.includes(t.code) ? 'checked' : '';
    return `<label><input type="checkbox" class="tab-chk" value="${t.code}" ${checked}> <strong>${t.code}</strong></label>`;
  }).join('');

  document.getElementById('edit-modal').classList.add('open');
}

function closeModal() {
  document.getElementById('edit-modal').classList.remove('open');
}

async function saveCustomer() {
  const id = document.getElementById('edit-id').value;
  const c = CUSTOMERS.find(x => x.id === id);
  if (!c) return;
  c.klantnaam = document.getElementById('edit-klantnaam').value.trim();
  c.klantnr = document.getElementById('edit-klantnr').value.trim();
  c.contactpersoon = document.getElementById('edit-contact').value.trim();
  c.email = document.getElementById('edit-email').value.trim();
  c.aanspreking = document.getElementById('edit-aanspreking').value.split('/').map(s => s.trim()).join('\n');
  c.taal = document.getElementById('edit-taal').value;
  c.korting_rp = parseFloat(document.getElementById('edit-rp').value || '0') / 100.0;
  c.korting_ep_lt10 = parseFloat(document.getElementById('edit-ep1').value || '0') / 100.0;
  c.korting_ep_ge10 = parseFloat(document.getElementById('edit-ep10').value || '0') / 100.0;
  c.korting_band_pallet = parseFloat(document.getElementById('edit-band-vp').value || '0') / 100.0;
  c.meha_mode = document.getElementById('edit-meha').value;
  c.toebeh_palletprijzen = document.getElementById('edit-toebeh-pal').checked;
  c.special_vpu = document.getElementById('edit-spec-vpu').checked;
  c.geen_meerkost_7016 = document.getElementById('edit-ral7016').checked;
  c.actief = document.getElementById('edit-actief').checked;
  c.tabbladen = Array.from(document.querySelectorAll('.tab-chk:checked')).map(el => el.value);
  c.pk_apart = c.tabbladen.includes('PK');

  closeModal();
  const res = await fetch('/api/save_and_generate', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(c)
  });
  const out = await res.json();
  updateStats();
  renderTable();
  showToast(`✅ Opgeslagen & gegenereerd: ${out.xlsx_file} + ${out.eml_file}`);
}

/* Markdown Template Editor functies */
function openTemplateModal() {
  const sel = document.getElementById('tpl-select');
  sel.innerHTML = TEMPLATES.map(t => `<option value="${t.filename}">[${t.taal}] ${t.naam}</option>`).join('');

  const custSel = document.getElementById('preview-customer-select');
  const activeCusts = CUSTOMERS.filter(c => c.actief);
  custSel.innerHTML = activeCusts.map(c =>
    `<option value="${c.id}" ${c.klantnaam.includes('ANMABETON') ? 'selected' : ''}>${c.klantnaam} (${c.taal})</option>`
  ).join('');

  loadSelectedTemplateIntoEditor();
  document.getElementById('tpl-modal').classList.add('open');
}

function closeTemplateModal() {
  document.getElementById('tpl-modal').classList.remove('open');
}

function loadSelectedTemplateIntoEditor() {
  const fn = document.getElementById('tpl-select').value;
  const tpl = TEMPLATES.find(t => t.filename === fn) || TEMPLATES[0];
  if (!tpl) return;
  document.getElementById('tpl-filename').value = tpl.filename;
  document.getElementById('tpl-naam').value = tpl.naam;
  document.getElementById('tpl-taal').value = tpl.taal;
  document.getElementById('tpl-onderwerp').value = tpl.onderwerp;
  document.getElementById('tpl-body').value = tpl.body;
  updateLiveMailPreview();
}

function newBlankTemplate() {
  document.getElementById('tpl-filename').value = '';
  document.getElementById('tpl-naam').value = 'Nieuwe Mailtemplate';
  document.getElementById('tpl-taal').value = 'NL';
  document.getElementById('tpl-onderwerp').value = 'Prijslijst vanaf {{ingangsdatum_kort}} - {{klantnaam}}';
  document.getElementById('tpl-body').value = `{{aanspreking}}\n\nHierbij bezorgen wij u onze nieuwe prijslijst, geldig vanaf **{{ingangsdatum_lang}}**.\n\n- **Bijlage:** \`{{bijlage_naam}}\`\n\n{{handtekening}}`;
  document.getElementById('tpl-naam').focus();
  updateLiveMailPreview();
}

function insertVar(varText) {
  const ta = document.getElementById('tpl-body');
  const start = ta.selectionStart;
  const end = ta.selectionEnd;
  ta.value = ta.value.substring(0, start) + varText + ta.value.substring(end);
  ta.focus();
  ta.selectionStart = ta.selectionEnd = start + varText.length;
  updateLiveMailPreview();
}

async function updateLiveMailPreview() {
  const cid = document.getElementById('preview-customer-select').value;
  const onderwerp = document.getElementById('tpl-onderwerp').value;
  const body = document.getElementById('tpl-body').value;
  const taal = document.getElementById('tpl-taal').value;
  const rawMd = `---\nnaam: Preview\ntaal: ${taal}\nonderwerp: ${onderwerp}\n---\n${body}`;

  const res = await fetch('/api/preview_mail', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ customer_id: cid, raw_markdown: rawMd })
  });
  const out = await res.json();
  document.getElementById('prev-subj').textContent = out.subject;
  document.getElementById('prev-att').textContent = out.attachment;
  document.getElementById('prev-html').innerHTML = out.html;
}

async function saveTemplateFromEditor() {
  const payload = {
    filename: document.getElementById('tpl-filename').value.trim(),
    naam: document.getElementById('tpl-naam').value.trim(),
    taal: document.getElementById('tpl-taal').value,
    onderwerp: document.getElementById('tpl-onderwerp').value.trim(),
    body: document.getElementById('tpl-body').value
  };
  const res = await fetch('/api/save_template', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload)
  });
  const out = await res.json();
  TEMPLATES = out.templates;
  populateCampaignControls();
  document.getElementById('tpl-filename').value = out.saved;
  document.getElementById('tpl-select').innerHTML = TEMPLATES.map(t =>
    `<option value="${t.filename}" ${t.filename === out.saved ? 'selected' : ''}>[${t.taal}] ${t.naam}</option>`
  ).join('');
  showToast(`✅ Markdown template opgeslagen: ${payload.naam}`);
}

async function generateSingle(id) {
  document.getElementById('status-text').textContent = 'Bezig met genereren...';
  const res = await fetch(`/api/generate_single?id=${encodeURIComponent(id)}`, { method: 'POST' });
  const out = await res.json();
  document.getElementById('status-text').textContent = `Laatst gegenereerd: ${out.xlsx_file}`;
  showToast(`✅ Gegenereerd: ${out.xlsx_file} + Outlook .eml (${out.elapsed_ms} ms)`);
}

async function generateAllActive() {
  const btn = document.getElementById('btn-gen-all');
  btn.disabled = true;
  btn.textContent = '⏳ Bezig...';
  document.getElementById('status-text').textContent = 'Bulk-generatie gestart...';
  const res = await fetch('/api/generate_all', { method: 'POST' });
  const out = await res.json();
  btn.disabled = false;
  btn.textContent = '⚡ Genereer';
  document.getElementById('status-text').textContent = `✅ ${out.count} prijslijsten + ${out.count} .eml mails klaar (${out.elapsed_sec}s)`;
  showToast(`🎉 ${out.count} prijslijsten + ${out.count} Outlook mails gegenereerd in ${out.elapsed_sec}s!`);
}

async function generateNavision() {
  const res = await fetch('/api/generate_navision', { method: 'POST' });
  const out = await res.json();
  showToast(`📊 Navision importbestand aangemaakt: ${out.file}`);
}

async function exportMasterExcel() {
  const res = await fetch('/api/export_master', { method: 'POST' });
  const out = await res.json();
  showToast(`📋 Master Config Excel bijgewerkt: ${out.file}`);
}

async function openOutputFolder() {
  await fetch('/api/open_folder', { method: 'POST' });
}

init();
</script>
</body>
</html>
""".replace("__APP_VERSION__", APP_VERSION)


JUST_UPDATED_FILES = []


class ReusableHTTPServer(HTTPServer):
    allow_reuse_address = True


class StudioHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def _send_json(self, obj, status=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        global JUST_UPDATED_FILES
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = HTML_UI.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif parsed.path == "/assets/vdm_logo.png":
            if os.path.exists(LOGO_PATH):
                with open(LOGO_PATH, "rb") as f:
                    img = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Content-Length", str(len(img)))
                self.end_headers()
                self.wfile.write(img)
            else:
                self.send_error(404)
        elif parsed.path == "/api/customers":
            customers = load_config_json()
            settings = load_campaign_settings()
            templates = list_mail_templates()
            dvars = get_date_variables(settings.get("ingangsdatum_iso", "2026-06-01"))
            fstatus = get_files_status()
            all_tabs = [{"code": k, "label": v} for k, v in ALL_TAB_LABELS.items()]
            updated_now = list(JUST_UPDATED_FILES)
            JUST_UPDATED_FILES = []
            self._send_json({
                "customers": customers,
                "all_tabs": all_tabs,
                "settings": settings,
                "templates": templates,
                "date_vars": dvars,
                "files_status": fstatus,
                "just_updated": updated_now,
            })
        else:
            self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/save_settings":
            length = int(self.headers.get("Content-Length", "0"))
            new_settings = json.loads(self.rfile.read(length).decode("utf-8"))
            save_campaign_settings(new_settings)
            dvars = get_date_variables(new_settings.get("ingangsdatum_iso", "2026-06-01"))
            self._send_json({"ok": True, "date_vars": dvars})

        elif parsed.path == "/api/save_template":
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            saved_fn = save_mail_template(
                payload.get("filename", "template.md"),
                payload.get("naam", "Template"),
                payload.get("taal", "NL"),
                payload.get("onderwerp", "Prijslijst {{ingangsdatum_kort}}"),
                payload.get("body", ""),
            )
            self._send_json({"ok": True, "saved": saved_fn, "templates": list_mail_templates()})

        elif parsed.path == "/api/preview_mail":
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            customers = load_config_json()
            if not customers:
                cust = {
                    "id": "DEMO",
                    "klantnaam": "VOORBEELD KLANT BV",
                    "klantnr": "K00000001",
                    "taal": "NL",
                    "contactpersoon": "Jan",
                    "aanspreking": "Goedemiddag Jan,",
                }
            else:
                cust = next((c for c in customers if c["id"] == payload.get("customer_id")), customers[0])
            settings = load_campaign_settings()
            dvars = get_date_variables(settings.get("ingangsdatum_iso", "2026-06-01"))
            lang_tag = "_FRANS" if cust.get("taal") == "FR" else ""
            att_name = f"Art. Prijzen_{cust['klantnaam']}{lang_tag}_{dvars['datum_code']}.xlsx"
            subj, html_body, _ = render_customer_mail_content(
                cust, att_name, settings=settings, template_override=payload.get("raw_markdown")
            )
            html_preview = html_body.replace(
                'src="cid:vdm_logo_001"',
                'src="/assets/vdm_logo.png"'
            )
            self._send_json({"subject": subj, "attachment": att_name, "html": html_preview})

        elif parsed.path == "/api/upload_file":
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            ftype = payload.get("file_type")
            fname = os.path.basename(payload.get("filename", "bestand.xlsx"))
            raw_bytes = base64.b64decode(payload.get("content_base64", ""))

            os.makedirs(RAW_DIR, exist_ok=True)
            settings = load_campaign_settings()

            if ftype == "bron_nl":
                dest = os.path.join(RAW_DIR, fname)
                with open(dest, "wb") as f:
                    f.write(raw_bytes)
                settings["bron_nl"] = fname
                save_campaign_settings(settings)
                self._send_json({"ok": True, "message": f"Nederlandse Bronfile '{fname}' opgeslagen in confidential/!"})

            elif ftype == "bron_fr":
                dest = os.path.join(RAW_DIR, fname)
                with open(dest, "wb") as f:
                    f.write(raw_bytes)
                settings["bron_fr"] = fname
                save_campaign_settings(settings)
                self._send_json({"ok": True, "message": f"Franstalige Bronfile '{fname}' opgeslagen in confidential/!"})

            elif ftype == "master_config":
                # Check of het een 00_MASTER_KLANTENCONFIG.xlsx is of het oude Overzicht speciale voorwaarden
                with open(MASTER_EXCEL_PATH, "wb") as f:
                    f.write(raw_bytes)
                custs = parse_master_klantenconfig_complete(MASTER_EXCEL_PATH)
                if not custs:
                    # Misschien uploadde Nick zijn originele 'Overzicht speciale voorwaarden klanten.xlsx'
                    with open(VOORWAARDEN_PATH, "wb") as f:
                        f.write(raw_bytes)
                    custs = parse_voorwaarden_excel(VOORWAARDEN_PATH)
                    generate_master_klantenconfig_excel(custs, MASTER_EXCEL_PATH)
                save_config_json(custs, CONFIG_JSON_PATH)
                self._send_json({"ok": True, "message": f"Klantenconfiguratie ingeladen ({len(custs)} klanten)!"})
            else:
                self._send_json({"ok": False, "error": "Onbekend bestandstype"}, 400)

        elif parsed.path == "/api/generate_single":
            qs = parse_qs(parsed.query)
            cid = qs.get("id", [""])[0]
            customers = load_config_json()
            cust = next((c for c in customers if c["id"] == cid), None)
            if not cust:
                self._send_json({"error": "Klant niet gevonden"}, 404)
                return
            t0 = time.time()
            settings = load_campaign_settings()
            xlsx_path = generate_customer_workbook(cust, settings=settings)
            eml_path = generate_customer_eml(cust, xlsx_path, settings=settings)
            audit = audit_workbook_security(xlsx_path)
            ms = int((time.time() - t0) * 1000)
            self._send_json({
                "xlsx_file": os.path.basename(xlsx_path),
                "eml_file": os.path.basename(eml_path),
                "audit": audit,
                "elapsed_ms": ms,
            })

        elif parsed.path == "/api/save_and_generate":
            length = int(self.headers.get("Content-Length", "0"))
            updated = json.loads(self.rfile.read(length).decode("utf-8"))
            customers = load_config_json()
            for i, c in enumerate(customers):
                if c["id"] == updated["id"]:
                    customers[i] = updated
                    break
            save_config_json(customers)
            generate_master_klantenconfig_excel(customers)
            settings = load_campaign_settings()
            xlsx_path = generate_customer_workbook(updated, settings=settings)
            eml_path = generate_customer_eml(updated, xlsx_path, settings=settings)
            self._send_json({
                "xlsx_file": os.path.basename(xlsx_path),
                "eml_file": os.path.basename(eml_path),
            })

        elif parsed.path == "/api/generate_all":
            t0 = time.time()
            settings = load_campaign_settings()
            customers = [c for c in load_config_json() if c.get("actief")]
            for cust in customers:
                xlsx_path = generate_customer_workbook(cust, settings=settings)
                generate_customer_eml(cust, xlsx_path, settings=settings)
            generate_navision_export(customers, settings=settings)
            sec = round(time.time() - t0, 1)
            self._send_json({"count": len(customers), "elapsed_sec": sec})

        elif parsed.path == "/api/generate_navision":
            settings = load_campaign_settings()
            customers = [c for c in load_config_json() if c.get("actief")]
            nav_path = generate_navision_export(customers, settings=settings)
            self._send_json({"file": os.path.basename(nav_path)})

        elif parsed.path == "/api/export_master":
            customers = load_config_json()
            m_path = generate_master_klantenconfig_excel(customers)
            self._send_json({"file": os.path.basename(m_path)})

        elif parsed.path == "/api/check_updates":
            res = check_and_apply_updates()
            self._send_json(res)

        elif parsed.path == "/api/open_folder":
            os.makedirs(OUT_DIR, exist_ok=True)
            if sys.platform == "darwin":
                subprocess.Popen(["open", OUT_DIR])
            elif sys.platform.startswith("win"):
                subprocess.Popen(["explorer", OUT_DIR])
            self._send_json({"ok": True})
        else:
            self.send_error(404)


def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0


def main():
    global JUST_UPDATED_FILES

    # Als we herstart zijn na een silent update, lees welke bestanden geüpdatet zijn
    for arg in sys.argv[1:]:
        if arg.startswith("--updated="):
            JUST_UPDATED_FILES = [x for x in arg.split("=", 1)[1].split(",") if x]

    # Silent auto-update check bij het opstarten (alleen als we nog niet net herstart zijn)
    if not JUST_UPDATED_FILES:
        res = check_and_apply_updates()
        if res.get("ok") and res.get("updated"):
            upd_str = ",".join(res["updated"])
            # Herstart het proces direct met de nieuw binnengehaalde code
            os.execv(sys.executable, [sys.executable, os.path.abspath(__file__), f"--updated={upd_str}"])

    if not os.path.exists(MASTER_EXCEL_PATH) and os.path.exists(VOORWAARDEN_PATH):
        customers = parse_voorwaarden_excel(VOORWAARDEN_PATH)
        save_config_json(customers)
        generate_master_klantenconfig_excel(customers)

    url = f"http://127.0.0.1:{PORT}"
    if is_port_in_use(PORT):
        print("=" * 68)
        print(f"  VDM Prijslijst Studio draait al op {url}!")
        print("  Het scherm wordt nu geopend in je browser...")
        print("=" * 68)
        try:
            webbrowser.open(url)
        except Exception:
            pass
        return

    print("=" * 68)
    print(f"  VDM Prijslijst Studio v{APP_VERSION} (100% Lokaal & Offline)")
    print(f"  Open in je browser: {url}")
    print("  Druk op Ctrl+C om af te sluiten.")
    print("=" * 68)
    try:
        webbrowser.open(url)
    except Exception:
        pass
    server = ReusableHTTPServer(("127.0.0.1", PORT), StudioHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nAfgesloten.")


if __name__ == "__main__":
    main()
