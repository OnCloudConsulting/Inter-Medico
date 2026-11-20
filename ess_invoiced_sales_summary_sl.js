/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 *
 * Executive Invoiced Sales Summary
 * ---------------------------------------------------------------------------
 * Blends two sources into one summary, so totals run across go-live:
 *   1. Historic (pre-go-live) daily sales: a CSV in the File Cabinet, exported
 *      from the client's Excel workbook (Date, Daily Sales).
 *   2. Live sales: the client's saved search (default internal ID 3244).
 *
 * Any date before the go-live date comes from the CSV; any date on or after
 * it comes from the saved search, so a day is never counted twice.
 *
 * Pages (all GET):
 *   (default)          summary: Daily | MTD | LY MTD | QTD | LY QTD | YTD | LY YTD + variances
 *   action=detail      day-by-day breakdown for one period (click any amount)
 *   action=txn         saved-search rows for a date range (click a NetSuite day)
 *   &export=csv        CSV download of the page being viewed
 *   &asof=YYYY-MM-DD   report date (default: business day before today)
 *   view=portlet       compact, chrome-free summary for the dashboard portlet (iframe);
 *                      every link opens the full report in the main window
 *
 * Script parameters (all optional; defaults in DEFAULTS):
 *   custscript_ess_search_id     Saved search internal ID or script ID   (3244)
 *   custscript_ess_hist_file     File Cabinet path or internal ID of the historic CSV
 *   custscript_ess_golive_date   First date that comes from NetSuite (Date). Blank = day after last CSV date
 *   custscript_ess_amount_label  Label of the amount column in the saved search. Blank = auto-detect
 *   custscript_ess_date_field    Date field to filter/group on                (trandate)
 *   custscript_ess_title         Page title                                   (Invoiced Sales Summary)
 */
define(['N/search', 'N/file', 'N/runtime', 'N/format', 'N/url', 'N/ui/serverWidget', 'N/log'],
    (search, file, runtime, format, url, serverWidget, log) => {

        const PARAM = {
            SEARCH: 'custscript_ess_search_id',
            HIST_FILE: 'custscript_ess_hist_file',
            GO_LIVE: 'custscript_ess_golive_date',
            AMOUNT_LABEL: 'custscript_ess_amount_label',
            DATE_FIELD: 'custscript_ess_date_field',
            TITLE: 'custscript_ess_title'
        };

        const DEFAULTS = {
            SEARCH: '3244',
            HIST_FILE: 'SuiteScripts/ExecSalesSummary/historic_daily_sales.csv',
            DATE_FIELD: 'trandate',
            TITLE: 'Invoiced Sales Summary'
        };

        const TXN_PAGE_SIZE = 500;
        const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
            'August', 'September', 'October', 'November', 'December'];

        class UserError extends Error { }

        // =====================================================================
        // Entry point
        // =====================================================================
        const onRequest = (ctx) => {
            const params = ctx.request.parameters || {};
            const wantsCsv = params.export === 'csv';
            const isPortlet = params.view === 'portlet';
            let cfg = null;
            try {
                cfg = getConfig();
                const action = params.action || 'summary';
                const asOf = fromKey(params.asof) || prevBusinessDay(userToday());

                if (isPortlet) return renderPortlet(ctx, cfg, asOf);
                if (action === 'detail') return renderDetail(ctx, cfg, asOf, params, wantsCsv);
                if (action === 'txn') return renderTransactions(ctx, cfg, asOf, params, wantsCsv);
                return renderSummary(ctx, cfg, asOf, wantsCsv);
            } catch (e) {
                log.error({ title: 'Executive Sales Summary error', details: (e && e.stack) || e });
                const msg = (e instanceof UserError) ? e.message
                    : 'Something went wrong while building the report: ' + ((e && e.message) || e);
                if (isPortlet) {
                    return writeBare(ctx, `<div class="essp-err"><b>Sales summary unavailable.</b> ${esc(msg)}</div>`);
                }
                if (wantsCsv) {
                    ctx.response.setHeader({ name: 'Content-Type', value: 'text/plain' });
                    ctx.response.write(msg);
                    return;
                }
                writePage(ctx, (cfg && cfg.title) || DEFAULTS.TITLE,
                    `<div class="ess"><div class="ess-err"><b>Report unavailable.</b><br>${esc(msg)}</div></div>`);
            }
        };

        // =====================================================================
        // Pages
        // =====================================================================

        /** Main summary page, laid out like Invoiced_Sales_Summary_Layout.xlsx */
        function renderSummary(ctx, cfg, asOf, wantsCsv) {
            const P = buildPeriods(asOf);
            const data = getDailySales(cfg, P.lyytd.from, asOf);
            const t = {};
            Object.keys(P).forEach(k => { t[k] = sumRange(data, P[k].from, P[k].to); });

            const v = {
                mtdPct: pct(t.mtd, t.lymtd),
                qtdPct: pct(t.qtd, t.lyqtd),
                ytdAmt: t.ytd - t.lyytd,
                ytdPct: pct(t.ytd, t.lyytd)
            };

            if (wantsCsv) {
                const rows = [
                    ['Invoiced Sales Summary', 'As of ' + usDate(asOf)],
                    [],
                    ['Sales', 'Daily', 'MTD', 'LY MTD', 'MTD Var %', 'QTD', 'LY QTD', 'QTD Variance',
                        'YTD', 'LY YTD', 'YTD Sales Variance', 'YTD Var %'],
                    ['Total', r2(t.daily), r2(t.mtd), r2(t.lymtd), pctCsv(v.mtdPct), r2(t.qtd), r2(t.lyqtd),
                        pctCsv(v.qtdPct), r2(t.ytd), r2(t.lyytd), r2(v.ytdAmt), pctCsv(v.ytdPct)],
                    ['Range', rng(P.daily), rng(P.mtd), rng(P.lymtd), '', rng(P.qtd), rng(P.lyqtd), '',
                        rng(P.ytd), rng(P.lyytd), '', '']
                ];
                return writeCsv(ctx, `Invoiced_Sales_Summary_${key(asOf)}.csv`, rows);
            }

            const dl = (pid, val) => `<a href="${esc(link({ action: 'detail', p: pid, asof: key(asOf) }))}" title="${esc(P[pid].label + ' ' + rng(P[pid]))} – click for detail">${money(val, false, true)}</a>`;

            const card = (pid, lyPid, label) => {
                const p = lyPid ? pct(t[pid], t[lyPid]) : null;
                const cmp = lyPid
                    ? `${varChip(p)} <span>vs LY ${money(t[lyPid])}</span>`
                    : `<span>${esc(longDate(asOf))}</span>`;
                return `<a class="ess-card" href="${esc(link({ action: 'detail', p: pid, asof: key(asOf) }))}">
                    <div class="lbl">${esc(label)}</div>
                    <div class="val">${money(t[pid])}</div>
                    <div class="cmp">${cmp}</div></a>`;
            };

            const body = `
<div class="ess">
  ${header(cfg, 'INVOICED SALES SUMMARY', asOf, link({ asof: key(asOf), export: 'csv' }))}
  ${data.warnings.map(w => `<div class="ess-warn">${esc(w)}</div>`).join('')}
  <div class="ess-cards">
    ${card('daily', null, 'Daily')}
    ${card('mtd', 'lymtd', 'Month to Date')}
    ${card('qtd', 'lyqtd', 'Quarter to Date')}
    ${card('ytd', 'lyytd', 'Year to Date')}
  </div>
  <div class="ess-scroll">
  <table class="ess-t">
    <thead><tr>
      <th class="l">Sales</th><th>Daily</th><th>MTD</th><th>LY MTD</th><th class="vh">MTD Var %</th>
      <th>QTD</th><th>LY QTD</th><th class="vh">QTD Variance</th>
      <th>YTD</th><th>LY YTD</th><th class="vh">YTD Sales Variance</th><th class="vh">YTD Var %</th>
    </tr></thead>
    <tbody>
      <tr class="tot">
        <td class="l">Total</td>
        <td>${dl('daily', t.daily)}</td>
        <td>${dl('mtd', t.mtd)}</td><td>${dl('lymtd', t.lymtd)}</td><td class="vc">${pctCell(v.mtdPct)}</td>
        <td>${dl('qtd', t.qtd)}</td><td>${dl('lyqtd', t.lyqtd)}</td><td class="vc">${pctCell(v.qtdPct)}</td>
        <td>${dl('ytd', t.ytd)}</td><td>${dl('lyytd', t.lyytd)}</td>
        <td class="vc ${v.ytdAmt < 0 ? 'down' : 'up'}">${money(v.ytdAmt, true, true)}</td><td class="vc">${pctCell(v.ytdPct)}</td>
      </tr>
      <tr class="rng">
        <td class="l">Period</td>
        <td>${rng(P.daily)}</td><td>${rng(P.mtd)}</td><td>${rng(P.lymtd)}</td><td class="vc"></td>
        <td>${rng(P.qtd)}</td><td>${rng(P.lyqtd)}</td><td class="vc"></td>
        <td>${rng(P.ytd)}</td><td>${rng(P.lyytd)}</td><td class="vc"></td><td class="vc"></td>
      </tr>
    </tbody>
  </table>
  </div>
  ${footer(cfg, data)}
</div>`;
            writePage(ctx, cfg.title, body);
        }

        /**
         * Compact dashboard view (loaded in an iframe by the portlet script).
         * Rows = Daily / MTD / QTD / YTD; columns = CY, LY, Var $, Var %.
         * Amounts are abbreviated ($1.92M) with the exact figure on hover.
         */
        function renderPortlet(ctx, cfg, asOf) {
            const P = buildPeriods(asOf);
            const data = getDailySales(cfg, P.lyytd.from, asOf);
            const t = {};
            Object.keys(P).forEach(k => { t[k] = sumRange(data, P[k].from, P[k].to); });

            const top = (href, inner, title) =>
                `<a href="${esc(href)}" target="_top"${title ? ` title="${esc(title)}"` : ''}>${inner}</a>`;
            const amt = (pid) => top(link({ action: 'detail', p: pid, asof: key(asOf) }),
                short(t[pid]), `${P[pid].label} ${rng(P[pid])}: ${money(t[pid])}`);

            const row = (label, pid, lyPid) => {
                const diff = lyPid ? t[pid] - t[lyPid] : null;
                const p = lyPid ? pct(t[pid], t[lyPid]) : null;
                return `<tr>
                    <td class="l"><b>${label}</b><span class="rg">${rng(P[pid])}</span></td>
                    <td class="cy">${amt(pid)}</td>
                    <td>${lyPid ? amt(lyPid) : '<span class="muted">—</span>'}</td>
                    <td class="vd ${diff === null ? '' : (diff < 0 ? 'down' : 'up')}">${diff === null ? '<span class="muted">—</span>' : esc(short(diff, true))}</td>
                    <td>${lyPid ? varChip(p) : '<span class="muted">—</span>'}</td>
                </tr>`;
            };

            const warn = data.warnings.length
                ? `<div class="essp-warn" title="${esc(data.warnings.join(' '))}">⚠ Data warning – open the full report for details.</div>` : '';

            writeBare(ctx, `
<div class="essp-head">
  <span>As of <b>${esc(longDate(asOf))}</b></span>
  ${top(link({ asof: key(asOf) }), 'Full report ›')}
</div>
${warn}
<table class="essp-t">
  <thead><tr><th class="l">Period</th><th>Sales</th><th>LY</th><th class="vd">Var $</th><th>Var %</th></tr></thead>
  <tbody>
    ${row('Daily', 'daily', null)}
    ${row('MTD', 'mtd', 'lymtd')}
    ${row('QTD', 'qtd', 'lyqtd')}
    ${row('YTD', 'ytd', 'lyytd')}
  </tbody>
</table>`);
        }

        /** Day-by-day breakdown of one period, with month subtotals for long ranges. */
        function renderDetail(ctx, cfg, asOf, params, wantsCsv) {
            const P = buildPeriods(asOf);
            const per = P[params.p];
            if (!per) throw new UserError('Unknown period "' + (params.p || '') + '".');

            const lyMap = { mtd: 'lymtd', qtd: 'lyqtd', ytd: 'lyytd', lymtd: 'mtd', lyqtd: 'qtd', lyytd: 'ytd' };
            const other = lyMap[params.p] ? P[lyMap[params.p]] : null;
            const from = other && other.from < per.from ? other.from : per.from;
            const data = getDailySales(cfg, from, asOf);

            const days = listDays(data, per.from, per.to);
            const total = days.reduce((s, d) => s + d.amount, 0);
            const otherTotal = other ? sumRange(data, other.from, other.to) : null;

            if (wantsCsv) {
                const rows = [[per.label + ' ' + rng(per)], [], ['Date', 'Day', 'Source', 'Sales']];
                days.forEach(d => rows.push([usDate(d.date), DOW[d.date.getDay()], d.source === 'ns' ? 'NetSuite' : 'Historic (Excel)', r2(d.amount)]));
                rows.push(['Total', '', '', r2(total)]);
                return writeCsv(ctx, `Sales_${per.id}_${key(per.from)}_${key(per.to)}.csv`, rows);
            }

            const nsFrom = maxDate(per.from, data.goLive);
            const hasNs = nsFrom <= per.to;
            const multiMonth = per.from.getMonth() !== per.to.getMonth() || per.from.getFullYear() !== per.to.getFullYear();

            let rowsHtml = '';
            let curMonth = null, monthSum = 0, monthDays = 0;
            const flushMonth = () => {
                if (multiMonth && curMonth !== null) {
                    rowsHtml += `<tr class="sub"><td class="l" colspan="3">${esc(curMonth)} subtotal · ${monthDays} day${monthDays === 1 ? '' : 's'}</td><td>${money(monthSum)}</td><td></td></tr>`;
                }
            };
            days.forEach(d => {
                const m = MONTHS[d.date.getMonth()] + ' ' + d.date.getFullYear();
                if (m !== curMonth) { flushMonth(); curMonth = m; monthSum = 0; monthDays = 0; }
                monthSum += d.amount; monthDays++;
                const txn = d.source === 'ns'
                    ? `<a href="${esc(link({ action: 'txn', from: key(d.date), to: key(d.date), asof: key(asOf), back: params.p }))}">Transactions ›</a>`
                    : '<span class="muted">Daily total only</span>';
                rowsHtml += `<tr><td class="l">${usDate(d.date)}</td><td class="l">${DOW[d.date.getDay()]}</td>
                    <td class="l">${srcTag(d.source)}</td><td>${money(d.amount)}</td><td>${txn}</td></tr>`;
            });
            flushMonth();
            if (!days.length) rowsHtml = '<tr><td class="l muted" colspan="5">No sales recorded in this period.</td></tr>';

            const cmpHtml = other
                ? `<div class="ess-stat"><div class="lbl">${esc(other.label)} (${rng(other)})</div>
                   <div class="val sm"><a href="${esc(link({ action: 'detail', p: other.id, asof: key(asOf) }))}">${money(otherTotal)}</a></div></div>
                   <div class="ess-stat"><div class="lbl">Variance</div>
                   <div class="val sm">${per.id.indexOf('ly') === 0 ? varChip(pct(otherTotal, total)) + ' <span class="muted">(CY vs LY)</span>' : varChip(pct(total, otherTotal))}</div></div>`
                : '';

            const body = `
<div class="ess">
  <div class="crumbs"><a href="${esc(link({ asof: key(asOf) }))}">‹ Summary</a> / ${esc(per.label)}</div>
  ${header(cfg, per.label.toUpperCase() + ' DETAIL', null, link({ action: 'detail', p: per.id, asof: key(asOf), export: 'csv' }), rng(per))}
  ${data.warnings.map(w => `<div class="ess-warn">${esc(w)}</div>`).join('')}
  <div class="ess-stats">
    <div class="ess-stat"><div class="lbl">Total</div><div class="val">${money(total)}</div></div>
    <div class="ess-stat"><div class="lbl">Selling days</div><div class="val sm">${days.length}</div></div>
    <div class="ess-stat"><div class="lbl">Avg / selling day</div><div class="val sm">${money(days.length ? total / days.length : 0)}</div></div>
    ${cmpHtml}
  </div>
  ${hasNs ? `<p><a class="ess-btn" href="${esc(link({ action: 'txn', from: key(nsFrom), to: key(per.to), asof: key(asOf), back: per.id }))}">View all NetSuite transactions (${usDate(nsFrom)} – ${usDate(per.to)}) ›</a></p>` : ''}
  <div class="ess-scroll">
  <table class="ess-t">
    <thead><tr><th class="l">Date</th><th class="l">Day</th><th class="l">Source</th><th>Sales</th><th></th></tr></thead>
    <tbody>${rowsHtml}
      <tr class="tot"><td class="l" colspan="3">Total · ${days.length} day${days.length === 1 ? '' : 's'}</td><td>${money(total)}</td><td></td></tr>
    </tbody>
  </table>
  </div>
  ${footer(cfg, data)}
</div>`;
            writePage(ctx, cfg.title, body);
        }

        /** Saved-search rows for a date range (only the NetSuite portion of the range). */
        function renderTransactions(ctx, cfg, asOf, params, wantsCsv) {
            let from = fromKey(params.from), to = fromKey(params.to);
            if (!from || !to || from > to) throw new UserError('Invalid date range.');

            const hist = loadHistoric(cfg);
            const goLive = cfg.goLive || hist.defaultGoLive;
            if (to < goLive) throw new UserError(`${rngDates(from, to)} is before go-live (${usDate(goLive)}). Only daily totals exist for that period, in the historic Excel data.`);
            from = maxDate(from, goLive);

            const s = loadSearch(cfg);
            const cols = s.columns;
            const amtCol = findAmountColumn(cols, cfg.amountLabel);
            const isSummary = cols.some(c => !!c.summary);
            const dateWarn = dateFilterWarning(s, cfg.dateField);
            s.filterExpression = withDateFilter(s, cfg.dateField, from, to);

            const pageSize = wantsCsv ? 1000 : TXN_PAGE_SIZE;
            const paged = s.runPaged({ pageSize });
            const count = paged.count;
            const headers = cols.map(c => c.label || c.name);

            const readRow = (r) => cols.map(c => {
                let txt = null;
                try { txt = r.getText(c); } catch (e) { /* not a select field */ }
                const val = r.getValue(c);
                return (txt !== null && txt !== undefined && txt !== '') ? txt : (val === null || val === undefined ? '' : val);
            });

            if (wantsCsv) {
                const rows = [headers];
                paged.pageRanges.forEach(pr => {
                    paged.fetch({ index: pr.index }).data.forEach(r => rows.push(readRow(r)));
                });
                return writeCsv(ctx, `Transactions_${key(from)}_${key(to)}.csv`, rows);
            }

            const pageCount = paged.pageRanges.length;
            const page = Math.min(Math.max(parseInt(params.page, 10) || 0, 0), Math.max(pageCount - 1, 0));
            const results = pageCount ? paged.fetch({ index: page }).data : [];

            // Period total comes from the same search, summed by day (matches the summary page exactly).
            const nsTotal = Object.values(queryNetSuiteDaily(cfg, from, to)).reduce((a, b) => a + b, 0);

            const amtIdx = cols.indexOf(amtCol);
            const urlCache = {};
            const recUrl = (r) => {
                if (isSummary || !r.id || !r.recordType) return null;
                if (!(r.recordType in urlCache)) {
                    try { urlCache[r.recordType] = url.resolveRecord({ recordType: r.recordType, recordId: 0 }).replace(/id=0\b/, 'id={ID}'); }
                    catch (e) { urlCache[r.recordType] = '/app/accounting/transactions/transaction.nl?id={ID}'; }
                }
                return urlCache[r.recordType].replace('{ID}', encodeURIComponent(r.id));
            };

            let pageAmt = 0;
            const rowsHtml = results.map(r => {
                const vals = readRow(r);
                if (amtIdx >= 0) pageAmt += num(r.getValue(cols[amtIdx]));
                const u = recUrl(r);
                return '<tr>' + (isSummary ? '' : `<td class="l">${u ? `<a href="${esc(u)}" target="_blank">Open</a>` : ''}</td>`) +
                    vals.map((v, i) => {
                        const isNum = i === amtIdx || /^-?\d+(\.\d+)?$/.test(String(v));
                        return isNum && i === amtIdx ? `<td>${money(num(v))}</td>`
                            : `<td class="${isNum ? '' : 'l'}">${esc(v)}</td>`;
                    }).join('') + '</tr>';
            }).join('') || `<tr><td class="l muted" colspan="${headers.length + 1}">No transactions in this range.</td></tr>`;

            const base = { action: 'txn', from: key(from), to: key(to), asof: key(asOf), back: params.back || '' };
            const pager = pageCount > 1 ? `<div class="ess-pager">
                ${page > 0 ? `<a href="${esc(link(Object.assign({}, base, { page: page - 1 })))}">‹ Prev</a>` : '<span class="muted">‹ Prev</span>'}
                <span>Page ${page + 1} of ${pageCount}</span>
                ${page < pageCount - 1 ? `<a href="${esc(link(Object.assign({}, base, { page: page + 1 })))}">Next ›</a>` : '<span class="muted">Next ›</span>'}
              </div>` : '';

            const backLink = params.back && /^(daily|mtd|lymtd|qtd|lyqtd|ytd|lyytd)$/.test(params.back)
                ? ` / <a href="${esc(link({ action: 'detail', p: params.back, asof: key(asOf) }))}">${esc(buildPeriods(asOf)[params.back].label)}</a>` : '';

            const firstRow = count ? page * TXN_PAGE_SIZE + 1 : 0;
            const lastRow = Math.min((page + 1) * TXN_PAGE_SIZE, count);

            const body = `
<div class="ess">
  <div class="crumbs"><a href="${esc(link({ asof: key(asOf) }))}">‹ Summary</a>${backLink} / Transactions</div>
  ${header(cfg, 'TRANSACTIONS', null, link(Object.assign({}, base, { export: 'csv' })), rngDates(from, to))}
  ${dateWarn ? `<div class="ess-warn">${esc(dateWarn)}</div>` : ''}
  <div class="ess-stats">
    <div class="ess-stat"><div class="lbl">Total sales</div><div class="val">${money(nsTotal)}</div></div>
    <div class="ess-stat"><div class="lbl">${isSummary ? 'Result rows' : 'Lines'}</div><div class="val sm">${count.toLocaleString('en-US')}</div></div>
    ${pageCount > 1 ? `<div class="ess-stat"><div class="lbl">This page (${firstRow}–${lastRow})</div><div class="val sm">${money(pageAmt)}</div></div>` : ''}
  </div>
  ${pager}
  <div class="ess-scroll">
  <table class="ess-t ess-txn">
    <thead><tr>${isSummary ? '' : '<th class="l"></th>'}${headers.map((h, i) => `<th class="${i === amtIdx ? '' : 'l'}">${esc(h)}</th>`).join('')}</tr></thead>
    <tbody>${rowsHtml}</tbody>
  </table>
  </div>
  ${pager}
  <div class="ess-note">Rows from saved search ${esc(cfg.searchId)} (“${esc(s.title || '')}”), filtered to ${rngDates(from, to)} on ${esc(cfg.dateField)}.</div>
</div>`;
            writePage(ctx, cfg.title, body);
        }

        // =====================================================================
        // Data
        // =====================================================================

        /**
         * Daily sales for [from, to]: historic CSV before go-live, saved search on/after.
         * Returns { hist, ns, goLive, histMin, histMax, warnings }.
         */
        function getDailySales(cfg, from, to) {
            const hist = loadHistoric(cfg);
            const goLive = cfg.goLive || hist.defaultGoLive;
            const warnings = hist.warnings.slice();

            let ns = {};
            const nsFrom = maxDate(from, goLive);
            if (nsFrom <= to) {
                const w = dateFilterWarning(loadSearch(cfg), cfg.dateField);
                if (w) warnings.push(w);
                ns = queryNetSuiteDaily(cfg, nsFrom, to);
            }
            return { hist: hist.byDay, ns, goLive, histMin: hist.min, histMax: hist.max, warnings };
        }

        /** Parse the historic CSV (Date, Daily Sales). Accepts M/D/YYYY or YYYY-MM-DD; $ , and (neg) OK. */
        function loadHistoric(cfg) {
            let f;
            try {
                f = file.load({ id: cfg.histFile });
            } catch (e) {
                throw new UserError(`Historic sales file "${cfg.histFile}" could not be loaded. Upload the CSV to the File Cabinet and set the script parameter "${PARAM.HIST_FILE}" to its path or internal ID.`);
            }
            if (/\.xlsx?$/i.test(f.name || '')) {
                throw new UserError(`"${f.name}" is an Excel file. In Excel use File > Save As > CSV (Comma delimited), upload the .csv and point "${PARAM.HIST_FILE}" at it.`);
            }

            const byDay = {};
            let min = null, max = null, bad = 0, lineNo = 0;
            f.lines.iterator().each(line => {
                lineNo++;
                const raw = String(line.value || '').replace(/^﻿/, '').replace(/\r$/, '');
                if (!raw.trim()) return true;
                const cells = splitCsv(raw);
                const d = parseLooseDate(cells[0]);
                const a = parseAmount(cells[1]);
                if (!d || a === null) {
                    if (lineNo > 1) bad++;           // line 1 is the header
                    return true;
                }
                const k = key(d);
                byDay[k] = (byDay[k] || 0) + a;
                if (!min || d < min) min = d;
                if (!max || d > max) max = d;
                return true;
            });

            if (!max) throw new UserError(`No usable rows in "${f.name}". Expected two columns: Date (MM/DD/YYYY) and Daily Sales.`);
            const warnings = bad ? [`${bad} line(s) in the historic file could not be read and were skipped. Check "${f.name}" for blank dates or text in the amount column.`] : [];
            return { byDay, min, max, defaultGoLive: addDays(max, 1), warnings };
        }

        /** Sum of the saved search's amount column per day for [from, to]. Returns { 'YYYY-MM-DD': amount }. */
        function queryNetSuiteDaily(cfg, from, to) {
            const s = loadSearch(cfg);
            const amt = findAmountColumn(s.columns, cfg.amountLabel);

            const def = { name: amt.name, summary: search.Summary.SUM };
            if (amt.join) def.join = amt.join;
            if (amt.formula) def.formula = amt.formula;
            const sumCol = search.createColumn(def);
            const dateCol = search.createColumn({ name: cfg.dateField, summary: search.Summary.GROUP, sort: search.Sort.ASC });

            s.filterExpression = withDateFilter(s, cfg.dateField, from, to);
            s.columns = [dateCol, sumCol];

            const out = {};
            const paged = s.runPaged({ pageSize: 1000 });
            paged.pageRanges.forEach(pr => {
                paged.fetch({ index: pr.index }).data.forEach(r => {
                    const d = parseNsDate(r.getValue(dateCol));
                    if (!d) return;
                    const k = key(d);
                    out[k] = (out[k] || 0) + num(r.getValue(sumCol));
                });
            });
            return out;
        }

        function loadSearch(cfg) {
            try {
                return search.load({ id: cfg.searchId });
            } catch (e) {
                throw new UserError(`Saved search "${cfg.searchId}" could not be loaded (${e.message || e.name}). Check the "${PARAM.SEARCH}" parameter and that the search is shared with the executives' roles.`);
            }
        }

        /** The saved search's own criteria AND our date range. */
        function withDateFilter(s, field, from, to) {
            const dateExpr = [field, 'within', nsDate(from), nsDate(to)];
            const expr = s.filterExpression || [];
            return expr.length ? [expr, 'AND', dateExpr] : [dateExpr];
        }

        /** A date filter inside the saved search would clip MTD/QTD/YTD; flag it on screen. */
        function dateFilterWarning(s, field) {
            const hit = (s.filters || []).some(f => (f.name || '').toLowerCase() === field && !f.join);
            return hit ? `Saved search ${s.id || ''} has its own "${field}" filter. That filter still applies, so MTD/QTD/YTD may be understated. Remove the date criteria from the saved search; this page supplies the dates.` : null;
        }

        /** Which saved-search column holds the sales amount. */
        function findAmountColumn(cols, label) {
            const lc = s => String(s || '').trim().toLowerCase();
            if (label) {
                const c = cols.find(x => lc(x.label) === lc(label)) || cols.find(x => lc(x.name) === lc(label));
                if (c) return c;
                throw new UserError(`No column labelled "${label}" in the saved search. Available: ${cols.map(x => x.label || x.name).join(', ')}.`);
            }
            const known = ['amount', 'netamount', 'netamountnotax', 'fxamount', 'foreignamount', 'total', 'amountnet'];
            const c = cols.find(x => lc(x.summary) === 'sum')
                || cols.find(x => known.indexOf(lc(x.name)) >= 0 && !x.join)
                || cols.find(x => /^formula(currency|numeric)$/.test(lc(x.name)))
                || cols.find(x => /amount|sales|total/.test(lc(x.label)));
            if (c) return c;
            throw new UserError(`Couldn't tell which saved-search column is the sales amount. Set "${PARAM.AMOUNT_LABEL}" to its label. Columns: ${cols.map(x => x.label || x.name).join(', ')}.`);
        }

        function amountFor(data, k, d) {
            return d < data.goLive ? (data.hist[k] || 0) : (data.ns[k] || 0);
        }

        function sumRange(data, from, to) {
            let s = 0;
            for (let d = from; d <= to; d = addDays(d, 1)) s += amountFor(data, key(d), d);
            return s;
        }

        function listDays(data, from, to) {
            const out = [];
            for (let d = from; d <= to; d = addDays(d, 1)) {
                const k = key(d);
                const src = d < data.goLive ? 'hist' : 'ns';
                const map = src === 'hist' ? data.hist : data.ns;
                if (Object.prototype.hasOwnProperty.call(map, k)) out.push({ date: d, amount: map[k], source: src });
            }
            return out;
        }

        // =====================================================================
        // Periods & dates
        // =====================================================================

        function buildPeriods(asOf) {
            const y = asOf.getFullYear(), m = asOf.getMonth(), q = Math.floor(m / 3) * 3;
            const ly = sameDayLastYear(asOf);
            const mk = (id, label, from, to) => ({ id, label, from, to });
            return {
                daily: mk('daily', 'Daily', asOf, asOf),
                mtd: mk('mtd', 'MTD', new Date(y, m, 1), asOf),
                lymtd: mk('lymtd', 'LY MTD', new Date(y - 1, m, 1), ly),
                qtd: mk('qtd', 'QTD', new Date(y, q, 1), asOf),
                lyqtd: mk('lyqtd', 'LY QTD', new Date(y - 1, q, 1), ly),
                ytd: mk('ytd', 'YTD', new Date(y, 0, 1), asOf),
                lyytd: mk('lyytd', 'LY YTD', new Date(y - 1, 0, 1), ly)
            };
        }

        function sameDayLastYear(d) {
            const y = d.getFullYear() - 1, m = d.getMonth();
            const last = new Date(y, m + 1, 0).getDate();      // Feb 29 -> Feb 28
            return new Date(y, m, Math.min(d.getDate(), last));
        }

        function prevBusinessDay(d) {
            let x = addDays(d, -1);
            while (x.getDay() === 0 || x.getDay() === 6) x = addDays(x, -1);
            return x;
        }

        /** Today's date in the user's time zone (the server runs on Pacific time). */
        function userToday() {
            try {
                const tz = runtime.getCurrentUser().getPreference({ name: 'TIMEZONE' });
                const s = format.format({ value: new Date(), type: format.Type.DATETIME, timezone: tz });
                const d = format.parse({ value: s, type: format.Type.DATETIME });
                if (d instanceof Date && !isNaN(d.getTime())) return trunc(d);
            } catch (e) {
                log.debug({ title: 'userToday fallback', details: e.message });
            }
            return trunc(new Date());
        }

        const pad = n => (n < 10 ? '0' : '') + n;
        const trunc = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
        const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
        const maxDate = (a, b) => (a > b ? a : b);
        const key = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
        const usDate = d => `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
        const longDate = d => `${['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][d.getDay()]}, ${MONTHS[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
        const rngDates = (a, b) => (a.getTime() === b.getTime() ? usDate(a) : `${usDate(a)} – ${usDate(b)}`);
        const rng = p => rngDates(p.from, p.to);

        function fromKey(k) {
            const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(k || ''));
            if (!m) return null;
            const d = new Date(+m[1], +m[2] - 1, +m[3]);
            return d.getMonth() === +m[2] - 1 ? d : null;
        }

        function nsDate(d) {
            return format.format({ value: d, type: format.Type.DATE });
        }

        function parseNsDate(s) {
            if (!s) return null;
            try {
                const d = format.parse({ value: s, type: format.Type.DATE });
                return (d instanceof Date && !isNaN(d.getTime())) ? trunc(d) : null;
            } catch (e) {
                return parseLooseDate(s);
            }
        }

        function parseLooseDate(s) {
            s = String(s || '').trim();
            let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(s);
            if (m) {
                let y = +m[3]; if (y < 100) y += 2000;
                const d = new Date(y, +m[1] - 1, +m[2]);
                return d.getMonth() === +m[1] - 1 ? d : null;
            }
            m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
            if (m) {
                const d = new Date(+m[1], +m[2] - 1, +m[3]);
                return d.getMonth() === +m[2] - 1 ? d : null;
            }
            return null;
        }

        // =====================================================================
        // Parsing & formatting
        // =====================================================================

        function splitCsv(line) {
            const out = []; let cur = '', q = false;
            for (let i = 0; i < line.length; i++) {
                const ch = line[i];
                if (q) {
                    if (ch === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else q = false; }
                    else cur += ch;
                } else if (ch === '"') q = true;
                else if (ch === ',') { out.push(cur); cur = ''; }
                else cur += ch;
            }
            out.push(cur);
            return out;
        }

        function parseAmount(s) {
            s = String(s === undefined || s === null ? '' : s).trim();
            if (!s) return null;
            const neg = /^\(.*\)$/.test(s) || /^-/.test(s.replace(/[$\s]/g, ''));
            const n = parseFloat(s.replace(/[^0-9.]/g, ''));
            if (isNaN(n)) return null;
            return neg ? -n : n;
        }

        /** $1.92M / $92.8K / $512 – for the narrow dashboard portlet. */
        function short(n, signed) {
            const a = Math.abs(n);
            let s;
            if (a >= 1e6) s = (a / 1e6).toFixed(2) + 'M';
            else if (a >= 1e3) s = (a / 1e3).toFixed(1) + 'K';
            else s = a.toFixed(0);
            return (n < 0 ? '-$' : (signed && n > 0 ? '+$' : '$')) + s;
        }

        const num = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
        const r2 = n => (Math.round(n * 100) / 100).toFixed(2);
        const pct = (a, b) => (b ? (a - b) / Math.abs(b) : null);
        const pctCsv = p => (p === null ? '' : (p * 100).toFixed(1) + '%');

        function money(n, signed, noSymbol) {
            const abs = Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            const sym = noSymbol ? '' : '$';
            if (n < 0) return '-' + sym + abs;
            return (signed && n > 0 ? '+' : '') + sym + abs;
        }

        function pctCell(p) {
            if (p === null) return '<span class="muted">n/a</span>';
            return `<span class="${p < 0 ? 'down' : 'up'}">${p > 0 ? '+' : ''}${(p * 100).toFixed(1)}%</span>`;
        }

        function varChip(p) {
            if (p === null) return '<span class="chip">n/a</span>';
            return `<span class="chip ${p < 0 ? 'down' : 'up'}">${p < 0 ? '▼' : '▲'} ${Math.abs(p * 100).toFixed(1)}%</span>`;
        }

        const srcTag = s => (s === 'ns' ? '<span class="tag ns">NetSuite</span>' : '<span class="tag hist">Historic</span>');

        function esc(s) {
            return String(s === null || s === undefined ? '' : s)
                .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
        }

        // =====================================================================
        // Output
        // =====================================================================

        function getConfig() {
            const sc = runtime.getCurrentScript();
            const p = n => {
                const v = sc.getParameter({ name: n });
                return (v === null || v === undefined || v === '') ? null : v;
            };
            const gl = p(PARAM.GO_LIVE);
            let goLive = null;
            if (gl instanceof Date) goLive = trunc(gl);
            else if (gl) goLive = parseNsDate(String(gl));

            return {
                searchId: String(p(PARAM.SEARCH) || DEFAULTS.SEARCH).trim(),
                histFile: String(p(PARAM.HIST_FILE) || DEFAULTS.HIST_FILE).trim(),
                goLive,
                amountLabel: p(PARAM.AMOUNT_LABEL) ? String(p(PARAM.AMOUNT_LABEL)).trim() : null,
                dateField: String(p(PARAM.DATE_FIELD) || DEFAULTS.DATE_FIELD).trim().toLowerCase(),
                title: String(p(PARAM.TITLE) || DEFAULTS.TITLE),
                scriptId: sc.id,
                deploymentId: sc.deploymentId
            };
        }

        let _base = null;
        function link(params) {
            if (_base === null) {
                const sc = runtime.getCurrentScript();
                _base = url.resolveScript({ scriptId: sc.id, deploymentId: sc.deploymentId });
            }
            const q = Object.keys(params || {})
                .filter(k => params[k] !== undefined && params[k] !== null && params[k] !== '')
                .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k])).join('&');
            return q ? _base + (_base.indexOf('?') >= 0 ? '&' : '?') + q : _base;
        }

        function header(cfg, title, asOf, csvHref, subtitle) {
            const picker = asOf ? `
      <label class="asof">As of
        <input type="date" id="ess-asof" value="${key(asOf)}" onchange="if(this.value){window.location.href='${esc(link({}))}${link({}).indexOf('?') >= 0 ? '&' : '?'}asof='+this.value;}">
      </label>
      <a class="ess-btn ghost" href="${esc(link({}))}">Last business day</a>` : '';
            return `
  <div class="ess-head">
    <div>
      <h1>${esc(title)}</h1>
      <div class="sub">${asOf ? esc(longDate(asOf)) : esc(subtitle || '')}</div>
    </div>
    <div class="ess-actions">${picker}
      <a class="ess-btn ghost" href="${esc(csvHref)}">Export CSV</a>
    </div>
  </div>`;
        }

        function footer(cfg, data) {
            return `<div class="ess-note">Sources: historic daily sales from Excel (${usDate(data.histMin)} – ${usDate(addDays(data.goLive, -1))}); NetSuite saved search ${esc(cfg.searchId)} from ${usDate(data.goLive)}.
            LY figures compare the same calendar dates one year earlier. Click any amount for its daily breakdown; click a NetSuite day for the transactions.</div>`;
        }

        function writeCsv(ctx, name, rows) {
            const cell = v => {
                const s = String(v === null || v === undefined ? '' : v);
                return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
            };
            const contents = rows.map(r => r.map(cell).join(',')).join('\r\n');
            ctx.response.writeFile({ file: file.create({ name, fileType: file.Type.CSV, contents }), isInline: false });
        }

        /** Chrome-free HTML page for the dashboard iframe; reports its height to the portlet. */
        function writeBare(ctx, bodyHtml) {
            ctx.response.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top">${PORTLET_STYLE}</head>
<body><div class="essp" id="essp">${bodyHtml}</div>
<script>
(function(){function h(){try{window.parent.postMessage({essHeight:document.getElementById('essp').offsetHeight+4},'*');}catch(e){}}
window.addEventListener('load',h);window.addEventListener('resize',h);setTimeout(h,300);})();
</script></body></html>`);
        }

        const PORTLET_STYLE = `<style>
html,body{margin:0;padding:0;background:transparent}
.essp{font-family:-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:#1f2937;font-size:12.5px;padding:2px 2px 4px}
.essp a{color:#0b5cad;text-decoration:none}.essp a:hover{text-decoration:underline}
.essp .muted{color:#9ca3af}.essp .up{color:#047857}.essp .down{color:#b91c1c}
.essp-head{display:flex;justify-content:space-between;align-items:center;gap:8px;margin:0 0 8px;color:#6b7280;flex-wrap:wrap}
.essp-head b{color:#0b2545}
.essp-t{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums}
.essp-t th{background:#0b2545;color:#fff;font-weight:600;padding:6px 8px;text-align:right;font-size:11.5px;white-space:nowrap}
.essp-t th:first-child{border-radius:6px 0 0 0}.essp-t th:last-child{border-radius:0 6px 0 0}
.essp-t td{padding:7px 8px;border-bottom:1px solid #eef0f3;text-align:right;white-space:nowrap;vertical-align:middle}
.essp-t tr:last-child td{border-bottom:0}
.essp-t .l{text-align:left}
.essp-t td.cy a{font-weight:700;font-size:14px;color:#0b2545}
.essp-t td.vd{font-weight:600}
.essp-t .rg{display:block;font-size:10.5px;color:#9ca3af;font-weight:400;margin-top:1px}
.essp .chip{display:inline-block;padding:1px 6px;border-radius:999px;font-size:11.5px;font-weight:600;background:#f3f4f6;color:#6b7280}
.essp .chip.up{background:#ecfdf5;color:#047857}.essp .chip.down{background:#fef2f2;color:#b91c1c}
.essp-warn{background:#fffbeb;border:1px solid #fcd34d;color:#92400e;padding:5px 8px;border-radius:6px;margin-bottom:8px;font-size:11.5px}
.essp-err{background:#fef2f2;border:1px solid #fca5a5;color:#991b1b;padding:10px;border-radius:6px;line-height:1.45}
@media (max-width:420px){.essp-t .vd{display:none}.essp-t .rg{display:none}}
</style>`;

        function writePage(ctx, title, bodyHtml) {
            const form = serverWidget.createForm({ title });
            const fld = form.addField({ id: 'custpage_ess_html', type: serverWidget.FieldType.INLINEHTML, label: ' ' });
            fld.defaultValue = STYLE + bodyHtml;
            ctx.response.writePage(form);
        }

        const STYLE = `<style>
.ess{font-family:-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:#1f2937;max-width:1500px;padding:4px 2px 24px}
.ess a{color:#0b5cad;text-decoration:none}.ess a:hover{text-decoration:underline}
.ess .muted{color:#9ca3af}.ess .up{color:#047857}.ess .down{color:#b91c1c}
.ess-head{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap;border-bottom:3px solid #0b2545;padding-bottom:10px;margin-bottom:18px}
.ess-head h1{font-size:22px;letter-spacing:.06em;margin:0;color:#0b2545;font-weight:700}
.ess-head .sub{color:#6b7280;font-size:13px;margin-top:4px}
.ess-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.ess .asof{font-size:13px;color:#374151;display:flex;align-items:center;gap:6px}
.ess .asof input{font:inherit;padding:5px 8px;border:1px solid #d1d5db;border-radius:6px}
.ess-btn{display:inline-block;background:#0b2545;color:#fff!important;padding:7px 12px;border-radius:6px;font-size:13px;text-decoration:none!important}
.ess-btn.ghost{background:#fff;color:#0b2545!important;border:1px solid #0b2545}
.ess-cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-bottom:22px}
@media (max-width:900px){.ess-cards{grid-template-columns:repeat(2,minmax(0,1fr))}}
.ess-card{display:block;background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;color:inherit!important;text-decoration:none!important;box-shadow:0 1px 2px rgba(0,0,0,.04);transition:border-color .15s}
.ess-card:hover{border-color:#0b2545}
.ess-card .lbl,.ess-stat .lbl{font-size:11.5px;text-transform:uppercase;letter-spacing:.08em;color:#6b7280}
.ess-card .val{font-size:26px;font-weight:600;margin:6px 0 6px;color:#0b2545;font-variant-numeric:tabular-nums}
.ess-card .cmp{font-size:12.5px;color:#6b7280;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.ess .chip{display:inline-block;padding:2px 7px;border-radius:999px;font-size:12px;font-weight:600;background:#f3f4f6;color:#6b7280}
.ess .chip.up{background:#ecfdf5;color:#047857}.ess .chip.down{background:#fef2f2;color:#b91c1c}
.ess-stats{display:flex;gap:28px;flex-wrap:wrap;margin:4px 0 16px}
.ess-stat .val{font-size:24px;font-weight:600;color:#0b2545;margin-top:4px;font-variant-numeric:tabular-nums}
.ess-stat .val.sm{font-size:18px}
.ess-scroll{overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px}
table.ess-t{border-collapse:collapse;width:100%;font-size:12.5px;font-variant-numeric:tabular-nums;background:#fff}
.ess-t th{background:#0b2545;color:#fff;font-weight:600;padding:9px 10px;text-align:right;white-space:normal;line-height:1.25;font-size:12.5px;vertical-align:bottom}
.ess-t th.vh{background:#1d3a63}
.ess-t td{padding:8px 10px;border-bottom:1px solid #eef0f3;text-align:right;white-space:nowrap}
.ess-t .l{text-align:left}
.ess-t td.vc{background:#f8fafc;font-weight:600}
.ess-t tr.tot td{font-weight:700;font-size:14px}
.ess-t tbody tr.tot:last-child td{border-top:2px solid #0b2545;border-bottom:0}
.ess-t tr.rng td{font-size:11px;color:#6b7280;background:#f9fafb;font-weight:400;white-space:normal;line-height:1.35}
.ess-t tr.sub td{background:#f3f4f6;font-weight:600;color:#374151}
.ess-t tbody tr:not(.sub):not(.tot):not(.rng):hover td{background:#f9fbff}
.ess-txn td{white-space:nowrap;max-width:360px;overflow:hidden;text-overflow:ellipsis}
.ess .tag{display:inline-block;font-size:11px;padding:1px 7px;border-radius:4px;font-weight:600}
.ess .tag.ns{background:#e0ecff;color:#0b5cad}.ess .tag.hist{background:#f3f4f6;color:#6b7280}
.crumbs{font-size:13px;margin-bottom:10px;color:#6b7280}
.ess-pager{display:flex;gap:14px;align-items:center;font-size:13px;margin:8px 0}
.ess-note{font-size:12px;color:#6b7280;margin-top:14px;line-height:1.5}
.ess-warn{background:#fffbeb;border:1px solid #fcd34d;color:#92400e;padding:10px 12px;border-radius:8px;margin-bottom:14px;font-size:13px}
.ess-err{background:#fef2f2;border:1px solid #fca5a5;color:#991b1b;padding:14px 16px;border-radius:8px;font-size:14px;line-height:1.5}
</style>`;

        return { onRequest };
    });