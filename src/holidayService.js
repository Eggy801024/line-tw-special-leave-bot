import { cellA1, rangeA1, singleCellA1 } from "./a1.js";
import { formatDateForReply, normalizeDateValue, parseDateFromText } from "./dateParser.js";
import { padTable, wholeSheetRange } from "./googleSheets.js";
import { RULES_TEXT } from "./rulesText.js";

const LOG_HEADERS = [
  "時間",
  "狀態",
  "工號",
  "姓名",
  "組別",
  "申請日期",
  "LINE名稱",
  "LINE userId",
  "LINE groupId",
  "原始訊息",
  "備註",
];

const BINDING_HEADERS = [
  "LINE userId",
  "工號",
  "姓名",
  "組別",
  "LINE名稱",
  "LINE groupId",
  "建立時間",
  "更新時間",
];

function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeWorkerId(value) {
  return cleanText(value).toUpperCase();
}

function toHalfWidth(input) {
  return String(input).replace(/[０-９]/g, (char) =>
    String("０１２３４５６７８９".indexOf(char)),
  );
}

function normalizeInputText(text) {
  return toHalfWidth(text)
    .replace(/[／]/g, "/")
    .replace(/[．。]/g, ".")
    .replace(/[，、；;]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function nowText(timeZone) {
  return new Intl.DateTimeFormat("zh-TW", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).format(new Date());
}

function getGroupId(source) {
  return source?.groupId || source?.roomId || "";
}

function statusText(status) {
  return {
    ACCEPTED: "申請成功",
    REJECTED_FULL: "名額已滿",
    REJECTED_DUPLICATE: "已申請過",
    REJECTED_UNKNOWN_WORKER: "查無工號",
    REJECTED_WRONG_GROUP: "群組不符",
    NEED_WORKER_ID: "需要工號",
    NEED_DATE: "需要日期",
  }[status] || status;
}

function extractWorkerId(text, config) {
  const match = cleanText(text).match(config.rules.workerIdPattern);
  return match ? normalizeWorkerId(match[0]) : null;
}

function uniqueBy(items, keyFn) {
  const seen = new Set();
  const result = [];

  for (const item of items) {
    const key = keyFn(item);
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(item);
  }

  return result;
}

function findHeaderRows(values) {
  const headers = [];

  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex] || [];
    const idCol = row.findIndex((cell) => cleanText(cell) === "工號");
    const nameCol = row.findIndex((cell) => cleanText(cell) === "姓名");
    const groupCol = row.findIndex((cell) => {
      const text = cleanText(cell);
      return text.includes("組別") || text.includes("群組");
    });

    if (idCol >= 0 && nameCol >= 0) {
      headers.push({ rowIndex, idCol, nameCol, groupCol });
    }
  }

  return headers;
}

function discoverRoster(values, config) {
  const headers = findHeaderRows(values);
  const employees = [];
  const blocks = [];

  for (let i = 0; i < headers.length; i += 1) {
    const header = headers[i];
    const nextHeaderRow = headers[i + 1]?.rowIndex ?? values.length;
    const dateStartCol = (header.groupCol >= 0 ? header.groupCol : header.nameCol) + 1;
    const dateEndCol = dateStartCol + config.rules.maxDateColumnsWithoutOriginal - 1;

    const block = {
      ...header,
      firstDataRow: header.rowIndex + 1,
      lastDataRow: nextHeaderRow - 1,
      dateStartCol,
      dateEndCol,
    };
    blocks.push(block);

    for (
      let rowIndex = block.firstDataRow;
      rowIndex <= block.lastDataRow && rowIndex < values.length;
      rowIndex += 1
    ) {
      const row = values[rowIndex] || [];
      const workerId = normalizeWorkerId(row[header.idCol]);
      if (!workerId || !config.rules.workerIdPattern.test(workerId)) continue;

      employees.push({
        workerId,
        name: cleanText(row[header.nameCol]),
        team: header.groupCol >= 0 ? cleanText(row[header.groupCol]) : "台籍",
        rowIndex,
        block,
      });
    }
  }

  return { headers, blocks, employees };
}

export class HolidayService {
  constructor({ sheetsClient, config }) {
    this.sheets = sheetsClient;
    this.config = config;
    this.mainSheetName = null;
    this.messageQueue = Promise.resolve();
  }

  async handleTextMessage(message) {
    const queued = this.messageQueue.then(() => this.processTextMessage(message));
    this.messageQueue = queued.catch(() => {});
    return queued;
  }

  async resolveMainSheetName() {
    if (this.mainSheetName) return this.mainSheetName;

    const preferred = this.config.sheets.mainSheetName;
    const spreadsheet = await this.sheets.getSpreadsheet();
    const sheetNames = spreadsheet.sheets.map((sheet) => sheet.properties.title);

    if (!sheetNames.includes(preferred)) {
      throw new Error(`找不到主表分頁：${preferred}`);
    }

    this.mainSheetName = preferred;
    return this.mainSheetName;
  }

  async ensureSupportSheets() {
    await this.sheets.ensureSheet(this.config.sheets.logSheetName, LOG_HEADERS);
    await this.sheets.ensureSheet(this.config.sheets.bindingSheetName, BINDING_HEADERS);
  }

  async loadMainSheet() {
    const sheetName = await this.resolveMainSheetName();
    const values = padTable(await this.sheets.getValues(wholeSheetRange(sheetName)));
    const roster = discoverRoster(values, this.config);

    if (roster.employees.length === 0) {
      throw new Error(`找不到人員資料，請確認 ${sheetName} 有工號 / 姓名欄位。`);
    }

    return { sheetName, values, roster };
  }

  async loadBindings() {
    await this.ensureSupportSheets();
    const rows = await this.sheets.getValues(
      `'${this.config.sheets.bindingSheetName}'!A2:H500`,
    );

    return rows.map((row, index) => ({
      rowIndex: index + 1,
      userId: cleanText(row[0]),
      workerId: normalizeWorkerId(row[1]),
      name: cleanText(row[2]),
      team: cleanText(row[3]),
      displayName: cleanText(row[4]),
      groupId: cleanText(row[5]),
    }));
  }

  async findBindingByUserId(userId) {
    if (!userId) return null;
    const bindings = await this.loadBindings();
    return bindings.find((binding) => binding.userId === userId) || null;
  }

  async upsertBinding({ userId, employee, displayName, groupId }) {
    if (!userId) return;

    const bindings = await this.loadBindings();
    const existing = bindings.find((binding) => binding.userId === userId);
    const createdAt = existing ? "" : nowText(this.config.timeZone);
    const updatedAt = nowText(this.config.timeZone);
    const row = [
      userId,
      employee.workerId,
      employee.name,
      employee.team,
      displayName || "",
      groupId || "",
      createdAt,
      updatedAt,
    ];

    if (existing) {
      const updateRange = rangeA1(
        this.config.sheets.bindingSheetName,
        existing.rowIndex,
        0,
        existing.rowIndex,
        BINDING_HEADERS.length - 1,
      );
      await this.sheets.updateValues(updateRange, [row]);
      return;
    }

    await this.sheets.appendValues(`'${this.config.sheets.bindingSheetName}'!A:H`, [row]);
  }

  async appendLog({ status, employee, dateIso, displayName, source, text, note }) {
    await this.ensureSupportSheets();
    await this.sheets.appendValues(`'${this.config.sheets.logSheetName}'!A:K`, [
      [
        nowText(this.config.timeZone),
        statusText(status),
        employee?.workerId || "",
        employee?.name || "",
        employee?.team || "",
        dateIso ? formatDateForReply(dateIso) : "",
        displayName || "",
        source?.userId || "",
        getGroupId(source),
        text || "",
        note || "",
      ],
    ]);
  }

  getDateIsoAtColumn(snapshot, block, colIndex) {
    const sameRow = normalizeDateValue(
      snapshot.values[block.rowIndex]?.[colIndex],
      this.config.timeZone,
    );
    if (sameRow) return sameRow;

    return normalizeDateValue(
      snapshot.values[block.rowIndex + 1]?.[colIndex],
      this.config.timeZone,
    );
  }

  collectDateColumns(snapshot) {
    const dates = [];

    for (const block of snapshot.roster.blocks) {
      for (let colIndex = block.dateStartCol; colIndex <= block.dateEndCol; colIndex += 1) {
        const iso = this.getDateIsoAtColumn(snapshot, block, colIndex);
        if (!iso) continue;

        const [, month, day] = iso.split("-").map(Number);
        dates.push({ colIndex, iso, month, day });
      }
    }

    return uniqueBy(dates, (date) => `${date.colIndex}:${date.iso}`).sort((a, b) =>
      a.iso.localeCompare(b.iso),
    );
  }

  resolveDateColumn(parsedDate, text, snapshot) {
    const dates = this.collectDateColumns(snapshot);
    const exact = dates.find((date) => date.iso === parsedDate.iso);
    if (exact) return exact;

    const sameMonthDay = dates.filter(
      (date) => date.month === parsedDate.month && date.day === parsedDate.day,
    );
    if (sameMonthDay.length === 1) return sameMonthDay[0];

    if (!/(\d)\s*(\/|-|\.|月)/.test(cleanText(text))) {
      const sameDay = dates.filter((date) => date.day === parsedDate.day);
      if (sameDay.length === 1) return sameDay[0];
    }

    return null;
  }

  resolveRequestedDates(text, snapshot) {
    const normalized = normalizeInputText(text);
    const requested = [];
    const occupied = [];
    const monthDayPattern = /(?:^|[^\dA-Za-z])(\d{1,2})\s*(?:\/|月)\s*(\d{1,2})((?:\s*(?:\.|\s)\s*\d{1,2})*)/g;
    let match;

    while ((match = monthDayPattern.exec(normalized))) {
      const month = Number(match[1]);
      const firstDay = Number(match[2]);
      const tail = match[3] || "";
      const days = [firstDay, ...[...tail.matchAll(/\d{1,2}/g)].map((item) => Number(item[0]))];

      occupied.push([match.index, monthDayPattern.lastIndex]);
      for (const day of days) {
        const parsed = parseDateFromText(`${month}/${day}`, this.config.timeZone);
        if (!parsed) continue;
        const matchedDate = this.resolveDateColumn(parsed, `${month}/${day}`, snapshot);
        requested.push({ input: `${month}/${day}`, parsed, matchedDate });
      }
    }

    let rest = normalized;
    for (const [start, end] of occupied.reverse()) {
      rest = `${rest.slice(0, start)} ${rest.slice(end)}`;
    }

    const workerId = extractWorkerId(rest, this.config);
    if (workerId) rest = rest.replace(workerId, " ");

    for (const item of rest.matchAll(/(?:^|[^\dA-Za-z])(\d{1,2})(?=$|[^\dA-Za-z])/g)) {
      const parsed = parseDateFromText(item[1], this.config.timeZone);
      if (!parsed) continue;
      const matchedDate = this.resolveDateColumn(parsed, item[1], snapshot);
      requested.push({ input: item[1], parsed, matchedDate });
    }

    return uniqueBy(requested, (item) => item.matchedDate?.iso || item.parsed.iso);
  }

  formatAvailableDateRange(snapshot) {
    const dates = this.collectDateColumns(snapshot);
    if (dates.length === 0) return "";

    const first = formatDateForReply(dates[0].iso);
    const last = formatDateForReply(dates[dates.length - 1].iso);
    return first === last ? first : `${first}-${last}`;
  }

  findEmployee(snapshot, workerId) {
    return (
      snapshot.roster.employees.find(
        (employee) => employee.workerId === normalizeWorkerId(workerId),
      ) || null
    );
  }

  findExistingSelection(snapshot, employee) {
    const row = snapshot.values[employee.rowIndex] || [];
    const block = employee.block;

    for (let colIndex = block.dateStartCol; colIndex <= block.dateEndCol; colIndex += 1) {
      if (cleanText(row[colIndex]) !== this.config.rules.newMark) continue;

      const iso = this.getDateIsoAtColumn(snapshot, block, colIndex);
      return {
        colIndex,
        iso,
        label: iso ? formatDateForReply(iso) : cellA1(employee.rowIndex, colIndex),
      };
    }

    return null;
  }

  countForDate(snapshot, colIndex) {
    return snapshot.roster.employees.filter((employee) => {
      const value = cleanText(snapshot.values[employee.rowIndex]?.[colIndex]);
      return value === this.config.rules.newMark;
    }).length;
  }

  async writeAcceptedSelections(snapshot, employee, colIndexes) {
    const sheetName = snapshot.sheetName;
    await this.sheets.batchUpdateValues(
      colIndexes.map((colIndex) => ({
        range: singleCellA1(sheetName, employee.rowIndex, colIndex),
        values: [[this.config.rules.newMark]],
      })),
    );

    if (!this.sheets.formatCells) return;

    await this.sheets.formatCells(
      sheetName,
      colIndexes.map((colIndex) => ({
        startRowIndex: employee.rowIndex,
        endRowIndex: employee.rowIndex + 1,
        startColumnIndex: colIndex,
        endColumnIndex: colIndex + 1,
      })),
      {
        backgroundColor: { red: 1, green: 1, blue: 0 },
        textFormat: {
          foregroundColor: { red: 1, green: 0, blue: 0 },
          bold: true,
        },
      },
    );
  }

  async processTextMessage({ text, source, displayName }) {
    const normalizedText = cleanText(text);
    const groupId = getGroupId(source);

    if (!normalizedText) return null;

    if (/^(規則|說明|help)$/i.test(normalizedText)) {
      return RULES_TEXT;
    }

    if (/^(群組資料|groupid|group id)$/i.test(normalizedText)) {
      const mappedTeam = this.config.rules.groupTeamMap[groupId] || "未設定";
      return [`groupId：${groupId || "無群組 ID"}`, `對應組別：${mappedTeam}`].join("\n");
    }

    const snapshot = await this.loadMainSheet();
    const requestedDates = this.resolveRequestedDates(normalizedText, snapshot);
    const workerIdInText = extractWorkerId(normalizedText, this.config);
    const isStatusQuery = /^(查詢|狀態|status|query)/i.test(normalizedText);

    if (requestedDates.length === 0 && isStatusQuery) {
      return "請輸入要查詢的日期，例如：查詢 6/11 或 查詢 11";
    }

    if (requestedDates.length === 0) return null;

    const invalidDate = requestedDates.find((item) => !item.matchedDate);
    if (invalidDate) {
      const rangeText = this.formatAvailableDateRange(snapshot) || "目前表格日期範圍";
      await this.appendLog({
        status: "NEED_DATE",
        dateIso: invalidDate.parsed.iso,
        displayName,
        source,
        text,
        note: "輸入日期不在表格指定範圍內",
      });
      return `日期 ${formatDateForReply(invalidDate.parsed.iso)} 不在表格指定範圍內，請依照表格日期輸入。目前可申請日期：${rangeText}`;
    }

    if (isStatusQuery) {
      return requestedDates
        .map((item) => this.buildDateStatusReply(snapshot, item.matchedDate.iso, item.matchedDate.colIndex))
        .join("\n\n");
    }

    let workerId = workerIdInText;
    if (!workerId && source?.userId) {
      const binding = await this.findBindingByUserId(source.userId);
      workerId = binding?.workerId || null;
    }

    if (!workerId) {
      const first = requestedDates[0].matchedDate;
      await this.appendLog({
        status: "NEED_WORKER_ID",
        dateIso: first.iso,
        displayName,
        source,
        text,
        note: "輸入日期但缺少工號",
      });
      return `請輸入工號和日期，例如：P0949 ${formatDateForReply(first.iso)}。完成第一次申請後，下次可只輸入日期。`;
    }

    const employee = this.findEmployee(snapshot, workerId);
    if (!employee) {
      await this.appendLog({
        status: "REJECTED_UNKNOWN_WORKER",
        dateIso: requestedDates[0].matchedDate.iso,
        displayName,
        source,
        text,
        note: `查無工號 ${workerId}`,
      });
      return `查無工號 ${workerId}，請確認工號是否正確。`;
    }

    const mappedTeam = this.config.rules.groupTeamMap[groupId];
    if (mappedTeam && mappedTeam !== employee.team) {
      await this.appendLog({
        status: "REJECTED_WRONG_GROUP",
        employee,
        dateIso: requestedDates[0].matchedDate.iso,
        displayName,
        source,
        text,
        note: `群組設定為 ${mappedTeam}，人員組別為 ${employee.team}`,
      });
      return `${employee.name} 屬於 ${employee.team}，此群組設定為 ${mappedTeam}，請到正確群組申請。`;
    }

    const existingBinding = source?.userId ? await this.findBindingByUserId(source.userId) : null;
    if (
      existingBinding &&
      existingBinding.workerId &&
      existingBinding.workerId !== employee.workerId
    ) {
      await this.appendLog({
        status: "REJECTED_UNKNOWN_WORKER",
        employee,
        dateIso: requestedDates[0].matchedDate.iso,
        displayName,
        source,
        text,
        note: `LINE帳號已綁定 ${existingBinding.workerId}`,
      });
      return `此 LINE 帳號已綁定工號 ${existingBinding.workerId}，如需變更請聯絡組長。`;
    }

    const existingSelection = this.findExistingSelection(snapshot, employee);
    if (existingSelection && !this.config.rules.allowChange) {
      await this.appendLog({
        status: "REJECTED_DUPLICATE",
        employee,
        dateIso: requestedDates[0].matchedDate.iso,
        displayName,
        source,
        text,
        note: `已申請 ${existingSelection.label}`,
      });
      return `${employee.name} 已申請 ${existingSelection.label}，如需變更請聯絡組長。`;
    }

    const accepted = [];
    const rejected = [];

    for (const item of requestedDates) {
      const { matchedDate } = item;
      const currentCellValue = cleanText(snapshot.values[employee.rowIndex]?.[matchedDate.colIndex]);

      if (currentCellValue !== this.config.rules.eligibleShiftMark) {
        rejected.push(
          `${formatDateForReply(matchedDate.iso)} 原班別不是 ${this.config.rules.eligibleShiftMark}`,
        );
        continue;
      }

      const count = this.countForDate(snapshot, matchedDate.colIndex);
      if (count >= this.config.rules.maxPerDate) {
        rejected.push(`${formatDateForReply(matchedDate.iso)} 名額已滿`);
        continue;
      }

      accepted.push(matchedDate);
      snapshot.values[employee.rowIndex][matchedDate.colIndex] = this.config.rules.newMark;
    }

    if (accepted.length === 0) {
      await this.appendLog({
        status: "REJECTED_FULL",
        employee,
        dateIso: requestedDates[0].matchedDate.iso,
        displayName,
        source,
        text,
        note: rejected.join("；"),
      });
      return [`${employee.name} 未完成申請。`, ...rejected].join("\n");
    }

    await this.writeAcceptedSelections(
      snapshot,
      employee,
      accepted.map((date) => date.colIndex),
    );
    await this.upsertBinding({
      userId: source?.userId,
      employee,
      displayName,
      groupId,
    });

    for (const date of accepted) {
      await this.appendLog({
        status: "ACCEPTED",
        employee,
        dateIso: date.iso,
        displayName,
        source,
        text,
        note: rejected.length ? `部分失敗：${rejected.join("；")}` : "",
      });
    }

    const successText = accepted.map((date) => formatDateForReply(date.iso)).join("、");
    const lines = [`${employee.name} 已成功申請特休：${successText}`];
    if (rejected.length) {
      lines.push("未完成：");
      lines.push(...rejected);
    }
    return lines.join("\n");
  }

  buildDateStatusReply(snapshot, dateIso, colIndex) {
    const names = [];

    for (const employee of snapshot.roster.employees) {
      if (this.getDateIsoAtColumn(snapshot, employee.block, colIndex) !== dateIso) continue;
      const cell = cleanText(snapshot.values[employee.rowIndex]?.[colIndex]);
      if (cell !== this.config.rules.newMark) continue;
      names.push(`${employee.name}(${employee.workerId})`);
    }

    return (
      `${formatDateForReply(dateIso)} 特休申請狀況\n` +
      `台籍: ${names.length}/${this.config.rules.maxPerDate}` +
      (names.length ? ` - ${names.join(", ")}` : "")
    );
  }
}
