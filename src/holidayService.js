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

function multiLang({ zh, en, vi }) {
  return [`中文：${zh}`, `English: ${en}`, `Tiếng Việt: ${vi}`].join("\n");
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

function dateInputHasMonth(text) {
  return /(\d)\s*(\/|-|\.|月)/.test(cleanText(text));
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
      originalCol: -1,
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

    if (sheetNames.includes(preferred)) {
      this.mainSheetName = preferred;
      return this.mainSheetName;
    }

    throw new Error(`找不到主表分頁：${preferred}`);
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
    const createdAt = existing ? existing.createdAt || "" : nowText(this.config.timeZone);
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

    if (!dateInputHasMonth(text)) {
      const sameDay = dates.filter((date) => date.day === parsedDate.day);
      if (sameDay.length === 1) return sameDay[0];
    }

    return null;
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

  countForDate(snapshot, team, colIndex) {
    return snapshot.roster.employees.filter((employee) => {
      if (employee.team !== team) return false;
      const value = cleanText(snapshot.values[employee.rowIndex]?.[colIndex]);
      return value === this.config.rules.newMark;
    }).length;
  }

  async writeAcceptedSelection(snapshot, employee, colIndex) {
    const sheetName = snapshot.sheetName;
    await this.sheets.batchUpdateValues([
      {
        range: singleCellA1(sheetName, employee.rowIndex, colIndex),
        values: [[this.config.rules.newMark]],
      },
    ]);

    if (!this.sheets.formatCells) return;

    await this.sheets.formatCells(
      sheetName,
      [
        {
          startRowIndex: employee.rowIndex,
          endRowIndex: employee.rowIndex + 1,
          startColumnIndex: colIndex,
          endColumnIndex: colIndex + 1,
        },
      ],
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

    if (/^(規則|說明|help|rule|rules)$/i.test(normalizedText)) {
      return RULES_TEXT;
    }

    if (/^(群組資料|groupid|group id)$/i.test(normalizedText)) {
      const mappedTeam = this.config.rules.groupTeamMap[groupId] || "未設定";
      return multiLang({
        zh: `groupId：${groupId || "無群組 ID"}\n對應組別：${mappedTeam}`,
        en: `groupId: ${groupId || "No group ID"}\nMapped team: ${mappedTeam}`,
        vi: `groupId: ${groupId || "Không có mã nhóm"}\nNhóm tương ứng: ${mappedTeam}`,
      });
    }

    const parsedDate = parseDateFromText(normalizedText, this.config.timeZone);
    const workerIdInText = extractWorkerId(normalizedText, this.config);
    const isStatusQuery = /^(查詢|狀態|status|query)/i.test(normalizedText);

    if (!parsedDate && isStatusQuery) {
      return multiLang({
        zh: "請輸入要查詢的日期，例如：查詢 6/3",
        en: "Please enter the date to check, for example: status 6/3",
        vi: "Vui lòng nhập ngày cần kiểm tra, ví dụ: kiem tra 6/3",
      });
    }

    if (!parsedDate) return null;

    const snapshot = await this.loadMainSheet();
    const matchedDate = this.resolveDateColumn(parsedDate, normalizedText, snapshot);

    if (!matchedDate) {
      const rangeText = this.formatAvailableDateRange(snapshot) || "目前表格日期範圍";
      await this.appendLog({
        status: "NEED_DATE",
        dateIso: parsedDate.iso,
        displayName,
        source,
        text,
        note: "輸入日期不在表格指定範圍內",
      });
      return multiLang({
        zh: `此日期不在表格指定範圍內，請依照表格日期輸入。目前可申請日期：${rangeText}`,
        en: `This date is not in the sheet date range. Please enter a date shown in the sheet. Available dates: ${rangeText}`,
        vi: `Ngày này không nằm trong phạm vi ngày của bảng. Vui lòng nhập ngày có trong bảng. Ngày có thể đăng ký: ${rangeText}`,
      });
    }

    let workerId = workerIdInText;
    if (!workerId && source?.userId) {
      const binding = await this.findBindingByUserId(source.userId);
      workerId = binding?.workerId || null;
    }

    if (isStatusQuery) {
      return this.buildDateStatusReply(snapshot, matchedDate.iso, matchedDate.colIndex);
    }

    if (!workerId) {
      await this.appendLog({
        status: "NEED_WORKER_ID",
        dateIso: matchedDate.iso,
        displayName,
        source,
        text,
        note: "輸入日期但缺少工號",
      });
      return multiLang({
        zh: `請輸入工號和日期，例如：P0257 ${formatDateForReply(matchedDate.iso)}。完成第一次申請後，下次可只輸入日期。`,
        en: `Please enter your worker ID and date, for example: P0257 ${formatDateForReply(matchedDate.iso)}. After the first request, you can enter only the date next time.`,
        vi: `Vui lòng nhập mã nhân viên và ngày, ví dụ: P0257 ${formatDateForReply(matchedDate.iso)}. Sau lần đăng ký đầu tiên, lần sau chỉ cần nhập ngày.`,
      });
    }

    const employee = this.findEmployee(snapshot, workerId);
    if (!employee) {
      await this.appendLog({
        status: "REJECTED_UNKNOWN_WORKER",
        dateIso: matchedDate.iso,
        displayName,
        source,
        text,
        note: `查無工號 ${workerId}`,
      });
      return multiLang({
        zh: `查無工號 ${workerId}，請確認工號是否正確。`,
        en: `Worker ID ${workerId} was not found. Please check the worker ID.`,
        vi: `Không tìm thấy mã nhân viên ${workerId}. Vui lòng kiểm tra lại mã nhân viên.`,
      });
    }

    const mappedTeam = this.config.rules.groupTeamMap[groupId];
    if (mappedTeam && mappedTeam !== employee.team) {
      await this.appendLog({
        status: "REJECTED_WRONG_GROUP",
        employee,
        dateIso: matchedDate.iso,
        displayName,
        source,
        text,
        note: `群組設定為 ${mappedTeam}，人員組別為 ${employee.team}`,
      });
      return multiLang({
        zh: `${employee.name} 屬於 ${employee.team}，此群組設定為 ${mappedTeam}，請到正確群組申請。`,
        en: `${employee.name} belongs to ${employee.team}. This group is set to ${mappedTeam}. Please request in the correct group.`,
        vi: `${employee.name} thuộc nhóm ${employee.team}. Nhóm này được đặt là ${mappedTeam}. Vui lòng đăng ký đúng nhóm.`,
      });
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
        dateIso: matchedDate.iso,
        displayName,
        source,
        text,
        note: `LINE帳號已綁定 ${existingBinding.workerId}`,
      });
      return multiLang({
        zh: `此 LINE 帳號已綁定工號 ${existingBinding.workerId}，如需變更請聯絡組長。`,
        en: `This LINE account is already linked to worker ID ${existingBinding.workerId}. Please contact the leader if it needs to be changed.`,
        vi: `Tài khoản LINE này đã liên kết với mã nhân viên ${existingBinding.workerId}. Nếu cần thay đổi, vui lòng liên hệ tổ trưởng.`,
      });
    }

    const existingSelection = this.findExistingSelection(snapshot, employee);
    if (existingSelection && !this.config.rules.allowChange) {
      await this.appendLog({
        status: "REJECTED_DUPLICATE",
        employee,
        dateIso: matchedDate.iso,
        displayName,
        source,
        text,
        note: `已申請 ${existingSelection.label}`,
      });
      return multiLang({
        zh: `${employee.name} 已申請 ${existingSelection.label}，如需變更請聯絡組長。`,
        en: `${employee.name} has already requested ${existingSelection.label}. Please contact the leader if it needs to be changed.`,
        vi: `${employee.name} đã đăng ký ngày ${existingSelection.label}. Nếu cần thay đổi, vui lòng liên hệ tổ trưởng.`,
      });
    }

    const currentCellValue = cleanText(snapshot.values[employee.rowIndex]?.[matchedDate.colIndex]);
    if (currentCellValue !== this.config.rules.eligibleShiftMark) {
      await this.appendLog({
        status: "NEED_DATE",
        employee,
        dateIso: matchedDate.iso,
        displayName,
        source,
        text,
        note: `該日期原班別不是 ${this.config.rules.eligibleShiftMark}`,
      });
      return multiLang({
        zh: `${employee.name} 在 ${formatDateForReply(matchedDate.iso)} 原班別不是 ${this.config.rules.eligibleShiftMark}，不可申請該日特休。`,
        en: `${employee.name}'s original shift on ${formatDateForReply(matchedDate.iso)} is not ${this.config.rules.eligibleShiftMark}, so this date cannot be requested.`,
        vi: `Ca ban đầu của ${employee.name} vào ngày ${formatDateForReply(matchedDate.iso)} không phải ${this.config.rules.eligibleShiftMark}, nên không thể đăng ký ngày này.`,
      });
    }

    const count = this.countForDate(snapshot, employee.team, matchedDate.colIndex);
    if (count >= this.config.rules.maxPerDate) {
      await this.appendLog({
        status: "REJECTED_FULL",
        employee,
        dateIso: matchedDate.iso,
        displayName,
        source,
        text,
        note: `${formatDateForReply(matchedDate.iso)} 已有 ${count} 人`,
      });
      return multiLang({
        zh: `${formatDateForReply(matchedDate.iso)} 已有 ${count} 人申請，名額已滿，請改選其他日期。`,
        en: `${formatDateForReply(matchedDate.iso)} already has ${count} people requesting special leave. The limit is full. Please choose another date.`,
        vi: `Ngày ${formatDateForReply(matchedDate.iso)} đã có ${count} người đăng ký, đã hết chỗ. Vui lòng chọn ngày khác.`,
      });
    }

    await this.writeAcceptedSelection(snapshot, employee, matchedDate.colIndex);
    await this.upsertBinding({
      userId: source?.userId,
      employee,
      displayName,
      groupId,
    });
    await this.appendLog({
      status: "ACCEPTED",
      employee,
      dateIso: matchedDate.iso,
      displayName,
      source,
      text,
      note: "",
    });

    const remaining = Math.max(0, this.config.rules.maxPerDate - (count + 1));
    return multiLang({
      zh: `${employee.name} 已成功申請特休 ${formatDateForReply(matchedDate.iso)}。\n剩餘名額：${remaining}`,
      en: `${employee.name} has successfully requested special leave for ${formatDateForReply(matchedDate.iso)}.\nRemaining slots: ${remaining}`,
      vi: `${employee.name} đã đăng ký thành công ngày ${formatDateForReply(matchedDate.iso)}.\nSố chỗ còn lại: ${remaining}`,
    });
  }

  buildDateStatusReply(snapshot, dateIso, colIndex) {
    const lines = [
      multiLang({
        zh: `${formatDateForReply(dateIso)} 特休申請狀況`,
        en: `${formatDateForReply(dateIso)} special leave request status`,
        vi: `Tình trạng đăng ký ngày ${formatDateForReply(dateIso)}`,
      }),
    ];
    const names = [];

    for (const employee of snapshot.roster.employees) {
      if (this.getDateIsoAtColumn(snapshot, employee.block, colIndex) !== dateIso) continue;

      const cell = cleanText(snapshot.values[employee.rowIndex]?.[colIndex]);
      if (cell !== this.config.rules.newMark) continue;
      names.push(`${employee.name}(${employee.workerId})`);
    }

    lines.push(`台籍: ${names.length}/${this.config.rules.maxPerDate}` + (names.length ? ` - ${names.join(", ")}` : ""));
    return lines.join("\n");
  }
}
