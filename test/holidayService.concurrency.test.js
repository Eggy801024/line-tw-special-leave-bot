import assert from "node:assert/strict";
import test from "node:test";
import { HolidayService } from "../src/holidayService.js";

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function delay(ms = 2) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function colToIndex(label) {
  let value = 0;
  for (const char of label) {
    value = value * 26 + (char.charCodeAt(0) - 64);
  }
  return value - 1;
}

function parseSingleCell(range) {
  const match = range.match(/!([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`Unsupported range: ${range}`);
  return {
    rowIndex: Number(match[2]) - 1,
    colIndex: colToIndex(match[1]),
  };
}

class FakeSheetsClient {
  constructor(values) {
    this.values = values;
    this.logs = [];
  }

  async getSpreadsheet() {
    return { sheets: [{ properties: { title: "台籍特休" } }] };
  }

  async ensureSheet() {}

  async getValues(range) {
    await delay();
    if (range.includes("Line綁定")) return [];
    return clone(this.values);
  }

  async appendValues(range, rows) {
    await delay();
    this.logs.push({ range, rows });
  }

  async updateValues() {}

  async batchUpdateValues(data) {
    await delay();
    for (const update of data) {
      const { rowIndex, colIndex } = parseSingleCell(update.range);
      this.values[rowIndex][colIndex] = update.values[0][0];
    }
  }
}

function makeConfig() {
  return {
    timeZone: "Asia/Taipei",
    sheets: {
      mainSheetName: "台籍特休",
      logSheetName: "台籍特休回覆紀錄",
      bindingSheetName: "Line綁定",
    },
    rules: {
      maxPerDate: 3,
      allowChange: false,
      workerIdPattern: /[A-Z]{1,3}\d{3,4}/i,
      groupTeamMap: {},
      newMark: "特休",
      oldMark: "O",
      eligibleShiftMark: "N1",
      maxDateColumnsWithoutOriginal: 31,
    },
  };
}

function makeSheetValues() {
  const values = [
    ["", "", "", "預排特休日期", "", "", "", "", "", ""],
    ["工號", "姓名", "組別\n日期", "A", "A", "B", "B", "A", "A", "B"],
    ["", "", "", "6/7", "6/8", "6/9", "6/10", "6/11", "6/12", "6/13"],
  ];

  for (let i = 1; i <= 50; i += 1) {
    values.push([
      `TA${String(i).padStart(3, "0")}`,
      `台籍人員${i}`,
      "TW_A",
      "N1",
      "N1",
      "N1",
      "N1",
      "N1",
      "N1",
      "N1",
    ]);
  }

  return values;
}

test("serializes concurrent special leave requests so max per date is not exceeded", async () => {
  const sheets = new FakeSheetsClient(makeSheetValues());
  const service = new HolidayService({ sheetsClient: sheets, config: makeConfig() });

  const messages = Array.from({ length: 50 }, (_, index) =>
    service.handleTextMessage({
      text: `TA${String(index + 1).padStart(3, "0")} 8`,
      source: { type: "group", groupId: "G1" },
      displayName: `user-${index + 1}`,
    }),
  );

  const replies = await Promise.all(messages);
  const accepted = replies.filter((reply) => reply.includes("已成功申請特休")).length;
  const full = replies.filter((reply) => reply.includes("名額已滿")).length;
  const selectedCount = sheets.values
    .slice(3, 53)
    .filter((row) => String(row[4] || "") === "特休").length;

  assert.equal(accepted, 3);
  assert.equal(full, 47);
  assert.equal(selectedCount, 3);
});
