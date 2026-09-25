import { describe, expect, it } from "vitest";
import {
  buildGoogleSheetLookup,
  cleanGrid,
  cleanMobile,
  cleanRoll,
  formatForWebPortal,
  maskEmail,
  maskPhone,
  normalizeRoll,
  safeBaseName,
  toGoogleSheetCsvUrl
} from "./cleaning";

describe("cleaning rules", () => {
  it("masks email local parts while preserving domains", () => {
    expect(maskEmail("jasminalam261@gmail.com")).toBe("jas*******261@gmail.com");
    expect(maskEmail("kona@gmail.com")).toBe("kona@gmail.com");
    expect(maskEmail("invalid")).toBe("invalid");
  });

  it("masks variable-length phone numbers", () => {
    expect(maskPhone("01815916577")).toBe("018*****577");
    expect(maskPhone("1768131018")).toBe("176****018");
    expect(maskPhone("N/A")).toBe("N/A");
  });

  it("normalizes rolls with hyphens, 6-digits padding, and zeros", () => {
    expect(normalizeRoll("28-28508-1")).toBe("28285081");
    expect(normalizeRoll("00-26040-1")).toBe("00260401");
    expect(normalizeRoll("260401")).toBe("00260401"); // 6 digits gets 00 in front
    expect(normalizeRoll("100000")).toBe("00100000"); // 6 digits gets 00 in front
    expect(normalizeRoll("8-8001-1")).toBe("00880011"); // 8-8001-1 -> 880011 (6 digits) -> 00880011
    expect(normalizeRoll("24-24001-1")).toBe("24240011");
  });

  it("cleans and safely pads rolls", () => {
    expect(cleanRoll("27-26048-1").value).toBe("27260481");
    expect(cleanRoll(200000).value).toBe("00200000");
    expect(cleanRoll("123456789").value).toBe("123456789");
    expect(cleanRoll("123456789").warning).toBeTruthy();
  });

  it("cleans phone numbers reliably", () => {
    expect(cleanMobile("+8801716958830")).toBe("01716958830");
    expect(cleanMobile("8801837306200")).toBe("01837306200");
    expect(cleanMobile("01712-729130")).toBe("01712729130");
    expect(cleanMobile("1716958830")).toBe("01716958830");
  });

  it("finds headers case-insensitively and removes Paid Amount", () => {
    const result = cleanGrid([
      [" Student Name ", " PHONE", "Email ", "ROLL", " paid amount "],
      ["A", "01815916577", "jasminalam261@gmail.com", "200000", 3999]
    ]);
    expect(result.rows[0]).toEqual([" Student Name ", " PHONE", "Email ", "ROLL"]);
    expect(result.rows[1]).toEqual(["A", "018*****577", "jas*******261@gmail.com", "00200000"]);
  });

  it("sanitizes download names", () => expect(safeBaseName("my:file.xlsx")).toBe("my-file"));

  it("converts Google Sheet URL to CSV export URL", () => {
    const editUrl = "https://docs.google.com/spreadsheets/d/1d3d2CF4HzIpDyz2I4QRf4zeNGLH7wL2H-mOUPS8BQww/edit?gid=0#gid=0";
    expect(toGoogleSheetCsvUrl(editUrl)).toBe(
      "https://docs.google.com/spreadsheets/d/1d3d2CF4HzIpDyz2I4QRf4zeNGLH7wL2H-mOUPS8BQww/export?format=csv&gid=0"
    );
  });

  it("builds lookup map from Google Sheet CSV with secondary phone prioritized", () => {
    const gSheetCsv = `SL,Date,MR,Roll,Name,College,,Phone,Secondary Phone,Batch,Total,1st Inst,2nd Inst,WEB ENTRY,WEB 2nd,Due,Status,Num,Verify,WEB Roll*,Email
1,21/5/26,3001,281001,Bimukdha Mallick,Govt High,,01716958830,01772780997,4 PM,10000,6000,4000,10000,,0,Full,01716958830,Verified,28-28047-1,adroo@gmail.com
2,22/5/26,3002,281002,Anwesha Das,School,,01764396004,,3 PM,10000,5000,5000,10000,,0,Full,01764396004,Verified,28-28145-1,ashis@gmail.com`;

    const lookup = buildGoogleSheetLookup(gSheetCsv, "test-sheet");
    // Web roll normalized: 28-28047-1 -> 28280471
    const student1 = lookup.get("28280471");
    expect(student1).toBeDefined();
    // Col J (Secondary Phone) is present, so chosenMobile must be secondary!
    expect(student1.chosenMobile).toBe("01772780997");
    expect(student1.phoneSource).toBe("gsheet_secondary");

    // Student 2 has blank Col J (Secondary), so chosenMobile must be Col I (Phone)!
    const student2 = lookup.get("28281451");
    expect(student2).toBeDefined();
    expect(student2.chosenMobile).toBe("01764396004");
    expect(student2.phoneSource).toBe("gsheet_primary");
  });

  it("formats correctly for web portal replacing phone from Google Sheet", () => {
    const gSheetCsv = `SL,Date,MR,Roll,Name,College,,Phone,Secondary Phone,Batch,Total,1st Inst,2nd Inst,WEB ENTRY,WEB 2nd,Due,Status,Num,Verify,WEB Roll*,Email
1,21/5/26,3001,281001,Bimukdha Mallick,Govt High,,01716958830,01772780997,4 PM,10000,6000,4000,10000,,0,Full,01716958830,Verified,28-28047-1,adroo@gmail.com
2,22/5/26,3002,281002,Anwesha Das,School,,01764396004,,3 PM,10000,5000,5000,10000,,0,Full,01764396004,Verified,28-28145-1,ashis@gmail.com`;
    const lookup = buildGoogleSheetLookup(gSheetCsv);

    const grid = [
      ["Student Name", "Phone", "Email", "Roll", "Paid Amount"],
      ["Kanchan Mallick", "01716958830", "adroo.hr@gmail.com", "28-28047-1", 12000],
      ["Ashis Kumar Das", "01999999999", "ashisdae@gmail.com", "28-28145-1", 12000],
      ["Unknown Student", "01811111111", "unknown@gmail.com", "100000", 0]
    ];

    const result = formatForWebPortal(grid, "Morning", lookup);
    expect(result.rows[0]).toEqual(["roll_number", "name", "mobile", "batch"]);

    // Kanchan Mallick: GSheet has secondary phone 01772780997
    expect(result.rows[1]).toEqual(["28280471", "Kanchan Mallick", "01772780997", "Morning"]);
    expect(result.phoneSources[1]).toBe("gsheet_secondary");

    // Ashis Kumar: GSheet has blank secondary phone, so takes primary phone 01764396004
    expect(result.rows[2]).toEqual(["28281451", "Ashis Kumar Das", "01764396004", "Morning"]);
    expect(result.phoneSources[2]).toBe("gsheet_primary");

    // Unknown Student: not in GSheet, 6 digit roll 100000 normalized to 00100000, keeps uploaded phone
    expect(result.rows[3]).toEqual(["00100000", "Unknown Student", "01811111111", "Morning"]);
    expect(result.phoneSources[3]).toBe("uploaded_fallback");
  });
});
