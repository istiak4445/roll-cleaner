import { describe, expect, it } from "vitest";
import { cleanGrid, cleanRoll, formatForWebPortal, maskEmail, maskPhone, safeBaseName } from "./cleaning";

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
  it("cleans and safely pads rolls", () => {
    expect(cleanRoll("27-26048-1").value).toBe("27260481");
    expect(cleanRoll(200000).value).toBe("00200000");
    expect(cleanRoll("123456789").value).toBe("123456789");
    expect(cleanRoll("123456789").warning).toBeTruthy();
  });
  it("finds headers case-insensitively and removes Paid Amount", () => {
    const result = cleanGrid([[" Student Name ", " PHONE", "Email ", "ROLL", " paid amount "], ["A", "01815916577", "jasminalam261@gmail.com", "200000", 3999]]);
    expect(result.rows[0]).toEqual([" Student Name ", " PHONE", "Email ", "ROLL"]);
    expect(result.rows[1]).toEqual(["A", "018*****577", "jas*******261@gmail.com", "00200000"]);
  });
  it("sanitizes download names", () => expect(safeBaseName("my:file.xlsx")).toBe("my-file"));
  it("formats correctly for web portal without masking and with robust cleanups", () => {
    const grid = [
      [" Student Name ", " PHONE", "Email ", "ROLL", " paid amount "],
      ["Mostafa Kamal", " +88-01712-345 678 ", "jasminalam261@gmail.com", "27-26048-1", 3999],
      ["Katha Chowdhury undefined", "1864365040", "katha@gmail.com", "00261143", 0],
      ["Udoy Kar Ayon", "UdoyKarAyon", "udoy@gmail.com", "00271190", 0],
      ["Samiha Tafannum Tursa", "01845801236/01817726343/01878896804", "samiha@gmail.com", "00271136", 0]
    ];
    const result = formatForWebPortal(grid, "Morning");
    expect(result.rows[0]).toEqual(["roll_number", "name", "mobile", "batch"]);
    expect(result.rows[1]).toEqual(["27260481", "Mostafa Kamal", "01712345678", "Morning"]);
    expect(result.rows[2]).toEqual(["00261143", "Katha Chowdhury", "01864365040", "Morning"]); // trailing undefined stripped, 10-digit padded
    expect(result.rows[3]).toEqual(["00271190", "Udoy Kar Ayon", "", "Morning"]); // text stripped
    expect(result.rows[4]).toEqual(["00271136", "Samiha Tafannum Tursa", "01845801236", "Morning"]); // multiple split
  });
});

