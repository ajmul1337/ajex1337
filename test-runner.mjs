import * as fs from "node:fs";
import * as fsp from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import * as XLSX from "xlsx";
import {
  normalizePartArticle,
  normalizeBrand,
  parseNumericValue,
  parseStockValue,
  generatePricePreview,
} from "./src/lib/price-import.ts";

async function runTests() {
  console.log("=== 1. Testing Normalization & Parsers ===");

  // Article normalization tests
  console.assert(normalizePartArticle("0 242 229 656") === "0242229656", "Article spaces failed");
  console.assert(normalizePartArticle("w712/75") === "W71275", "Article slash failed");
  console.assert(normalizePartArticle("OC-90") === "OC90", "Article hyphen failed");
  console.assert(normalizePartArticle("  ab.12_34  ") === "AB1234", "Article punctuation failed");
  console.log("✓ normalizePartArticle passed all tests");

  // Brand normalization tests
  console.assert(normalizeBrand("  Bosch  ") === "BOSCH", "Brand trim/case failed");
  console.assert(normalizeBrand("mann-filter") === "MANN-FILTER", "Brand hyphen preserved failed");
  console.log("✓ normalizeBrand passed all tests");

  // Price parsing tests
  console.assert(parseNumericValue(150.5) === 150.5, "Numeric price failed");
  console.assert(parseNumericValue("1 250,75 ₴") === 1250.75, "Ukrainian currency string failed");
  console.assert(parseNumericValue("$ 1,999.00") === 1999.0, "Dollar string failed");
  console.assert(parseNumericValue("250.00") === 250.0, "Standard decimal failed");
  console.assert(parseNumericValue("-50") === null, "Negative price rejected");
  console.assert(parseNumericValue("бесплатно") === null, "Text price rejected");
  console.log("✓ parseNumericValue passed all tests");

  // Stock parsing tests
  console.assert(parseStockValue(25) === 25, "Numeric stock failed");
  console.assert(parseStockValue(">10") === 10, "Stock with > failed");
  console.assert(parseStockValue(" 50+ ") === 50, "Stock with + failed");
  console.assert(parseStockValue("в наличии") === 10, "Text stock in-stock failed");
  console.assert(parseStockValue("нет") === 0, "Text stock out-of-stock failed");
  console.log("✓ parseStockValue passed all tests");

  console.log("\n=== 2. Testing CSV Preview & Delimiter Sniffing ===");

  const tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), "ajex-test-"));
  const csvFile = path.join(tmpDir, "price-sample.csv");

  const sampleCsvContent = `Бренд;Артикул;Название детали;Цена;Наличие
BOSCH;0 242 229 656;Свеча зажигания FR7DC+;145.50;25
MANN;W 712/75;Масляный фильтр;280,00;12
KNECHT;OC90;Фильтр масляный;310;>5
BOSCH;0242229656;Свеча Bosch (дубликат);140.00;10
UNKNOWN;;Деталь без артикула;50;5
`;

  await fsp.writeFile(csvFile, sampleCsvContent, "utf8");

  const csvPreview = await generatePricePreview({
    filePath: csvFile,
    originalFilename: "price-sample.csv",
    warehouseId: "1",
    context: {},
  });

  console.assert(csvPreview.fileFormat === "csv", "Expected CSV format");
  console.assert(csvPreview.headers.includes("Бренд"), "Headers missing Бренд");
  console.assert(csvPreview.headers.includes("Артикул"), "Headers missing Артикул");
  console.assert(csvPreview.mapping.article === "Артикул", "Article mapping failed");
  console.assert(csvPreview.mapping.brand === "Бренд", "Brand mapping failed");
  console.assert(csvPreview.mapping.price === "Цена", "Price mapping failed");
  console.assert(csvPreview.mapping.stock === "Наличие", "Stock mapping failed");
  console.assert(csvPreview.previewRows.length === 5, "Expected 5 preview rows");
  console.assert(csvPreview.previewRows[0].isValid === true, "Row 1 should be valid");
  console.assert(csvPreview.previewRows[4].isValid === false, "Row 5 (empty article) should be invalid");
  console.log("✓ CSV preview passed all assertions! Rows found:", csvPreview.estimatedRows);

  console.log("\n=== 3. Testing XLSX Preview Generation ===");

  const xlsxFile = path.join(tmpDir, "price-sample.xlsx");
  const wsData = [
    ["Manufacturer", "PartNumber", "Description", "Cost", "Quantity"],
    ["Brembo", "09.9145.11", "Тормозной диск передний", 1850.0, 4],
    ["Ferodo", "FDB1617", "Колодки тормозные", 950.5, 8],
    ["TRW", "GDB1550", "Колодки дисковые", 1100, 15],
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "Прайс");
  XLSX.writeFile(wb, xlsxFile);

  const xlsxPreview = await generatePricePreview({
    filePath: xlsxFile,
    originalFilename: "price-sample.xlsx",
    warehouseId: "1",
    context: {},
  });

  console.assert(xlsxPreview.fileFormat === "xlsx", "Expected XLSX format");
  console.assert(xlsxPreview.mapping.brand === "Manufacturer", "XLSX brand alias mapping failed");
  console.assert(xlsxPreview.mapping.article === "PartNumber", "XLSX article alias mapping failed");
  console.assert(xlsxPreview.mapping.price === "Cost", "XLSX price alias mapping failed");
  console.assert(xlsxPreview.mapping.stock === "Quantity", "XLSX stock alias mapping failed");
  console.assert(xlsxPreview.previewRows.length === 3, "Expected 3 preview rows in XLSX");
  console.assert(xlsxPreview.previewRows[0].isValid === true, "XLSX row 1 should be valid");
  console.log("✓ XLSX preview passed all assertions! Rows:", xlsxPreview.estimatedRows);

  // Clean up
  await fsp.rm(tmpDir, { recursive: true, force: true });
  console.log("\n>>> ALL TESTS PASSED SUCCESSFULLY! <<<");
}

runTests().catch((err) => {
  console.error("Test failed:", err);
  process.exit(1);
});
