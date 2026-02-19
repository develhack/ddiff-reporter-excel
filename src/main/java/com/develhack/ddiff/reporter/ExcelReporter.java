package com.develhack.ddiff.reporter;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.text.Normalizer;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.develhack.ddiff.Diff;
import com.develhack.ddiff.LogicalPath.LogicalPathElement;
import com.develhack.ddiff.Reporter;
import com.github.difflib.text.DiffRow;
import com.github.difflib.text.DiffRow.Tag;
import com.github.difflib.text.DiffRowGenerator;

public class ExcelReporter implements Reporter {

    private static final String SUMMARY_SHEET_NAME = "Summary";
    private static final String CHANGES_TEMPLATE_SHEET_NAME = "_Changes_";

    private static final String INLINE_DIFF_DELIMITER = String.valueOf((char) 0);

    private static final Map<Tag, String> TAG_MARK_MAP = new HashMap<>();

    private static final int MAX_SHEET_NAME_LENGTH = 31;

    static {
        TAG_MARK_MAP.put(Tag.CHANGE, "*");
        TAG_MARK_MAP.put(Tag.INSERT, "+");
        TAG_MARK_MAP.put(Tag.DELETE, "-");
    }

    private final DiffRowGenerator generator = DiffRowGenerator.create()
            .showInlineDiffs(true)
            .lineNormalizer(Function.identity())
            .oldTag(f -> INLINE_DIFF_DELIMITER)
            .newTag(f -> INLINE_DIFF_DELIMITER)
            .build();

    @Override
    public String getFormat() {
        return "excel";
    }

    @Override
    public void report(Path originalRoot, Path revisedRoot, List<Diff> diffs, OutputStream os) throws IOException {

        try (XSSFWorkbook workbook = new XSSFWorkbook(getClass().getResourceAsStream("template.xlsx"))) {

            XSSFSheet summarySheet = workbook.getSheet(SUMMARY_SHEET_NAME);
            CellRangeAddress rowCellRangeAddress = getNamedCellRangeAddress(workbook, summarySheet, "Row");

            // prepare summary rows
            int firstRowIndex = rowCellRangeAddress.getFirstRow();
            CellCopyPolicy cellCopyPolicy = new CellCopyPolicy();
            for (int i = 1; i < diffs.size(); ++i) {
                summarySheet.copyRows(firstRowIndex, firstRowIndex, firstRowIndex + i, cellCopyPolicy);
            }

            // write summary rows
            int firstColIndex = rowCellRangeAddress.getLastColumn();
            if (firstColIndex < 0) { // range is row
                firstColIndex = 0;
            }
            int offset = 0;
            for (Diff diff : diffs) {
                XSSFRow row = getRow(summarySheet, firstRowIndex + offset);
                getCell(summarySheet, row, firstColIndex).setCellValue(diff.getPath().toString());
                XSSFCell statusCell = getCell(summarySheet, row, firstColIndex + 1);
                statusCell.setCellValue(diff.getStatus().toString());
                if (diff.getStatus() == Diff.Status.CHANGED) {
                    XSSFSheet changesSheet = createChangeSheet(workbook, diff);
                    XSSFHyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
                    hyperlink.setAddress("'" + changesSheet.getSheetName() + "'!A1");
                    statusCell.setHyperlink(hyperlink);
                }
                ++offset;
            }
            summarySheet.autoSizeColumn(firstColIndex);

            // set filter
            summarySheet.setAutoFilter(getNamedCellRangeAddress(workbook, summarySheet, "Header"));

            // write title
            CellRangeAddress titleCellRangeAddress = getNamedCellRangeAddress(workbook, summarySheet, "Title");
            XSSFCell titleCell = getCell(summarySheet, titleCellRangeAddress.getFirstRow(),
                    titleCellRangeAddress.getFirstColumn());
            titleCell.setCellValue(String.format("Compare %s with %s", originalRoot, revisedRoot));

            // remove template sheet
            workbook.removeSheetAt(workbook.getSheetIndex(CHANGES_TEMPLATE_SHEET_NAME));

            workbook.write(os);
        }
    }

    String generateSheetName(XSSFWorkbook workbook, Diff diff) {
        List<LogicalPathElement> logicalPathElements = diff.getPath().logicalPathElements;
        String fileName = logicalPathElements.get(logicalPathElements.size() - 1).toString();
        String normalizeName = Normalizer.normalize(fileName, Normalizer.Form.NFKC).replaceAll("[\\[\\]\\\\/*?:]", "_");
        if (normalizeName.length() > MAX_SHEET_NAME_LENGTH) {
            normalizeName = normalizeName.substring(0, MAX_SHEET_NAME_LENGTH);
        }
        if (workbook.getSheetIndex(normalizeName) < 0) {
            return normalizeName;
        }

        for (int i = 0;; ++i) {
            String suffixedName = normalizeName + String.format("_%d", i);
            if (suffixedName.length() > MAX_SHEET_NAME_LENGTH) {
                suffixedName = suffixedName.substring(
                        suffixedName.length() - MAX_SHEET_NAME_LENGTH,
                        suffixedName.length());
            }
            if (workbook.getSheetIndex(suffixedName) < 0) {
                return suffixedName;
            }
        }
    }

    XSSFSheet createChangeSheet(XSSFWorkbook workbook, Diff diff) {

        XSSFSheet changesTemplateSheet = workbook.getSheet(CHANGES_TEMPLATE_SHEET_NAME);
        XSSFSheet changesSheet = workbook.cloneSheet(workbook.getSheetIndex(CHANGES_TEMPLATE_SHEET_NAME),
                generateSheetName(workbook, diff));

        List<DiffRow> diffRows = generator.generateDiffRows(diff.getOriginalLines(), diff.getPatch());

        CellRangeAddress rowCellRangeAddress = getNamedCellRangeAddress(workbook, changesTemplateSheet, "Row");

        // prepare changes rows
        int firstRowIndex = rowCellRangeAddress.getFirstRow();
        CellCopyPolicy cellCopyPolicy = new CellCopyPolicy();
        if (diffRows.size() > 1) {
            changesSheet.copyRows(firstRowIndex + 1, firstRowIndex + 1, firstRowIndex + diffRows.size(), cellCopyPolicy);
        }
        for (int i = 1; i < diffRows.size(); ++i) {
            changesSheet.copyRows(firstRowIndex, firstRowIndex, firstRowIndex + i, cellCopyPolicy);
        }

        // write changes rows
        int firstColIndex = rowCellRangeAddress.getLastColumn();
        if (firstColIndex < 0) { // range is row
            firstColIndex = 0;
        }
        XSSFFont defaultFont = getCell(changesSheet, firstRowIndex, firstColIndex).getCellStyle().getFont();
        XSSFFont highlightFonr = getCell(changesSheet, firstRowIndex - 1, firstColIndex).getCellStyle().getFont();

        int offset = 0;
        for (DiffRow diffRow : diffRows) {
            XSSFRow row = getRow(changesSheet, firstRowIndex + offset);
            getCell(changesSheet, row, firstColIndex).setCellValue(TAG_MARK_MAP.get(diffRow.getTag()));
            if (diffRow.getTag() == Tag.EQUAL) {
                getCell(changesSheet, row, firstColIndex + 1).setCellValue(diffRow.getOldLine());
                getCell(changesSheet, row, firstColIndex + 2).setCellValue(diffRow.getNewLine());
            } else {
                getCell(changesSheet, row, firstColIndex + 1)
                        .setCellValue(highlightInlineDiff(diffRow.getOldLine(), defaultFont, highlightFonr));
                getCell(changesSheet, row, firstColIndex + 2)
                        .setCellValue(highlightInlineDiff(diffRow.getNewLine(), defaultFont, highlightFonr));
            }
            ++offset;
        }
        changesSheet.autoSizeColumn(firstColIndex + 1);
        changesSheet.autoSizeColumn(firstColIndex + 2);

        // write title
        CellRangeAddress titleCellRangeAddress = getNamedCellRangeAddress(workbook, changesTemplateSheet, "Title");
        XSSFCell titleCell = getCell(changesSheet, titleCellRangeAddress.getFirstRow(),
                titleCellRangeAddress.getFirstColumn());
        titleCell.setCellValue(diff.getPath().toString());

        return changesSheet;
    }

    XSSFName getNameInSheet(XSSFWorkbook workbook, XSSFSheet sheet, String name) {
        int sheetIndex = workbook.getSheetIndex(sheet);
        return workbook.getNames(name).stream()
                .filter(n -> n.getSheetIndex() == sheetIndex)
                .findFirst()
                .get();
    }

    CellRangeAddress getNamedCellRangeAddress(XSSFWorkbook workbook, XSSFSheet sheet, String name) {
        return CellRangeAddress.valueOf(getNameInSheet(workbook, sheet, name).getRefersToFormula());
    }

    XSSFRow getRow(XSSFSheet sheet, int rowIndex) {
        XSSFRow row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    XSSFRow getRow(XSSFSheet sheet, CellReference cellReference) {
        return getRow(sheet, cellReference.getRow());
    }

    XSSFCell getCell(XSSFSheet sheet, int rowIndex, int colIndex) {
        XSSFRow row = getRow(sheet, rowIndex);
        XSSFCell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        return cell;
    }

    XSSFCell getCell(XSSFSheet sheet, XSSFRow row, int colIndex) {
        XSSFCell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        return cell;
    }

    XSSFCell getCell(XSSFSheet sheet, CellReference cellReference) {
        return getCell(sheet, cellReference.getRow(), cellReference.getCol());
    }

    XSSFRichTextString highlightInlineDiff(String plain, XSSFFont defaultFont, XSSFFont highlightFonr) {
        XSSFRichTextString rich = new XSSFRichTextString();
        boolean highlighting = false;
        for (String token : plain.split(INLINE_DIFF_DELIMITER)) {
            if (!token.isEmpty()) {
                if (highlighting) {
                    rich.append(token, highlightFonr);
                } else {
                    rich.append(token, defaultFont);
                }
            }
            highlighting = !highlighting;
        }
        return rich;
    }
}
