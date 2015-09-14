package utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Using Apache POI 3.12
 * 
 * @author Cyber365
 *
 */
public class ExcelUtils {

	/**
	 * 
	 * Insert new row into worksheet with style as style of above row
	 * 
	 * @author Cyber365
	 * 
	 * @param workbook
	 * @param worksheet
	 * @param sourceRowNum
	 * @param rowNum
	 */

	public static void insertNewRowWithStyleAsAbove(Sheet worksheet, int rowNum) {

		worksheet.shiftRows(rowNum, worksheet.getLastRowNum(), 1);
		Row aboveRow = worksheet.getRow(rowNum - 1);

		if (aboveRow == null) {
			return;

		}

		Row newRow = worksheet.createRow(rowNum);
		newRow.setHeightInPoints(aboveRow.getHeightInPoints());

		for (int i = 0; i < aboveRow.getLastCellNum(); i++) {
			Cell cellAbove = aboveRow.getCell(i);
			if (cellAbove == null) {
				continue;

			}

			Cell newCell = newRow.createCell(i);
			newCell.setCellStyle(cellAbove.getCellStyle());
			newCell.setCellType(cellAbove.getCellType());
			newCell.setCellComment(cellAbove.getCellComment());

		}

		for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
			CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
			if (cellRangeAddress.getFirstRow() == aboveRow.getRowNum()) {
				CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
						(newRow.getRowNum() + (cellRangeAddress.getFirstRow() - cellRangeAddress.getLastRow())), cellRangeAddress.getFirstColumn(),
						cellRangeAddress.getLastColumn());
				worksheet.addMergedRegion(newCellRangeAddress);
			}
		}

	}

}
