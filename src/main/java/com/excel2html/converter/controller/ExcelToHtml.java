package com.excel2html.converter.controller;

import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * @author Gurcharan
 *
 */
public class ExcelToHtml {

	final private StringBuilder out = new StringBuilder(65536);
	final private SimpleDateFormat sdf;
	final private XSSFWorkbook book;
	final private FormulaEvaluator evaluator;
	private int colIndex;
	private int rowIndex, mergeStart, mergeEnd;

	/**
	 * Generates HTML from the InputStream of an Excel file. Generates sheet name in
	 * HTML h1 element.
	 * 
	 * @param in InputStream of the Excel file.
	 * @throws IOException When POI cannot read from the input stream.
	 */
	public ExcelToHtml(final InputStream in) throws IOException {
		sdf = new SimpleDateFormat("dd/MM/yyyy");
		if (in == null) {
			book = null;
			evaluator = null;
			return;
		}
		book = new XSSFWorkbook(in);
		evaluator = book.getCreationHelper().createFormulaEvaluator();
		for (int i = 0; i < book.getNumberOfSheets(); ++i) {		
			table(book.getSheetAt(i));
			out.append("<br>");
			out.append("<div>");
			addImage(book.getSheetAt(i));
			out.append("</div>");
		}
		
	}

	/**
	 * (Each Excel sheet produces an HTML table) Generates an HTML table with no
	 * cell, border spacing or padding.
	 * 
	 * @param sheet The Excel sheet.
	 */
	private void table(final XSSFSheet sheet) {
		if (sheet == null) {
			return;
		}
		out.append("<table cellspacing='0' style='border-spacing:0; border-collapse:collapse;'>\n");
		for (rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); ++rowIndex) {
			tr(sheet.getRow(rowIndex));
		}
		out.append("</table>\n");
	}

	/**
	 * (Each Excel sheet row becomes an HTML table row) Generates an HTML table row
	 * which has the same height as the Excel row.
	 * 
	 * @param xssfRow The Excel row.
	 */
	private void tr(final XSSFRow xssfRow) {
		if (xssfRow == null) {
			return;
		}
		out.append("<tr ");
		// Find merged cells in current row.
		for (int i = 0; i < xssfRow.getSheet().getNumMergedRegions(); ++i) {
			final CellRangeAddress merge = xssfRow.getSheet().getMergedRegion(i);
			if (rowIndex >= merge.getFirstRow() && rowIndex <= merge.getLastRow()) {
				mergeStart = merge.getFirstColumn();
				mergeEnd = merge.getLastColumn();
				break;
			}
		}
		out.append("style='");
		if (xssfRow.getHeight() != -1) {
			out.append("height: ").append(Math.round(xssfRow.getHeight() / 20.0 * 1.33333)).append("px; ");
		}
		out.append("'>\n");
		for (colIndex = 0; colIndex < xssfRow.getLastCellNum(); ++colIndex) {
			td(xssfRow.getCell(colIndex));
		}
		out.append("</tr>\n");
	}

	/**
	 * (Each Excel sheet cell becomes an HTML table cell) Generates an HTML table
	 * cell which has the same font styles, alignments, colours and borders as the
	 * Excel cell.
	 * 
	 * @param xssfCell The Excel cell.
	 */
	private void td(final XSSFCell xssfCell) {
		int colspan = 1;
		/*
		 * if (colIndex == mergeStart) { // First cell in the merging region - set
		 * colspan. colspan = mergeEnd - mergeStart + 1; } else if (colIndex ==
		 * mergeEnd) { // Last cell in the merging region - no more skipped cells.
		 * mergeStart = -1; mergeEnd = -1; return; } else if (mergeStart != -1 &&
		 * mergeEnd != -1 && colIndex > mergeStart && colIndex < mergeEnd) { // Within
		 * the merging region - skip the cell. return; }
		 */
		out.append("<td ");
		/*
		 * if (colspan > 1) { out.append("colspan='").append(colspan).append("' "); }
		 */
		if (xssfCell == null) {
			out.append("/>\n");
			return;
		}
		out.append("style='");
		final XSSFCellStyle style = xssfCell.getCellStyle();
		// Text alignment
		switch (style.getAlignment()) {
		case CellStyle.ALIGN_LEFT:
			out.append("text-align: left; ");
			break;
		case CellStyle.ALIGN_RIGHT:
			out.append("text-align: right; ");
			break;
		case CellStyle.ALIGN_CENTER:
			out.append("text-align: center; ");
			break;
		default:
			break;
		}
		// Font style, size and weight
		final XSSFFont font = style.getFont();
		if (font.getBoldweight() == XSSFFont.BOLDWEIGHT_BOLD) {
			out.append("font-weight: bold; ");
		}
		if (font.getItalic()) {
			out.append("font-style: italic; ");
		}
		if (font.getUnderline() != XSSFFont.U_NONE) {
			out.append("text-decoration: underline; ");
		}
		out.append("font-size: ").append(Math.floor(font.getFontHeightInPoints() * 0.8)).append("pt; ");

		// Border
		if (style.getBorderTop() != XSSFCellStyle.BORDER_NONE) {
			out.append("border-top-style: solid; ");
		}
		if (style.getBorderRight() != XSSFCellStyle.BORDER_NONE) {
			out.append("border-right-style: solid; ");
		}
		if (style.getBorderBottom() != XSSFCellStyle.BORDER_NONE) {
			out.append("border-bottom-style: solid; ");
		}
		if (style.getBorderLeft() != XSSFCellStyle.BORDER_NONE) {
			out.append("border-left-style: solid; ");
		}
		out.append("'>");
		String val = "";
		try {
			switch (xssfCell.getCellType()) {
			case XSSFCell.CELL_TYPE_STRING:
				val = xssfCell.getStringCellValue();
				break;
			case XSSFCell.CELL_TYPE_NUMERIC:
				// POI does not distinguish between integer and double, thus:
				final double original = xssfCell.getNumericCellValue(), rounded = Math.round(original);
				if (Math.abs(rounded - original) < 0.00000000000000001) {
					val = String.valueOf((int) rounded);
				} else {
					val = String.valueOf(original);
				}
				break;
			case XSSFCell.CELL_TYPE_FORMULA:
				final CellValue cv = evaluator.evaluate(xssfCell);
				switch (cv.getCellType()) {
				case Cell.CELL_TYPE_BOOLEAN:
					out.append(cv.getBooleanValue());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					out.append(cv.getNumberValue());
					break;
				case Cell.CELL_TYPE_STRING:
					out.append(cv.getStringValue());
					break;
				case Cell.CELL_TYPE_BLANK:
					break;
				case Cell.CELL_TYPE_ERROR:
					break;
				default:
					break;
				}
				break;
			default:
				// Neither string or number? Could be a date.
				try {
					val = sdf.format(xssfCell.getDateCellValue());
				} catch (final Exception e1) {
				}
			}
		} catch (final Exception e) {
			val = e.getMessage();
		}
		if ("null".equals(val)) {
			val = "";
		}

		out.append(val);
		out.append("</td>\n");
	}

	public void addImage(XSSFSheet sheet) {

		XSSFPictureData data = null;

		XSSFDrawing drawing = sheet.createDrawingPatriarch();

		// loop through all of the shapes in the drawing area
		for (XSSFShape shape : drawing.getShapes()) {
			if (shape instanceof Picture) {
				// convert the shape into a picture
				XSSFPicture picture = (XSSFPicture) shape;
				data = picture.getPictureData();
			}
		}
		out.append("<img alt='Image in Excel sheet' src='data:");
		
		out.append(data.getMimeType());
		out.append(";base64,");
		try {
			out.append(new String(Base64.encodeBase64(data.getData()), "US-ASCII"));
		} catch (final UnsupportedEncodingException e) {
			throw new RuntimeException(e);
		}
		out.append("'/>");
	}

	public String getHTML() {
		return out.toString();
	}
}
