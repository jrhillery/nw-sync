/*
 * Created on Nov 18, 2017
 */
package com.moneydance.modules.features.nwsync;

import static java.util.Calendar.DAY_OF_MONTH;
import static java.util.Calendar.JANUARY;
import static java.util.Calendar.MONTH;
import static java.util.Calendar.YEAR;
import static org.odftoolkit.odfdom.dom.attribute.office.OfficeValueTypeAttribute.Value.CURRENCY;
import static org.odftoolkit.odfdom.dom.attribute.office.OfficeValueTypeAttribute.Value.DATE;
import static org.odftoolkit.odfdom.dom.attribute.office.OfficeValueTypeAttribute.Value.FLOAT;

import java.text.NumberFormat;
import java.util.Calendar;

import org.odftoolkit.odfdom.dom.attribute.office.OfficeValueTypeAttribute;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Row;

/**
 * Abstract read and write access to numeric spreadsheet cells.
 */
public abstract class CellHandler {

	/**
	 * Provide read and write access to currency spreadsheet cells.
	 */
	public static class CurrencyCellHandler extends CellHandler {
		public CurrencyCellHandler(Cell cell) {
			super(cell);

		} // end (Cell) constructor

		public Double getValue() {
			return this.cell.getCurrencyValue();
		} // end getValue()

		public void setValue(Number value) {
			this.cell.setCurrencyValue((Double) value, this.cell.getCurrencyCode());

		} // end setValue(Number)

		public NumberFormat getNumberFormat() {
			return NumberFormat.getCurrencyInstance();
		} // end getNumberFormat()

	} // end class CurrencyCellHandler

	/**
	 * Provide read and write access to floating point spreadsheet cells.
	 */
	public static class FloatCellHandler extends CellHandler {
		public FloatCellHandler(Cell cell) {
			super(cell);

		} // end (Cell) constructor

		public Double getValue() {
			return this.cell.getDoubleValue();
		} // end getValue()

		public void setValue(Number value) {
			this.cell.setDoubleValue((Double) value);

		} // end setValue(Number)

		public NumberFormat getNumberFormat() {
			return NumberFormat.getNumberInstance();
		} // end getNumberFormat()

	} // end class FloatCellHandler

	/**
	 * Provide read and write access to date spreadsheet cells.
	 */
	public static class DateCellHandler extends CellHandler {
		public DateCellHandler(Cell cell) {
			super(cell);

		} // end (Cell) constructor

		public Integer getValue() {
			Calendar cal = this.cell.getDateValue();
			int dateInt = cal.get(YEAR) * 10000 + (cal.get(MONTH) - JANUARY + 1) * 100
					+ cal.get(DAY_OF_MONTH);

			return dateInt;
		} // end getValue()

		public void setValue(Number value) {
			int dateInt = value.intValue();
			Calendar date = Calendar.getInstance();
			date.clear();
			date.set(YEAR, dateInt / 10000);
			date.set(MONTH, ((dateInt % 10000) / 100) - 1 + JANUARY);
			date.set(DAY_OF_MONTH, dateInt % 100);
			this.cell.setDateValue(date);

		} // end setValue(Number)

		public NumberFormat getNumberFormat() {
			return NumberFormat.getIntegerInstance();
		} // end getNumberFormat()

	} // end class DateCellHandler

	protected Cell cell;
	private Number newValue = null;

	/**
	 * Sole constructor.
	 *
	 * @param cell ods Cell instance to be handled
	 */
	public CellHandler(Cell cell) {
		this.cell = cell;

	} // end (Cell) constructor

	/**
	 * @return the numeric value of this cell
	 */
	public abstract Number getValue();

	/**
	 * @param value the value to save in this cell
	 */
	public abstract void setValue(Number value);

	/**
	 * @return a formatter for values in this cell
	 */
	public abstract NumberFormat getNumberFormat();

	/**
	 * @return the formula string of this cell
	 */
	public String getFormula() {
		return this.cell.getFormula();
	} // end getFormula()

	/**
	 * @return the text displayed in this cell
	 */
	public String getDisplayText() {
		return this.cell.getDisplayText();
	} // end getDisplayText()

	/**
	 * @param newValue new value to save for later application
	 */
	public void setNewValue(Number newValue) {
		this.newValue = newValue;

	} // end setNewValue(Number)

	/**
	 * Apply any previously set new value.
	 */
	public void applyUpdate() {
		if (this.newValue != null) {
			setValue(this.newValue);
		}

	} // end applyUpdate()

	/**
	 * @return string representation of this CellHandler
	 */
	public String toString() {

		return this.cell.getOdfElement().toString();
	} // end toString()

	/**
	 * @param cell
	 * @param valueType
	 * @return true when the supplied cell has the specified value type
	 */
	public static boolean isValueType(Cell cell, OfficeValueTypeAttribute.Value valueType) {

		return cell != null && valueType.toString().equals(cell.getValueType());
	} // end isValueType(Cell, Value)

	/**
	 * @param row ods Row instance
	 * @param index
	 * @return concrete CellHandler instance for the specified cell
	 */
	public static CellHandler getCellHandlerByIndex(Row row, int index) {
		Cell cell = row.getCellByIndex(index);

		if (isValueType(cell, CURRENCY))
			return new CurrencyCellHandler(cell);

		if (isValueType(cell, FLOAT))
			return new FloatCellHandler(cell);

		if (isValueType(cell, DATE))
			return new DateCellHandler(cell);

		return null;
	} // end getCellHandlerByIndex(Row, int)

} // end class CellHandler
