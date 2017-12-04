/*
 * Created on Nov 10, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.infinitekind.moneydance.model.Account.AccountType.CREDIT_CARD;
import static com.sun.star.table.CellContentType.FORMULA;
import static com.sun.star.table.CellContentType.TEXT;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.ResourceBundle;

import com.infinitekind.moneydance.model.Account;
import com.infinitekind.moneydance.model.AccountBook;
import com.infinitekind.moneydance.model.AcctFilter;
import com.infinitekind.moneydance.model.CurrencyTable;
import com.infinitekind.moneydance.model.CurrencyType;
import com.sun.star.container.XEnumeration;
import com.sun.star.frame.XDesktop2;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.uno.UnoRuntime;

/**
 * Provides read/write access to an ods (OpenOffice/LibreOffice) spreadsheet
 * document.
 */
public class OdsAccessor {
	private Account root;
	private CurrencyTable securities;
	private StringWriter msgBuffer;
	private PrintWriter msgWriter;

	private ArrayList<CellHandler> changes = new ArrayList<CellHandler>();
	private CalcDoc spreadSheetDoc = null;
	private XCellRange dateRow = null;
	private int latestColumn = 0;
	private int latestDate = 0;
	private int numPricesSet = 0;
	private int numBalancesSet = 0;

	private static ResourceBundle msgBundle = null;

	private static final String baseMessageBundleName = "com.moneydance.modules.features.nwsync.NwSyncMessages";
	private static final double [] centMult = {1, 10, 100, 1000, 10000};

	/**
	 * Sole constructor.
	 *
	 * @param accountBook Moneydance account book
	 */
	public OdsAccessor(AccountBook accountBook) {
		this.root = accountBook.getRootAccount();
		this.securities = accountBook.getCurrencies();
		this.msgBuffer = new StringWriter();
		this.msgWriter = new PrintWriter(this.msgBuffer);

	} // end (AccountBook) constructor

	/**
	 * Synchronize data between a spreadsheet document and Moneydance.
	 */
	public void syncNwData() throws OdsException {
		XEnumeration rowItr = getCalcDoc().getFirstSheetRowIterator();
		findDateRow(rowItr);
		findLatestDate();

		while (rowItr.hasMoreElements()) {
			XCellRange row = CalcDoc.next(XCellRange.class, rowItr);
			XCell key = CalcDoc.getCellByIndex(row, 0);

			if (CalcDoc.isValueType(key, TEXT) || CalcDoc.isValueType(key, FORMULA)) {
				String keyVal = CellHandler.asDisplayText(key);
				CellHandler val = getCalcDoc().getCellHandlerByIndex(row, this.latestColumn);

				if (val != null) {
					CurrencyType security = this.securities.getCurrencyByTickerSymbol(keyVal);

					if (security != null) {
						// found this row's ticker symbol in Moneydance securities
						Number price = val.getValue();

						if (price instanceof Double && price.doubleValue() > 0) {
							// set the new security price
							setPrice(security, price.doubleValue());
						}
					} else {
						Account account = getAccount(keyVal);

						if (account != null) {
							// found this row's account in Moneydance
							setBalanceIfDiff(val, account, keyVal);
						} else {
							System.err.println("Ignoring row " + keyVal);
						}
					}
				}
			}
		} // end while

	} // end syncNwData()

	/**
	 * @param keyVal
	 * @return the Moneydance account corresponding to keyVal
	 */
	private Account getAccount(String keyVal) {
		final String[] actNames = keyVal.split(":");
		Account account = this.root.getAccountByName(actNames[0]);

		if (account != null && actNames.length > 1) {
			List<Account> subs = account.getSubAccounts(new AcctFilter() {

				public boolean matches(Account acct) {
					String accountName = acct.getAccountName();

					return accountName.equalsIgnoreCase(actNames[1]);
				} // end matches(Account)

				public String format(Account acct) {

					return acct.getFullAccountName();
				} // end format(Account)
			});
			account = subs == null || subs.isEmpty() ? null : subs.get(0);
		}

		return account;
	} // end getAccount(String)

	/**
	 * Set the Moneydance security price for the latest date found in the
	 * spreadsheet.
	 *
	 * @param security
	 * @param price
	 */
	private void setPrice(CurrencyType security, double price) {
		// Get the old price rounded to the tenth place past the decimal point
		BigDecimal bd = BigDecimal.valueOf(1 / security.getUserRateByDateInt(this.latestDate));
		double oldPrice = bd.setScale(10, RoundingMode.HALF_EVEN).doubleValue();

		if (price != oldPrice) {
			double rate = 1 / price;
			security.setSnapshotInt(this.latestDate, rate);
			security.setUserRate(rate); // this is usually the newest snapshot
			security.syncItem();
			++this.numPricesSet;

			// Change %s price from $%.2f to $%.2f (%+.2f%%).%n
			writeFormatted("NWSYNC10", security.getName(), oldPrice, price,
					(price / oldPrice - 1) * 100);
		}

	} // end setPrice(CurrencyType, double)

	/**
	 * Set the spreadsheet account balance if it differs from Moneydance.
	 *
	 * @param val
	 * @param account
	 * @param keyVal
	 */
	private void setBalanceIfDiff(CellHandler val, Account account, String keyVal) {
		int decimalPlaces = account.getCurrencyType().getDecimalPlaces();
		double balance = account.getUserCurrentBalance() / centMult[decimalPlaces];

		if (account.getAccountType() == CREDIT_CARD) {
			balance = -balance;
		}
		Number oldBalance = val.getValue();

		if ((oldBalance instanceof Double) && balance != oldBalance.doubleValue()) {
			val.setNewValue(balance);
			this.changes.add(val);
			++this.numBalancesSet;

			// Change %s balance from %s to %s.%n
			NumberFormat nf = val.getNumberFormat();
			nf.setMaximumFractionDigits(decimalPlaces);
			writeFormatted("NWSYNC11", keyVal, nf.format(oldBalance.doubleValue()),
					nf.format(balance));
		}

	} // end setBalanceIfDiff(CellHandler, Account, String)

	/**
	 * Capture row with 'Date' in first column.
	 *
	 * @param rowIterator
	 */
	private void findDateRow(XEnumeration rowIterator) throws OdsException {
		while (rowIterator.hasMoreElements()) {
			XCellRange row = CalcDoc.next(XCellRange.class, rowIterator);
			XCell c = CalcDoc.getCellByIndex(row, 0);

			if (CalcDoc.isValueType(c, TEXT)
					&& "Date".equalsIgnoreCase(CellHandler.asDisplayText(c))) {
				this.dateRow = row;

				return;
			}
		} // end while

		// Unable to find row with 'Date' in first column in %s.
		throw new OdsException(null, "NWSYNC01", getCalcDoc());
	} // end findDateRow(XEnumeration)

	/**
	 * Capture index of the rightmost date in the supplied row. Also capture this
	 * date's YYYYMMDD value.
	 */
	private void findLatestDate() throws OdsException {
		int cellIndex = 0;
		CellHandler c = null, latestCell;

		do {
			latestCell = c;
			c = getCalcDoc().getCellHandlerByIndex(this.dateRow, ++cellIndex);
		} while (c instanceof CellHandler.DateCellHandler);

		if (cellIndex == 1)
			// Unable to find any dates in the row with 'Date' in first column in %s.
			throw new OdsException(null, "NWSYNC02", getCalcDoc());

		this.latestColumn = cellIndex - 1;

		// capture date value in decimal form YYYYMMDD
		this.latestDate = latestCell.getValue().intValue();

		// Found rightmost date: %s.%n
		writeFormatted("NWSYNC13", latestCell.getDisplayText());

	} // end findLatestDate()

	/**
	 * Commit any changes to the spreadsheet document.
	 */
	public void commitChanges() {
		if (isModified()) {
			for (CellHandler cHandler : this.changes) {
				cHandler.applyUpdate();
			}
			this.changes.clear();

			// write a summary line
			writeSummary();
		}

	} // end commitChanges()

	/**
	 * Write a summary line.
	 */
	private void writeSummary() {
		String date;
		try {
			date = CellHandler
					.asDisplayText(CalcDoc.getCellByIndex(this.dateRow, this.latestColumn));
		} catch (Exception e) {
			date = "unknown date";
		}

		// For %s changed %d security price%s and %d account balance%s.%n
		writeFormatted("NWSYNC12", date, this.numPricesSet, this.numPricesSet == 1 ? "" : "s",
				this.numBalancesSet, this.numBalancesSet == 1 ? "" : "s");

	} // end writeSummary()

	/**
	 * @return True when the spreadsheet has uncommitted changes in memory.
	 */
	public boolean isModified() {

		return !this.changes.isEmpty();
	} // end isModified()

	/**
	 * @return a string representing accumulated messages
	 */
	public String getMessages() {

		return this.msgBuffer.toString();
	} // end getMessages()

	/**
	 * @return the currently open spreadsheet document
	 */
	private CalcDoc getCalcDoc() throws OdsException {
		if (this.spreadSheetDoc == null) {
			XSpreadsheetDocument doc = null;
			int numSpreadsheetDocs = 0;
			XDesktop2 libreOfficeDesktop = CalcDoc.getOfficeDesktop();
			XEnumeration compItr = libreOfficeDesktop.getComponents().createEnumeration();

			if (!compItr.hasMoreElements()) {
				// no components so we probably started the desktop => terminate it
				libreOfficeDesktop.terminate();
			} else {
				do {
					XServiceInfo comp = CalcDoc.next(XServiceInfo.class, compItr);

					if (comp.supportsService("com.sun.star.sheet.SpreadsheetDocument")) {
						doc = UnoRuntime.queryInterface(XSpreadsheetDocument.class, comp);
						++numSpreadsheetDocs;
					}
				} while (compItr.hasMoreElements());
			}

			if (doc == null)
				// No open spreadsheet documents found.
				throw new OdsException(null, "NWSYNC05");

			if (numSpreadsheetDocs > 1)
				// Found %d open spreadsheet documents. Can only work with one.
				throw new OdsException(null, "NWSYNC04", numSpreadsheetDocs);

			this.spreadSheetDoc = new CalcDoc(doc);
		}

		return this.spreadSheetDoc;
	} // end getCalcDoc()

	/**
	 * Release any resources we acquired. This includes closing the connection to
	 * the office process.
	 *
	 * @return null
	 */
	public OdsAccessor releaseResources() {
		CalcDoc.closeOfficeConnection();

		return null;
	} // end releaseResources()

	private static ResourceBundle getMsgBundle() {
		if (msgBundle == null) {
			try {
				msgBundle = ResourceBundle.getBundle(baseMessageBundleName);
			} catch (Exception e) {
				System.err.format("Unable to load message bundle %s. %s%n", baseMessageBundleName, e);
				msgBundle = new ResourceBundle() {
					protected Object handleGetObject(String key) {
						// just use the key since we have no message bundle
						return key;
					}

					public Enumeration<String> getKeys() {
						return null;
					}
				};
			} // end catch
		}

		return msgBundle;
	} // end getMsgBundle()

	/**
	 * Inner class to house exceptions.
	 */
	public static class OdsException extends Exception {

		private static final long serialVersionUID = -9019657124886615063L;

		/**
		 * @param cause Exception that caused this (null if none)
		 * @param key The resource bundle key (or message)
		 * @param params Optional parameters for the detail message
		 */
		public OdsException(Throwable cause, String key, Object... params) {
			super(String.format(retrieveMessage(key), params), cause);

		} // end (Throwable, String, Object...) constructor

	} // end class OdsException

	/**
	 * @param key The resource bundle key (or message)
	 * @return message for this key
	 */
	private static String retrieveMessage(String key) {
		try {
			return getMsgBundle().getString(key);
		} catch (Exception e) {
			// just use the key when not found
			return key;
		}
	} // end retrieveMessage(String)

	/**
	 * @param key The resource bundle key (or message)
	 * @param params Optional array of parameters for the message
	 */
	private void writeFormatted(String key, Object... params) {
		this.msgWriter.format(retrieveMessage(key), params);

	} // end writeFormatted(String, Object...)

} // end class OdsAccessor
