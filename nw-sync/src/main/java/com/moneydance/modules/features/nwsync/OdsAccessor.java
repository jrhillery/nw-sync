/*
 * Created on Nov 10, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.infinitekind.moneydance.model.Account.AccountType.CREDIT_CARD;
import static com.leastlogic.swing.util.HTMLPane.CL_DECREASE;
import static com.leastlogic.swing.util.HTMLPane.CL_INCREASE;
import static com.sun.star.table.CellContentType.FORMULA;
import static com.sun.star.table.CellContentType.TEXT;
import static java.time.format.FormatStyle.MEDIUM;

import java.text.NumberFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Map.Entry;
import java.util.ResourceBundle;
import java.util.TreeMap;

import com.infinitekind.moneydance.model.Account;
import com.infinitekind.moneydance.model.AccountBook;
import com.infinitekind.moneydance.model.CurrencySnapshot;
import com.infinitekind.moneydance.model.CurrencyTable;
import com.infinitekind.moneydance.model.CurrencyType;
import com.leastlogic.moneydance.util.MdUtil;
import com.leastlogic.moneydance.util.MduException;
import com.moneydance.modules.features.nwsync.CellHandler.DateCellHandler;
import com.sun.star.container.XEnumeration;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;

/**
 * Provides read/write access to an ods (OpenOffice/LibreOffice) spreadsheet
 * document.
 */
public class OdsAccessor implements MessageBundleProvider {
	private MessageWindow messageWindow;
	private Locale locale;
	private Account root;
	private CurrencyTable securities;

	private CalcDoc calcDoc = null;
	private XCellRange dateRow = null;
	private int latestColumn = 0;
	private DateCellHandler latestDateCell = null;
	private int numPricesSet = 0;
	private int numBalancesSet = 0;
	private Map<LocalDate, List<String>> securitySnapshots = new TreeMap<>();
	private ResourceBundle msgBundle = null;

	private static final DateTimeFormatter dateFmt = DateTimeFormatter.ofLocalizedDate(MEDIUM);

	/**
	 * Sole constructor.
	 *
	 * @param messageWindow
	 * @param accountBook Moneydance account book
	 */
	public OdsAccessor(MessageWindow messageWindow, AccountBook accountBook) {
		this.messageWindow = messageWindow;
		this.locale = messageWindow.getLocale();
		this.root = accountBook.getRootAccount();
		this.securities = accountBook.getCurrencies();

	} // end (MessageWindow, AccountBook) constructor

	/**
	 * Synchronize data between a spreadsheet document and Moneydance.
	 */
	public void syncNwData() throws MduException {
		CalcDoc calcDoc = getCalcDoc();

		if (calcDoc != null && calcDoc.getSheets() == null) {
			// can't access the sheets, force a reconnect
			this.calcDoc = null;
			calcDoc = getCalcDoc();
		}
		if (calcDoc == null)
			return; // nothing to synchronize

		XEnumeration rowItr = calcDoc.getFirstSheetRowIterator();

		if (!findDateRow(rowItr, calcDoc) || !findLatestDate())
			return; // can't synchronize without a date row and latest date

		while (rowItr.hasMoreElements()) {
			XCellRange row = CalcDoc.next(XCellRange.class, rowItr); // get next row
			XCell key = calcDoc.getCellByIndex(row, 0); // get its first column

			if (CalcDoc.isContentType(key, TEXT) || CalcDoc.isContentType(key, FORMULA)) {
				String keyVal = CellHandler.asDisplayText(key);
				CellHandler val = calcDoc.getCellHandlerByIndex(row, this.latestColumn);

				if (val != null) {
					CurrencyType security = this.securities.getCurrencyByTickerSymbol(keyVal);

					if (security != null) {
						// found this row's ticker symbol in Moneydance securities
						setPriceIfDiff(val, security);
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
		analyzeSecurityDates();

		if (!isModified()) {
			// No new price or balance data found.
			writeFormatted("NWSYNC03");
		}

	} // end syncNwData()

	/**
	 * @param keyVal
	 * @return The Moneydance account corresponding to keyVal
	 */
	private Account getAccount(String keyVal) {
		final String[] actNames = keyVal.split(":");
		Account account = this.root.getAccountByName(actNames[0]);

		if (account != null && actNames.length > 1) {
			account = MdUtil.getSubAccountByName(account, actNames[1]);
		}

		return account;
	} // end getAccount(String)

	/**
	 * @param security
	 * @return The price in the last snapshot of the supplied security
	 */
	private double getLatestPrice(CurrencyType security) {
		CurrencySnapshot latestSnapshot = MdUtil.getLatestSnapshot(security);

		// add this snapshot to our collection
		getSecurityListForDate(latestSnapshot.getDateInt())
			.add(security.getName() + " (" + security.getTickerSymbol() + ')');

		return MdUtil.validateCurrentUserRate(security, latestSnapshot);
	} // end getLatestPrice(CurrencyType, NumberFormat)

	/**
	 * @param dateInt
	 * @return The list of security names for the specified date integer
	 */
	private List<String> getSecurityListForDate(int dateInt) {
		LocalDate localDate = MdUtil.convDateIntToLocal(dateInt);
		List<String> daysSecurities = this.securitySnapshots.get(localDate);

		if (daysSecurities == null) {
			daysSecurities = new ArrayList<>();
			this.securitySnapshots.put(localDate, daysSecurities);
		}

		return daysSecurities;
	} // end getSecurityListForDate(int)

	/**
	 * Analyze security dates to see if they are all the same.
	 */
	private void analyzeSecurityDates() throws MduException {
		Iterator<Entry<LocalDate, List<String>>> snapshotsIterator =
			this.securitySnapshots.entrySet().iterator();

		if (snapshotsIterator.hasNext()) {
			Entry<LocalDate, List<String>> snapShotsEntry = snapshotsIterator.next();
			LocalDate localDate = snapShotsEntry.getKey();

			if (!snapshotsIterator.hasNext()) {
				// must be a single date => use it
				setDateIfDiff(localDate);
			} else {
				// have multiple latest dates for security prices
				writeFormatted("NWSYNC17", localDate.format(dateFmt), snapShotsEntry.getValue());

				while (snapshotsIterator.hasNext()) {
					snapShotsEntry = snapshotsIterator.next();
					localDate = snapShotsEntry.getKey();

					// The following security prices were last updated on %s: %s
					writeFormatted("NWSYNC17", localDate.format(dateFmt),
						snapShotsEntry.getValue());
				} // end while
			}
		}

	} // end analyzeSecurityDates()

	/**
	 * @param localDate the new date to use
	 */
	private void setDateIfDiff(LocalDate localDate) throws MduException {
		LocalDate oldLocalDate = this.latestDateCell.getDateValue();

		if (!localDate.equals(oldLocalDate)) {

			if (localDate.getMonthValue() == oldLocalDate.getMonthValue()
					&& localDate.getYear() == oldLocalDate.getYear()) {
				// Change rightmost date from %s to %s.
				writeFormatted("NWSYNC15", oldLocalDate.format(dateFmt),
					localDate.format(dateFmt));

				this.latestDateCell.setNewValue(MdUtil.convLocalToDateInt(localDate));
			} else if (localDate.isAfter(oldLocalDate)) {
				// A new month column is needed to change date from %s to %s.
				writeFormatted("NWSYNC16", oldLocalDate.format(dateFmt),
					localDate.format(dateFmt));

				getCalcDoc().forgetChanges();
			}
		}

	} // end setDateIfDiff(LocalDate)

	/**
	 * Set the spreadsheet security price if it differs from Moneydance for the
	 * latest date found in the spreadsheet.
	 *
	 * @param val the cell to potentially change
	 * @param security the corresponding Moneydance security data
	 */
	private void setPriceIfDiff(CellHandler val, CurrencyType security) throws MduException {
		double price = getLatestPrice(security);
		Number oldVal = val.getValue();

		if (oldVal instanceof Double) {
			double oldPrice = MdUtil.roundPrice(oldVal.doubleValue());

			if (price != oldPrice) {
				// Change %s (%s) price from %s to %s (<span class="%s">%+.2f%%</span>).
				NumberFormat priceFmt = val.getNumberFormat();
				String spanCl = price < oldPrice ? CL_DECREASE : CL_INCREASE;
				writeFormatted("NWSYNC10", security.getName(), security.getTickerSymbol(),
					val.getDisplayText(), priceFmt.format(price), spanCl,
					(price / oldPrice - 1) * 100);

				val.setNewValue(price);
				++this.numPricesSet;
			}
		}

	} // end setPriceIfDiff(CellHandler, CurrencyType)

	/**
	 * Set the spreadsheet account balance if it differs from Moneydance.
	 *
	 * @param val the cell to potentially change
	 * @param account the corresponding Moneydance account
	 * @param keyVal the spreadsheet name of this account
	 */
	private void setBalanceIfDiff(CellHandler val, Account account, String keyVal)
			throws MduException {
		double balance = MdUtil.getCurrentBalance(account);

		if (account.getAccountType() == CREDIT_CARD) {
			balance = -balance;
		}
		Number oldBalance = val.getValue();

		if ((oldBalance instanceof Double) && balance != oldBalance.doubleValue()) {
			// Change %s balance from %s to %s.
			writeFormatted("NWSYNC11", keyVal, val.getDisplayText(),
				val.getNumberFormat().format(balance));

			val.setNewValue(balance);
			++this.numBalancesSet;
		}

	} // end setBalanceIfDiff(CellHandler, Account, String)

	/**
	 * Capture row with 'Date' in first column.
	 *
	 * @param rowIterator
	 * @param calcDoc
	 * @return True when found
	 */
	private boolean findDateRow(XEnumeration rowIterator, CalcDoc calcDoc)
			throws MduException {
		while (rowIterator.hasMoreElements()) {
			XCellRange row = CalcDoc.next(XCellRange.class, rowIterator); // get next row
			XCell c = calcDoc.getCellByIndex(row, 0); // get its first column

			if (CalcDoc.isContentType(c, TEXT)
					&& "Date".equalsIgnoreCase(CellHandler.asDisplayText(c))) {
				this.dateRow = row;

				return true;
			}
		} // end while

		// Unable to find row with 'Date' in first column in %s.
		writeFormatted("NWSYNC01", getCalcDoc());

		return false;
	} // end findDateRow(XEnumeration)

	/**
	 * Capture index of the rightmost date in the date row. Also capture the
	 * corresponding cell handler.
	 *
	 * @return True when found
	 */
	private boolean findLatestDate() throws MduException {
		int cellIndex = 0;
		CellHandler c = null;

		do {
			// capture the rightmost date cell handler so far
			this.latestDateCell = (DateCellHandler) c;
			c = getCalcDoc().getCellHandlerByIndex(this.dateRow, ++cellIndex);
		} while (c instanceof DateCellHandler);

		if (cellIndex == 1) {
			// Unable to find any dates in the row with 'Date' in first column in %s.
			writeFormatted("NWSYNC02", getCalcDoc());

			return false;
		}

		// capture index of the rightmost date in the date row
		this.latestColumn = cellIndex - 1;

		// Found date [%s] in %s.
		writeFormatted("NWSYNC13", this.latestDateCell.getDateValue().format(dateFmt),
			getCalcDoc());

		return true;
	} // end findLatestDate()

	/**
	 * Commit any changes to the spreadsheet document.
	 */
	public void commitChanges() {
		if (isModified()) {
			this.calcDoc.commitChanges();

			// Changed %d security price%s and %d account balance%s.
			writeFormatted("NWSYNC12", this.numPricesSet, this.numPricesSet == 1 ? "" : "s",
				this.numBalancesSet, this.numBalancesSet == 1 ? "" : "s");
		}

		forgetChanges();

	} // end commitChanges()

	/**
	 * Clear out any pending changes.
	 */
	public void forgetChanges() {
		if (this.calcDoc != null) {
			this.calcDoc.forgetChanges();
		}
		this.numPricesSet = 0;
		this.numBalancesSet = 0;

	} // end forgetChanges()

	/**
	 * @return True when we have uncommitted changes in memory
	 */
	public boolean isModified() {

		return this.calcDoc != null && this.calcDoc.isModified();
	} // end isModified()

	/**
	 * @return The currently open spreadsheet document
	 */
	private CalcDoc getCalcDoc() throws MduException {
		if (this.calcDoc == null) {
			List<XSpreadsheetDocument> docList = CalcDoc.getSpreadsheetDocs(this);

			switch (docList.size()) {
			case 0:
				// No open spreadsheet documents found.
				writeFormatted("NWSYNC05");
				break;

			case 1:
				// found one => use it
				this.calcDoc = new CalcDoc(docList.get(0), this);
				break;

			default:
				// Found %d open spreadsheet documents. Can only work with one.
				writeFormatted("NWSYNC04", docList.size());
				break;
			}
		}

		return this.calcDoc;
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

	/**
	 * @return Our message bundle
	 */
	private ResourceBundle getMsgBundle() {
		if (this.msgBundle == null) {
			this.msgBundle = MdUtil.getMsgBundle(MessageWindow.baseMessageBundleName,
				this.locale);
		}

		return this.msgBundle;
	} // end getMsgBundle()

	/**
	 * @param key The resource bundle key (or message)
	 * @return Message for this key
	 */
	public String retrieveMessage(String key) {
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
		this.messageWindow.addText(String.format(this.locale, retrieveMessage(key), params));

	} // end writeFormatted(String, Object...)

} // end class OdsAccessor
