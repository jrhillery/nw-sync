/*
 * Created on Nov 10, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.sun.star.table.CellContentType.FORMULA;
import static com.sun.star.table.CellContentType.TEXT;
import static com.sun.star.uno.UnoRuntime.queryInterface;
import static java.time.format.FormatStyle.MEDIUM;

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

import com.infinitekind.moneydance.model.Account;
import com.infinitekind.moneydance.model.AccountBook;
import com.infinitekind.moneydance.model.CurrencySnapshot;
import com.infinitekind.moneydance.model.CurrencyTable;
import com.infinitekind.moneydance.model.CurrencyType;
import com.leastlogic.moneydance.util.*;
import com.leastlogic.swing.util.HTMLPane;
import com.moneydance.modules.features.nwsync.CellHandler.DateCellHandler;
import com.sun.star.bridge.XBridge;
import com.sun.star.bridge.XBridgeFactory;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XEnumeration;
import com.sun.star.frame.XDesktop2;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.uno.XComponentContext;

import ooo.connector.BootstrapSocketConnector;
import ooo.connector.server.OOoServer;

/**
 * Provides read/write access to an ods (OpenOffice/LibreOffice) spreadsheet
 * document.
 */
public class OdsAccessor implements StagedInterface, AutoCloseable {
	private final NwSyncWorker syncWorker;
	private final Locale locale;
	private final Account root;
	private final CurrencyTable securities;

	private CalcDoc calcDoc = null;
	private XCellRange dateRow = null;
	private int latestColumn = 0;
	private DateCellHandler latestDateCell = null;
	private int[] earlierDates = null;
	private int numPricesSet = 0;
	private int numBalancesSet = 0;
	private int numDatesSet = 0;
	private final TreeMap<LocalDate, List<String>> securitySnapshots = new TreeMap<>();
	private Properties nwSyncProps = null;

	private static final String propertiesFileName = "nw-sync.properties";
	private static final DateTimeFormatter dateFmt = DateTimeFormatter.ofLocalizedDate(MEDIUM);

	/**
	 * Sole constructor.
	 *
	 * @param syncWorker  The worker we can use to send messages to the event dispatch thread
	 * @param locale      Our message window's Locale
	 * @param accountBook Moneydance account book
	 */
	public OdsAccessor(NwSyncWorker syncWorker, Locale locale, AccountBook accountBook) {
		this.syncWorker = syncWorker;
		this.locale = locale;
		this.root = accountBook.getRootAccount();
		this.securities = accountBook.getCurrencies();

	} // end constructor

	/**
	 * Synchronize data between a spreadsheet document and Moneydance.
	 */
	public void syncNwData() throws MduException {
		loadCalcDoc();

		if (this.calcDoc == null)
			return; // nothing to synchronize

		XEnumeration rowItr = this.calcDoc.getFirstSheetRowIterator();

		if (!findDateRow(rowItr) || !findLatestDate())
			return; // can't synchronize without a date row and latest date

		while (rowItr.hasMoreElements()) {
			XCellRange row = next(XCellRange.class, rowItr); // get next row
			XCell key = this.calcDoc.getCellByIndex(row, 0); // get its first column

			if (CalcDoc.isContentType(key, TEXT) || CalcDoc.isContentType(key, FORMULA)) {
				String keyVal = CellHandler.asDisplayText(key);
				CellHandler val = this.calcDoc.getCellHandlerByIndex(row, this.latestColumn);

				if (val != null) {
					CurrencyType security = this.securities.getCurrencyByTickerSymbol(keyVal);

					if (security != null) {
						// found this row's ticker symbol in Moneydance securities
						SnapshotList ssList = new SnapshotList(security);
						setTodaysPriceIfDiff(val, ssList);
						setEarlierPricesIfDiff(row, ssList);
					} else {
						getAccount(keyVal).ifPresentOrElse(account -> {
							// found this row's account in Moneydance
							setTodaysBalIfDiff(val, account, keyVal);
							setEarlierBalsIfDiff(row, account, keyVal);
						}, () ->
							MdLog.all("Ignoring row %s".formatted(keyVal)));
					}
				}
			}

			if (this.syncWorker.isCancelled())
				return;
		} // end while
		analyzeSecurityDates();

		if (!isModified()) {
			this.syncWorker.display("No new price or balance data found");
		}

	} // end syncNwData()

	/**
	 * @param keyVal Account name:subaccount name
	 * @return The Moneydance account corresponding to keyVal
	 */
	private Optional<Account> getAccount(String keyVal) {
		final String[] actNames = keyVal.split(":");
		Optional<Account> account = Optional.ofNullable(this.root.getAccountByName(actNames[0]));

		if (account.isPresent() && actNames.length > 1) {
			account = MdUtil.getSubAccountByName(account.get(), actNames[1]);
		}

		return account;
	} // end getAccount(String)

	/**
	 * @param snapshotList The list of snapshots to use
	 * @return Today's price in the snapshot list supplied
	 */
	private BigDecimal getTodaysPrice(SnapshotList snapshotList) {
		CurrencyType security = snapshotList.getSecurity();
		Optional<CurrencySnapshot> currentSnapshot = snapshotList.getTodaysSnapshot();

		if (currentSnapshot.isEmpty())
			return BigDecimal.ONE; // default price to 1 when no snapshot

		if (!MdUtil.isIBondTickerPrefix(security.getTickerSymbol())) {
			// add this snapshot to our collection
			getSecurityListForDate(currentSnapshot.get())
				.add(security.getName() + " (" + security.getTickerSymbol() + ')');
		}

		return MdUtil.getAndValidateCurrentSnapshotPrice(security, currentSnapshot.get(),
			this.locale, this.syncWorker::display);
	} // end getTodaysPrice(SnapshotList)

	/**
	 * @param snapshotList The list of snapshots to use
	 * @param asOfDates    The dates to obtain the price for
	 * @return Security prices as of the end of each date in asOfDates
	 */
	private BigDecimal[] getPricesAsOfDates(SnapshotList snapshotList, int[] asOfDates) {
		BigDecimal[] prices = new BigDecimal[asOfDates.length];

		for (int i = prices.length - 1; i >= 0; --i) {
			prices[i] = snapshotList.getSnapshotForDate(asOfDates[i])
					.map(SnapshotList::getPrice).orElse(BigDecimal.ONE);
		} // end for

		return prices;
	} // end getPricesAsOfDates(SnapshotList, int[])

	/**
	 * @param currentSnapshot Last currency snapshot before, or on, today
	 * @return The list of security names with the same date
	 */
	private List<String> getSecurityListForDate(CurrencySnapshot currentSnapshot) {
		LocalDate marketDate = MdUtil.convDateIntToLocal(currentSnapshot.getDateInt());

		return this.securitySnapshots.computeIfAbsent(marketDate, k -> new ArrayList<>());
	} // end getSecurityListForDate(int)

	/**
	 * Analyze security dates to see if they are all the same.
	 */
	private void analyzeSecurityDates() {
		if (this.securitySnapshots.size() == 1) {
			// just a single date => use it
			setDateIfDiff(this.securitySnapshots.firstKey());
		} else {
			// have multiple latest dates for security prices
			this.securitySnapshots.forEach(this::reportOneOfMultipleDates);
		}

	} // end analyzeSecurityDates()

	/**
	 * @param marketDate     The date these securities were updated
	 * @param daysSecurities The list of security names updated on market date
	 */
	private void reportOneOfMultipleDates(LocalDate marketDate, List<String> daysSecurities) {
		this.syncWorker.display("Prices last updated on %s: %s"
			.formatted(marketDate.format(dateFmt), daysSecurities));
		LocalDate oldDate = this.latestDateCell.getDateValue();

		if (marketDate.isAfter(oldDate)
				&& (marketDate.getMonthValue() != oldDate.getMonthValue()
						|| marketDate.getYear() != oldDate.getYear())) {
			handleNewMonth(marketDate, oldDate);
		}

	} // end reportOneOfMultipleDates(LocalDate, List<String>)

	/**
	 * @param marketDate The new date to use
	 */
	private void setDateIfDiff(LocalDate marketDate) {
		LocalDate oldDate = this.latestDateCell.getDateValue();

		if (!marketDate.equals(oldDate)) {

			if (marketDate.getMonthValue() == oldDate.getMonthValue()
					&& marketDate.getYear() == oldDate.getYear()) {
				this.syncWorker.display("Change the rightmost date from %s to %s"
					.formatted(oldDate.format(dateFmt), marketDate.format(dateFmt)));

				this.latestDateCell.setNewValue(MdUtil.convLocalToDateInt(marketDate));
				++this.numDatesSet;
			} else if (marketDate.isAfter(oldDate)) {
				handleNewMonth(marketDate, oldDate);
			}
		}

	} // end setDateIfDiff(LocalDate)

	/**
	 * @param marketDate The new date to use
	 * @param oldDate    The date from the spreadsheet
	 */
	private void handleNewMonth(LocalDate marketDate, LocalDate oldDate) {
		this.syncWorker.display("A new month column is needed to change date from %s to %s"
			.formatted(oldDate.format(dateFmt), marketDate.format(dateFmt)));

		this.calcDoc.forgetChanges();

	} // end handleNewMonth(LocalDate, LocalDate)

	/**
	 * Set the spreadsheet security price if it differs from Moneydance for the
	 * latest date column found in the spreadsheet.
	 *
	 * @param val          The cell to potentially change
	 * @param snapshotList The list of snapshots to use
	 */
	private void setTodaysPriceIfDiff(CellHandler val, SnapshotList snapshotList) {
		BigDecimal price = getTodaysPrice(snapshotList);
		setPriceIfDiff(val, price, snapshotList.getSecurity(), "today");

	} // end setTodaysPriceIfDiff(CellHandler, SnapshotList)

	/**
	 * @param val      The cell to potentially change
	 * @param price    The new price
	 * @param security The corresponding Moneydance security data
	 * @param dayStr   The applicable day
	 */
	private void setPriceIfDiff(CellHandler val, BigDecimal price, CurrencyType security,
			String dayStr) {
		Number oldVal = val.getValue();

		if (oldVal instanceof Double) {
			BigDecimal oldPrice = MdUtil.roundPrice(oldVal.doubleValue());

			if (price.compareTo(oldPrice) != 0) {
				NumberFormat priceFmt = MdUtil.getCurrencyFormat(this.locale, oldPrice, price);
				this.syncWorker.display(
					"Change %s (%s) price for %s from %s to %s (<span class=\"%s\">%+.2f%%</span>)"
					.formatted(security.getName(), security.getTickerSymbol(), dayStr,
					priceFmt.format(oldPrice), priceFmt.format(price), HTMLPane.getSpanCl(price, oldPrice),
					(price.doubleValue() / oldPrice.doubleValue() - 1) * 100));

				val.setNewValue(price);
				++this.numPricesSet;
			}
		}

	} // end setPriceIfDiff(CellHandler, BigDecimal, CurrencyType, String)

	/**
	 * Set the spreadsheet security prices if any differ from Moneydance.
	 *
	 * @param row          The row with cells to potentially change
	 * @param snapshotList The list of snapshots to use
	 */
	private void setEarlierPricesIfDiff(XCellRange row, SnapshotList snapshotList) {
		BigDecimal[] prices = getPricesAsOfDates(snapshotList, this.earlierDates);

		for (int i = 0; i < prices.length; ++i) {
			CellHandler val = this.calcDoc.getCellHandlerByIndex(row, i + 1);

			if (val != null) {
				String dayStr = MdUtil.convDateIntToLocal(this.earlierDates[i]).format(dateFmt);
				setPriceIfDiff(val, prices[i], snapshotList.getSecurity(), dayStr);
			}
		} // end for

	} // end setEarlierPricesIfDiff(XCellRange, SnapshotList)

	/**
	 * Set the spreadsheet account balance if it differs from Moneydance for the
	 * latest date column found in the spreadsheet.
	 *
	 * @param val     The cell to potentially change
	 * @param account The corresponding Moneydance account
	 * @param keyVal  The spreadsheet name of this account
	 */
	private void setTodaysBalIfDiff(CellHandler val, Account account, String keyVal) {
		BigDecimal balance = MdUtil.getCurrentBalance(account);
		setBalanceIfDiff(val, balance, keyVal, "today");

	} // end setTodaysBalIfDiff(CellHandler, Account, String)

	/**
	 * @param val     The cell to potentially change
	 * @param balance The new balance
	 * @param keyVal  The spreadsheet name of this account
	 * @param dayStr  The applicable day
	 */
	private void setBalanceIfDiff(CellHandler val, BigDecimal balance, String keyVal,
			String dayStr) {
		Number oldBalance = val.getValue();

		if (oldBalance instanceof Double) {
			// compare balance rounded to 13 digit precision
			BigDecimal oldBal = MdUtil.roundPrice(oldBalance.doubleValue());

			if (balance.compareTo(oldBal) != 0) {
				NumberFormat nf = val.isCurrency()
					? MdUtil.getCurrencyFormat(this.locale, oldBal, balance)
					: MdUtil.getNumberFormat(this.locale, oldBal, balance);

				this.syncWorker.display("Change %s balance for %s from %s to %s"
					.formatted(keyVal, dayStr, nf.format(oldBal), nf.format(balance)));

				val.setNewValue(balance);
				++this.numBalancesSet;
			}
		}

	} // end setBalanceIfDiff(CellHandler, BigDecimal, String, String)

	/**
	 * Set the spreadsheet account balances if any differ from Moneydance.
	 *
	 * @param row     The row with cells to potentially change
	 * @param account The corresponding Moneydance account
	 * @param keyVal  The spreadsheet name of this account
	 */
	private void setEarlierBalsIfDiff(XCellRange row, Account account, String keyVal) {
		BigDecimal[] balances = MdUtil.getBalancesAsOfDates(this.root.getBook(), account,
			this.earlierDates);

		for (int i = 0; i < balances.length; ++i) {
			CellHandler val = this.calcDoc.getCellHandlerByIndex(row, i + 1);

			if (val != null) {
				String dayStr = MdUtil.convDateIntToLocal(this.earlierDates[i]).format(dateFmt);
				setBalanceIfDiff(val, balances[i], keyVal, dayStr);
			}
		} // end for

	} // end setEarlierBalsIfDiff(XCellRange, Account, String)

	/**
	 * Capture row with 'Date' in first column.
	 *
	 * @param rowIterator Spreadsheet row iterator
	 * @return True when found
	 */
	private boolean findDateRow(XEnumeration rowIterator)
			throws MduException {
		while (rowIterator.hasMoreElements()) {
			XCellRange row = next(XCellRange.class, rowIterator); // get next row
			XCell c = this.calcDoc.getCellByIndex(row, 0); // get its first column

			if (CalcDoc.isContentType(c, TEXT)
					&& "Date".equalsIgnoreCase(CellHandler.asDisplayText(c))) {
				this.dateRow = row;

				return true;
			}
		} // end while

		this.syncWorker.display("Unable to find row with 'Date' in first column in %s"
			.formatted(this.calcDoc));

		return false;
	} // end findDateRow(XEnumeration)

	/**
	 * Capture index of the rightmost date in the date row. Also capture the
	 * corresponding cell handler.
	 *
	 * @return True when found
	 */
	private boolean findLatestDate() {
		int cellIndex = 0;
		CellHandler c = null;
		ArrayList<LocalDate> dates = new ArrayList<>();

		do {
			// capture the rightmost date cell handler so far
			this.latestDateCell = (DateCellHandler) c;

			if (this.latestDateCell != null) {
				dates.add(this.latestDateCell.getDateValue());
			}
			c = this.calcDoc.getCellHandlerByIndex(this.dateRow, ++cellIndex);
		} while (c instanceof DateCellHandler);

		if (cellIndex == 1) {
			this.syncWorker.display("Unable to find any dates in the row with 'Date' in first column in %s"
				.formatted(this.calcDoc));

			return false;
		}

		// capture index of the rightmost date in the date row
		this.latestColumn = cellIndex - 1;

		// save each earlier date
		this.earlierDates = new int[dates.size() - 1];

		for (int i = 0; i < this.earlierDates.length; ++i) {
			this.earlierDates[i] = MdUtil.convLocalToDateInt(dates.get(i));
		}

		this.syncWorker.display("Found date [%s] in %s"
			.formatted(this.latestDateCell.getDateValue().format(dateFmt), this.calcDoc));

		return true;
	} // end findLatestDate()

	/**
	 * Commit any changes to the spreadsheet document.
	 *
	 * @return Optional summary of the changes committed
	 */
	public Optional<String> commitChanges() {
		Optional<String> commitText = Optional.empty();

		if (isModified()) {
			this.calcDoc.commitChanges();
			String msg = (this.numDatesSet == 1)
				? "Changed %d security price%s, %d account balance%s and the rightmost date"
				: "Changed %d security price%s, %d account balance%s and %d dates";
			commitText = Optional.of(msg.formatted(
				this.numPricesSet, this.numPricesSet == 1 ? "" : "s",
				this.numBalancesSet, this.numBalancesSet == 1 ? "" : "s", this.numDatesSet));
		}

		forgetChanges();

		return commitText;
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
		this.numDatesSet = 0;
		this.securitySnapshots.clear();

	} // end forgetChanges()

	/**
	 * @return True when we have uncommitted changes in memory
	 */
	public boolean isModified() {

		return this.calcDoc != null && this.calcDoc.isModified();
	} // end isModified()

	/**
	 * Load the currently open spreadsheet document into this instance.
	 */
	private void loadCalcDoc() throws MduException {
		CalcDoc calcDoc = getCalcDoc();

		if (calcDoc != null && calcDoc.getSheets() == null) {
			// can't access the sheets, force a reconnection
			calcDoc = getCalcDoc();
		}
		this.calcDoc = calcDoc;

	} // end loadCalcDoc()

	/**
	 * @return The currently open spreadsheet document
	 */
	private CalcDoc getCalcDoc() throws MduException {
		CalcDoc calcDoc = null;
		List<XSpreadsheetDocument> docList = getSpreadsheetDocs();

		switch (docList.size()) {
			case 0 ->
					this.syncWorker.display("No open spreadsheet documents found");
			case 1 ->
					// found one => use it
					calcDoc = new CalcDoc(docList.getFirst());
			default ->
					this.syncWorker.display("Found %d open spreadsheet documents; Can only work with one"
						.formatted(docList.size()));
		}

		return calcDoc;
	} // end getCalcDoc()

	/**
	 * @return A list of currently open spreadsheet documents
	 */
	private List<XSpreadsheetDocument> getSpreadsheetDocs() throws MduException {
		List<XSpreadsheetDocument> docList = new ArrayList<>();
		XDesktop2 libreOfficeDesktop = getOfficeDesktop();
		XEnumeration compItr = libreOfficeDesktop.getComponents().createEnumeration();

		if (!compItr.hasMoreElements()) {
			// no components so we probably started the desktop => terminate it
			libreOfficeDesktop.terminate();
		} else {
			do {
				XServiceInfo comp = next(XServiceInfo.class, compItr);

				if (comp.supportsService("com.sun.star.sheet.SpreadsheetDocument")) {
					docList.add(queryInterface(XSpreadsheetDocument.class, comp));
				}
			} while (compItr.hasMoreElements());
		}

		return docList;
	} // end getSpreadsheetDocs()

	/**
	 * @return A LibreOffice desktop interface
	 */
	private XDesktop2 getOfficeDesktop() throws MduException {
		final String OFFICE_PATH = "office.install.path";
		String officeInstallPath = getNwSyncProps().getProperty(OFFICE_PATH);
		if (officeInstallPath == null)
			throw new MduException(null, "Unable to obtain %s from %s on the class path",
				OFFICE_PATH, propertiesFileName);

		XComponentContext remoteContext;
		try {
			List<String> oooOptions = Arrays.asList(Bootstrap.getDefaultOptions());
			remoteContext = new BootstrapSocketConnector(
				new OOoServer(officeInstallPath, oooOptions)).connect();
		} catch (Throwable e) {
			throw new MduException(e, "Exception obtaining office context");
		}
		if (remoteContext == null)
			throw new MduException(null, "Unable to obtain office context");

		XMultiComponentFactory remoteServiceMgr = remoteContext.getServiceManager();
		if (remoteServiceMgr == null)
			throw new MduException(null, "Unable to obtain office service manager");

		XDesktop2 libreOfficeDesktop;
		try {
			libreOfficeDesktop = queryInterface(XDesktop2.class, remoteServiceMgr
				.createInstanceWithContext("com.sun.star.frame.Desktop", remoteContext));
		} catch (Exception e) {
			throw new MduException(e, "Exception obtaining office desktop");
		}
		if (libreOfficeDesktop == null)
			throw new MduException(null, "Unable to obtain office desktop");

		return libreOfficeDesktop;
	} // end getOfficeDesktop()

	/**
	 * @param zInterface The type to return
	 * @param iterator   The iterator to use
	 * @return The next zInterface from the iterator
	 */
	private static <T> T next(Class<T> zInterface, XEnumeration iterator) throws MduException {
		T result;
		try {
			result = queryInterface(zInterface, iterator.nextElement());
		} catch (Exception e) {
			throw new MduException(e, "Exception iterating to next %s",
				zInterface.getSimpleName());
		}
		if (result == null)
			throw new MduException(null, "Unable to obtain next %s",
				zInterface.getSimpleName());

		return result;
	} // end next(Class<T>, XEnumeration)

	/**
	 * Release any resources we acquired. This includes closing the connection to
	 * the office process.
	 */
	public void close() {
		closeOfficeConnection();
		this.calcDoc = null;

	} // end close()

	/**
	 * Close our connection to the office process.
	 */
	public static void closeOfficeConnection() {
		try {
			// get the bridge factory from the local service manager
			XBridgeFactory bridgeFactory = queryInterface(XBridgeFactory.class,
				Bootstrap.createSimpleServiceManager()
					.createInstance("com.sun.star.bridge.BridgeFactory"));

			if (bridgeFactory != null) {
				for (XBridge bridge : bridgeFactory.getExistingBridges()) {
					// dispose of this bridge after closing its connection
					queryInterface(XComponent.class, bridge).dispose();
				}
			}
			MdLog.debug("Office connection closed");
		} catch (Throwable e) {
			MdLog.all("Problem disposing office process connection bridge", e);
		}

	} // end closeOfficeConnection()

	/**
	 * @return Our properties
	 */
	public Properties getNwSyncProps() throws MduException {
		if (this.nwSyncProps == null) {
			this.nwSyncProps = MdUtil.loadProps(propertiesFileName, getClass());
		}

		return this.nwSyncProps;
	} // end getNwSyncProps()

} // end class OdsAccessor
