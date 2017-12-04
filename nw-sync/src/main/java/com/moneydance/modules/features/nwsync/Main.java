/*
 * Created on Nov 10, 2017
 */
package com.moneydance.modules.features.nwsync;

import java.io.File;

import com.moneydance.apps.md.controller.FeatureModule;
import com.sun.star.lib.util.NativeLibraryLoader;

/**
 * Module used to synchronize John's NW spreadsheet document with Moneydance.
 */
public class Main extends FeatureModule {
	private MessageWindow messageWindow = null;
	private OdsAccessor odsAcc = null;

	/**
	 * Register this module to be invoked via the extensions menu.
	 *
	 * @see com.moneydance.apps.md.controller.FeatureModule#init()
	 */
	public void init() {
		getContext().registerFeature(this, "donwsync", null, getName());

	} // end init()

	/**
	 * This is called when this extension is invoked.
	 *
	 * @see com.moneydance.apps.md.controller.FeatureModule#invoke(java.lang.String)
	 */
	public void invoke(String uri) {
		System.err.println(getName() + " invoked with uri [" + uri + ']');
		showConsole();
		File fOffice = NativeLibraryLoader.getResource(Main.class.getClassLoader(),
				"soffice.exe");
		System.err.println("Office exe is [" + fOffice + ']');

		this.odsAcc = new OdsAccessor(getContext().getCurrentAccountBook());
		try {
			this.odsAcc.syncNwData();
			this.messageWindow.setText(this.odsAcc.getMessages());
			this.messageWindow.enableCommitButton(this.odsAcc.isModified());
		} catch (Throwable e) {
			e.printStackTrace(System.err);
		}

	} // end invoke(String)

	/**
	 * This is called when the commit button is selected.
	 */
	void commitChanges() {
		try {
			this.odsAcc.commitChanges();
			this.messageWindow.setText(this.odsAcc.getMessages());
			this.messageWindow.enableCommitButton(this.odsAcc.isModified());
		} catch (Throwable e) {
			e.printStackTrace(System.err);
		}

	} // end commitChanges()

	public void cleanup() {
		closeConsole();

		if (this.odsAcc != null)
			this.odsAcc = this.odsAcc.releaseResources();

	} // end cleanup()

	public String getName() {

		return "NW Sync";
	} // end getName()

	private synchronized void showConsole() {
		if (this.messageWindow == null) {
			this.messageWindow = new MessageWindow(this);
			this.messageWindow.setVisible(true);
		} else {
			this.messageWindow.setVisible(true);
			this.messageWindow.toFront();
			this.messageWindow.requestFocus();
		}

	} // end showConsole()

	synchronized void closeConsole() {
		if (this.messageWindow != null)
			this.messageWindow = this.messageWindow.goAway();

		if (this.odsAcc != null)
			this.odsAcc = this.odsAcc.releaseResources();

	} // end closeConsole()

} // end class Main
