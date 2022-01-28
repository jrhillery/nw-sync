/*
 * Created on Nov 10, 2017
 */
package com.moneydance.modules.features.nwsync;

import com.moneydance.apps.md.controller.FeatureModule;

/**
 * Module used to synchronize John's NW spreadsheet document with Moneydance.
 */
@SuppressWarnings("unused")
public class Main extends FeatureModule implements AutoCloseable {
	private NwSyncConsole syncConsole = null;
	private NwSyncWorker syncWorker = null;

	/**
	 * Register this module to be invoked via the Extensions menu.
	 *
	 * @see com.moneydance.apps.md.controller.FeatureModule#init()
	 */
	public void init() {
		getContext().registerFeature(this, "do:nw:sync", null, getName());

	} // end init()

	/**
	 * This is called when this extension is invoked.
	 *
	 * @see com.moneydance.apps.md.controller.FeatureModule#invoke(java.lang.String)
	 */
	public void invoke(String uri) {
		System.err.format("%s invoked with uri [%s].%n", getName(), uri);

		try {
			if (this.syncWorker != null) {
				this.syncWorker.stopExecute();
			}
			showConsole();
			this.syncConsole.clearText();

			// SwingWorker instances are not reusable, so make a new one
			this.syncWorker = new NwSyncWorker(this.syncConsole, getName(), getContext());
			this.syncWorker.execute();
		} catch (Throwable e) {
			handleException(e);
		}

	} // end invoke(String)

	private void handleException(Throwable e) {
		this.syncConsole.addText(e.toString());
		this.syncConsole.enableCommitButton(false);
		e.printStackTrace(System.err);

	} // end handleException(Throwable)

	/**
	 * Stop execution, close our console window and release resources.
	 */
	public synchronized void cleanup() {
		if (this.syncConsole != null)
			this.syncConsole = this.syncConsole.goAway();

		if (this.syncWorker != null)
			this.syncWorker = this.syncWorker.stopExecute();

	} // end cleanup()

	public String getName() {

		return "NW Sync";
	} // end getName()

	/**
	 * Show our console window.
	 */
	private synchronized void showConsole() {
		if (this.syncConsole == null) {
			this.syncConsole = new NwSyncConsole(getName());
			this.syncConsole.addCloseableResource(this);
			this.syncConsole.setVisible(true);
		} else {
			this.syncConsole.setVisible(true);
			this.syncConsole.toFront();
			this.syncConsole.requestFocus();
		}

	} // end showConsole()

	/**
	 * Closes this resource, relinquishing any underlying resources.
	 */
	public void close() {
		this.syncConsole = null;
		this.syncWorker = null;

	} // end close()

} // end class Main
