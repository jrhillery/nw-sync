/*
 * Created on Nov 12, 2017
 */
package com.moneydance.modules.features.nwsync;

import com.leastlogic.moneydance.util.StagedInterface;
import com.leastlogic.swing.util.HTMLPane;

import javax.swing.*;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.event.WindowEvent;
import java.io.Serial;
import java.util.ResourceBundle;

import static javax.swing.GroupLayout.DEFAULT_SIZE;

public class NwSyncConsole extends JFrame {
	private final Main feature;
	private JButton btnCommit;
	private HTMLPane pnOutputLog;
	private StagedInterface staged = null;
	private AutoCloseable closeableResource = null;

	static final String baseMessageBundleName = "com.moneydance.modules.features.nwsync.NwSyncMessages"; //$NON-NLS-1$
	private static final ResourceBundle msgBundle = ResourceBundle.getBundle(baseMessageBundleName);
	@Serial
	private static final long serialVersionUID = 8224939513161266369L;

	/**
	 * Create the frame.
	 *
	 * @param feature Our main class
	 */
	public NwSyncConsole(Main feature) {
		super((feature == null ? msgBundle.getString("NwSyncConsole.window.title.default") //$NON-NLS-1$
			: feature.getName()) + msgBundle.getString("NwSyncConsole.window.title.suffix")); //$NON-NLS-1$
		this.feature = feature;

		initComponents();
		wireEvents();
		readIconImage();

	} // end (Main) constructor

	/**
	 * Initialize the swing components.
	 */
	private void initComponents() {
		setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
		setSize(692, 428);
		JPanel contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);

		this.btnCommit = new JButton(msgBundle.getString("NwSyncConsole.btnCommit.text")); //$NON-NLS-1$
		this.btnCommit.setEnabled(false);
		reducePreferredHeight(this.btnCommit);
		this.btnCommit.setToolTipText(msgBundle.getString("NwSyncConsole.btnCommit.toolTipText")); //$NON-NLS-1$

		this.pnOutputLog = new HTMLPane();
		JScrollPane scrollPane = new JScrollPane(this.pnOutputLog);
		GroupLayout gl_contentPane = new GroupLayout(contentPane);
		gl_contentPane.setHorizontalGroup(
			gl_contentPane.createParallelGroup(Alignment.TRAILING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addContainerGap(403, Short.MAX_VALUE)
					.addComponent(this.btnCommit))
				.addComponent(scrollPane, DEFAULT_SIZE, 532, Short.MAX_VALUE)
		);
		gl_contentPane.setVerticalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addComponent(this.btnCommit)
					.addPreferredGap(ComponentPlacement.RELATED)
					.addComponent(scrollPane, DEFAULT_SIZE, 271, Short.MAX_VALUE))
		);
		contentPane.setLayout(gl_contentPane);

	} // end initComponents()

	/**
	 * @param button The button to adjust
	 */
	private void reducePreferredHeight(JComponent button) {
		HTMLPane.reduceHeight(button, 20);

	} // end reducePreferredHeight(JComponent)

	/**
	 * Wire in our event listeners.
	 */
	private void wireEvents() {
		this.btnCommit.addActionListener(event -> {
			// invoked when Commit is selected
			if (this.staged != null) {
				try {
					String changeSummary = this.staged.commitChanges();

					if (changeSummary != null)
						addText(changeSummary);
					enableCommitButton(this.staged.isModified());
				} catch (Exception e) {
					addText(e.toString());
					enableCommitButton(false);
					e.printStackTrace(System.err);
				}
			}
		}); // end btnCommit.addActionListener

	} // end wireEvents()

	/**
	 * Read in and set our icon image.
	 */
	private void readIconImage() {
		setIconImage(HTMLPane.readResourceImage("update-icon24.png", getClass())); //$NON-NLS-1$

	} // end readIconImage()

	/**
	 * @param text HTML-text to append to the output log text area
	 */
	public void addText(String text) {
		this.pnOutputLog.addText(text);

	} // end addText(String)

	/**
	 * Clear the output log text area.
	 */
	public void clearText() {
		this.pnOutputLog.clearText();

	} // end clearText()

	/**
	 * @param b true to enable the button, otherwise false
	 */
	public void enableCommitButton(boolean b) {
		this.btnCommit.setEnabled(b);

	} // end enableCommitButton(boolean)

	/**
	 * Store the object to manage staged changes.
	 *
	 * @param staged The object managing staged changes
	 */
	public void setStaged(StagedInterface staged) {
		this.staged = staged;

	} // end setStaged(StagedInterface)

	/**
	 * Store the object with resources to close.
	 *
	 * @param closable The object managing closeable resources
	 */
	public void setCloseableResource(AutoCloseable closable) {
		this.closeableResource = closable;

	} // end setCloseableResource(AutoCloseable)

	/**
	 * Processes events on this window.
	 *
	 * @param event The event to be processed
	 */
	protected void processEvent(AWTEvent event) {
		if (event.getID() == WindowEvent.WINDOW_CLOSING) {
			if (this.feature != null) {
				this.feature.closeConsole();
			} else {
				goAway();
			}
		} else {
			super.processEvent(event);
		}

	} // end processEvent(AWTEvent)

	/**
	 * Remove this frame.
	 *
	 * @return null
	 */
	public NwSyncConsole goAway() {
		Dimension winSize = getSize();
		System.err.format(getLocale(), "Closing %s with width=%.0f, height=%.0f.%n",
			getTitle(), winSize.getWidth(), winSize.getHeight());
		setVisible(false);
		dispose();

		if (this.closeableResource != null) {
			// Release any resources we acquired.
			try {
				this.closeableResource.close();
			} catch (Exception e) {
				e.printStackTrace(System.err);
			}
		}

		return null;
	} // end goAway()

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(() -> {
			try {
				NwSyncConsole frame = new NwSyncConsole(null);
				frame.setVisible(true);
				frame.enableCommitButton(true);
			} catch (Exception e) {
				e.printStackTrace();
			}
		});

	} // end main(String[])

} // end class NwSyncConsole
