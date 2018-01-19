/*
 * Created on Nov 12, 2017
 */
package com.moneydance.modules.features.nwsync;

import static javax.swing.GroupLayout.DEFAULT_SIZE;

import java.awt.AWTEvent;
import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowEvent;
import java.io.InputStream;

import javax.imageio.ImageIO;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.WindowConstants;
import javax.swing.border.EmptyBorder;

import com.johns.swing.util.HTMLPane;

public class MessageWindow extends JFrame implements ActionListener {
	private Main feature;
	private JButton btnCommit;
	private HTMLPane pnOutputLog;

	private static final long serialVersionUID = -3229995555109327972L;

	/**
	 * Create the frame.
	 *
	 * @param feature
	 */
	public MessageWindow(Main feature) {
		super((feature == null ? "Message" : feature.getName()) + " Console");
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
		setSize(576, 356);
		JPanel contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);

		this.btnCommit = new JButton("Commit Changes");
		this.btnCommit.setEnabled(false);
		reducePreferredHeight(this.btnCommit);
		this.btnCommit.setToolTipText("Commit changes to the spreadsheet");

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
	 * @param button
	 */
	private void reducePreferredHeight(JComponent button) {
		HTMLPane.reduceHeight(button, 20);

	} // end reducePreferredHeight(JComponent)

	/**
	 * Wire in our event listeners.
	 */
	private void wireEvents() {
		this.btnCommit.addActionListener(this);

	} // end wireEvents()

	/**
	 * Read in and set our icon image.
	 */
	private void readIconImage() {
		InputStream stream = getClass().getResourceAsStream("update-icon24.png");

		if (stream != null) {
			try {
				setIconImage(ImageIO.read(stream));
			} catch (Exception e) {
				System.err.println("Exception reading icon image" + e);
			} finally {
				try {
					stream.close();
				} catch (Exception e) { /* ignore */ }
			}
		}

	} // end readIconImage()

	/**
	 * Invoked when an action occurs.
	 *
	 * @param event
	 */
	public void actionPerformed(ActionEvent event) {
		Object source = event.getSource();

		if (source == this.btnCommit && this.feature != null) {
			this.feature.commitChanges();
		}

	} // end actionPerformed(ActionEvent)

	/**
	 * @param text HTML text to append to the output log text area
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
	 * Processes events on this window.
	 *
	 * @param event
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
	public MessageWindow goAway() {
		setVisible(false);
		dispose();

		return null;
	} // end goAway()

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					MessageWindow frame = new MessageWindow(null);
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});

	} // end main(String[])

} // end class MessageWindow
