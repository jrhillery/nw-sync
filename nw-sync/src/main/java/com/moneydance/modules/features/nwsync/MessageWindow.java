/*
 * Created on Nov 12, 2017
 */
package com.moneydance.modules.features.nwsync;

import java.awt.AWTEvent;
import java.awt.GridBagLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowEvent;

import javax.swing.Box;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.WindowConstants;
import javax.swing.border.EmptyBorder;

import com.moneydance.awt.AwtUtil;

public class MessageWindow extends JFrame implements ActionListener {
	private Main feature;
	private JTextArea messageArea;
	private JButton commitButton;
	private JButton closeButton;

	private static final long serialVersionUID = -3229995555109327972L;

	public MessageWindow(Main feature) {
		super(feature.getName() + " Console");
		this.feature = feature;

		this.messageArea = new JTextArea(feature.getName() + " loading...");
		this.messageArea.setEditable(false);
		this.commitButton = new JButton("Commit Changes");
		this.closeButton = new JButton("Done");

		JPanel p = new JPanel(new GridBagLayout());
		p.setBorder(new EmptyBorder(10, 10, 10, 10));
		p.add(new JScrollPane(this.messageArea), AwtUtil.getConstraints(0, 0, 1, 1, 4, 1, true, true));
		p.add(Box.createVerticalStrut(8), AwtUtil.getConstraints(0, 2, 0, 0, 1, 1, false, false));
		p.add(this.commitButton, AwtUtil.getConstraints(0, 3, 1, 0, 1, 1, false, true));
		p.add(this.closeButton, AwtUtil.getConstraints(1, 3, 1, 0, 1, 1, false, true));
		getContentPane().add(p);

		setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
		enableEvents(WindowEvent.WINDOW_CLOSING);
		this.commitButton.addActionListener(this);
		this.closeButton.addActionListener(this);
		setSize(560, 360);

	} // end (Main) constructor

	public void setText(String text) {
		this.messageArea.setText(text);

	} // end setText(String)

	public void enableCommitButton(boolean b) {
		this.commitButton.setEnabled(b);

	} // end enableCommitButton(boolean)

	public void actionPerformed(ActionEvent event) {
		Object source = event.getSource();

		if (source == this.commitButton) {
			this.feature.commitChanges();
		}

		if (source == this.closeButton) {
			this.feature.closeConsole();
		}

	} // end actionPerformed(ActionEvent)

	protected void processEvent(AWTEvent event) {
		if (event.getID() == WindowEvent.WINDOW_CLOSING) {
			this.feature.closeConsole();
		} else {
			super.processEvent(event);
		}

	} // end processEvent(AWTEvent)

	public MessageWindow goAway() {
		setVisible(false);
		dispose();

		return null;
	} // end goAway()

} // end class MessageWindow
