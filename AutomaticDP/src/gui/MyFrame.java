package gui;

import java.awt.Dimension;

import javax.swing.JFrame;

@SuppressWarnings("serial")
public class MyFrame extends JFrame {

	private MyPanel startPanel;

	public MyFrame() {
		startPanel = new MyPanel();

		setupFrame();
	}

	private void setupFrame() {
		setTitle("AutomaticDP");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		getContentPane().add(startPanel);
		setSize(new Dimension(400, 200));
		setVisible(true);
	}

}
