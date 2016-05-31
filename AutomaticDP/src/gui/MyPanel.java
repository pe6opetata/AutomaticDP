package gui;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;

import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.LayoutStyle.ComponentPlacement;

import classes.ProcessExcelFile;

@SuppressWarnings("serial")
public class MyPanel extends JPanel {

	private JTextField txtFileName;
	private JLabel lblFileName;
	private JLabel lblColWidth;
	private GroupLayout groupLayout;
	private JButton btnOk;
	private JButton btnBrowse;
	private JButton btnExit;
	private JFileChooser fileChooser;
	private JTextField txtColWidth;

	public MyPanel() {

		lblFileName = new JLabel("File name:");
		lblColWidth = new JLabel("Col width");
		
		txtFileName = new JTextField();

		btnOk = new JButton("Submit");
		btnBrowse = new JButton("Browse...");
		btnExit = new JButton("Exit");

		fileChooser = new JFileChooser();
		
		txtColWidth = new JTextField();
		


		groupLayout = new GroupLayout(this);
		groupLayout.setHorizontalGroup(
			groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
					.addGroup(groupLayout.createParallelGroup(Alignment.LEADING, false)
						.addGroup(groupLayout.createSequentialGroup()
							.addGap(16)
							.addComponent(lblFileName)
							.addPreferredGap(ComponentPlacement.RELATED)
							.addComponent(txtFileName, GroupLayout.PREFERRED_SIZE, 266, GroupLayout.PREFERRED_SIZE))
						.addGroup(Alignment.TRAILING, groupLayout.createSequentialGroup()
							.addContainerGap()
							.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
								.addComponent(btnOk)
								.addComponent(btnBrowse))
							.addPreferredGap(ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
							.addComponent(lblColWidth)
							.addPreferredGap(ComponentPlacement.RELATED)
							.addGroup(groupLayout.createParallelGroup(Alignment.LEADING, false)
								.addComponent(txtColWidth, 0, 0, Short.MAX_VALUE)
								.addComponent(btnExit, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
					.addContainerGap(115, Short.MAX_VALUE))
		);
		groupLayout.setVerticalGroup(
			groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
					.addGap(19)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(txtFileName, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(lblFileName, GroupLayout.PREFERRED_SIZE, 33, GroupLayout.PREFERRED_SIZE))
					.addPreferredGap(ComponentPlacement.RELATED)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(btnBrowse)
						.addComponent(txtColWidth, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(lblColWidth))
					.addGap(31)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(btnOk)
						.addComponent(btnExit))
					.addGap(135))
		);

		setupPanel();
	}

	private void setupPanel() {
		lblFileName.setLabelFor(txtFileName);
		txtColWidth.setText("4");
		txtColWidth.setColumns(10);

		txtFileName.setToolTipText("Insert the file name here");
		txtFileName.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String fileName = txtFileName.getText();
				txtFileName.setText("");
				try {
					ProcessExcelFile processExcelFile = new ProcessExcelFile(
							fileName,Integer.parseInt(txtColWidth.getText()));
					processExcelFile.process();
					JOptionPane.showMessageDialog(btnOk,
							processExcelFile.getLog(),
							"There were no problems in the process!",
							JOptionPane.INFORMATION_MESSAGE);
				} catch (Throwable t) {
					StringWriter sw = new StringWriter();
					PrintWriter pw = new PrintWriter(sw);
					pw.println(t.getMessage());
					t.printStackTrace(pw);
					JOptionPane.showMessageDialog(btnOk, sw.toString(),
							"An error occurred!", JOptionPane.ERROR_MESSAGE);
				}
			}
		});

		btnOk.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {
				if (txtFileName.getText().equals("")) {
					txtFileName.requestFocus();
				} else {
					String fileName = txtFileName.getText();
					try {
						ProcessExcelFile processExcelFile = new ProcessExcelFile(
								fileName,Integer.parseInt(txtColWidth.getText()));
						processExcelFile.process();
						// create a JTextArea
						JTextArea textArea = new JTextArea(20, 100);
						textArea.setText(processExcelFile.getLog());
						textArea.setEditable(false);

						// wrap a scrollpane around it
						JScrollPane scrollPane = new JScrollPane(textArea);
						JOptionPane.showMessageDialog(btnOk, scrollPane,
								"There were no problems in the process!",
								JOptionPane.INFORMATION_MESSAGE);
					} catch (Throwable t) {
						StringWriter sw = new StringWriter();
						PrintWriter pw = new PrintWriter(sw);
						pw.println(t.getMessage());
						t.printStackTrace(pw);
						JOptionPane.showMessageDialog(btnOk, sw.toString(),
								"An error occurred!", JOptionPane.ERROR_MESSAGE);
					}
				}
			}
		});

		btnBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				fileChooser.setCurrentDirectory(new File("."));
				int returnValue = fileChooser.showOpenDialog(null);

				if (returnValue == JFileChooser.APPROVE_OPTION) {
					String fileName = fileChooser.getSelectedFile().toString();
					txtFileName.setText(fileName);
					txtFileName.requestFocus();
				}

			}
		});

		btnExit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(1);
			}
		});

		fileChooser.changeToParentDirectory();

		setLayout(groupLayout);
	}
}
