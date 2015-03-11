package gui;

import javax.swing.JFileChooser;
import javax.swing.JPanel;
import javax.swing.JLabel;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JTextField;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.JButton;

import classes.ProcessExcelFile;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;


@SuppressWarnings("serial")
public class MyPanel extends JPanel {


	private JTextField txtFileName;
	private JLabel lblFileName;
	private GroupLayout groupLayout;
	private JButton btnOk;
	private JButton btnBrowse;
	private JButton btnExit;
	private JFileChooser fileChooser;
	
	
	public MyPanel() {

		lblFileName = new JLabel("File name:");
		
		txtFileName = new JTextField();
		
		btnOk = new JButton("Submit");
		btnBrowse = new JButton("Browse...");
		btnExit = new JButton("Exit");		

		fileChooser = new JFileChooser();

		groupLayout = new GroupLayout(this);
			
		
		setupPanel();

	}

	private void setupPanel() {
		
		
		lblFileName.setLabelFor(txtFileName);
		
		txtFileName.setToolTipText("Insert the file name here");
		txtFileName.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String fileName = txtFileName.getText();
				txtFileName.setText("");
				
				try {
					new ProcessExcelFile(fileName);
					return;
				} catch (IOException e) {
					
					// Show appropriate message on the screen
				}
			}
		});
		
		btnOk.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {
				if (txtFileName.getText().equals("")) {
					txtFileName.requestFocus();
					}
				else {
					String fileName = txtFileName.getText();
					
					try {
						new ProcessExcelFile(fileName);
						return;
					} catch (IOException e) {
						
						// Show appropriate message on the screen
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
		
		groupLayout.setHorizontalGroup(
				groupLayout.createParallelGroup(Alignment.LEADING)
					.addGroup(groupLayout.createSequentialGroup()
						.addContainerGap()
						.addComponent(btnBrowse)
						.addContainerGap(270, Short.MAX_VALUE))
					.addGroup(groupLayout.createSequentialGroup()
						.addGroup(groupLayout.createParallelGroup(Alignment.TRAILING, false)
							.addGroup(Alignment.LEADING, groupLayout.createSequentialGroup()
								.addContainerGap()
								.addComponent(btnOk)
								.addPreferredGap(ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
								.addComponent(btnExit))
							.addGroup(Alignment.LEADING, groupLayout.createSequentialGroup()
								.addGap(16)
								.addComponent(lblFileName)
								.addPreferredGap(ComponentPlacement.RELATED)
								.addComponent(txtFileName, GroupLayout.PREFERRED_SIZE, 266, GroupLayout.PREFERRED_SIZE)))
						.addContainerGap(24, Short.MAX_VALUE))
			);
			groupLayout.setVerticalGroup(
				groupLayout.createParallelGroup(Alignment.LEADING)
					.addGroup(groupLayout.createSequentialGroup()
						.addGap(19)
						.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
							.addComponent(txtFileName, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
							.addComponent(lblFileName, GroupLayout.PREFERRED_SIZE, 33, GroupLayout.PREFERRED_SIZE))
						.addPreferredGap(ComponentPlacement.RELATED)
						.addComponent(btnBrowse)
						.addGap(31)
						.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
							.addComponent(btnOk)
							.addComponent(btnExit))
						.addGap(135))
			);
		
		
		setLayout(groupLayout);
		
		

	}
}
