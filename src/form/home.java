package form;

import bin.uxcel;
import bin.xc_ZipFile;
import bin.xc_extension;
import bin.xc_extractFiles;
import bin.xc_sheetList;
import bin.xc_xmlRemove;
import java.io.File;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

/**
 *
 * @author cacing
 */
public class home extends javax.swing.JFrame {
    ClassLoader loader = uxcel.class.getClassLoader();
    /**
     * Creates new form home
     */
    uxcel excel = new uxcel();
    
    
    
    public home() {
        initComponents();
        
        final String dir = System.getProperty("user.dir");
        //JOptionPane.showMessageDialog(rootPane, System.getProperty("java.io.tmpdir"));
        File var = new File(System.getProperty("java.io.tmpdir") + "var");
        File tmp = new File(System.getProperty("java.io.tmpdir") + "tmp");
        if ((var.exists()) && (tmp.exists())) {
            
        } else {
            var.mkdir();
            tmp.mkdir();
        }
        this.setLocationRelativeTo(this);
        tampilTabel();
    }
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        fileChooser = new javax.swing.JFileChooser();
        dirChooser = new javax.swing.JFileChooser();
        btn_Browse = new javax.swing.JButton();
        txt_File = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenuItem1 = new javax.swing.JMenuItem();
        jMenuItem4 = new javax.swing.JMenuItem();
        jMenu2 = new javax.swing.JMenu();
        jMenu3 = new javax.swing.JMenu();
        jMenuItem2 = new javax.swing.JMenuItem();
        jMenuItem3 = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Excel UnProtect by COEM");
        setResizable(false);
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowOpened(java.awt.event.WindowEvent evt) {
                formWindowOpened(evt);
            }
            public void windowActivated(java.awt.event.WindowEvent evt) {
                formWindowActivated(evt);
            }
        });

        btn_Browse.setText("Save to");
        btn_Browse.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btn_BrowseActionPerformed(evt);
            }
        });

        jLabel1.setFont(new java.awt.Font("Ubuntu", 2, 10)); // NOI18N
        jLabel1.setText("...");

        jMenu1.setText("File");

        jMenuItem1.setAccelerator(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_O, java.awt.event.InputEvent.CTRL_MASK));
        jMenuItem1.setText("Open");
        jMenuItem1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem1ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem1);

        jMenuItem4.setText("Check");
        jMenuItem4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem4ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem4);

        jMenuBar1.add(jMenu1);

        jMenu2.setText("Edit");
        jMenuBar1.add(jMenu2);

        jMenu3.setText("Help");

        jMenuItem2.setText("Help Contents");
        jMenu3.add(jMenuItem2);

        jMenuItem3.setText("About");
        jMenu3.add(jMenuItem3);

        jMenuBar1.add(jMenu3);

        setJMenuBar(jMenuBar1);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(txt_File, javax.swing.GroupLayout.PREFERRED_SIZE, 186, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(btn_Browse, javax.swing.GroupLayout.PREFERRED_SIZE, 72, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(26, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txt_File, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btn_Browse))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jLabel1)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents
    
    private void tampilTabel(){
        
        
    }
    
    private void jMenuItem1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem1ActionPerformed
        // TODO add your handling code here:
        
        final File f = new File(home.class.getProtectionDomain().getCodeSource().getLocation().getPath());
        
        if (fileChooser.showOpenDialog(rootPane)== JFileChooser.APPROVE_OPTION) {
            uxcel.setFileLocation(fileChooser.getSelectedFile().getPath());
            uxcel.setFileNameExt(fileChooser.getSelectedFile().getName());
            
            jLabel1.setText(uxcel.getFileName());
            System.out.println(uxcel.getFileLocation());
            try {
                uxcel.copyExcel(new File(uxcel.getFileLocation()), new File(System.getProperty("java.io.tmpdir") + "tmp/"+uxcel.getFileNameExt()));
                try {
                    xc_extension.renameFileExtension(System.getProperty("java.io.tmpdir") + "tmp/"+uxcel.getFileNameExt(), "zip");
                    try {
                        xc_extractFiles.unzip(uxcel.getFileLocation());
                        System.out.println();
                        try {
                            System.out.println(xc_sheetList.getSheetName());
                        } catch (Exception e) {
                            JOptionPane.showMessageDialog(rootPane, e, "4", 0);
                        }
                    } catch (Exception e) {
                        JOptionPane.showMessageDialog(rootPane, e, "3", 0);
                    }
                } catch (Exception e) {
                    JOptionPane.showMessageDialog(rootPane, e, "2", 0);
                }
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(rootPane, ex, "1", 0);
            }
        }
    }//GEN-LAST:event_jMenuItem1ActionPerformed

    private void formWindowOpened(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowOpened
        // TODO add your handling code here:
    }//GEN-LAST:event_formWindowOpened

    private void formWindowActivated(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowActivated
        // TODO add your handling code here:
        
    }//GEN-LAST:event_formWindowActivated

    private void btn_BrowseActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btn_BrowseActionPerformed
        int result;
        dirChooser.setCurrentDirectory(new java.io.File("."));
        dirChooser.setDialogTitle("Save To");
        dirChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        //
        // disable the "All files" option.
        //
        dirChooser.setAcceptAllFileFilterUsed(false);
        //    
        if (dirChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) { 
            uxcel.setSaveLocation(dirChooser.getSelectedFile().toString());
            if (xc_extension.getFileExtension(uxcel.getFileLocation()).equals("xlsx")) { //office 2010
                if (uxcel.getSaveLocation().length() > 0) {
                    try {
                        for (int i = 0; i < xc_sheetList.getSheetName().size(); i++) {
                            System.out.println(xc_sheetList.getSheetName().values().toArray()[i]);
                            try {
                                xc_xmlRemove.xmlRemovePassword(xc_sheetList.getSheetName().values().toArray()[i]+".xml", "sheetProtection");
                            } catch (Exception e) {
                                JOptionPane.showMessageDialog(null, (String) xc_sheetList.getSheetName().values().toArray()[i] + " canot be unprotect, because dont have password", (String) xc_sheetList.getSheetName().values().toArray()[i], i);
                            }
                        }
                        try {
                            File dir = new File(System.getProperty("java.io.tmpdir") +"/var/"+uxcel.getFileName());
                            String zipDirName = uxcel.getSaveLocation()+ "\\" +uxcel.getFileName() + "_uxcel.zip";
                            xc_ZipFile.zipDirectory(dir, zipDirName);
                            try {
                                xc_extension.renameFileExtension(zipDirName, xc_extension.getFileExtension(uxcel.getFileNameExt()));
                                    try {
                                        System.out.println(xc_sheetList.getSheetName());
                                        JOptionPane.showMessageDialog(null, "Excel " + uxcel.getFileName() + " UnProtected sheet success!");
                                        try {
                                            uxcel.deleteDir(new File(System.getProperty("java.io.tmpdir") + "var/" + uxcel.getFileName()));
                                            uxcel.deleteDir(new File(System.getProperty("java.io.tmpdir") + "tmp/" + uxcel.getFileName() + ".zip"));
                                            try {
                                                uxcel.setFileNameExt("");
                                                uxcel.setFileName("");

                                                uxcel.setFileLocation("");
                                                uxcel.setSaveLocation("");
                                                xc_ZipFile.filesListInDir.clear();
                                            } catch (Exception e) {
                                            }
                                        } catch (Exception e) {
                                        }
                                    } catch (Exception e) {
                                    }
                                } catch (Exception e) {
                          System.out.println(e);
                          }
                      } catch (Exception e) {
                          System.out.println(e);
                      }
                    } catch (Exception e) {
                      System.out.println(e);
                    }
                }
            } else if(xc_extension.getFileExtension(uxcel.getFileLocation()).equals("xls")){
                System.out.println("Running <= 2007");
            }
        }
        else {
          System.out.println("No Selection ");
        }
        

    }//GEN-LAST:event_btn_BrowseActionPerformed

    private void jMenuItem4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem4ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jMenuItem4ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if (("Windows".equals(info.getName())) || ("GTK+".equals(info.getName()))) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(home.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(home.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(home.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(home.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new home().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btn_Browse;
    private javax.swing.JFileChooser dirChooser;
    private javax.swing.JFileChooser fileChooser;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenu jMenu3;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JMenuItem jMenuItem2;
    private javax.swing.JMenuItem jMenuItem3;
    private javax.swing.JMenuItem jMenuItem4;
    private javax.swing.JTextField txt_File;
    // End of variables declaration//GEN-END:variables
}
