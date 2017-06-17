package form;

import bin.uxcel;
import bin.xc_ZipFile;
import bin.xc_extension;
import bin.xc_extractFiles;
import bin.xc_sheetList;
import bin.xc_xmlRemove;
import bin.wBook;
import java.awt.Color;
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
    int xMouse,yMouse;
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
        jLabel5 = new javax.swing.JLabel();
        btnClose = new javax.swing.JLabel();
        btnOpen = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        txt_File = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        btnSave = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jPanel1 = new javax.swing.JPanel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Excel UnProtect by COEM");
        setBackground(new java.awt.Color(47, 47, 47));
        setForeground(java.awt.Color.lightGray);
        setName("UXCEL"); // NOI18N
        setUndecorated(true);
        setResizable(false);
        setSize(new java.awt.Dimension(100, 30));
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowActivated(java.awt.event.WindowEvent evt) {
                formWindowActivated(evt);
            }
            public void windowOpened(java.awt.event.WindowEvent evt) {
                formWindowOpened(evt);
            }
        });
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel5.setForeground(new java.awt.Color(51, 51, 51));
        jLabel5.setText("COEM | 2017");
        getContentPane().add(jLabel5, new org.netbeans.lib.awtextra.AbsoluteConstraints(433, 242, 70, -1));

        btnClose.setBackground(new java.awt.Color(34, 168, 109));
        btnClose.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        btnClose.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/closex.png"))); // NOI18N
        btnClose.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(51, 51, 51)));
        btnClose.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        btnClose.setOpaque(true);
        btnClose.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                btnCloseMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                btnCloseMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                btnCloseMouseExited(evt);
            }
        });
        getContentPane().add(btnClose, new org.netbeans.lib.awtextra.AbsoluteConstraints(393, 120, 85, 32));

        btnOpen.setBackground(new java.awt.Color(34, 168, 109));
        btnOpen.setForeground(new java.awt.Color(51, 51, 51));
        btnOpen.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        btnOpen.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/open.png"))); // NOI18N
        btnOpen.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(51, 51, 51)));
        btnOpen.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        btnOpen.setOpaque(true);
        btnOpen.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                btnOpenMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                btnOpenMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                btnOpenMouseExited(evt);
            }
        });
        getContentPane().add(btnOpen, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 120, 85, 32));

        jLabel4.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/text_1.png"))); // NOI18N
        getContentPane().add(jLabel4, new org.netbeans.lib.awtextra.AbsoluteConstraints(172, 80, 170, -1));

        jLabel2.setForeground(new java.awt.Color(51, 51, 51));
        jLabel2.setText("Copyright© Cacing");
        getContentPane().add(jLabel2, new org.netbeans.lib.awtextra.AbsoluteConstraints(14, 242, 100, -1));
        jLabel2.getAccessibleContext().setAccessibleName("Copyright© Cacing");

        txt_File.setEditable(false);
        txt_File.setBackground(new java.awt.Color(255, 255, 255));
        txt_File.setFont(new java.awt.Font("MS Reference Sans Serif", 0, 18)); // NOI18N
        txt_File.setForeground(new java.awt.Color(51, 51, 51));
        txt_File.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        txt_File.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(51, 51, 51)));
        txt_File.setCaretColor(new java.awt.Color(51, 51, 51));
        txt_File.setDisabledTextColor(new java.awt.Color(51, 51, 51));
        txt_File.setEnabled(false);
        getContentPane().add(txt_File, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 165, 450, 30));

        jLabel1.setFont(new java.awt.Font("Ubuntu", 2, 10)); // NOI18N
        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/logo.png"))); // NOI18N
        getContentPane().add(jLabel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(173, 20, 160, 70));

        btnSave.setBackground(new java.awt.Color(34, 168, 109));
        btnSave.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        btnSave.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/process.png"))); // NOI18N
        btnSave.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(51, 51, 51)));
        btnSave.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        btnSave.setOpaque(true);
        btnSave.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                btnSaveMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                btnSaveMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                btnSaveMouseExited(evt);
            }
        });
        getContentPane().add(btnSave, new org.netbeans.lib.awtextra.AbsoluteConstraints(175, 210, 160, 32));

        jLabel3.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/bg.jpg"))); // NOI18N
        jLabel3.setText("jLabel3");
        jLabel3.addMouseMotionListener(new java.awt.event.MouseMotionAdapter() {
            public void mouseDragged(java.awt.event.MouseEvent evt) {
                jLabel3MouseDragged(evt);
            }
        });
        jLabel3.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mousePressed(java.awt.event.MouseEvent evt) {
                jLabel3MousePressed(evt);
            }
        });
        getContentPane().add(jLabel3, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 10, 490, 250));

        jPanel1.setBackground(new java.awt.Color(47, 47, 47));
        jPanel1.setForeground(new java.awt.Color(47, 47, 47));
        jPanel1.addMouseMotionListener(new java.awt.event.MouseMotionAdapter() {
            public void mouseDragged(java.awt.event.MouseEvent evt) {
                jPanel1MouseDragged(evt);
            }
        });
        jPanel1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mousePressed(java.awt.event.MouseEvent evt) {
                jPanel1MousePressed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 510, Short.MAX_VALUE)
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 270, Short.MAX_VALUE)
        );

        getContentPane().add(jPanel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, -1, -1));

        pack();
    }// </editor-fold>//GEN-END:initComponents
    
    private void tampilTabel(){
        
        
    }
    
    private void formWindowOpened(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowOpened
        // TODO add your handling code here:
    }//GEN-LAST:event_formWindowOpened

    private void formWindowActivated(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowActivated
        // TODO add your handling code here:
        
    }//GEN-LAST:event_formWindowActivated

    private void crack2010(){
        if (uxcel.getSaveLocation().length() > 0) {
            try {
                for (int i = 0; i < xc_sheetList.getSheetName().size(); i++) {
                    System.out.println(xc_sheetList.getSheetName().values().toArray()[i]);
                    try {
                        xc_xmlRemove.xmlRemovePassword(xc_sheetList.getSheetName().values().toArray()[i]+".xml", "sheetProtection");
                    } catch (Exception e) {
                        System.out.println(e.getLocalizedMessage().toString());
                    }
                }
                try {
                    xc_xmlRemove.xmlRemovePassword(wBook.workBook().toString(), "workbookProtection");

                } catch (Exception e) {
                    JOptionPane.showMessageDialog(null, uxcel.getFileName()+" not protected ");
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
    } 
    
    private void btnOpenMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_btnOpenMouseEntered
        // TODO add your handling code here:
        btnOpen.setBackground(new Color(29,149,96));
    }//GEN-LAST:event_btnOpenMouseEntered

    private void btnOpenMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_btnOpenMouseExited
        // TODO add your handling code here:
        btnOpen.setBackground(new Color(34,168,109));
    }//GEN-LAST:event_btnOpenMouseExited

    private void btnOpenMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_btnOpenMouseClicked
        // TODO add your handling code here:
        btnOpen.setBackground(new Color(27,136,88));
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
                            txt_File.setText(uxcel.getFileName());
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
    }//GEN-LAST:event_btnOpenMouseClicked

    private void btnCloseMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_btnCloseMouseClicked
        btnClose.setBackground(new Color(27,136,88));
        System.exit(0);
    }//GEN-LAST:event_btnCloseMouseClicked

    private void btnSaveMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_btnSaveMouseClicked
        // TODO add your handling code here:
        btnSave.setBackground(new Color(27,136,88));
        int result;
        dirChooser.setCurrentDirectory(new java.io.File("."));
        dirChooser.setDialogTitle("Save To");
        dirChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        dirChooser.setAcceptAllFileFilterUsed(false);
        if (uxcel.getFileLocation() != null) {
           if (dirChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) { 
            uxcel.setSaveLocation(dirChooser.getSelectedFile().toString());
            if (xc_extension.getFileExtension(uxcel.getFileLocation()).equals("xlsx")) { //office 2010
                crack2010();
            } else if(xc_extension.getFileExtension(uxcel.getFileLocation()).equals("xls")){
                System.out.println("Running <= 2007");
            }
        }
            else {
          JOptionPane.showMessageDialog(null, "No file selected!");
        }
        } else {
            JOptionPane.showMessageDialog(null, "File to unprotect cannot be null, select excel file first!");
        }
    }//GEN-LAST:event_btnSaveMouseClicked

    private void btnSaveMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_btnSaveMouseEntered
        // TODO add your handling code here:
        btnSave.setBackground(new Color(29,149,96));
    }//GEN-LAST:event_btnSaveMouseEntered

    private void btnSaveMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_btnSaveMouseExited
        // TODO add your handling code here:
        btnSave.setBackground(new Color(34,168,109));
    }//GEN-LAST:event_btnSaveMouseExited

    private void btnCloseMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_btnCloseMouseExited
        // TODO add your handling code here:
        btnClose.setBackground(new Color(34,168,109));
    }//GEN-LAST:event_btnCloseMouseExited

    private void btnCloseMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_btnCloseMouseEntered
        // TODO add your handling code here:
        btnClose.setBackground(new Color(29,149,96));
    }//GEN-LAST:event_btnCloseMouseEntered

    private void jLabel3MousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel3MousePressed
        // TODO add your handling code here:
        xMouse = evt.getX();
        yMouse = evt.getY();
    }//GEN-LAST:event_jLabel3MousePressed

    private void jLabel3MouseDragged(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel3MouseDragged
        // TODO add your handling code here:
        int x = evt.getXOnScreen();
        int y = evt.getYOnScreen();
        
        this.setLocation(x-xMouse, y-yMouse);
    }//GEN-LAST:event_jLabel3MouseDragged

    private void jPanel1MousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jPanel1MousePressed
        // TODO add your handling code here:
        xMouse = evt.getX();
        yMouse = evt.getY();
    }//GEN-LAST:event_jPanel1MousePressed

    private void jPanel1MouseDragged(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jPanel1MouseDragged
        // TODO add your handling code here:
        int x = evt.getXOnScreen();
        int y = evt.getYOnScreen();
        
        this.setLocation(x-xMouse, y-yMouse);
    }//GEN-LAST:event_jPanel1MouseDragged

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
                home gui = new home();
//                gui.setUndecorated(true);
//                gui.getRootPane().setWindowDecorationStyle(JRootPane.NONE);
                //gui.setUndecorated(true);
                gui.setSize(510,270);
                gui.setVisible(true);
                gui.setLocationRelativeTo(null);
                
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JLabel btnClose;
    private javax.swing.JLabel btnOpen;
    private javax.swing.JLabel btnSave;
    private javax.swing.JFileChooser dirChooser;
    private javax.swing.JFileChooser fileChooser;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JTextField txt_File;
    // End of variables declaration//GEN-END:variables
}
