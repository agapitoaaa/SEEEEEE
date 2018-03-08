/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapplication2;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 *
 * @author Ace Agapito
 */
public class JavaApplication2 {

    public static String url = "jdbc:sqlite:C:\\Users\\Public\\TeamPapsie.db";
    private static String EXCEL_FILE_LOCATION = "C:\\Users\\Ace Agapito\\Desktop\\ace\\ORGANIZATION.xls";
    private static String EXCEL_FILE_LOCATION2 = "C:\\Users\\Ace Agapito\\Desktop\\ace\\OFFICERS.xls";
    public static Connection cons = null;
    public static PreparedStatement ps = null;
    public static ResultSet rs = null;

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        cons = con();

        WritableWorkbook myFirstWbook = null;
        WritableWorkbook myFirstWbook2 = null;
        try {
            myFirstWbook = Workbook.createWorkbook(new File(EXCEL_FILE_LOCATION));
            myFirstWbook2 = Workbook.createWorkbook(new File(EXCEL_FILE_LOCATION2));
        } catch (IOException ex) {
            Logger.getLogger(JavaApplication2.class.getName()).log(Level.SEVERE, null, ex);
        }

        try {
            ps = cons.prepareStatement("SELECT * FROM ORGANIZATION");
            rs = ps.executeQuery();
            try {
        WritableSheet excelSheet = myFirstWbook.createSheet("Sheet 1", 0);

                Label label = new Label(0, 0, "ORG_ID");
                excelSheet.addCell(label);
                label = new Label(1, 0, "NAME ");
                excelSheet.addCell(label);
                label = new Label(2, 0, "YEAR_ESTABLISHED");
                excelSheet.addCell(label);
                label = new Label(3, 0, "IS_UWIDE ");
                excelSheet.addCell(label);
                label = new Label(4, 0, "college_code ");
                excelSheet.addCell(label);
                label = new Label(5, 0, "CODE_COLLEGE ");
                excelSheet.addCell(label);
                int ctr = 1;
                int xz = 0;
                while (rs.next()) {
                    label = new Label(0, ctr, rs.getString(1));
                    excelSheet.addCell(label);
                    label = new Label(1, ctr, rs.getString(2));
                    excelSheet.addCell(label);
                    label = new Label(2, ctr, rs.getString(3));
                    excelSheet.addCell(label);
                    label = new Label(3, ctr, rs.getString(4));
                    excelSheet.addCell(label);
                    label = new Label(4, ctr, rs.getString(5));
                    excelSheet.addCell(label);
                    label = new Label(5, ctr, rs.getString(6));
                    excelSheet.addCell(label);
                    ctr = ctr + 1;
                }
                    myFirstWbook.write();
                    myFirstWbook.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            rs.close();
            ps.close();
        } catch (SQLException ex) {
            Logger.getLogger(JavaApplication2.class.getName()).log(Level.SEVERE, null, ex);
        }

        try {
            ps = cons.prepareStatement("SELECT * FROM OFFICERS");
            rs = ps.executeQuery();
            try {

        WritableSheet excelSheet2 = myFirstWbook2.createSheet("Sheet 1", 0);
                Label label = new Label(0, 0, "ID");
                excelSheet2.addCell(label);
                label = new Label(1, 0, "studno");
                excelSheet2.addCell(label);
                label = new Label(2, 0, "firstname");
                excelSheet2.addCell(label);
                label = new Label(3, 0, "middlename");
                excelSheet2.addCell(label);
                label = new Label(4, 0, "dateofBirth");
                excelSheet2.addCell(label);
                label = new Label(5, 0, "emailAddress");
                excelSheet2.addCell(label);
                label = new Label(6, 0, "degree");
                excelSheet2.addCell(label);
                label = new Label(7, 0, "college");
                excelSheet2.addCell(label);
                label = new Label(8, 0, "yearsec");
                excelSheet2.addCell(label);
                label = new Label(9, 0, "organization");
                excelSheet2.addCell(label);
                label = new Label(10, 0, "orgPosition");
                excelSheet2.addCell(label);
                label = new Label(11, 0, "acadYear");
                excelSheet2.addCell(label);
                label = new Label(12, 0, "is_uniwide");
                excelSheet2.addCell(label);
                int ctr = 1;
                int xz = 0;
                while (rs.next()) {
                    label = new Label(0, ctr, rs.getString(1));
                    excelSheet2.addCell(label);
                    label = new Label(1, ctr, rs.getString(2));
                    excelSheet2.addCell(label);
                    label = new Label(2, ctr, rs.getString(3));
                    excelSheet2.addCell(label);
                    label = new Label(3, ctr, rs.getString(4));
                    excelSheet2.addCell(label);
                    label = new Label(4, ctr, rs.getString(5));
                    excelSheet2.addCell(label);
                    label = new Label(5, ctr, rs.getString(6));
                    excelSheet2.addCell(label);
                    label = new Label(6, ctr, rs.getString(7));
                    excelSheet2.addCell(label);
                    label = new Label(7, ctr, rs.getString(8));
                    excelSheet2.addCell(label);
                    label = new Label(8, ctr, rs.getString(9));
                    excelSheet2.addCell(label);
                    label = new Label(9, ctr, rs.getString(10));
                    excelSheet2.addCell(label);
                    label = new Label(10, ctr, rs.getString(11));
                    excelSheet2.addCell(label);
                    label = new Label(11, ctr, rs.getString(12));
                    excelSheet2.addCell(label);
                    label = new Label(12, ctr, rs.getString(13));
                    excelSheet2.addCell(label);
                    ctr = ctr + 1;
                }
                    myFirstWbook2.write();
                    myFirstWbook2.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            rs.close();
            ps.close();

        } catch (Exception ex) {
            Logger.getLogger(JavaApplication2.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    public static Connection con() {
        try {
            cons = DriverManager.getConnection(url);
            return cons;
        } catch (SQLException ex) {
            Logger.getLogger(JavaApplication2.class.getName()).log(Level.SEVERE, null, ex);
            return null;
        }
    }

}
