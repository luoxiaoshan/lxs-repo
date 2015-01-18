/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package workdaycalculator;

import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 *
 * @author Mike
 */
public class Excel {

    private final static String e_filename = "input.xls";
    private File e_excelfile = null;                            //瀵瑰簲EXCEL鏂囦欢
    private Workbook e_workbook = null;
    private WritableWorkbook e_writableworkbook = null;
    private WritableSheet e_writablesheet = null;
    private Sheet e_sheet = null;
    private Cell e_cell = null;
//    private boolean e_writeflag = true;
    private boolean e_openflag = false;

    public Excel() {
    }

    public int openFile(String p_filename) {
        if (p_filename != null && !p_filename.isEmpty()) {
            this.e_excelfile = new File(p_filename);
        } else {
            this.e_excelfile = new File(this.e_filename);
        }
        try {
            this.e_workbook = Workbook.getWorkbook(e_excelfile);
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (BiffException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
        try {
            this.e_writableworkbook = Workbook.createWorkbook(e_excelfile, e_workbook);
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }





        //String readfile = "銆怐EMO銆慩XX鐮斿彂閮ㄤ骇鍝佺嚎娴嬭瘯璐ㄩ噺鏁版嵁鎶ュ憡锛�12骞�堬級.xls";
        //Open file at "C:/wangming.xls"

        return 0;
    }

    public int getSheet (String p_sheetname) {
        if (this.e_writableworkbook == null) {
            System.out.println("鏈墦寮�鏂囦欢");
            return -1;
        }
        if (p_sheetname != null && !p_sheetname.isEmpty()) {
            this.e_writablesheet = this.e_writableworkbook.getSheet(p_sheetname);
            this.e_sheet = this.e_workbook.getSheet(p_sheetname);
        } else {
            this.e_writablesheet = this.e_writableworkbook.getSheet(0);
            this.e_sheet = this.e_workbook.getSheet(0);
        }

        e_openflag = true;
        return 0;
    }

//    public void write() {
//        try {
//            if (!this.e_openflag) {
//                return;
//            }
//            this.e_writableworkbook.write();
//            this.e_writeflag = true;
//        } catch (IOException ex) {
//            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
//        }
//    }

    public void close() {
        try {
            this.e_writableworkbook.write();
            this.e_writableworkbook.close();
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WriteException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
        this.e_workbook.close();
        this.e_openflag = false;
    }

    public String getString(int x, int y) {
        if (!this.e_openflag) {
            System.out.println("鏈墦寮�宸ヤ綔琛�");
            return null;
        }
        this.e_cell = this.e_sheet.getCell(x-1, y-1);
        String res = this.e_cell.getContents();
        return res;
    }

    public void setString(int x, int y, String str) {
        if (!this.e_openflag) {
            System.out.println("鏈墦寮�宸ヤ綔琛�");
            return;
        }
        if (x < 1 || y < 1) {
            System.out.println("鍧愭爣闈炴硶");
            return;
        }
        Label label = new Label(x-1, y-1, str);
        try {
            this.e_writablesheet.addCell(label);
//            this.e_writeflag = false;
        } catch (WriteException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public int getRows() {
    	if (!this.e_openflag) {
            return 0;
        } else {
        	return this.e_sheet.getRows();
        }
    }

}
