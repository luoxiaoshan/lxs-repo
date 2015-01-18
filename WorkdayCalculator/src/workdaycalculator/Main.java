/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package workdaycalculator;

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author Mike
 */
public class Main {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
    	String targetpos = "A";
        String startpos = "E";
        String endpos = "J";
        String respos = "K";
        int startrow = 2;
        Tools t = new Tools();
        Excel e = new Excel();
        Workday w = new Workday("input.xls");
        DateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm");
        DecimalFormat df = new DecimalFormat("0.000");
        e.openFile("gongdan.xls");
        e.getSheet("data");
        while (!e.getString(t.convert(targetpos), startrow).isEmpty()) {
            String startdate = e.getString(t.convert(startpos), startrow);
            String enddate = e.getString(t.convert(endpos), startrow);
            if (!enddate.isEmpty()&&!startdate.isEmpty()) {
                try {
                    e.setString(t.convert(respos), startrow, String.valueOf(df.format(w.getWorkdays(sdf.parse(startdate), sdf.parse(enddate)))));
                    print("test brach2 merger");
                } catch (ParseException ex) {
                    Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
            startrow++;
            if (startrow>e.getRows()) break;
        }
        e.close();

        // TODO code application logic here
    }

	private static void print(String string) {
		// TODO Auto-generated method stub
		
	}
}
