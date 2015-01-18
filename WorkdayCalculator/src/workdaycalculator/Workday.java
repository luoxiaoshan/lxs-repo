/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package workdaycalculator;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 *
 * @author Mike
 */
public class Workday {

    private static DateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    private List<Date> workDay = new ArrayList<Date>();
    private List<Date> festival = new ArrayList<Date>();
    private File excel_file;                            //对应EXCEL文件
    private Workbook work_book;                         //JXL接口对应的EXCEL文件
    private Sheet main_sheet;                           //EXCEL文件中MAIN sheet
    private Cell cell;                      //单元格

    public Workday(String file_name) {


        this.excel_file = new File(file_name);
        try {
            this.work_book = Workbook.getWorkbook(excel_file);
        } catch (BiffException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        main_sheet = this.work_book.getSheet(0);
        int row_num = main_sheet.getRows();
        for (int i = 0; i < row_num; i++) {
            if (main_sheet.getCell(1, i).getContents().equals("w")) {
                try {
                    workDay.add(sdf.parse(main_sheet.getCell(0, i).getContents()));
                } catch (ParseException ex) {
                    Logger.getLogger(Workday.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
            if (main_sheet.getCell(1, i).getContents().equals("f")) {
                try {
                    festival.add(sdf.parse(main_sheet.getCell(0, i).getContents()));
                } catch (ParseException ex) {
                    Logger.getLogger(Workday.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }
    }

    public double getWorkdays(Date startdate, Date enddate) {
        try {
            Date tempdate = new Date();
            Date f_sd = sdf.parse(sdf.format(startdate));//开始日期的零点
            Date f_ed = sdf.parse(sdf.format(enddate));//结束日期的零点
            long days = (f_ed.getTime() - f_sd.getTime()) / 86400000 + 1;//横跨多少天
            long sttime = 0;
            long edtime = 0;
            long worktimes = 0;
            sttime = startdate.getTime();

            for (int i = 0; i < days; i++) {
                tempdate.setTime(f_sd.getTime() + (long) 24 * 3600000 * i); //相对开始日期后i天零点的时间

                if (this.isWorkDay(tempdate)) {
                    edtime = tempdate.getTime() + (long) 24 * 3600000;//相对开始日期后i+1天零点的时间
                    if (edtime <= enddate.getTime()) {
                        worktimes += edtime - sttime;
                        sttime = edtime;
                    } else {
                        worktimes += enddate.getTime() - sttime;
                    }
                } else {
                    sttime = tempdate.getTime() + (long) 24 * 3600000;
                }
            }
            double res = round((double) worktimes / 3600000 / 24, 4, BigDecimal.ROUND_HALF_UP);//(double)worktimes/3600000/24);
            return res;

            //return 0;
        } catch (ParseException ex) {
            Logger.getLogger(Workday.class.getName()).log(Level.SEVERE, null, ex);
            return -1;
        }
    }

    public double getWorkdays(String startdate, String enddate) {
        try {
            return getWorkdays(sdf.parse(startdate), sdf.parse(enddate));
        } catch (ParseException ex) {
            Logger.getLogger(Workday.class.getName()).log(Level.SEVERE, null, ex);
            return -1;
        }
    }

    public double getWorkdays(long startdate, long enddate) {
        Date sd = new Date();
        Date ed = new Date();
        sd.setTime(startdate);
        ed.setTime(enddate);
        return getWorkdays(sd, ed);

    }

    /*******************************************************************************/
    public List getFestival() {
        return this.festival;
    }

    public List getSpecialWorkDay() {
        return this.workDay;
    }

    /**
     * 判断一个日期是否日节假日
     * 法定节假日只判断月份和天，不判断年
     * @param date
     * @return
     */
    public boolean isFestival(Date date) {
        boolean festival = false;
        Calendar fcal = Calendar.getInstance();
        Calendar dcal = Calendar.getInstance();
        dcal.setTime(date);
        List<Date> list = this.getFestival();
        for (Date dt : list) {
            fcal.setTime(dt);

            //法定节假日判断
            if (fcal.get(Calendar.MONTH) == dcal.get(Calendar.MONTH)
                    && fcal.get(Calendar.DATE) == dcal.get(Calendar.DATE)) {
                festival = true;
            }
        }
        return festival;
    }

    public boolean isWeekend(Date date) {
        boolean weekend = false;
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        if (cal.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY
                || cal.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
            weekend = true;
        }
        return weekend;
    }

    public boolean isWorkDay(Date date) {
        boolean workday = true;
        if (this.isFestival(date) || this.isWeekend(date)) {
            workday = false;
        }

        /*特殊工作日判断*/
        Calendar cal1 = Calendar.getInstance();
        cal1.setTime(date);
        Calendar cal2 = Calendar.getInstance();
        for (Date dt : this.workDay) {
            cal2.setTime(dt);
            if (cal1.get(Calendar.YEAR) == cal2.get(Calendar.YEAR)
                    && cal1.get(Calendar.MONTH) == cal2.get(Calendar.MONTH)
                    && cal1.get(Calendar.DATE) == cal2.get(Calendar.DATE)) {
                //年月日相等为特殊工作日
                workday = true;
            }

        }
        return workday;
    }

    private double round(double value, int scale, int roundingMode) {

        BigDecimal bd = new BigDecimal(value);

        bd = bd.setScale(scale, roundingMode);

        double d = bd.doubleValue();

        bd = null;

        return d;

    }
}
