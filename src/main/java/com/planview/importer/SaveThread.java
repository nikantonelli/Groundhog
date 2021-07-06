package com.planview.importer;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SaveThread {
    ArrayList<UpdateRecord> upds = new ArrayList<>();
    String xlsxfn;
    XSSFWorkbook wb;
    Boolean debugPrint;

    private void dpf(String fmt, Object... parms) {
        if (debugPrint) {
            System.out.printf(fmt, parms);
        }
    }

    public SaveThread( String filename, XSSFWorkbook wrkb, Boolean dpf) {
        xlsxfn = filename;
        wb = wrkb;
        debugPrint = dpf;
    }
    
    public synchronized Integer addUpdate(UpdateRecord upd){
        upds.add(upd);
        return upds.size();
    }

    public synchronized void removeUpdate(UpdateRecord upd){
        upds.remove(upd);

        if (upds.size() == 0) {
            saveFile();
        }
        return;
    }

    private void saveFile() {
        Boolean donePrint = true;
        Integer loopCnt = 12;
        while (loopCnt > 0) {
            FileOutputStream oStr = null;
            try {
                oStr = new FileOutputStream(xlsxfn);
                try {
                    wb.write(oStr);
                    try {
                        oStr.close();
                        oStr = null;
                        loopCnt = 0;
                    } catch (IOException e) {
                        dpf("Error %s while closing file %s\n", e, xlsxfn);
                    }
                } catch (IOException e) {
                    dpf("Error %s while writing file %s\n", e, xlsxfn);
                    oStr.close(); // If this fails, just give up!
                }
            } catch (IOException e) {
                dpf("Error %s while opening/closing file %s\n", e, xlsxfn);
            }
            if (loopCnt == 0) {
                break;
            }

            Calendar now = Calendar.getInstance();
            Calendar then = Calendar.getInstance();
            then.add(Calendar.SECOND, 5);
            Long timeDiff = then.getTimeInMillis() - now.getTimeInMillis();
            if (donePrint) {
                dpf("File \"%s\" in use. Please close to let this program continue\n", xlsxfn);
                donePrint = false;
            }
            try {
                Thread.sleep(timeDiff);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            --loopCnt;
        }
    }
    
}
