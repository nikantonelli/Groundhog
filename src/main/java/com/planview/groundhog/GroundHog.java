package com.planview.groundhog;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

import com.planview.groundhog.Leankit.CardLongRead;
import com.planview.groundhog.Leankit.CardType;
import com.planview.groundhog.Leankit.Id;
import com.planview.groundhog.Leankit.LeanKitAccess;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

public class GroundHog {

    Integer refreshPeriod = 14; // Number of days to run for by default
    Configuration config = new Configuration();
    String xlsxfn = "";
    FileInputStream xlsxfis = null;
    XSSFWorkbook wb = null;
    static Boolean useCron = false;
    static String statusFile = "";
    static String moveLane = null;
    static Boolean deleteItems = false;

    /**
     * One line sheet that contains the credentials to access the Leankit Server
     * Must contain columns "url", "username", "password" and "apiKey", but not
     * necessarily have data in all of them - see getConfig()
     */
    XSSFSheet configSht = null;
    /**
     * Lines in this sheet refer to changes to make per day
     */
    XSSFSheet changesSht = null;

    /**
     * Sheets that contain the artifacts that are going to be changed. They can be
     * organised into logical pages of stuff, or all in the same sheet. I did it
     * this way, so that you can keep a sheet for each team and separate ones for
     * portfolio changes.
     */
    ArrayList<XSSFSheet> teamShts = null;

    public static void main(String[] args) {

        GroundHog hog = new GroundHog();

        hog.getCommandLine(args);
        hog.parseXlsx();
        hog.getConfig();
        Integer day = 0;

        // Check to see if there is command line option to loop or use cron
        if (!useCron) {
            while (true) {
                try {
                    Calendar now = Calendar.getInstance();
                    Calendar then = Calendar.getInstance();
                    then.add(Calendar.HOUR, 1);
                    // then.set(Calendar.HOUR_OF_DAY, 3); // Set to three in the morning
                    // then.set(Calendar.MINUTE, 0);
                    // then.set(Calendar.SECOND, 0);
                    Long timeDiff = then.getTimeInMillis() - now.getTimeInMillis();
                    // Do todays activity and then sleep
                    hog.activity(day++);
                    System.out.println("Sleeping until: " + then.getTime());
                    Thread.sleep(timeDiff);
                } catch (InterruptedException e) {
                    System.out.println("Early wake from daily timer");
                    System.exit(1);
                }
            }
        } else {
            // Create a local file to hold the day
            File statusFs = new File(statusFile);
            JSONObject settings = new JSONObject();

            try {

                if (!statusFs.exists()) {
                    statusFs.createNewFile();
                    setUpStatusFile(statusFs);
                } // If already exists, then that is OK.
                BufferedReader fis = new BufferedReader(new FileReader(statusFs));
                // On the first go at this, we might have a file with nothing in it.
                // If so, allow to continue as it will get corrected further on.
                // Json is all on one line in the file, so it makes it easier to parse
                String jsonLine = fis.readLine();
                settings = new JSONObject((jsonLine == null) ? "{}" : jsonLine);
                // Now check we have the correct 'day' field as the file may have already
                // existed
                Integer fDay = null;
                if (settings.has("day")) {
                    fDay = settings.getInt("day");
                    if ((fDay == null) || (fDay < 0) || (fDay >= hog.getRefresh())) {
                        setUpStatusFile(statusFs); // Summat wrong, so try to re-initialise
                        fis.close();
                        fis = new BufferedReader(new FileReader(statusFs));
                        settings = new JSONObject(fis.readLine());
                        fDay = settings.getInt("day");
                    }

                } else {
                    // Once again,file must be corrupted, so try again
                    setUpStatusFile(statusFs); // Summat wrong, so try to re-initialise
                    fis.close();
                    fis = new BufferedReader(new FileReader(statusFs));
                    settings = new JSONObject(fis.readLine());
                    fDay = settings.getInt("day");
                }
                fis.close();
                hog.activity(fDay++);
                // Reset file after the time period has expired
                if (fDay > hog.getRefresh()) {
                    // We need to reset the day to zero
                    setUpStatusFile(statusFs);
                } else {
                    settings.put("day", fDay);
                    FileWriter fos = new FileWriter(statusFs);
                    settings.write(fos);
                    fos.flush();
                    fos.close();
                }

            } catch (IOException e) {
                System.out.println(e.getMessage()); // Any other error, we barf.
                System.exit(1);
            }

        }
    }

    public void updateStatusDay(int day) {

    }

    public static void setUpStatusFile(File statusFs) {
        // Newly created, so add the default settings in

        JSONObject settings = new JSONObject();
        try {
            FileWriter fos = new FileWriter(statusFs);
            settings.put("day", 0);
            settings.write(fos);
            fos.flush();
            fos.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }
    }

    public void getCommandLine(String[] args) {
        Options opts = new Options();
        Option xlsxname = new Option("f", "filename", true, "source XLSX spreadsheet");
        xlsxname.setRequired(true);
        opts.addOption(xlsxname);
        Option cronIt = new Option("c", "cron", false, "Program is called by cron job");
        cronIt.setRequired(false);
        opts.addOption(cronIt);
        Option statusFn = new Option("s", "status", true, "Status file used when called by cron job");
        statusFn.setRequired(false);
        opts.addOption(statusFn);
        Option deleteCycle = new Option("d", "delete", false, "Delete all artifacts on end of cycle");
        deleteCycle.setRequired(false);
        opts.addOption(deleteCycle);

        Option moveCycle = new Option("m", "move", true, "Move all artifacts on end of cycle to this lane");
        moveCycle.setRequired(false);
        opts.addOption(moveCycle);

        CommandLineParser p = new DefaultParser();
        HelpFormatter hf = new HelpFormatter();
        CommandLine cl = null;
        try {
            cl = p.parse(opts, args);
        } catch (ParseException e) {
            System.out.println(e.getMessage());
            hf.printHelp(" ", opts);
            System.exit(1);
        }

        xlsxfn = cl.getOptionValue("filename");
        useCron = cl.hasOption("cron");
        if (cl.hasOption("status")) {
            statusFile = cl.getOptionValue("status");
        }

        // Move items to a lane and then delete them if needed - or just delete.
        if (cl.hasOption("move")) {
            moveLane = cl.getOptionValue("move");
        }
        deleteItems = cl.hasOption("delete");

    }

    /**
     * Check if the XLSX file provided has the correct sheets and we can parse the
     * details we need
     */
    public void parseXlsx() {
        // Check we can open the file
        try {
            xlsxfis = new FileInputStream(new File(xlsxfn));
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage());

            System.exit(1);

        }
        try {
            wb = new XSSFWorkbook(xlsxfis);
        } catch (IOException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }

        // These two should come first in the file and must be present

        configSht = wb.getSheet("Config");
        // This is a common sheet listing all the progression changes for all boards
        changesSht = wb.getSheet("Changes");

        Integer shtCount = wb.getNumberOfSheets();
        if ((shtCount < 3) || (configSht == null) || (changesSht == null)) {
            System.out.println(
                    "Did not detect correct sheets in the spreadsheet: \"Config\",\"Changes\" and one, or more, board(s)");
            System.exit(1);
        }

        // Grab the rest of the sheets that describe boards
        teamShts = new ArrayList<XSSFSheet>();
        Iterator<Sheet> si = wb.iterator();

        // Then all the level info comes next
        while (si.hasNext()) {
            XSSFSheet n = (XSSFSheet) si.next();
            if (!((n.getSheetName().equals("Config")) || (n.getSheetName().equals("Changes")))) {
                teamShts.add(n);
            }

        }
    }

    /**
     * To get access to the Leankit server, we need either a url/apiKey pair or a
     * username/password pair We will use the apiKey in preference if it is present.
     * Fallback is Base64 encoding of username/password If neither, then barf!
     * 
     * @return None
     */
    private void getConfig() {
        // Make the contents of the file lower case before comparing with these.
        Field[] p = (new Configuration()).getClass().getDeclaredFields();

        ArrayList<String> cols = new ArrayList<String>();
        for (int i = 0; i < p.length; i++) {
            p[i].setAccessible(true); // Set this up for later
            cols.add(p[i].getName().toLowerCase());
        }
        HashMap<String, Object> hm = new HashMap<>();

        // Assume that the titles are the first row
        Iterator<Row> ri = configSht.iterator();
        Row hdr = ri.next();
        Iterator<Cell> cptr = hdr.cellIterator();

        while (cptr.hasNext()) {
            Cell cell = cptr.next();
            Integer idx = cell.getColumnIndex();
            String cellName = cell.getStringCellValue().trim().toLowerCase();
            if (cols.contains(cellName)) {
                hm.put(cellName, idx); // Store the column index of the field
            }
        }

        if (hm.size() != cols.size()) {
            System.out.println("Did not detect correct columns on Config sheet: " + cols.toString());
            System.exit(1);
        }

        // Now we know which columns contain the data, scan down the sheet looking for a
        // row with data in the 'url' cell
        while (ri.hasNext()) {
            Row drRow = ri.next();
            String cv = drRow.getCell((int) (hm.get(cols.get(0)))).getStringCellValue();
            if (cv != null) {

                for (int i = 0; i < cols.size(); i++) {

                    try {
                        String idx = cols.get(i);
                        Object obj = hm.get(idx);
                        String val = obj.toString();
                        Cell cell = drRow.getCell(Integer.parseInt(val));

                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case STRING:
                                    p[i].set(config,
                                            (cell != null ? drRow.getCell(Integer.parseInt(val)).getStringCellValue()
                                                    : ""));
                                    break;
                                case NUMERIC:
                                    p[i].set(config,
                                            (cell != null ? drRow.getCell(Integer.parseInt(val)).getNumericCellValue()
                                                    : ""));
                                    break;
                                default:
                                    break;
                            }
                        } else {
                            p[i].set(config, (p[i].getType().equals(String.class)) ? "" : 0.0);
                        }

                    } catch (IllegalArgumentException | IllegalAccessException e) {
                        System.out.println(e.getMessage());
                        System.exit(1);
                    }

                }

                break; // Exit out of while loop as we are only interested in the first one we find
            }
        }

        // Creds are now found and set. If not, you're buggered.
        // We can opt to use username/password or apikey.

        if ((config.cyclelength != null) && (config.cyclelength != 0.0)) {
            refreshPeriod = config.cyclelength.intValue();
        }
        return;
    }

    private Integer findColumnFromName(Row firstRow, String name) {
        Iterator<Cell> frtc = firstRow.iterator();
        // First, find the column that the "Day Delta" info is in
        int dayCol = -1;
        int td = 0;
        while (frtc.hasNext()) {
            Cell tc = frtc.next();
            if (!tc.getStringCellValue().equals(name)) {
                td++;
            } else {
                dayCol = td;
                break;
            }
        }
        return dayCol;
    }

    Integer dayCol = null;
    Integer itemShtCol = null;
    Integer rowCol = null;
    Integer actionCol = null;
    Integer fieldCol = null;
    Integer valueCol = null;

    Integer findColumnFromSheet(XSSFSheet sht, String name) {
        Iterator<Row> row = sht.iterator();
        Row firstRow = row.next(); // Get the header row
        Integer col = findColumnFromName(firstRow, name);
        if (col < 0) {
            System.out.printf("Could not locate column %s in sheet %s", name, sht.getSheetName());
        }
        return col;
    }

    private void activity(Integer day) {
        // Find all the change records for today
        Iterator<Row> row = changesSht.iterator();
        ArrayList<Row> todaysChanges = new ArrayList<Row>();

        dayCol = findColumnFromSheet(changesSht, "Day Delta");
        itemShtCol = findColumnFromSheet(changesSht, "Item Sheet");
        rowCol = findColumnFromSheet(changesSht, "Item Row");
        actionCol = findColumnFromSheet(changesSht, "Action");
        fieldCol = findColumnFromSheet(changesSht, "Field");
        valueCol = findColumnFromSheet(changesSht, "Final Value");

        if ((dayCol == null) || (itemShtCol == null) || (rowCol == null) || (actionCol == null) || (fieldCol == null)
                || (valueCol == null)) {
            System.out.printf(
                    "Could nto find all required columns in %s sheet: \"Day Delta\", \"Item Sheet\", \"Item Row\", \"Action\", \"Field\", \"Final Value\"");
            System.exit(1);
        }
        // Nw add the rows of today to an array
        row.next(); // Skip first row with headers
        while (row.hasNext()) {
            Row tr = row.next();
            if (tr.getCell(dayCol).getNumericCellValue() == day) {
                todaysChanges.add(tr);
            }
        }

        if (todaysChanges.size() == 0) {
            System.out.printf("No actions to take on day %d", day);
            return;
        }

        // Now scan through the changes doing the actions
        Iterator<Row> cItor = todaysChanges.iterator();
        Row item = null;
        while (cItor.hasNext()) {
            Row change = cItor.next();
            // Get the item that this change refers to
            XSSFSheet iSht = findSheet(change.getCell(itemShtCol).getStringCellValue());
            Integer idCol = findColumnFromSheet(iSht, "ID");
            Integer titleCol = findColumnFromSheet(iSht, "title");
            item = iSht.getRow((int) (change.getCell(rowCol).getNumericCellValue()));

            // If unset, it has a null value for the Leankit ID
            if (item.getCell(idCol) == null) {
                // Check if this is a 'create' operation. If not, ignore and continue past.
                if (!change.getCell(actionCol).getStringCellValue().equals("Create")) {
                    System.out.printf("Ignoring action %s on item %s (no ID present in row %d)\n",
                            change.getCell(actionCol).getStringCellValue(), item.getCell(titleCol).getStringCellValue(),
                            item.getRowNum());
                    break; // Break out and try next change
                }
            } else {
                // Check if this is a 'create' operation. If it is, ignore and continue past.
                if (change.getCell(actionCol).getStringCellValue().equals("Create")) {
                    System.out.printf("Ignoring action %s on item %s (attempting create on existing ID in row %d)\n",
                            change.getCell(actionCol).getStringCellValue(), item.getCell(titleCol).getStringCellValue(),
                            item.getRowNum());
                    break; // Break out and try next change
                }

            }
            System.out.printf("Committing to change %s on item %s\n", change.getCell(actionCol).getStringCellValue(),
                    item.getCell(titleCol).getStringCellValue());
            doAction(change, item);
        }

    }

    private String doAction(Row change, Row item) {
        // We need to find the ID of the board that this is targetting for a card
        // creation
        LeanKitAccess lka = new LeanKitAccess(config);
        XSSFSheet iSht = findSheet(change.getCell(itemShtCol).getStringCellValue());
        String boardNumber = lka
                .fetchBoardId(item.getCell(findColumnFromSheet(iSht, "Board Name")).getStringCellValue());

        /**
         * We need to either 'create' if ID == null && action == 'Create' or update if
         * ID != null && action == 'Modify'
         * 
         */
        if (change.getCell(actionCol).getStringCellValue().equals("Create")) {
            /**
             * We need to get the header row for this sheet and work out which columns the
             * fields are in. It is possible that fields could be different between sheets,
             * so we have to do this every 'change'
             */

            Iterator<Row> iRow = iSht.iterator();
            Row iFirst = iRow.next();
            /**
             * Now iterate across the cells finding out which fields need to be set
             */
            JSONObject fieldLst = new JSONObject();
            Iterator<Cell> cItor = iFirst.iterator();
            while (cItor.hasNext()) {
                Cell cl = cItor.next();
                String nm = cl.getStringCellValue();
                if ((nm.toLowerCase().equals("id") || (nm.toLowerCase().equals("board name")))) {
                    continue;
                }
                fieldLst.put(nm, cl.getColumnIndex());
            }

            // Now 'translate' the spreadsheet name:col pairs to fieldName:value pairs
            Iterator<String> keys = fieldLst.keys();
            JSONObject flds = new JSONObject();

            while (keys.hasNext()) {
                String key = keys.next();
                if (item.getCell(fieldLst.getInt(key)) != null) {
                    switch (key) {
                        case "Type": {
                            //Convert Type string to typeId
                        }
                        default: {

                            if (item.getCell(fieldLst.getInt(key)).getCellType() == CellType.STRING) { 
                                flds.put(key, item.getCell(fieldLst.getInt(key)).getStringCellValue()); 
                            } else {
                                if (DateUtil.isCellDateFormatted(item.getCell(fieldLst.getInt(key)))){
                                    SimpleDateFormat dtf = new SimpleDateFormat("yyyy-MM-dd");
                                    Date date = item.getCell(fieldLst.getInt(key)).getDateCellValue();
                                    flds.put(key, dtf.format(date).toString());
                                } else {
                                    flds.put(key, (int) item.getCell(fieldLst.getInt(key)).getNumericCellValue());
                                }
                            }
                        }

                    }
                }
            }

            Id card = createCard(lka, boardNumber, flds); // Change from human readable to API fields on
                                                                        // the way
            if (card == null) {
                System.out.printf("Could not create card on board %s with details: %s", boardNumber,
                        fieldLst.toString());
                System.exit(1);
            }
            return card.id;

        }
        return null;
    }

    public Id createCard(LeanKitAccess lka, String bNum, JSONObject fieldLst) {

        // Get available types so we can convert type string to typeid string
        ArrayList<CardType> cTypes = lka.fetchCardTypes(bNum);

        // First create an empty card and get back the full structure as a string

        CardLongRead newCard = lka.createCard(bNum, new JSONObject());

        // If for any reason the network fails, we get a null back.
        if (newCard == null) {
            return null;
        }

        JSONObject finalCard = new JSONObject();
        if (fieldLst.has("Type")) {
            // Find type in cTypes and add new. If not present, then 'create' will default
            for (int i = 0; i < cTypes.size(); i++) {
                if (cTypes.get(i).name.equals(fieldLst.get("Type"))) {
                    finalCard.put("typeId", cTypes.get(i).id);
                }
            }

        }
        // Then we need to check what sort of stuff we have back and can we update the
        // fields given.
        Iterator<String> fields = fieldLst.keys();
        while (fields.hasNext()) {
            String fldName = fields.next();
            switch (fldName.toLowerCase()) {
                case "id":
                case "board name":
                case "type":
                    break;
                default:
                    //Make Sure field names from speadsheet are part of the Card model.
                    Field[] fld = CardLongRead.class.getFields();
                    Boolean fieldFound = false;
                    for (int i =0 ; i < fld.length; i++) {
                        if (fld[i].getName().equals(fldName)) {
                            fieldFound = true;
                            break;
                        }
                    }
                    if (fieldFound) {
                        finalCard.put(fldName, fieldLst.get(fldName));
                    }
                    else {
                        System.out.printf("Incorrect field name \"%s\" provided for update on card %s",fldName, newCard.id);
                    }
                    break;
            }

        }
        Id id = new Id();
        id.id = lka.updateCardFromId(newCard.id, finalCard).id;
        return id;
    }

    private XSSFSheet findSheet(String name) {
        Iterator<XSSFSheet> s = teamShts.iterator();
        while (s.hasNext()) {
            XSSFSheet xs = s.next();
            if (xs.getSheetName().equals(name)) {
                return xs;
            }
        }
        return null;
    }

    private Integer getRefresh() {
        return refreshPeriod;
    }
}