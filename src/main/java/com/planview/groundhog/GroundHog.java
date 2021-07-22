package com.planview.groundhog;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Scanner;

import com.planview.groundhog.Leankit.Board;
import com.planview.groundhog.Leankit.Card;
import com.planview.groundhog.Leankit.CardType;
import com.planview.groundhog.Leankit.Comment;
import com.planview.groundhog.Leankit.CustomField;
import com.planview.groundhog.Leankit.Id;
import com.planview.groundhog.Leankit.Lane;
import com.planview.groundhog.Leankit.LeanKitAccess;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.http.impl.conn.PoolingHttpClientConnectionManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

public class GroundHog {

    private static final String OUR_GROUNDHOG_TAG = "groundhog2-generated";

    static Integer debugPrint = -1;

    Integer refreshPeriod = 14; // Number of days to run for by default
    Configuration config = new Configuration();
    String xlsxfn = "";
    FileInputStream xlsxfis = null;
    XSSFWorkbook wb = null;
    static Boolean useCron = false;
    static String statusFile = "";
    static String moveLane = null;
    static Boolean flagMove = false;
    static String deleteItems = "";
    static Integer flagDelete = -1;
    static Integer updatePeriod = 60 * 60 * 24;
    static Boolean useUpdatePeriod = false;
    static Integer startDay = -1;
    static Boolean cycleOnce = false;
    static PoolingHttpClientConnectionManager cm = null;
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

    public static void main(String[] args) throws IOException, InterruptedException {

        cm = new PoolingHttpClientConnectionManager();
        cm.setMaxTotal(10);  //Recommendation for LK is 5, but the 429 handler should keep us in check
        
        GroundHog hog = new GroundHog();
        hog.getCommandLine(args);
        hog.parseXlsx();
        hog.getConfig();
        Integer day = (startDay >= 0) ? startDay : 0;

        Scanner scanner = new Scanner(System.in);
        // Check to see if there is command line option to loop or use cron
        if (!useCron) {

            while (true) {
                try {

                    if (updatePeriod > 0) {
                        // Do todays activity and then sleep
                        hog.activity(day++);
                        Calendar now = Calendar.getInstance();
                        Calendar then = Calendar.getInstance();
                        if (useUpdatePeriod) {
                            then.add(Calendar.SECOND, updatePeriod);
                        } else {
                            then.add(Calendar.HOUR, 24);
                            then.set(Calendar.HOUR_OF_DAY, 3); // Set to three in the morning
                            then.set(Calendar.MINUTE, 0);
                            then.set(Calendar.SECOND, 0);
                        }
                        Long timeDiff = then.getTimeInMillis() - now.getTimeInMillis();

                        // Reset file after the time period has expired
                        day = checkWhatsNext(day, hog);
                        dpf(Debug.INFO, "%s\n", "Sleeping until: " + then.getTime());
                        Thread.sleep(timeDiff);
                    } else {
                        // Wait for a user input to continue

                        System.out.printf("Hit <CR> to continue to day %d... ", day);
                        String line = scanner.nextLine();
                        System.out.println(line); // Do this so that the compiler doesn't remove the unused 'line'
                        // Do todays activity and then sleep
                        hog.activity(day++);
                        // Reset file after the time period has expired
                        day = checkWhatsNext(day, hog);
                    }
                } catch (InterruptedException e) {
                    dpf(Debug.ERROR, "%s", "Early wake from daily timer");
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
                if (settings.has("day") && settings.has("flagDelete") && settings.has("flagMove")) {
                    fDay = settings.getInt("day");
                    if ((fDay == null) || (fDay < 0) || (fDay >= hog.getRefresh())) {
                        setUpStatusFile(statusFs); // Summat wrong, so try to re-initialise
                        fis.close();
                        fis = new BufferedReader(new FileReader(statusFs));
                        settings = new JSONObject(fis.readLine());
                    }
                } else {
                    // Once again,file must be corrupted, so try again
                    setUpStatusFile(statusFs); // Summat wrong, so try to re-initialise
                    fis.close();
                    fis = new BufferedReader(new FileReader(statusFs));
                    settings = new JSONObject(fis.readLine());
                }

                fDay = settings.getInt("day");
                flagDelete = settings.getInt("flagDelete");
                flagMove = settings.getBoolean("flagMove");

                fis.close();
                hog.activity(fDay++);
                // Reset file after the time period has expired
                fDay = checkWhatsNext(fDay, hog);
                settings.put("day", fDay);
                settings.put("flagDelete", flagDelete);
                settings.put("flagMove", flagMove);
                FileWriter fos = new FileWriter(statusFs);
                settings.write(fos);
                fos.flush();
                fos.close();

            } catch (IOException e) {
                dpf(Debug.ERROR, "(1): %s", e.getMessage()); // Any other error, we barf.
                System.exit(-2);
            }

        }
        scanner.close();
    }

    private static Integer checkWhatsNext(Integer day, GroundHog hog) {
        if (day >= hog.getRefresh()) {
            if (deleteItems.equalsIgnoreCase("cycle")) {
                flagDelete = day;
            } 
            if (moveLane != null) {
                flagMove = true;
            } 
            // We need to reset the day to zero
            if (cycleOnce == false) {
                return 0;
            } else {
                dpf(Debug.ERROR, "Completed cycle once as requested");
                System.exit(3);
            }
        }
        else if (deleteItems.equalsIgnoreCase("day")) {
            flagDelete = 0;
        }
        return day;
    }

    public static void setUpStatusFile(File statusFs) {
        // Newly created, so add the default settings in

        JSONObject settings = new JSONObject();
        try {
            FileWriter fos = new FileWriter(statusFs);
            settings.put("day", 0);
            settings.put("flagDelete", -1);
            settings.put("flagMove", false);
            settings.write(fos);
            fos.flush();
            fos.close();
        } catch (IOException e) {
            dpf(Debug.ERROR, "(2): %s", e.getMessage());
            System.exit(4);
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
        Option updateRate = new Option("u", "update", true, "Rate to process updates (in seconds). Defaults to 1 day");
        updateRate.setRequired(false);
        opts.addOption(updateRate);
        Option beginDay = new Option("b", "begin", true, "Day to begin updates. Defaults to day 0");
        beginDay.setRequired(false);
        opts.addOption(beginDay);
        Option doOnce = new Option("o", "once", false, "Run updates from Day 0 to end only once (do not loop)");
        doOnce.setRequired(false);
        opts.addOption(doOnce);
        Option deleteCycle = new Option("d", "delete", true, "Delete all artifacts on start of day or end of cycle");
        deleteCycle.setRequired(false);
        opts.addOption(deleteCycle);

        Option moveCycle = new Option("m", "move", true, "Move all artifacts on end of cycle to this lane");
        moveCycle.setRequired(false);
        opts.addOption(moveCycle);

        Option dbp = new Option("x", "debug", true, "Print out loads of helpful stuff: 0 - Info, 1 - And Errors, 2 - And Warnings, 3 - And Debugging, 4 - Verbose");
        dbp.setRequired(false);
        opts.addOption(dbp);

        CommandLineParser p = new DefaultParser();
        HelpFormatter hf = new HelpFormatter();
        CommandLine cl = null;
        try {
            cl = p.parse(opts, args);
        } catch (ParseException e) {
            dpf(Debug.ERROR, "(3): %s", e.getMessage());
            hf.printHelp(" ", opts);
            System.exit(5);
        }

        xlsxfn = cl.getOptionValue("filename");
        useCron = cl.hasOption("cron");
        if (cl.hasOption("status")) {
            statusFile = cl.getOptionValue("status");
        }

        if (cl.hasOption("update")) {
            useCron = false;
            useUpdatePeriod = true;
            updatePeriod = Integer.parseInt(cl.getOptionValue("update"));
        }

        if (cl.hasOption("begin")) {
            startDay = Integer.parseInt(cl.getOptionValue("begin"));
        }

        // Move items to a lane and then delete them if needed - or just delete.
        if (cl.hasOption("move")) {
            moveLane = cl.getOptionValue("move");
        }

        if (cl.hasOption("once")) {
            cycleOnce = true;
        }

        if (cl.hasOption("delete")) {
            deleteItems = cl.getOptionValue("delete");
        }

        if (cl.hasOption("debug")) {
            String optVal = cl.getOptionValue("debug");
            if (optVal != null) {
                debugPrint = Integer.parseInt(optVal);
            }
            else {
                debugPrint = 99;
            }
        }

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
            dpf(Debug.ERROR, "(4) %s", e.getMessage());

            System.exit(6);

        }
        try {
            wb = new XSSFWorkbook(xlsxfis);
            xlsxfis.close();
        } catch (IOException e) {
            dpf(Debug.ERROR, "(5) %s", e.getMessage());
            System.exit(7);
        }

        // These two should come first in the file and must be present

        configSht = wb.getSheet("Config");
        // This is a common sheet listing all the progression changes for all boards
        changesSht = wb.getSheet("Changes");

        Integer shtCount = wb.getNumberOfSheets();
        if ((shtCount < 3) || (configSht == null) || (changesSht == null)) {
            dpf(Debug.ERROR, "%s",
                    "Did not detect correct sheets in the spreadsheet: \"Config\",\"Changes\" and one, or more, board(s)");
            System.exit(8);
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
        HashMap<String, Object> fieldMap = new HashMap<>();

        // Assume that the titles are the first row
        Iterator<Row> ri = configSht.iterator();
        if (!ri.hasNext()) {
            dpf(Debug.ERROR, "%s", "Did not detect any header info on Config sheet (first row!)");
            System.exit(9);
        }
        Row hdr = ri.next();
        Iterator<Cell> cptr = hdr.cellIterator();

        while (cptr.hasNext()) {
            Cell cell = cptr.next();
            Integer idx = cell.getColumnIndex();
            String cellName = cell.getStringCellValue().trim().toLowerCase();
            if (cols.contains(cellName)) {
                fieldMap.put(cellName, idx); // Store the column index of the field
            }
        }

        if (fieldMap.size() != cols.size()) {
            dpf(Debug.ERROR, "%s", "Did not detect correct columns on Config sheet: " + cols.toString());
            System.exit(10);
        }

        if (!ri.hasNext()) {
            dpf(Debug.ERROR, "%s",
                    "Did not detect any field info on Config sheet (first cell must be non-blank, e.g. url to a real host)");
            System.exit(11);
        }
        // Now we know which columns contain the data, scan down the sheet looking for a
        // row with data in the 'url' cell
        while (ri.hasNext()) {
            Row drRow = ri.next();
            String cv = drRow.getCell((int) (fieldMap.get(cols.get(0)))).getStringCellValue();
            if (cv != null) {

                for (int i = 0; i < cols.size(); i++) {

                    try {
                        String idx = cols.get(i);
                        Object obj = fieldMap.get(idx);
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
                        dpf(Debug.ERROR, "(6) %s", e.getMessage());
                        System.exit(12);
                    }

                }

                break; // Exit out of while loop as we are only interested in the first one we find
            }
        }

        // Creds are now found and set. If not, you're buggered.
        /**
         * We can opt to use username/password or apikey. Unfortunately, we have to hard
         * code the field names in here, even though I was trying to use the fields from
         * the Configuration class.
         **/

        if ((config.apikey == null) && ((config.username == null) || (config.password == null))) {
            dpf(Debug.ERROR, "%s", "Did not detect enough user info: apikey or username/password pair");
            System.exit(13);
        }
        if ((config.cyclelength != null) && (config.cyclelength.intValue() != 0)) {
            refreshPeriod = config.cyclelength.intValue();
        }
        return;
    }

    private Integer findColumnFromName(Row firstRow, String name) {
        Iterator<Cell> frtc = firstRow.iterator();
        // First, find the column that the "Day Delta" info is in
        int dayCol = -1;
        int td = 0;
        if (!frtc.hasNext()) {
            return dayCol;
        }

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
    Integer value1Col = null;

    private Integer findColumnFromSheet(XSSFSheet sht, String name) {
        Iterator<Row> row = sht.iterator();
        if (!row.hasNext()) {
            return null;
        }
        Row firstRow = row.next(); // Get the header row
        Integer col = findColumnFromName(firstRow, name);
        if (col < 0) {
            return null;
        }
        return col;
    }

    private Boolean hasOurTag(Card cd) {
        LeanKitAccess lka = new LeanKitAccess(config, debugPrint, cm);
        ArrayList<Comment> cl = lka.fetchCommentsForCard(cd);
        if (cl != null) {
            for (int i = 0; i < cl.size(); i++) {
                if (cl.get(i).text.equals(OUR_GROUNDHOG_TAG)) {
                    return true;
                }
            }
        }
        return false;
    }

    private void deleteUserItems(Integer rewind) {

        // Collate all the boards in the items sheets and get their IDs
        // For each board, get the cards and check whether we have an ID that matches a
        // row in the items
        ArrayList<String> bLst = new ArrayList<>(); // Boards we have
        ArrayList<String> cLst = new ArrayList<>(); // Card list we have
        Iterator<Row> row = changesSht.iterator();

        row.next(); // Skip header
        while (row.hasNext()) {
            Row change = row.next();

            // XLS Spreadsheets are horrible!
            if ((change.getCell(itemShtCol) == null) || (change.getCell(rowCol) == null)) {
                continue;
            }
            XSSFSheet iSht = findSheet(change.getCell(itemShtCol).getStringCellValue());
            Row item = iSht.getRow((int) (change.getCell(rowCol).getNumericCellValue() - 1));
            if (item.getCell(findColumnFromSheet(iSht, "Board Name")) == null) {
                continue;
            }
            String boardName = (item.getCell(findColumnFromSheet(iSht, "Board Name")).getStringCellValue());
            String cardId = "";
            if (item.getCell(findColumnFromSheet(iSht, "ID")) != null) {
                cardId = item.getCell(findColumnFromSheet(iSht, "ID")).getStringCellValue();
            }
            if (!bLst.contains(boardName)) {
                bLst.add(boardName);
            }

            if ((cardId != "") && (cardId != null) && !cLst.contains(cardId)) {
                cLst.add(cardId);
            }
        }
        if ((bLst.size() == 0) || (cLst.size() == 0)) {
            return;
        }
        LeanKitAccess lka = new LeanKitAccess(config, debugPrint, cm);
        Iterator<String> bIter = bLst.iterator();
        ArrayList<Card> leftOvers = new ArrayList<>();
        while (bIter.hasNext()) {
            String bName = bIter.next();
            Board brd = lka.fetchBoard(bName);
            if (brd == null) {
                dpf(Debug.ERROR, "Could not locate board \"%s\"\n", bName);
                return;
            }
            dpf(Debug.DEBUG, "Requesting card IDs from board \"%s\"\n", brd.id);
            ArrayList<Card> cards = lka.fetchCardIDsFromBoard(brd.id, rewind);
            ArrayList<Card> removeCards = new ArrayList<>();
            if (cards != null) {
                Iterator<Card> cIter = cards.iterator();
                while (cIter.hasNext()) {
                    Card cd = cIter.next();
                    if (!hasOurTag(cd)) {
                        removeCards.add(cd);
                    }
                }
                leftOvers.addAll(removeCards);
            } else {
                return;
            }
        }
        // Should now have a list of cards that are not ours
        lka.deleteCards(leftOvers);
        dpf(Debug.DEBUG, "Deleted %d %s\n", leftOvers.size(), (leftOvers.size() != 1) ? "cards" : "card");
    }

    private void moveOurItems() {
        LeanKitAccess lka = new LeanKitAccess(config, debugPrint, cm);
        Iterator<Row> row = changesSht.iterator();
        row.next(); // Skip header
        while (row.hasNext()) {
            Row change = row.next();
            // XLS Spreadsheets are horrible!
            if ((change.getCell(itemShtCol) == null) || (change.getCell(rowCol) == null)) {
                continue;
            }
            XSSFSheet iSht = findSheet(change.getCell(itemShtCol).getStringCellValue());
            Row item = iSht.getRow((int) (change.getCell(rowCol).getNumericCellValue() - 1));
            String bName = (item.getCell(findColumnFromSheet(iSht, "Board Name")).getStringCellValue());

            // Does this actually work here?
            if (item.getCell(findColumnFromSheet(iSht, "ID")).getCellType() == CellType.BLANK) {
                continue;
            }
            // Need to verify, but in the meantime - belt'n'braces
            String cardId = (item.getCell(findColumnFromSheet(iSht, "ID")).getStringCellValue());
            if (cardId == null || cardId.equals("")) {
                continue;
            }
            Board brd = lka.fetchBoard(bName);
            Card card = lka.fetchCard(cardId);
            if (brd == null) {
                dpf(Debug.ERROR, "Could not locate board \"%s\"\n", bName);
                return;
            }

            if (card == null) {
                dpf(Debug.ERROR, "Could not locate card \"%s\"\n", cardId);
                break;
            }

            JSONObject fld = new JSONObject();
            JSONObject vals = new JSONObject();
            vals.put("value1", moveLane); // Lower levels check for the correct lane name.
            vals.put("value2", "Archiving boards from GroundHog cycle");
            fld.put("Lane", vals);
            Id id = updateCard(lka, brd, card, fld);
            if (id == null) {
                dpf(Debug.DEBUG, "Could not move card %s, on board \"%s\" to lane \"%s\"\n", cardId, bName, moveLane);
                System.exit(14);
            } else {
                dpf(Debug.DEBUG, "Moved card %s, on board \"%s\" to lane \"%s\"\n", cardId, bName, moveLane);
            }
            // Now that the item is moved, delete the ID from the item row so we can create
            // new ones the next time around
            item.getCell(findColumnFromSheet(iSht, "ID")).setCellValue("");
            writeFile(); // Do this everytime just in case of user Ctrl-C or network failure
        }
    }

    private static void dpf(Integer level, String fmt, Object... parms) {
        String lp = null;
        switch (level) {
            case 0: {
                lp = "INFO: ";
                break;
            }
            case 1: {
                lp = "ERROR: ";
                break;
            }
            case 2: {
                lp = "WARN: ";
                break;
            }
            case 3: {
                lp = "DEBUG: ";
                break;
            }
            case 4: {
                lp = "VERBOSE: ";
                break;
            }
        }
        if (level <= debugPrint) {
            System.out.printf(lp+fmt, parms);
        }
    }

    private void activity(Integer day) throws IOException, InterruptedException {
        // Find all the change records for today
        Iterator<Row> row = changesSht.iterator();
        ArrayList<Row> todaysChanges = new ArrayList<Row>();

        dayCol = findColumnFromSheet(changesSht, "Day Delta");
        itemShtCol = findColumnFromSheet(changesSht, "Item Sheet");
        rowCol = findColumnFromSheet(changesSht, "Item Row");
        actionCol = findColumnFromSheet(changesSht, "Action");
        fieldCol = findColumnFromSheet(changesSht, "Field");
        value1Col = findColumnFromSheet(changesSht, "Value");

        if ((dayCol == null) || (itemShtCol == null) || (rowCol == null) || (actionCol == null) || (fieldCol == null)
                || (value1Col == null)) {
            dpf(Debug.ERROR, "Could not find all required columns in %s sheet: \"Day Delta\", \"Item Sheet\", \"Item Row\", \"Action\", \"Field\", \"Value\"\n",
                    changesSht.getSheetName());
            System.exit(15);
        }
        // Nw add the rows of today to an array
        row.next(); // Skip first row with headers
        while (row.hasNext()) {

            Row tr = row.next();
            if (tr.getCell(dayCol) != null) {
                if (tr.getCell(dayCol).getNumericCellValue() == day) {
                    todaysChanges.add(tr);
                }
            }
        }

        if (flagDelete >= 0) {
            deleteUserItems(flagDelete);
        }

        if (flagMove == true) {
            moveOurItems();
            flagMove = false;
        }

        if (todaysChanges.size() == 0) {
            dpf(Debug.INFO, "No actions to take on day %d\n", day);
            return;
        } else {
            dpf(Debug.INFO, "%d actions to take on day %d\n", todaysChanges.size(), day);
        }

        // Now scan through the changes doing the actions
        Iterator<Row> cItor = todaysChanges.iterator();
        Row item = null;
        while (cItor.hasNext()) {
            Row change = cItor.next();
            // Get the item that this change refers to
            // First check the validity of the info
            if ((change.getCell(itemShtCol) == null) || (change.getCell(rowCol) == null)
                    || (change.getCell(actionCol) == null)) {
                dpf(Debug.WARN, "Cannot decode change info in row \"%d\" - skipping\n", change.getRowNum());
                continue;
            }
            XSSFSheet iSht = findSheet(change.getCell(itemShtCol).getStringCellValue());
            Integer idCol = findColumnFromSheet(iSht, "ID");
            Integer titleCol = findColumnFromSheet(iSht, "title");
            Integer boardCol = findColumnFromSheet(iSht, "Board Name");
            Integer typeCol = findColumnFromSheet(iSht, "Type");

            item = iSht.getRow((int) (change.getCell(rowCol).getNumericCellValue() - 1));

            if ((idCol == null) || (titleCol == null)) {
                dpf(Debug.WARN, "Cannot locate \"ID\" and \"title\" columns needed in sheet \"%s\" - skipping\n",
                        iSht.getSheetName());
                continue;
            }

            // Check board name is present for a Create
            if ((change.getCell(actionCol).getStringCellValue().equalsIgnoreCase("Create")) && ((boardCol == null)
                    || (item.getCell(boardCol) == null) || (item.getCell(boardCol).getStringCellValue().isEmpty()))) {
                dpf(Debug.WARN, "Cannot locate \"Board Name\" column needed in sheet \"%s\"  for a Create - skipping\n",
                        iSht.getSheetName());
                continue;
            }

            // Check title is present for a Create
            if ((change.getCell(actionCol).getStringCellValue().equalsIgnoreCase("Create"))
                    && ((item.getCell(titleCol) == null) || (item.getCell(titleCol).getStringCellValue().isEmpty()))) {
                dpf(Debug.WARN, "Required \"title\" column/data missing in sheet \"%s\", row: %d for a Create - skipping\n",
                        iSht.getSheetName(), item.getRowNum());
                continue;
            }

            if ((typeCol == null) || (item.getCell(typeCol) == null)) {
                dpf(Debug.WARN, "Cannot locate \"Type\" column on row:  %d  - using default for board\n", item.getRowNum());
            }

            // If unset, it has a null value for the Leankit ID
            if ((item.getCell(idCol) == null) || (item.getCell(idCol).getStringCellValue() == "")) {
                // Check if this is a 'create' operation. If not, ignore and continue past.
                if (!change.getCell(actionCol).getStringCellValue().equalsIgnoreCase("Create")) {
                    dpf(Debug.WARN, "Ignoring action \"%s\" on item \"%s\" (no ID present in row: %d)\n",
                            change.getCell(actionCol).getStringCellValue(), item.getCell(titleCol).getStringCellValue(),
                            item.getRowNum());
                    continue; // Break out and try next change
                }
            } else {
                // Check if this is a 'create' operation. If it is, ignore and continue past.
                if (change.getCell(actionCol).getStringCellValue().equalsIgnoreCase("Create")) {
                    dpf(Debug.WARN, "Ignoring action \"%s\" on item \"%s\" (attempting create on existing ID in %s row: %d)\n",
                            change.getCell(actionCol).getStringCellValue(), item.getCell(titleCol).getStringCellValue(),
                            iSht.getSheetName(), item.getRowNum());
                    continue; // Break out and try next change
                }

            }
            String id = null;
            if (change.getCell(actionCol).getStringCellValue().equalsIgnoreCase("Create")) {
                id = doAction(change, item);
                if (item.getCell(idCol) == null) {

                    item.createCell(idCol);
                }
                item.getCell(idCol).setCellValue(id);
                XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
                writeFile();
                dpf(Debug.DEBUG, "Create card \"%s\" on board \"%s\"\n", item.getCell(titleCol).getStringCellValue(),
                        item.getCell(boardCol).getStringCellValue());
            } else {
                id = doAction(change, item);
                dpf(Debug.DEBUG, "Modified card \"%s\" on board \"%s\"\n", item.getCell(titleCol).getStringCellValue(),
                        item.getCell(boardCol).getStringCellValue());
            }

            if (id == null) {

                dpf(Debug.ERROR, "%s", "Got null back from doAction(). Seek help!\n");
            }
        }
    }

    private void writeFile() {
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
                        dpf(Debug.ERROR, "%s while closing file %s\n", e, xlsxfn);
                    }
                } catch (IOException e) {
                    dpf(Debug.ERROR, "%s while writing file %s\n", e, xlsxfn);
                    oStr.close(); // If this fails, just give up!
                }
            } catch (IOException e) {
                dpf(Debug.ERROR, "%s while opening/closing file %s\n", e, xlsxfn);
            }
            if (loopCnt == 0) {
                break;
            }

            Calendar now = Calendar.getInstance();
            Calendar then = Calendar.getInstance();
            then.add(Calendar.SECOND, 5);
            Long timeDiff = then.getTimeInMillis() - now.getTimeInMillis();
            if (donePrint) {
                dpf(Debug.ERROR, "File \"%s\" in use. Please close to let this program continue\n", xlsxfn);
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

    private String doAction(Row change, Row item) {
        // We need to find the ID of the board that this is targetting for a card
        // creation
        LeanKitAccess lka = new LeanKitAccess(config, debugPrint, cm);
        XSSFSheet iSht = findSheet(change.getCell(itemShtCol).getStringCellValue());
        String boardName = (item.getCell(findColumnFromSheet(iSht, "Board Name")).getStringCellValue());
        Board brd = lka.fetchBoard(boardName);

        if (brd == null) {
            dpf(Debug.WARN, "Cannot find board with name %s for item on sheet %s, row %d\n", boardName,
                    change.getCell(itemShtCol).getStringCellValue(),
                    (int) change.getCell(rowCol).getNumericCellValue());
            return null;
        }
        String boardNumber = brd.id;
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
        Integer idCol = null;
        while (cItor.hasNext()) {
            Cell cl = cItor.next();
            String nm = cl.getStringCellValue();
            if (nm.toLowerCase().equals("id")) {
                idCol = cl.getColumnIndex();
                continue;
            } else if (nm.toLowerCase().equals("board name")) {
                continue;
            }
            fieldLst.put(nm, cl.getColumnIndex());
        }

        if (change.getCell(actionCol).getStringCellValue().equalsIgnoreCase("Create")) {
            // Now 'translate' the spreadsheet name:col pairs to fieldName:value pairs

            Iterator<String> keys = fieldLst.keys();
            JSONObject flds = new JSONObject();

            while (keys.hasNext()) {
                JSONObject vals = new JSONObject();
                String key = keys.next();
                if (item.getCell(fieldLst.getInt(key)) != null) {
                    vals.put("value1", convertCells(item, fieldLst.getInt(key)));
                    flds.put(key, vals);
                }
            }

            // Add out tag to every card created
            JSONObject dc = new JSONObject();
            dc.put("value1", OUR_GROUNDHOG_TAG);
            flds.put("comments", dc);

            /**
             * We need to either 'create' if ID == null && action == 'Create' or update if
             * ID != null && action == 'Modify'
             * 
             */

            Id card = createCard(lka, brd, flds); // Change from human readable to API fields on
                                                  // the way
            if (card == null) {
                dpf(Debug.ERROR, "Could not create card on board \"%s\" with details: \"%s\"", boardNumber, flds.toString());
                System.exit(16);
            }
            return card.id;

        } else if (change.getCell(actionCol).getStringCellValue().equalsIgnoreCase("Modify")) {
            // Fetch the ID from the item and then fetch that card
            Card card = lka.fetchCard(item.getCell(idCol).getStringCellValue());

            if (card == null) {
                dpf(Debug.ERROR, "Could not locate card \"%s\" on board \"%s\"\n", item.getCell(idCol).getStringCellValue(),
                        boardNumber);
            } else {
            JSONObject fld = new JSONObject();
            JSONObject vals = new JSONObject();

            vals.put("value1", convertCells(change, value1Col));

            fld.put(change.getCell(fieldCol).getStringCellValue(), vals);
            Id id = updateCard(lka, brd, card, fld);
            if (id == null) {
                dpf(Debug.ERROR, "Could not modify card \"%s\" on board %s with details: %s", card.id, boardNumber, fld.toString());
                System.exit(17);
            }
            return id.id;
        }
        }
        // Unknown option comes here
        return null;
    }

    private Object convertCells(Row change, Integer col) {

        if (change.getCell(col) != null) {
            // Need to get the correct type of field
            if (change.getCell(col).getCellType() == CellType.FORMULA) {
                if (change.getCell(col).getCachedFormulaResultType() == CellType.STRING) {
                    return change.getCell(col).getStringCellValue();
                } else if (change.getCell(col).getCachedFormulaResultType() == CellType.NUMERIC) {
                    if (DateUtil.isCellDateFormatted(change.getCell(col))) {
                        SimpleDateFormat dtf = new SimpleDateFormat("yyyy-MM-dd");
                        Date date = change.getCell(col).getDateCellValue();
                        return dtf.format(date).toString();
                    } else {
                        return (int) change.getCell(col).getNumericCellValue();
                    }
                }
            } else if (change.getCell(col).getCellType() == CellType.STRING) {
                return change.getCell(col).getStringCellValue();
            } else {
                if (DateUtil.isCellDateFormatted(change.getCell(col))) {
                    SimpleDateFormat dtf = new SimpleDateFormat("yyyy-MM-dd");
                    Date date = change.getCell(col).getDateCellValue();
                    return dtf.format(date).toString();
                } else {
                    return change.getCell(col).getNumericCellValue();
                }
            }
        }
        return null;
    }

    private Id createCard(LeanKitAccess lka, Board brd, JSONObject fieldLst) {

        // First create an empty card and get back the full structure

        Card newCard = lka.createCard(brd.id, new JSONObject());

        // If for any reason the network fails, we get a null back.
        if (newCard == null) {
            return null;
        }
        return updateCard(lka, brd, newCard, fieldLst);
    }

    private ArrayList<Lane> findLanesFromName(ArrayList<Lane> lanes, String name) {
        ArrayList<Lane> ln = new ArrayList<>();
        for (int i = 0; i < lanes.size(); i++) {
            if (lanes.get(i).name.equals(name)) {
                ln.add(lanes.get(i));
                break;
            }
        }
        return ln;
    }

    private ArrayList<Lane> findLanesFromParentId(Lane[] lanes, String id) {
        ArrayList<Lane> ln = new ArrayList<>();
        for (int i = 0; i < lanes.length; i++) {
            if (lanes[i].parentLaneId != null) {
                if (lanes[i].parentLaneId.equals(id)) {
                    ln.add(lanes[i]);
                }
            }
        }
        return ln;
    }

    private Lane findLaneFromString(Board brd, String name) {
        String[] lanes = name.split("\\|");

        // Is this VS Code failing to handle the '|' character gracefully....
        if (lanes.length == 1) {
            String[] vsLanes = name.split("%");
            if (vsLanes.length > 1) {
                lanes = vsLanes;
            }
        }
        ArrayList<Lane> searchLanes = new ArrayList<>(Arrays.asList(brd.lanes));
        int j = 0;
        ArrayList<Lane> lanesToCheck = findLanesFromName(searchLanes, lanes[j]);
        do {
            if (++j >= lanes.length) {
                searchLanes = lanesToCheck;
                break;
            }
            Iterator<Lane> lIter = lanesToCheck.iterator();
            while (lIter.hasNext()) {
                ArrayList<Lane> foundLanes = new ArrayList<>();
                Lane ln = lIter.next();
                ArrayList<Lane> childLanes = findLanesFromParentId(brd.lanes, ln.id);
                Iterator<Lane> clIter = childLanes.iterator();
                while (clIter.hasNext()) {
                    Lane cl = clIter.next();
                    if (cl.name.equals(lanes[j])) {
                        foundLanes.add(cl);
                    }
                }
                if (foundLanes.size() > 0) {
                    lanesToCheck = foundLanes;
                }
            }

        } while (true);

        if (searchLanes.size() == 0) {
            dpf(Debug.WARN, "Cannot find lane \"%s\"on board \"%s\"\n", name, brd.title);
        }
        if (searchLanes.size() > 1) {
            dpf(Debug.WARN, "Ambiguous lane name \"%s\"on board \"%s\"\n", name, brd.title);
        }

        return searchLanes.get(0);
    }

    private Id updateCard(LeanKitAccess lka, Board brd, Card card, JSONObject fieldLst) {
        // Get available types so we can convert type string to typeid string
        CardType[] cTypes = brd.cardTypes;

        JSONObject finalUpdates = new JSONObject();
        if (fieldLst.has("Type")) {
            JSONObject fldVals = (JSONObject) fieldLst.get("Type");
            // Find type in cTypes and add new. If not present, then 'create' will default
            for (int i = 0; i < cTypes.length; i++) {
                if (cTypes[i].name.equalsIgnoreCase(fldVals.get("value1").toString())) {
                    finalUpdates.put("typeId", cTypes[i].id);
                }
            }

        }
        // Then we need to check what sort of stuff we have back and can we update the
        // fields given.
        Iterator<String> fields = fieldLst.keys();
        while (fields.hasNext()) {
            String fldName = fields.next();
            JSONObject fldValues = (JSONObject) fieldLst.get(fldName);
            switch (fldName.toLowerCase()) {
                // Remove these regardless of what the user has set the case to
                // User might have put id, ID or Id, so just try to accommodate
                case "id":
                case "board name":
                    break;
                case "type": {
                    JSONObject fldVals = (JSONObject) fieldLst.get("Type");
                    // Find type in cTypes and add new. If not present, then 'create' will default
                    for (int i = 0; i < cTypes.length; i++) {
                        if (cTypes[i].name.equalsIgnoreCase(fldVals.get("value1").toString())) {
                            JSONObject result = new JSONObject();
                            result.put("value1", cTypes[i].id);
                            finalUpdates.put("typeId", result);
                        }
                    }
                    break;
                }

                // Add the parent straight in as it will be handled in the lower layer
                case "parent":
                    finalUpdates.put(fldName, fieldLst.get(fldName));
                    break;
                case "lane": {
                    Lane foundLane = null;
                    int idx = fldValues.get("value1").toString().indexOf(",");
                    String woc = null;
                    if (idx < 0) {
                        foundLane = findLaneFromString(brd, fldValues.get("value1").toString());
                    } else if (idx > 0) {
                        foundLane = findLaneFromString(brd, fldValues.get("value1").toString().substring(0, idx));
                        woc = fldValues.get("value1").toString().substring(idx + 1).trim();
                    } else {
                        return null;
                    }
                    if (foundLane != null) {
                        if (findLanesFromParentId(brd.lanes, foundLane.id).size() != 0) { // Cannot move to a parent
                                                                                          // lane
                            return null;
                        }
                        JSONObject result = new JSONObject();
                        result.put("value1", foundLane.id);
                        if (woc != null) {
                            result.put("value2", woc);
                        }
                        finalUpdates.put(fldName, result);
                    }
                    break;
                }
                // case "attachments":
                default:
                    // Make Sure field names from speadsheet are part of the Card model.... or....
                    Field[] fld = Card.class.getFields();
                    Boolean fieldFound = false;
                    for (int i = 0; i < fld.length; i++) {
                        if (fld[i].getName().equalsIgnoreCase(fldName)) {
                            fieldFound = true;
                            break;
                        }
                    }
                    if (fieldFound) {
                        finalUpdates.put(fldName, fieldLst.get(fldName));
                    } else {
                        // ....see if they are in the customField set
                        CustomField[] cfld = card.customFields;
                        for (int i = 0; i < cfld.length; i++) {
                            if (cfld[i].label.equalsIgnoreCase(fldName)) {
                                fieldFound = true;
                                break;
                            }
                        }
                        if (fieldFound) {
                            JSONObject result = new JSONObject();
                            result.put("value1", fldName);
                            result.put("value2", ((JSONObject) fieldLst.get(fldName)).get("value1"));
                            finalUpdates.put("CustomField", result);
                        } else {
                            dpf(Debug.WARN, "Incorrect field name \"%s\" provided for update on card %s\n", fldName, card.id);
                            return null;
                        }
                    }
                    break;
            }

        }
        Id id = new Id();
        Card upc = lka.updateCardFromId(brd, card, finalUpdates);
        if (upc != null) {
            id.id = upc.id;
            return id;
        } else {
            return null;
        }
    }

    private XSSFSheet findSheet(String name) {
        Iterator<XSSFSheet> s = teamShts.iterator();
        while (s.hasNext()) {
            XSSFSheet xs = s.next();
            if (xs.getSheetName().equalsIgnoreCase(name)) {
                return xs;
            }
        }
        return null;
    }

    private Integer getRefresh() {
        return refreshPeriod;
    }
}