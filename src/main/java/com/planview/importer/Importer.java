package com.planview.importer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

public class Importer {

    static Boolean debugPrint = false;
    Configuration config = new Configuration();
    String xlsxfn = "";
    static Integer group = 0;   //Grouping
    FileInputStream xlsxfis = null;
    XSSFWorkbook wb = null;

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

        Importer hog = new Importer();

        hog.getCommandLine(args);
        hog.parseXlsx();
        hog.getConfig();
    
        hog.activity(group);
       
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
            dpf("%s", e.getMessage());
            System.exit(1);
        }
    }

    public void getCommandLine(String[] args) {
        Options opts = new Options();
        Option xlsxname = new Option("f", "filename", true, "source XLSX spreadsheet");
        xlsxname.setRequired(true);
        opts.addOption(xlsxname);
        
        Option groupOpt = new Option("g", "group", true, "Identifier of group to process (if present)");
        groupOpt.setRequired(false);
        opts.addOption(groupOpt);

        Option dbp = new Option("x", "debug", false, "Print out loads of helpful stuff");
        dbp.setRequired(false);
        opts.addOption(dbp);

        CommandLineParser p = new DefaultParser();
        HelpFormatter hf = new HelpFormatter();
        CommandLine cl = null;
        try {
            cl = p.parse(opts, args);
        } catch (ParseException e) {
            dpf("%s", e.getMessage());
            hf.printHelp(" ", opts);
            System.exit(1);
        }

        xlsxfn = cl.getOptionValue("filename");

        if (cl.hasOption("group")) {
            group = Integer.parseInt(cl.getOptionValue("group"));
        }

        debugPrint = cl.hasOption("debug");

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
            dpf("%s", e.getMessage());

            System.exit(1);

        }
        try {
            wb = new XSSFWorkbook(xlsxfis);
            xlsxfis.close();
        } catch (IOException e) {
            dpf("%s", e.getMessage());
            System.exit(1);
        }

        // These two should come first in the file and must be present

        configSht = wb.getSheet("Config");
        // This is a common sheet listing all the progression changes for all boards
        changesSht = wb.getSheet("Changes");

        Integer shtCount = wb.getNumberOfSheets();
        if ((shtCount < 3) || (configSht == null) || (changesSht == null)) {
            dpf("%s",
                    "Did not detect correct sheets in the spreadsheet: \"Config\",\"Changes\" and one, or more, team(s)");
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
        HashMap<String, Object> fieldMap = new HashMap<>();

        // Assume that the titles are the first row
        Iterator<Row> ri = configSht.iterator();
        if (!ri.hasNext()) {
            dpf("%s", "Did not detect any header info on Config sheet (first row!)");
            System.exit(1);
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
            dpf("%s", "Did not detect correct columns on Config sheet: " + cols.toString());
            System.exit(1);
        }

        if (!ri.hasNext()) {
            dpf("%s",
                    "Did not detect any field info on Config sheet (first cell must be non-blank, e.g. url to a real host)");
            System.exit(1);
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
                        dpf("%s", e.getMessage());
                        System.exit(1);
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
            dpf("%s", "Did not detect enough user info: apikey or username/password pair");
            System.exit(1);
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

    private static void dpf(String fmt, Object... parms) {
        if (debugPrint) {
            System.out.printf(fmt, parms);
        }
    }

    private void activity(Integer day) throws IOException, InterruptedException {
        // Find all the change records for today
        Iterator<Row> row = changesSht.iterator();
        ArrayList<Row> todaysChanges = new ArrayList<Row>();
        SaveThread saver = null;

        dayCol = findColumnFromSheet(changesSht, "Group");
        itemShtCol = findColumnFromSheet(changesSht, "Item Sheet");
        rowCol = findColumnFromSheet(changesSht, "Item Row");
        actionCol = findColumnFromSheet(changesSht, "Action");
        fieldCol = findColumnFromSheet(changesSht, "Field");
        value1Col = findColumnFromSheet(changesSht, "Value");


        if ((dayCol == null) || (itemShtCol == null) || (rowCol == null) || (actionCol == null) || (fieldCol == null)
                || (value1Col == null)) {
            dpf("Could not find all required columns in %s sheet: \"Group\", \"Item Sheet\", \"Item Row\", \"Action\", \"Field\", \"Value\"\n",
                    changesSht.getSheetName());
            System.exit(1);
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

        if (todaysChanges.size() == 0) {
            dpf("No actions to take for group %d\n", day);
            return;
        } else {
            dpf("%d actions to take for group %d\n", todaysChanges.size(), day);
        }

        // Now scan through the changes doing the actions
        Iterator<Row> cItor = todaysChanges.iterator();
        Row item = null;

        if (cItor.hasNext()) {
            //Create a file save thread
            saver = new SaveThread(xlsxfn, wb, debugPrint);
        }

        while (cItor.hasNext()) {
            Row change = cItor.next();
            // Get the item that this change refers to
            // First check the validity of the info
            if ((change.getCell(itemShtCol) == null) || (change.getCell(rowCol) == null)
                    || (change.getCell(actionCol) == null)) {
                dpf("Cannot decode change info in row \"%d\" - skipping\n", change.getRowNum());
                continue;
            }
            XSSFSheet iSht = findSheet(change.getCell(itemShtCol).getStringCellValue());
            Integer idCol = findColumnFromSheet(iSht, "ID");
            Integer titleCol = findColumnFromSheet(iSht, "title");
            Integer boardCol = findColumnFromSheet(iSht, "Board Name");
            Integer typeCol = findColumnFromSheet(iSht, "Type");

            item = iSht.getRow((int) (change.getCell(rowCol).getNumericCellValue() - 1));

            if ((idCol == null) || (titleCol == null)) {
                dpf("Cannot locate \"ID\" and \"title\" columns needed in sheet \"%s\" - skipping\n",
                        iSht.getSheetName());
                continue;
            }

            // Check board name is present for a Create
            if ((change.getCell(actionCol).getStringCellValue().equals("Create")) && ((boardCol == null)
                    || (item.getCell(boardCol) == null) || (item.getCell(boardCol).getStringCellValue().isEmpty()))) {
                dpf("Cannot locate \"Board Name\" column needed in sheet \"%s\"  for a Create - skipping\n",
                        iSht.getSheetName());
                continue;
            }

            // Check title is present for a Create
            if ((change.getCell(actionCol).getStringCellValue().equals("Create"))
                    && ((item.getCell(titleCol) == null) || (item.getCell(titleCol).getStringCellValue().isEmpty()))) {
                dpf("Required \"title\" column/data missing in sheet \"%s\", row: %d for a Create - skipping\n",
                        iSht.getSheetName(), item.getRowNum());
                continue;
            }

            if ((typeCol == null) || (item.getCell(typeCol) == null)) {
                dpf("Cannot locate \"Type\" column on row:  %d  - using default for board\n", item.getRowNum());
            }

            // If unset, it has a null value for the Leankit ID
            if ((item.getCell(idCol) == null) || (item.getCell(idCol).getStringCellValue() == "")) {
                // Check if this is a 'create' operation. If not, ignore and continue past.
                if (!change.getCell(actionCol).getStringCellValue().equals("Create")) {
                    dpf("Ignoring action \"%s\" on item \"%s\" (no ID present in row: %d)\n",
                            change.getCell(actionCol).getStringCellValue(), item.getCell(titleCol).getStringCellValue(),
                            item.getRowNum());
                    continue; // Break out and try next change
                }
            } else {
                // Check if this is a 'create' operation. If it is, ignore and continue past.
                if (change.getCell(actionCol).getStringCellValue().equals("Create")) {
                    dpf("Ignoring action \"%s\" on item \"%s\" (attempting create on existing ID in row: %d)\n",
                            change.getCell(actionCol).getStringCellValue(), item.getCell(titleCol).getStringCellValue(),
                            item.getRowNum());
                    continue; // Break out and try next change
                }

            }
            saver.addUpdate(new UpdateRecord(iSht.getSheetName(),item.getRowNum(),""));
            ActionThread thr = new ActionThread(saver, config, changesSht, iSht, change, item, debugPrint);
            thr.start();
        }
        dpf("Done\n"); // Finish up the dotting.....
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
}