package com.planview.importer;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;

import com.planview.importer.Leankit.Board;
import com.planview.importer.Leankit.Card;
import com.planview.importer.Leankit.CardType;
import com.planview.importer.Leankit.CustomField;
import com.planview.importer.Leankit.Id;
import com.planview.importer.Leankit.Lane;
import com.planview.importer.Leankit.LeanKitAccess;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.json.JSONObject;

public class ActionThread extends Thread {

    Boolean debugPrint = false;
    Configuration config = new Configuration();
    XSSFSheet changesSht = null;
    XSSFSheet itemSht = null;
    Row change = null;
    Row item = null;

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

    private void dpf(String fmt, Object... parms) {
        if (debugPrint) {
            System.out.printf(fmt, parms);
        }
    }
    
    public ActionThread( Configuration cfg, XSSFSheet cSht, XSSFSheet iSht, Row chg, Row itm){
        changesSht = cSht;
        itemSht = iSht;
        change = chg;
        item = itm;
        config = cfg;
    }

    public void run(){
 
        Integer itemShtCol = findColumnFromSheet(changesSht, "Item Sheet");
        Integer rowCol = findColumnFromSheet(changesSht, "Item Row");
        Integer actionCol = findColumnFromSheet(changesSht, "Action");
        Integer fieldCol = findColumnFromSheet(changesSht, "Field");
        Integer value1Col = findColumnFromSheet(changesSht, "Value");

        // We need to find the ID of the board that this is targetting for a card
        // creation
        LeanKitAccess lka = new LeanKitAccess(config, debugPrint);
        String boardName = (item.getCell(findColumnFromSheet(itemSht, "Board Name")).getStringCellValue());
        Board brd = lka.fetchBoard(boardName);

        if (brd == null) {
            dpf("Cannot find board with name %s for item on sheet %s, row %d\n", boardName,
                    change.getCell(itemShtCol).getStringCellValue(),
                    (int) change.getCell(rowCol).getNumericCellValue());
            return;
        }
        String boardNumber = brd.id;
        /**
         * We need to get the header row for this sheet and work out which columns the
         * fields are in. It is possible that fields could be different between sheets,
         * so we have to do this every 'change'
         */

        Iterator<Row> iRow = itemSht.iterator();
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

        if (change.getCell(actionCol).getStringCellValue().equals("Create")) {
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

            /**
             * We need to either 'create' if ID == null && action == 'Create' or update if
             * ID != null && action == 'Modify'
             * 
             */

            Id card = createCard(lka, brd, flds); // Change from human readable to API fields on
                                                  // the way
            if (card == null) {
                dpf("Could not create card on board %s with details: %s", boardNumber, fieldLst.toString());
                System.exit(1);
            }
            return;

        } else if (change.getCell(actionCol).getStringCellValue().equals("Modify")) {
            // Fetch the ID from the item and then fetch that card
            Card card = lka.fetchCard(item.getCell(idCol).getStringCellValue());
            if (card == null) {
                dpf("Could not locate card \"%s\" on board \"%s\"\n", item.getCell(idCol).getStringCellValue(), boardNumber);
                return ;
            }
            JSONObject fld = new JSONObject();
            JSONObject vals = new JSONObject();

            vals.put("value1", convertCells(change, value1Col));

            fld.put(change.getCell(fieldCol).getStringCellValue(), vals);
            Id id = updateCard(lka, brd, card, fld);
            if (id == null) {
                dpf("Could not modify card on board %s with details:\"%s\"\n", boardNumber, fld.toString());
                System.exit(1);
            }
            return;
        }
        // Unknown option comes here
        return;
    }

    private Id updateCard(LeanKitAccess lka, Board brd, Card card, JSONObject fieldLst) {
        // Get available types so we can convert type string to typeid string
        CardType[] cTypes = brd.cardTypes;

        JSONObject finalUpdates = new JSONObject();
        if (fieldLst.has("Type")) {
            JSONObject fldVals = (JSONObject) fieldLst.get("Type");
            // Find type in cTypes and add new. If not present, then 'create' will default
            for (int i = 0; i < cTypes.length; i++) {
                if (cTypes[i].name.equals(fldVals.get("value1"))) {
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
                        if (cTypes[i].name.equals(fldVals.get("value1"))) {
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
                    if ( idx < 0) {
                        foundLane = findLaneFromString(brd, fldValues.get("value1").toString());
                    } else if (idx >0) {
                        foundLane = findLaneFromString(brd, fldValues.get("value1").toString().substring(0, idx));
                        woc = fldValues.get("value1").toString().substring(idx+1);
                    } else {
                        return null;
                    }
                    if (foundLane != null) {
                        if (foundLane.columns != 1) {   //Cannot move to a parent lane
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
                default:
                    // Make Sure field names from speadsheet are part of the Card model.... or....
                    Field[] fld = Card.class.getFields();
                    Boolean fieldFound = false;
                    for (int i = 0; i < fld.length; i++) {
                        if (fld[i].getName().equals(fldName)) {
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
                            if (cfld[i].label.equals(fldName)) {
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
                            dpf("Incorrect field name \"%s\" provided for update on card %s\n", fldName, card.id);
                            return null;
                        }
                    }
                    break;
            }

        }
        Id id = new Id();
        Card upc = lka.updateCardFromId(brd, card, finalUpdates);
        if (upc != null){
            id.id = upc.id;
            return id;
        }
        else {
            return null;
        }
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
                    break;
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
                    if (cl.name.equals(lanes[j])){
                        foundLanes.add(cl);
                    }
                }
                if ( foundLanes.size() > 0){
                    lanesToCheck = foundLanes;
                }
            }

        } while(true);

        if (searchLanes.size() == 0) {
            dpf("Cannot find lane \"%s\"on board \"%s\"\n", name, brd.title);
        }
        if (searchLanes.size() > 1) {
            dpf("Ambiguous lane name \"%s\"on board \"%s\"\n", name, brd.title);
        }

        return searchLanes.get(0);
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
}
