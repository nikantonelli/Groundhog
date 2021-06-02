package com.planview.groundhog;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.bcel.generic.FNEG;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GroundHog {
    public static void main(String[] args) {

        Options opts = new Options();
        Option xlsxname = new Option("f", "filename", true, "source XLSX spreadsheet");
        xlsxname.setRequired(true);
        opts.addOption(xlsxname);

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

        String fn = cl.getOptionValue("filename");

        // Check we can open the file
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(new File(fn));
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage());
            hf.printHelp(" ", opts);
            System.exit(1);

        }

        XSSFWorkbook wb = null;
        try {
            wb = new XSSFWorkbook(fis);
        } catch (IOException e) {
            System.out.println(e.getMessage());
            hf.printHelp(" ", opts);
            System.exit(1);
        }
       
        //These two come first in the file. 
        XSSFSheet configSht = wb.getSheet("Config");
        XSSFSheet changesSht = wb.getSheet("Changes");

        Integer shtCount = wb.getNumberOfSheets();
        if (shtCount < 3) {
            System.out.println("Did not detect enough sheets in the spreadsheet");
            System.exit(1);
        }

        ArrayList<XSSFSheet> shtArray = new ArrayList<XSSFSheet>();
        Iterator<Sheet> si = wb.iterator();
        
        //Then all the level info comes next
        while (si.hasNext()) {
            XSSFSheet n = (XSSFSheet) si.next();
            if (!((n.getSheetName().equals("Config")) || (n.getSheetName().equals("Changes")))) {
                shtArray.add(n);
            }

        }



        
        System.out.println(fn);

    }
}