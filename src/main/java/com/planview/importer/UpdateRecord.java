package com.planview.importer;

public class UpdateRecord {
    public UpdateRecord(String str, Integer rn, String upd){
        shtName = str;
        rowNum = rn;
        update = upd;
    }
    public String shtName;
    public Integer rowNum;
    public String update;
}
