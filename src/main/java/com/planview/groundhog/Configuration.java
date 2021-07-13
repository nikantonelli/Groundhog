package com.planview.groundhog;

import java.util.Base64;

//All fieldnames must be lowercase alphabetical
public class Configuration {
    public String url;  //Must be first in this object.
    public String username;
    public String password;
    public String apikey;
    public Double cyclelength;  //Excel numeric fields are doubles
    public String hash() {
        return Base64.getEncoder().encodeToString(
            (url+username+password+apikey+cyclelength.toString()).getBytes()
            );
    }
}
