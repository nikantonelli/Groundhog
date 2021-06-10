package com.planview.groundhog.Leankit;

import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.Iterator;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.jsonFormatVisitors.JsonObjectFormatVisitor;
import com.planview.groundhog.Configuration;

import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPatch;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.methods.HttpRequestBase;
import org.apache.http.client.methods.HttpUriRequest;
import org.apache.http.client.utils.URIBuilder;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.apache.xml.security.utils.Base64;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class LeanKitAccess {

    Configuration config = null;
    HttpUriRequest request = null;
    BoardLongRead[] boards = null;

    public LeanKitAccess(Configuration cfg) {
        config = cfg;
        // Check URL has a trailing '/'
        if (!config.url.endsWith("/")) {
            config.url += "/";
        }
    }

    public <T> ArrayList<T> read(Class<T> expectedResponseType) {

        String bd = processRequest();
        JSONObject jresp = new JSONObject(bd);
        // Convert to a type to return to caller.
        if (bd != null) {
            if (jresp.has("pageMeta")) {
                JSONObject pageMeta = new JSONObject(jresp.get("pageMeta").toString());

                int totalReturned = pageMeta.getInt("totalRecords");
                // Unfortunately, we need to know what sort of item to get out of the json
                // object. Doh!
                String fieldName = null;
                ArrayList<T> items = new ArrayList<T>();
                String[] typename = expectedResponseType.getName().split("\\.");
                switch (typename[typename.length - 1]) {
                    case "BoardLongRead":
                        fieldName = "boards";
                        break;
                    default:
                        System.out.println("Incorrect item type returned from server API");
                }
                if (fieldName != null) {
                    // Got something to return
                    ObjectMapper om = new ObjectMapper();
                    om.configure(DeserializationFeature.USE_JAVA_ARRAY_FOR_JSON_ARRAY, true);
                    JSONArray p = (JSONArray) jresp.get(fieldName);
                    for (int i = 0; i < totalReturned; i++) {
                        try {
                            items.add(om.readValue(p.get(i).toString(), expectedResponseType));
                        } catch (JsonProcessingException | JSONException e) {
                            e.printStackTrace();
                        }
                    }
                    return items;
                }

            } else {
                // Getting CardTypes comes here, for example.
                Iterator<String> sItor = jresp.keys();
                String iStr = sItor.next();
                ArrayList<T> items = new ArrayList<T>();
                ObjectMapper om = new ObjectMapper();
                JSONArray p = (JSONArray) jresp.get(iStr);
                for (int i = 0; i < p.length(); i++) {
                    try {
                        items.add(om.readValue(p.get(i).toString(), expectedResponseType));
                    } catch (JsonProcessingException | JSONException e) {
                        e.printStackTrace();
                    }
                }
                return items;
            }
        }
        return null;

    }

    /**
     * 
     * @param <T>
     * @param expectedResponseType
     * @return string value of Id
     * 
     *         Create something and return just the id to it.
     */
    public <T> T execute(Class<T> expectedResponseType) {
        String result = processRequest();
        ObjectMapper om = new ObjectMapper();
        try {
            return om.readValue(result, expectedResponseType);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }
        return null;
    }

    private String processRequest() {

        // Deal with delays, retries and timeouts
        HttpClient client = HttpClients.createDefault();
        HttpResponse httpResponse = null;
        String result = null;
        request.setHeader("Accept", "application/json");
        request.setHeader("Content-type", "application/json");
        try {
            // Add the user credentials to the request
            if (config.apikey != null) {
                request.addHeader("Authorization", "Bearer " + config.apikey);
            } else {
                String creds = config.username + ":" + config.password;
                request.addHeader("Authorization", "Basic " + Base64.encode(creds.getBytes()));
            }
            httpResponse = client.execute(request); // Get a json string back
            result = EntityUtils.toString(httpResponse.getEntity());
        } catch (IOException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }
        return result;
    }

    public ArrayList<CardType> fetchCardTypes(String boardId) {
        request = new HttpGet(config.url + "io/board/" + boardId + "/cardType");
        ArrayList<CardType> brd = read(CardType.class);
        if (brd.size() > 0) {
            return brd;
        }
        return null;
    }

    public String fetchBoardId(String name) {
        request = new HttpGet(config.url + "io/board/");
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("search", name).build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }

        // Once you get the boards, you could cache them. There may be loads, but
        // shouldn't max
        // out memory.
        ArrayList<BoardLongRead> brd = read(BoardLongRead.class);

        BoardLongRead bd = null;
        if (brd.size() > 0) {
            // We found one or more with this name search. First try to find an exact match
            Iterator<BoardLongRead> bItor = brd.iterator();
            while (bItor.hasNext()) {
                BoardLongRead b = bItor.next();
                if (b.title.equals(name)) {
                    bd = b;
                }
            }
            // Then take the first if that fails
            if (bd == null)
                bd = brd.get(0);
            return bd.id;
        }
        return null;
    }

    
    public CardLongRead fetchCard(String id) {
        request = new HttpGet(config.url + "/io/card/"+id);
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("returnFullRecord", "true").build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }
        return execute(CardLongRead.class);
    }

    public CardLongRead updateCardFromId(String id, JSONObject updates){

        //Create Leankit updates from the list
        JSONArray jsa = new JSONArray();
        Iterator<String> keys = updates.keys();
        while(keys.hasNext()) {
            String key = keys.next();
            JSONObject upd = new JSONObject();
            upd.put("op", "replace");
            upd.put("path", "/"+key);
            upd.put("value", updates.get(key));
            jsa.put(upd);
        }
        request = new HttpPatch(config.url + "io/card/"+id);
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("returnFullRecord", "false").build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }
        try {
            ((HttpPatch) request).setEntity(new StringEntity(jsa.toString()));
            return execute(CardLongRead.class);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
            return null;
        }
    }

    public CardLongRead createCard(String boardId, JSONObject jItem) {

        request = new HttpPost(config.url + "io/card/");
        jItem.put("boardId", boardId);
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("returnFullRecord", "true").build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }

        if (!jItem.has("title")) {
            jItem.put("title", "dummy title"); // Used when we are testing a create to get back the card structure
        }
        try {
            ((HttpPost) request).setEntity(new StringEntity(jItem.toString()));
            return execute(CardLongRead.class);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
            return null;
        }
    }

    public Id createCardID(String boardId) {
        JSONObject jItem = new JSONObject();
        request = new HttpPost(config.url + "io/card/");
        jItem.put("boardId", boardId);
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("returnFullRecord", "false").build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }
        try {
            ((HttpPost) request).setEntity(new StringEntity(jItem.toString()));
            return execute(Id.class);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
            return null;
        }

    }
}
