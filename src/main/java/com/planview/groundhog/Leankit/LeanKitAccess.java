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
    Board[] boards = null;

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
            if (jresp.has("error")) {
                System.out.printf("ERROR: \"%s\" gave response: \"%s\"", request.getRequestLine(), jresp.toString());
                System.exit(1);
            } else if (jresp.has("statusCode")) {
                System.out.printf("ERROR: \"%s\" gave response: \"%s\"", request.getRequestLine(), jresp.toString());
                System.exit(1);
            } else if (jresp.has("pageMeta")) {
                JSONObject pageMeta = new JSONObject(jresp.get("pageMeta").toString());

                int totalReturned = pageMeta.getInt("totalRecords");
                // Unfortunately, we need to know what sort of item to get out of the json
                // object. Doh!
                String fieldName = null;
                ArrayList<T> items = new ArrayList<T>();
                String[] typename = expectedResponseType.getName().split("\\.");
                switch (typename[typename.length - 1]) {
                    case "Board":
                        fieldName = "boards";
                        break;
                    case "User":
                        fieldName = "users";
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
                ArrayList<T> items = new ArrayList<T>();
                ObjectMapper om = new ObjectMapper();
                switch (expectedResponseType.getSimpleName()) {
                    case "CardType": {
                        // Getting CardTypes comes here, for example.
                        Iterator<String> sItor = jresp.keys();
                        String iStr = sItor.next();
                        JSONArray p = (JSONArray) jresp.get(iStr);
                        for (int i = 0; i < p.length(); i++) {
                            try {
                                items.add(om.readValue(p.get(i).toString(), expectedResponseType));
                            } catch (JsonProcessingException | JSONException e) {
                                e.printStackTrace();
                            }
                        }
                        break;
                    }
                    // Returning a single item from a search for example
                    case "Board": {
                        try {
                            JSONObject bdj = new JSONObject(bd);
                            // Cannot process one of these if we don't know what's going to be in there!!!!
                            if (bdj.has("userSettings")) {
                                bdj.remove("userSettings");
                            }
                            if (bdj.has("integrations")) {
                                bdj.remove("integrations");
                            }
                            items.add(om.readValue(bdj.toString(), expectedResponseType));
                        } catch (JsonProcessingException e) {
                            e.printStackTrace();
                        }
                        break;
                    }
                    default: {
                        System.out.println("oops! don't recognise requested item type");
                        System.exit(1);
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

    public Lane[] fetchLanes(String boardId) {
        request = new HttpGet(config.url + "io/board/" + boardId + "/");
        ArrayList<Board> brd = read(Board.class);
        if (brd.size() > 0) {
            return brd.get(0).lanes;
        }
        return null;
    }

    public ArrayList<Board> fetchBoardFromName(String name) {
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
        return read(Board.class);
    }

    public ArrayList<Board> fetchBoardFromId(String id) {
        request = new HttpGet(config.url + "io/board/" + id);
        return read(Board.class);
    }

    public String fetchBoardId(String name) {

        ArrayList<Board> brd = fetchBoardFromName(name);
        Board bd = null;
        if (brd.size() > 0) {
            // We found one or more with this name search. First try to find an exact match
            Iterator<Board> bItor = brd.iterator();
            while (bItor.hasNext()) {
                Board b = bItor.next();
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

    public String fetchUserId(String emailAddress) {
        request = new HttpGet(config.url + "io/user/");
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("search", emailAddress).build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }

        ArrayList<User> userd = read(User.class);

        User user = null;
        if (userd.size() > 0) {
            // We found one or more with this name search. First try to find an exact match
            Iterator<User> uItor = userd.iterator();
            while (uItor.hasNext()) {
                User u = uItor.next();
                if (u.emailAddress.equals(emailAddress)) {
                    user = u;
                }
            }
            // Then take the first if that fails
            if (user == null)
                user = userd.get(0);
            return user.id;
        }
        return null;
    }

    public Card fetchCard(String id) {
        request = new HttpGet(config.url + "/io/card/" + id);
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("returnFullRecord", "true").build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }
        return execute(Card.class);
    }

    public User fetchUser(String id) {
        request = new HttpGet(config.url + "/io/user/" + id);
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("returnFullRecord", "true").build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            System.out.println(e.getMessage());
            System.exit(1);
        }
        return execute(User.class);
    }

    private Boolean doCardMove(JSONObject cardMove) {
        request = new HttpPost(config.url + "/io/card/move");
        try {
            ((HttpPost) request).setEntity(new StringEntity(cardMove.toString()));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        String result = processRequest();

        // TODO: Need to fix this..... debug for now.
        if (result == null) {
            return false;
        } else {
            return true;
        }
    }

    private Boolean addCardParent(JSONObject cardChild) {
        request = new HttpPost(config.url + "/io/card/connections");
        try {
            ((HttpPost) request).setEntity(new StringEntity(cardChild.toString()));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        String result = processRequest();

        // TODO: Need to fix this..... debug for now.
        if (result == null) {
            return false;
        } else {
            return true;
        }
    }

    private Integer findTagIndex(Card card, String name){
        Integer index = -1;
        String[] names = card.tags;
        if (names != null) {
            for ( int i = 0; i < names.length; i++) {
                if (names[i] == name) {
                    index = i;
                }
            }
        }
        return index;
    }

    public Card updateCardFromId(String bNum, Card card, JSONObject updates) {

        // Create Leankit updates from the list
        JSONArray jsa = new JSONArray();
        Iterator<String> keys = updates.keys();
        while (keys.hasNext()) {
            String key = keys.next();
            switch (key) {
                case "Parent": {
                    if (updates.get(key).equals("0") || updates.get(key) == null || updates.get(key) == "") {
                        System.out.printf("Error trying to set parent of %s to value %s", card.id, updates.get(key));
                        break;
                    }
                    // Need to find the lane on the board and set the card to be in it.
                    JSONObject cardChild = new JSONObject();
                    JSONArray cardIds = new JSONArray();
                    cardIds.put(card.id);
                    cardChild.put("cardIds", cardIds);
                    JSONArray dParents = new JSONArray();
                    dParents.put(updates.get(key));
                    JSONObject connections = new JSONObject();
                    connections.put("parents", dParents);
                    cardChild.put("connections", connections);
                    addCardParent(cardChild);
                    break;
                }
                case "Lane": {
                    // Need to find the lane on the board and set the card to be in it.
                    JSONObject cardMove = new JSONObject();
                    JSONArray cardId = new JSONArray();
                    cardId.put(card.id);
                    cardMove.put("cardIds", cardId);
                    JSONObject dLane = new JSONObject();
                    dLane.put("laneId", updates.get("Lane"));
                    cardMove.put("destination", dLane);
                    doCardMove(cardMove);
                    break;
                }
                case "tags": {
                    // Need to add or remove based on what we already have?
                    // Or does add/remove ignore duplicate calls. Trying this first.....
                    if (updates.get(key).toString().startsWith("-")) {
                        Integer tIndex = findTagIndex(card, updates.get(key).toString().substring(1));
                        // If we found it, we can remove it
                        if (tIndex >= 0) {
                            JSONObject upd = new JSONObject();
                            upd.put("op", "remove");
                            upd.put("path", "/tags");
                            upd.put("value", tIndex);
                            jsa.put(upd);
                        }
                    } else {
                        JSONObject upd = new JSONObject();
                        upd.put("op", "add");
                        upd.put("path", "/tags/-");
                        upd.put("value", updates.get(key).toString());
                        jsa.put(upd);
                    }
                    break;
                }
                case "assignedUsers": {
                    // Need to add or remove based on what we already have?
                    // Or does add/remove ignore duplicate calls. Trying this first.....
                    if (updates.get(key).toString().startsWith("-")) {
                        JSONObject upd = new JSONObject();
                        upd.put("op", "remove");
                        upd.put("path", "/assignedUserIds");
                        upd.put("value", fetchUserId(updates.get(key).toString().substring(1)));
                        jsa.put(upd);
                    } else {
                        JSONObject upd = new JSONObject();
                        upd.put("op", "add");
                        upd.put("path", "/assignedUserIds/-");
                        upd.put("value", fetchUserId(updates.get(key).toString()));
                        jsa.put(upd);
                    }
                    break;
                }
                default: {
                    JSONObject upd = new JSONObject();
                    upd.put("op", "replace");
                    upd.put("path", "/" + key);
                    upd.put("value", updates.get(key));
                    jsa.put(upd);
                }
            }
        }
        request = new HttpPatch(config.url + "io/card/" + card.id);
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
            return execute(Card.class);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
            return null;
        }
    }

    public Card createCard(String boardId, JSONObject jItem) {

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
            return execute(Card.class);
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
