package com.planview.groundhog.Leankit;

import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Iterator;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.planview.groundhog.Configuration;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpDelete;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPatch;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.methods.HttpRequestBase;
import org.apache.http.client.methods.HttpUriRequest;
import org.apache.http.client.utils.URIBuilder;
import org.apache.http.entity.StringEntity;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.FileBody;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class LeanKitAccess {

    Configuration config = null;
    HttpUriRequest request = null;
    Board[] boards = null;

    Boolean debugPrint = false;

    public LeanKitAccess(Configuration cfg, Boolean dbp) {
        config = cfg;
        // Check URL has a trailing '/'
        if (!config.url.endsWith("/")) {
            config.url += "/";
        }
        debugPrint = dbp;
    }

    private void dpf(String fmt, Object... parms) {
        if (debugPrint) {
            System.out.printf(fmt, parms);
        }
    }

    public <T> ArrayList<T> read(Class<T> expectedResponseType) {

        request.setHeader("Accept", "application/json");
        request.setHeader("Content-type", "application/json");
        String bd = processRequest();
        JSONObject jresp = new JSONObject(bd);
        // Convert to a type to return to caller.
        if (bd != null) {
            if (jresp.has("error") || jresp.has("statusCode")) {
                dpf("ERROR: \"%s\" gave response: \"%s\"", request.getRequestLine(), jresp.toString());
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
                    case "Card":
                        fieldName = "cards";
                        break;
                    case "Comment":
                        fieldName = "comments";
                        break;
                    default:
                        dpf("Unsupported item type returned from server API\n");
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
                        dpf("%s", "oops! don't recognise requested item type");
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
        request.setHeader("Accept", "application/json");
        request.setHeader("Content-type", "application/json");

        String result = processRequest();
        if (result != null) {
            ObjectMapper om = new ObjectMapper();
            try {
                return om.readValue(result, expectedResponseType);
            } catch (JsonProcessingException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    private String processRequest() {

        // Deal with delays, retries and timeouts
        HttpClient client = HttpClients.createDefault();
        HttpResponse httpResponse = null;
        String result = null;
        try {
            // Add the user credentials to the request
            if (config.apikey != null) {
                request.addHeader("Authorization", "Bearer " + config.apikey);
            } else {
                String creds = config.username + ":" + config.password;
                request.addHeader("Authorization", "Basic " + Base64.getEncoder().encode(creds.getBytes()));
            }
            httpResponse = client.execute(request);
            switch (httpResponse.getStatusLine().getStatusCode()) {
                case 200: // Card updated
                case 201: // Card created
                {
                    result = EntityUtils.toString(httpResponse.getEntity());
                    break;
                }
                case 204: // No response expected
                {
                    break;
                }
                case 429: { // Flow control
                    Integer retryAfter = Integer.parseInt(httpResponse.getHeaders("retry-after")[0].getValue());
                    dpf("Received 429 status. waiting %.2f seconds\n", ((1.0 * retryAfter) / 1000.0));
                    try {
                        Thread.sleep(retryAfter);
                    } catch (InterruptedException e) {
                        dpf("%s", e.getMessage());
                    }
                    result = processRequest();
                    break;
                }
                case 422: { // Unprocessable Parameter
                    dpf("Parameter Error in request: %s\n", request.toString());
                    return null;
                }
                case 404: { // Item not found
                    dpf("Item not found: %s %s\n", httpResponse.getStatusLine().getStatusCode(),
                            httpResponse.getStatusLine().getReasonPhrase());
                    return null;
                }
                case 503: { // Service unavailable
                    dpf("Received 503 status. retrying in 5 seconds\n");
                    try {
                        Thread.sleep(5000);
                    } catch (InterruptedException e) {
                        dpf("%s", e.getMessage());
                    }
                    result = processRequest();
                    break;
                }
                default: {
                    dpf("Network fault: %s %s\n", httpResponse.getStatusLine().getStatusCode(),
                            httpResponse.getStatusLine().getReasonPhrase());
                    return null;
                }
            }
        } catch (IOException e) {
            dpf("%s", e.getMessage());
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

    public ArrayList<Board> fetchBoardsFromName(String name) {
        request = new HttpGet(config.url + "io/board/");
        URI uriA = null;
        try {
            uriA = new URIBuilder(request.getURI()).setParameter("search", name).build();
            ((HttpRequestBase) request).setURI(uriA);
        } catch (URISyntaxException e) {
            dpf("%s", e.getMessage());
            System.exit(1);
        }

        // Once you get the boards, you could cache them. There may be loads, but
        // shouldn't max
        // out memory.
        return read(Board.class);
    }

    public void deleteCards(ArrayList<Card> cards) {
        for (int i = 0; i < cards.size(); i++) {
            dpf("Deleting card %s\n", cards.get(i).id);
            request = new HttpDelete(config.url + "io/card/" + cards.get(i).id);
            processRequest();
        }
    }

    public ArrayList<Comment> fetchCommentsForCard(Card cd) {
        request = new HttpGet(config.url + "io/card/" + cd.id + "/comment");
        return read(Comment.class);
    }

    public ArrayList<Card> fetchCardIDsFromBoard(String id) {
        request = new HttpGet(config.url + "io/card?board=" + id + "&deleted=0&only=id");
        // Once you get the boards, you could cache them. There may be loads, but
        // shouldn't max
        // out memory.
        return read(Card.class);
    }

    public Board fetchBoardFromId(String id) {
        request = new HttpGet(config.url + "io/board/" + id);
        URI uriB = null;
        try {
            uriB = new URIBuilder(request.getURI()).setParameter("returnFullRecord", "true").build();
            ((HttpRequestBase) request).setURI(uriB);
        } catch (URISyntaxException e) {
            dpf("%s", e.getMessage());
            System.exit(1);
        }
        ArrayList<Board> results = read(Board.class);
        if (results != null) {
            return results.get(0);
        }
        return null;
    }

    public Board fetchBoard(String name) {

        ArrayList<Board> brd = fetchBoardsFromName(name);
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
            return fetchBoardFromId(bd.id);
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
            dpf("%s", e.getMessage());
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

    private String sendAttachment(String id, String filename) {
        request = new HttpPost(config.url + "/io/card/" + id + "/attachment");
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("returnFullRecord", "false").build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            dpf("%s", e.getMessage());
            System.exit(1);
        }
        File atchmt = new File(filename);
        FileBody fb = new FileBody(atchmt);
        MultipartEntityBuilder mpeb = MultipartEntityBuilder.create().setMode(HttpMultipartMode.BROWSER_COMPATIBLE)
                .addTextBody("Description", "Auto-generated from Script").addPart(filename, fb);
        HttpEntity ent = mpeb.build();
        ((HttpPost) request).setEntity(ent);
        String status = processRequest();
        return status;
    }

    private String postComment(String id, String comment) {
        request = new HttpPost(config.url + "/io/card/" + id + "/comment");
        JSONObject ent = new JSONObject();
        ent.put("text", comment);
        try {
            ((HttpPost) request).setEntity(new StringEntity(ent.toString()));
        } catch (UnsupportedEncodingException e) {
            return null;
        }
        Comment c = execute(Comment.class);
        return c.id;
    }

    public Card fetchCard(String id) {
        request = new HttpGet(config.url + "/io/card/" + id);
        URI uri = null;
        try {
            uri = new URIBuilder(request.getURI()).setParameter("returnFullRecord", "true").build();
            ((HttpRequestBase) request).setURI(uri);
        } catch (URISyntaxException e) {
            dpf("%s", e.getMessage());
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
            dpf("%s", e.getMessage());
            System.exit(1);
        }
        return execute(User.class);
    }

    private Integer findTagIndex(Card card, String name) {
        Integer index = -1;
        String[] names = card.tags;
        if (names != null) {
            for (int i = 0; i < names.length; i++) {
                if (names[i].equals(name)) {
                    index = i;
                }
            }
        }
        return index;
    }

    public Card updateCardFromId(Board brd, Card card, JSONObject updates) {

        // Create Leankit updates from the list
        JSONArray jsa = new JSONArray();
        Iterator<String> keys = updates.keys();
        while (keys.hasNext()) {
            String key = keys.next();
            JSONObject values = (JSONObject) updates.get(key);
            switch (key) {
                case "blockReason": {
                    if (values.get("value1").toString().length() <= 1) {

                        JSONObject upd = new JSONObject();
                        upd.put("op", "replace");
                        upd.put("path", "/isBlocked");
                        upd.put("value", false);
                        jsa.put(upd);

                    } else if (values.get("value1").toString().startsWith("-")) {
                        // Make it startsWith rather than equals just
                        // in case user forgets
                        JSONObject upd1 = new JSONObject();
                        upd1.put("op", "replace");
                        upd1.put("path", "/isBlocked");
                        upd1.put("value", false);
                        jsa.put(upd1);
                        JSONObject upd2 = new JSONObject();
                        upd2.put("op", "add");
                        upd2.put("path", "/blockReason");
                        upd2.put("value", values.get("value1").toString().substring(1));
                        jsa.put(upd2);
                    } else {
                        JSONObject upd1 = new JSONObject();
                        upd1.put("op", "replace");
                        upd1.put("path", "/isBlocked");
                        upd1.put("value", true);
                        jsa.put(upd1);
                        JSONObject upd2 = new JSONObject();
                        upd2.put("op", "add");
                        upd2.put("path", "/blockReason");
                        upd2.put("value", values.get("value1").toString());
                        jsa.put(upd2);
                    }
                    break;
                }
                case "Parent": {
                    if ((values.get("value1") == null) || (values.get("value1").toString() == "")
                            || (values.get("value1").toString() == "0")) {
                        dpf("Error trying to set parent of %s to value \"%s\"\n", card.id,
                                values.get("value1").toString());
                        break;
                    } else if (values.get("value1").toString().startsWith("-")) {
                        JSONObject upd2 = new JSONObject();
                        upd2.put("op", "remove");
                        upd2.put("path", "/parentCardId");
                        upd2.put("value", values.get("value1").toString().substring(1));
                        jsa.put(upd2);
                        break;
                    } else {
                        JSONObject upd2 = new JSONObject();
                        upd2.put("op", "add");
                        upd2.put("path", "/parentCardId");
                        upd2.put("value", values.get("value1").toString());
                        jsa.put(upd2);
                        break;
                    }
                }
                case "Lane": {
                    // Need to find the lane on the board and set the card to be in it.
                    JSONObject upd1 = new JSONObject();
                    upd1.put("op", "replace");
                    upd1.put("path", "/laneId");
                    upd1.put("value", values.get("value1").toString());
                    jsa.put(upd1);
                    if (values.has("value2")) {
                        JSONObject upd2 = new JSONObject();
                        upd2.put("op", "add");
                        upd2.put("path", "/wipOverrideComment");
                        upd2.put("value", values.get("value2").toString().trim());
                        jsa.put(upd2);
                    }
                    break;
                }
                case "tags": {
                    // Need to add or remove based on what we already have?
                    // Or does add/remove ignore duplicate calls. Trying this first.....
                    if (values.get("value1").toString().toString().startsWith("-")) {
                        Integer tIndex = findTagIndex(card, values.get("value1").toString().substring(1));
                        // If we found it, we can remove it
                        if (tIndex >= 0) {
                            JSONObject upd = new JSONObject();
                            upd.put("op", "remove");
                            upd.put("path", "/tags/" + tIndex);
                            jsa.put(upd);
                        }
                    } else {
                        JSONObject upd = new JSONObject();
                        upd.put("op", "add");
                        upd.put("path", "/tags/-");
                        upd.put("value", values.get("value1").toString());
                        jsa.put(upd);
                    }
                    break;
                }
                case "assignedUsers": {
                    // Need to add or remove based on what we already have?
                    // Or does add/remove ignore duplicate calls. Trying this first.....
                    if (values.get("value1").toString().startsWith("-")) {
                        JSONObject upd = new JSONObject();
                        upd.put("op", "remove");
                        upd.put("path", "/assignedUserIds");
                        upd.put("value", fetchUserId(values.get("value1").toString().substring(1)));
                        jsa.put(upd);
                    } else {
                        JSONObject upd = new JSONObject();
                        upd.put("op", "add");
                        upd.put("path", "/assignedUserIds/-");
                        upd.put("value", fetchUserId(values.get("value1").toString()));
                        jsa.put(upd);
                    }
                    break;
                }
                case "externalLink": {
                    JSONObject link = new JSONObject();
                    JSONObject upd = new JSONObject();
                    String[] bits = values.get("value1").toString().split(",");
                    if (bits.length != 2) {
                        dpf("Could not extract externalLink from %s (possible ',' in label?)",
                                values.get("value1").toString());
                        break;
                    }
                    link.put("label", bits[0]);
                    link.put("url", bits[1].trim());
                    upd.put("op", "replace");
                    upd.put("path", "/externalLink");
                    upd.put("value", link);
                    jsa.put(upd);
                    break;
                }
                case "customIcon": {
                    if (brd.classOfServiceEnabled) {
                        if (brd.classesOfService != null) {
                            for (int i = 0; i < brd.classesOfService.length; i++) {
                                if (brd.classesOfService[i].name.equals(values.get("value1"))) {
                                    JSONObject upd = new JSONObject();
                                    upd.put("op", "replace");
                                    upd.put("path", "/customIconId");
                                    upd.put("value", brd.classesOfService[i].id);
                                    jsa.put(upd);
                                    break;
                                }
                            }
                        }
                    }
                    break;
                }
                case "attachments": {
                    sendAttachment(card.id, values.get("value1").toString());
                    break;
                }

                case "comments": {
                    postComment(card.id, values.get("value1").toString());
                    break;
                }

                case "CustomField": {
                    CustomField[] cflds = brd.customFields;
                    if (cflds != null) {
                        for (int i = 0; i < cflds.length; i++) {
                            if (cflds[i].label.equals(values.get("value1"))) {
                                JSONObject upd = new JSONObject();
                                JSONObject val = new JSONObject();

                                val.put("fieldId", cflds[i].id);
                                val.put("value", values.get("value2"));

                                upd.put("op", "add");
                                upd.put("path", "/customFields/-");
                                upd.put("value", val);
                                jsa.put(upd);
                            }
                        }
                    }
                    break;
                }
                // Mismatch between UI and database in LK.
                case "priority": {
                    JSONObject upd = new JSONObject();
                    upd.put("op", "replace");
                    upd.put("path", "/" + key);
                    upd.put("value", values.get("value1").toString().toLowerCase());
                    jsa.put(upd);
                }
                default: {
                    JSONObject upd = new JSONObject();
                    upd.put("op", "replace");
                    upd.put("path", "/" + key);
                    upd.put("value", values.get("value1"));
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
            dpf("%s", e.getMessage());
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
            dpf("%s", e.getMessage());
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
            dpf("%s", e.getMessage());
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
