package com.planview.groundhog.Leankit;

import java.util.Date;

public class User {
    public String id, username, firstName, lastName, fullName, emailAddress, organizationId, dateFormat, timeZone, licenseType, externalUserName, avatar;
    public Boolean administrator, accountOwner, enabled, deleted, boardCreator;
    public BoardCount boardCount;
    public Date createdOn, lastAccess;
}
