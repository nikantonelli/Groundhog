package com.planview.groundhog.Leankit;

import java.util.Date;

public class Comment {
    public class Short {
        String id, text;
        Date createdOn;
        AssignedUser.Short createdBy;
    }
    public class Long {
        String cardId;
    }
}
