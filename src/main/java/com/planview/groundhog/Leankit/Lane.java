package com.planview.groundhog.Leankit;

public class Lane {
    public class Create {
    int cardLimit;
    String description;
    String id;
    int index;
    String laneClassType;
    String laneType;
    String orientation;
    String title;
    String taskBoard;
    }

    public class Read extends Create {
        Boolean active, isCollapsed, isDefaultDropLane, isConnectionDoneLane;
    }
}
