package com.planview.groundhog.Leankit;

import java.util.Date;

public class Card {
    public class Required {
        String boardId, // Required
        title; // Required
    }
    public class Create extends Required {
        
        String  typeId, description, laneId, mirrorSourceCardId, copiedFromCardId, blockReason, priority, customIconId,
                color, wipOverrideComment, customId;
        Connections connections;
        ExternalLink externalLink;
        String[] tags, assignedUserIds;
        int size, index;
        Date plannedStart, plannedFinish;
        CustomField[] customFields;
    }

    public class Read extends Create {
        CustomIcon customIcon;
        Date updatedOn, movedOn, createdOn, archivedOn;
        ConnectedCardStats connectedCardStats;
        ParentCard[] parentCards;
        AssignedUser.Long[] assignedUsers;
        AssignedUser.Short createdBy, updatedBy, movedBy, archivedBy;
        ItemType type;
        String version;

    }

}
