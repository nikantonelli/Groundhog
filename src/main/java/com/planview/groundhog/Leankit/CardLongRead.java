package com.planview.groundhog.Leankit;

import java.util.Date;

public class CardLongRead {
    public String id, title, typeId, description, laneId, mirrorSourceCardId, copiedFromCardId, 
            blockReason, priority, subscriptionId, version,
            customIconId, customIconLabel, iconPath, color, wipOverrideComment;
    public CustomId customId;
    public Connections connections;
    public ExternalLink externalLink;
    public String[] tags, assignedUserIds;
    public Integer commentsCount, childCommentsCount, size, index;
    public Date updatedOn, movedOn, createdOn, archivedOn, plannedStart, plannedFinish, actualStart, actualFinish;
    public CustomField[] customFields;
    public CustomIcon customIcon;
    public ConnectedCardStats connectedCardStats;
    public ParentCard[] parentCards;
    public AssignedUser[] assignedUsers;
    public AssignedUser createdBy, updatedBy, movedBy, archivedBy;
    public ItemType type;
    public BlockedStatus blockedStatus;
    public BoardCard board;
    public Lane lane;
    public Boolean isBlocked;
    public ExternalLink[] externalLinks;
    public TaskBoardStats taskBoardStats;
    public Attachment[] attachments;
    public Comment[] comments;
}
