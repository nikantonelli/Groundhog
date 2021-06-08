package com.planview.groundhog.Leankit;

import java.util.Date;

public class BoardLongRead {
    public String id;
    public String title, description, boardRole, version, subscriptionId, sharedBoardRole, effectiveBoardRole;
    public int boardRoleId, currentExternalCardId, cardColorField;
    public Boolean isWelcome, isArchived, isShared, isPermalinkEnabled, isExternalUrlEnabled, allowUsersToDeleteCards,
            allowPlanviewIntegration, classOfServiceEnabled, isCardIdEnabled, isHeaderEnabled, isHyperLinkEnabled,
            isPrefixEnabled;
    public ClassOfService[] classesOfService;
    public String[] tags;
    public CustomField[] customFields;
    public Date creationDate;
    public Level level;
}