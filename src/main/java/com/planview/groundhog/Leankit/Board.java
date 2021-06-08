package com.planview.groundhog.Leankit;

import java.util.Date;

public class Board {
    public class Create {
        String title, description;
    }

    public class Template extends Create {
        String templateId;
        Boolean includeCards;
    }

    public class Copy extends Create {
        Boolean includeCards, includeExistingUsers, excludeCompletedAndArchiveViolations, baseWipOnCardSize;
    }

    public class Read {
        String id;
    }

    public class Long extends Read {
        String title, description, boardRole, version, subscriptionId, sharedBoardRole, effectiveBoardRole;
        int boardRoleId, currentExternalCardId, cardColorField;
        Boolean isWelcome, isArchived, isShared, isPermalinkEnabled, isExternalUrlEnabled, allowUsersToDeleteCards,
                allowPlanviewIntegration, classOfServiceEnabled, isCardIdEnabled, isHeaderEnabled, isHyperLinkEnabled,
                isPrefixEnabled;
        ClassOfService[] classesOfService;
        String[] tags;
        CustomField[] customFields;
        Date creationDate;
    }
}
