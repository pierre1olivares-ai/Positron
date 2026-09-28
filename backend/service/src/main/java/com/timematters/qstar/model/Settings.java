package com.timematters.qstar.model;

import java.util.ArrayList;
import java.util.List;
import lombok.Data;

@Data
public class Settings {
    private String msFormUrl = "";
    private String flowId = "";
    private String spSiteUrl = "";
    private String spListName = "";
    private String tenantId = "";
    private String clientId = "";
    private Boolean connected = false;
    private String lastTested = "";
    private List<AccessEntry> access = new ArrayList<>();
}
