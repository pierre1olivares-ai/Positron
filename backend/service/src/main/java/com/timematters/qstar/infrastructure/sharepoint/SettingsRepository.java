package com.timematters.qstar.infrastructure.sharepoint;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.timematters.qstar.model.Settings;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

/**
 * Reads/writes the IT-settings tab's configuration from the single-item "Q-Star Config" list
 * (column SettingsJson holds the whole object as JSON) — the same list and shape
 * frontend/src/webparts/qstarIssueManager/services/SharePointDataService.ts uses when running
 * without this backend.
 */
@Repository
public class SettingsRepository {

    private static final String CONFIG_LIST_NAME = "Q-Star Config";
    private static final String SETTINGS_JSON_FIELD = "SettingsJson";

    private final SharePointGraphClient graphClient;
    private final ObjectMapper objectMapper = new ObjectMapper();

    @Autowired
    public SettingsRepository(SharePointGraphClient graphClient) {
        this.graphClient = graphClient;
    }

    public Settings load() {
        String listId = graphClient.getListId(CONFIG_LIST_NAME);
        List<Map<String, Object>> items = graphClient.getAllItems(listId);
        if (items.isEmpty()) {
            return new Settings();
        }
        Object json = items.get(0).get(SETTINGS_JSON_FIELD);
        if (json == null) {
            return new Settings();
        }
        try {
            return objectMapper.readValue(json.toString(), Settings.class);
        } catch (Exception e) {
            return new Settings();
        }
    }

    public Settings save(Settings settings) {
        String listId = graphClient.getListId(CONFIG_LIST_NAME);
        List<Map<String, Object>> items = graphClient.getAllItems(listId);
        String json;
        try {
            json = objectMapper.writeValueAsString(settings);
        } catch (Exception e) {
            throw new IllegalStateException("Could not serialize settings.", e);
        }
        Map<String, Object> fields = new HashMap<>();
        fields.put(SETTINGS_JSON_FIELD, json);
        if (items.isEmpty()) {
            graphClient.createItem(listId, fields);
        } else {
            graphClient.updateItemFields(listId, items.get(0).get("Id").toString(), fields);
        }
        return settings;
    }
}
