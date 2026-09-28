package com.timematters.qstar.infrastructure.sharepoint;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.timematters.qstar.model.Settings;
import java.util.Map;
import org.springframework.stereotype.Repository;

@Repository
public class SettingsRepository {
    private final SharePointRestClient client;
    private final SharePointProperties properties;
    private final ObjectMapper json = new ObjectMapper();

    public SettingsRepository(SharePointRestClient client, SharePointProperties properties) {
        this.client = client;
        this.properties = properties;
    }

    private Map<String, Object> config() {
        var items = client.getItems(properties.getConfigListName(), "$select=Id,SettingsJson");
        if (items.size() != 1)
            throw new IllegalStateException("Exactly one provisioned Config item is required.");
        return items.getFirst();
    }

    public Settings load() {
        Object stored = config().get("SettingsJson");
        if (stored == null || stored.toString().isBlank()) return new Settings();
        try {
            return json.readValue(stored.toString(), Settings.class);
        } catch (Exception e) {
            throw new IllegalStateException("Stored settings could not be read.", e);
        }
    }

    public Settings save(Settings settings) {
        var item = config();
        long id = IssueRepository.number(item.get("Id"));
        String eTag = IssueRepository.eTag(item);
        if (eTag == null)
            eTag =
                    IssueRepository.eTag(
                            client.getItem(properties.getConfigListName(), id, "$select=Id"));
        try {
            // Scalar ReferenceOffset is immutable and is never part of this update.
            client.updateItem(
                    properties.getConfigListName(),
                    id,
                    Map.of("SettingsJson", json.writeValueAsString(settings)),
                    eTag);
        } catch (com.fasterxml.jackson.core.JsonProcessingException e) {
            throw new IllegalArgumentException("Invalid settings.", e);
        }
        return settings;
    }
}
