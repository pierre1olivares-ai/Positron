package com.timematters.qstar.model.mapper;

import com.timematters.qstar.api.model.SettingsATO;
import com.timematters.qstar.model.Settings;

public class SettingsMapper extends SimpleATOMapper<Settings, SettingsATO> {
    public SettingsMapper() {
        super(Settings.class, SettingsATO.class);
    }
}
