package com.timematters.qstar.model.mapper;

import com.remondis.remap.Mapper;
import com.remondis.remap.Mapping;
import com.timematters.qstar.api.model.SettingsATO;
import com.timematters.qstar.model.Settings;

public class SettingsMapper implements ATOMapper<Settings, SettingsATO> {

    private final AccessEntryMapper accessEntryMapper;

    public SettingsMapper() {
        this.accessEntryMapper = new AccessEntryMapper();
    }

    @Override
    public Mapper<Settings, SettingsATO> getMapperToATO() {
        return Mapping.from(Settings.class)
                .to(SettingsATO.class)
                .useMapper(accessEntryMapper.getMapperToATO())
                .mapper();
    }

    @Override
    public Mapper<SettingsATO, Settings> getMapperFromATO() {
        return Mapping.from(SettingsATO.class)
                .to(Settings.class)
                .useMapper(accessEntryMapper.getMapperFromATO())
                .mapper();
    }

    @Override
    public Settings fromATO(SettingsATO ato) {
        return getMapperFromATO().map(ato);
    }

    @Override
    public SettingsATO toATO(Settings model) {
        return getMapperToATO().map(model);
    }
}
