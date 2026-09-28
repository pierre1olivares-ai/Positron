package com.timematters.qstar.model.mapper;

import com.timematters.qstar.api.model.AccessEntryATO;
import com.timematters.qstar.model.AccessEntry;

public class AccessEntryMapper extends SimpleATOMapper<AccessEntry, AccessEntryATO> {
    public AccessEntryMapper() {
        super(AccessEntry.class, AccessEntryATO.class);
    }
}
