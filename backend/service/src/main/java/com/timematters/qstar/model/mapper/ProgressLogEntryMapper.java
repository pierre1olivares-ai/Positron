package com.timematters.qstar.model.mapper;

import com.timematters.qstar.api.model.ProgressLogEntryATO;
import com.timematters.qstar.model.ProgressLogEntry;

public class ProgressLogEntryMapper extends SimpleATOMapper<ProgressLogEntry, ProgressLogEntryATO> {
    public ProgressLogEntryMapper() {
        super(ProgressLogEntry.class, ProgressLogEntryATO.class);
    }
}
